#include <XLPP/Workbook/Workbook.h>

#include "Formula/CalculationValue.h"
#include "Formula/FormulaEvaluationContext.h"
#include "Formula/FormulaFunctions.h"

#include <algorithm>
#include <cctype>
#include <chrono>
#include <cmath>
#include <cstdlib>
#include <iomanip>
#include <limits>
#include <optional>
#include <sstream>
#include <stdexcept>
#include <unordered_map>
#include <unordered_set>

namespace xlpp {
namespace {

using internal::formula::Scalar;
using internal::formula::EvalValue;
using internal::formula::isError;
using internal::formula::errorOr;
using internal::formula::upperAscii;
using internal::formula::trimAscii;
using internal::formula::parseNumberText;
using internal::formula::numberValue;
using internal::formula::boolValue;
using internal::formula::scalarText;
using internal::formula::firstScalar;

bool sameScalar(const Scalar& lhs, const Scalar& rhs) {
    if (lhs.index() != rhs.index()) return false;
    return lhs == rhs;
}

struct FormulaKey {
    Worksheet* sheet{};
    std::size_t row{};
    std::size_t column{};
    bool operator==(const FormulaKey&) const noexcept = default;
};
struct FormulaKeyHash {
    std::size_t operator()(const FormulaKey& key) const noexcept {
        return std::hash<void*>{}(key.sheet) ^ (key.row * 1315423911u) ^ (key.column * 2654435761u);
    }
};

class CalculationEngine : public internal::formula::FormulaEvaluationContext {
public:
    CalculationEngine(Workbook& workbook, const CalculationOptions& options)
        : workbook_(workbook), options_(options) {}

    CalculationReport run() {
        std::vector<FormulaKey> formulas;
        for (auto& sheet : workbook_.worksheets()) {
            for (const auto& [_, cell] : sheet.cells()) {
                if (cell.hasFormula()) formulas.push_back({&sheet, cell.row(), cell.column()});
            }
        }
        report_.formulaCellsVisited = formulas.size();
        if (!options_.changedCells.empty() && !formulas.empty()) {
            const auto graph = buildFormulaDependencyGraph(workbook_);
            std::vector<CalculationCell> queue = options_.changedCells;
            std::unordered_set<std::string> visitedRoots;
            std::unordered_set<std::string> selected;
            auto keyText = [](const std::string& sheet, const std::string& cell) {
                return upperAscii(sheet) + "\x1f" + upperAscii(cell);
            };
            for (std::size_t i = 0; i < queue.size(); ++i) {
                const auto current = queue[i];
                const auto rootKey = keyText(current.sheet, current.cell);
                if (!visitedRoots.insert(rootKey).second) continue;
                for (const auto& edge : graph.dependentsOf(current.sheet, current.cell)) {
                    const auto dependentKey = keyText(edge.dependentSheet, edge.dependentCell);
                    if (selected.insert(dependentKey).second)
                        queue.push_back({edge.dependentSheet, edge.dependentCell});
                }
                if (auto* sheet = workbook_.worksheet(current.sheet)) {
                    try {
                        const auto ref = CellReference::parse(current.cell);
                        const auto* cell = sheet->tryCell(ref.row, ref.column);
                        if (cell && cell->hasFormula()) selected.insert(rootKey);
                    } catch (...) {}
                }
            }
            report_.dirtyRoots = options_.changedCells.size();
            report_.dirtyFormulaCellsSelected = selected.size();
            formulas.erase(std::remove_if(formulas.begin(), formulas.end(), [&](const FormulaKey& key) {
                return selected.find(keyText(key.sheet->name(), CellReference(key.row, key.column).address())) == selected.end();
            }), formulas.end());
        }
        if (!options_.iterativeCalculation || formulas.empty()) {
            for (const auto& key : formulas) (void)evaluateCell(key, 0);
            return report_;
        }

        if (options_.maxIterations == 0) {
            ++report_.iterativeConvergenceFailures;
            report_.warnings.push_back("Iterative calculation requires maxIterations > 0");
            return report_;
        }

        previousIteration_.clear();
        for (const auto& key : formulas) {
            const auto* cell = key.sheet->tryCell(key.row, key.column);
            previousIteration_[key] = cell ? cell->value() : Scalar{std::monostate{}};
        }

        bool converged = false;
        iterativePass_ = true;
        for (std::size_t iteration = 1; iteration <= options_.maxIterations; ++iteration) {
            cache_.clear();
            visiting_.clear();
            for (const auto& key : formulas) (void)evaluateCell(key, 0);

            bool stable = true;
            double largestChange = 0.0;
            for (const auto& key : formulas) {
                const auto next = cache_.count(key) ? cache_.at(key) : Scalar{std::monostate{}};
                const auto prior = previousIteration_.count(key) ? previousIteration_.at(key) : Scalar{std::monostate{}};
                if (sameScalar(next, prior)) continue;
                const auto nextNumber = numberValue(next, workbook_.date1904());
                const auto priorNumber = numberValue(prior, workbook_.date1904());
                if (nextNumber && priorNumber) {
                    largestChange = std::max(largestChange, std::fabs(*nextNumber - *priorNumber));
                    if (std::fabs(*nextNumber - *priorNumber) > options_.maxChange) stable = false;
                } else {
                    stable = false;
                }
            }
            previousIteration_ = cache_;
            report_.iterativeIterations = iteration;
            if (stable && largestChange <= options_.maxChange) { converged = true; break; }
        }
        iterativePass_ = false;

        if (!converged) {
            ++report_.iterativeConvergenceFailures;
            report_.warnings.push_back("Iterative calculation did not converge within maxIterations");
        }

        // One final pass consumes the converged prior-iteration values for
        // circular back-edges while applying normal cached values/spills.
        cache_.clear();
        visiting_.clear();
        for (const auto& key : formulas) (void)evaluateCell(key, 0);
        return report_;
    }

    EvalValue resolveReference(Worksheet& context, const std::string& sheetName,
                               const std::string& firstText, const std::optional<std::string>& lastText,
                               std::size_t depth) override {
        Worksheet* target = &context;
        if (!sheetName.empty()) {
            target = workbook_.worksheet(sheetName);
            if (!target) return EvalValue::fromScalar(CellError::Reference);
        }
        CellReference first;
        CellReference last;
        try {
            first = CellReference::parse(firstText);
            last = lastText ? CellReference::parse(*lastText) : first;
        } catch (...) {
            return EvalValue::fromScalar(CellError::Reference);
        }
        if (first.row > last.row) std::swap(first.row, last.row);
        if (first.column > last.column) std::swap(first.column, last.column);
        std::vector<Scalar> values;
        const auto count = (last.row - first.row + 1) * (last.column - first.column + 1);
        if (count > 1'000'000) return EvalValue::fromScalar(CellError::Number);
        values.reserve(count);
        for (std::size_t row = first.row; row <= last.row; ++row) {
            for (std::size_t column = first.column; column <= last.column; ++column) {
                const auto* cell = target->tryCell(row, column);
                if (!cell) { values.push_back(std::monostate{}); continue; }
                if (cell->hasFormula() && options_.recursiveDependencies) {
                    ++report_.dependencyEvaluations;
                    const FormulaKey dependency{target, row, column};
                    if (options_.iterativeCalculation && visiting_.find(dependency) != visiting_.end()) {
                        const auto previous = previousIteration_.find(dependency);
                        values.push_back(previous != previousIteration_.end() ? previous->second : cell->value());
                    } else {
                        values.push_back(evaluateCell(dependency, depth + 1));
                    }
                } else values.push_back(cell->value());
            }
        }
        if (values.size() == 1) {
            auto out = EvalValue::fromScalar(values.front());
            return std::move(out.withReference(target, first.row, first.column));
        }
        auto out = EvalValue::fromRange(std::move(values), last.row - first.row + 1, last.column - first.column + 1);
        return std::move(out.withReference(target, first.row, first.column));
    }

    EvalValue resolveExternalReference(const std::string& workbookToken, const std::string& sheetName,
                                       const std::string& firstText, const std::optional<std::string>& lastText) override {
        if (!options_.externalReferenceResolver) {
            ++report_.unresolvedExternalReferences;
            report_.warnings.push_back("External workbook reference requires externalReferenceResolver: [" + workbookToken + "]" + sheetName + "!" + firstText);
            return EvalValue::fromScalar(CellError::Reference);
        }
        CellReference first, last;
        try { first = CellReference::parse(firstText); last = lastText ? CellReference::parse(*lastText) : first; }
        catch (...) { ++report_.unresolvedExternalReferences; return EvalValue::fromScalar(CellError::Reference); }
        if (first.row > last.row) std::swap(first.row, last.row);
        if (first.column > last.column) std::swap(first.column, last.column);
        const auto count = (last.row - first.row + 1) * (last.column - first.column + 1);
        if (count > 1'000'000) return EvalValue::fromScalar(CellError::Number);
        std::vector<Scalar> values; values.reserve(count);
        for (std::size_t row = first.row; row <= last.row; ++row) {
            for (std::size_t column = first.column; column <= last.column; ++column) {
                const auto address = CellReference(row, column).address();
                auto value = options_.externalReferenceResolver(workbookToken, sheetName, address);
                if (!value) {
                    ++report_.unresolvedExternalReferences;
                    report_.warnings.push_back("External workbook resolver returned no value for [" + workbookToken + "]" + sheetName + "!" + address);
                    values.push_back(CellError::Reference);
                } else {
                    ++report_.externalReferencesResolved;
                    values.push_back(std::move(*value));
                }
            }
        }
        if (values.size() == 1) return EvalValue::fromScalar(values.front());
        return EvalValue::fromRange(std::move(values), last.row - first.row + 1, last.column - first.column + 1);
    }

    EvalValue resolveDefinedName(Worksheet& context, const std::string& name, std::size_t depth) override {
        const DefinedName* chosen = nullptr;
        const auto contextIndex = workbook_.index(context);
        for (const auto& item : workbook_.definedNames()) {
            if (upperAscii(item.name()) != upperAscii(name)) continue;
            if (item.localSheetId() && *item.localSheetId() == contextIndex) { chosen = &item; break; }
            if (!item.localSheetId() && !chosen) chosen = &item;
        }
        if (!chosen) return EvalValue::fromScalar(CellError::Name);
        ++report_.definedNamesResolved;
        try {
            return internal::formula::parseFormula(*this, context, chosen->value(), depth + 1);
        } catch (...) {
            return EvalValue::fromScalar(CellError::Value);
        }
    }

    EvalValue resolveStructuredReference(Worksheet& context, const std::string& expression, std::size_t depth) override {
        const auto open = expression.find('[');
        if (open == std::string::npos || expression.back() != ']')
            return EvalValue::fromScalar(CellError::Reference);
        const auto tableName = expression.substr(0, open);
        const auto selector = expression.substr(open);
        Worksheet* tableSheet = nullptr;
        const Table* table = nullptr;
        if (!tableName.empty()) {
            for (auto& candidateSheet : workbook_.worksheets()) {
                for (const auto& candidate : static_cast<const Worksheet&>(candidateSheet).tables()) {
                    if (upperAscii(candidate.name()) == upperAscii(tableName) ||
                        upperAscii(candidate.displayName()) == upperAscii(tableName)) {
                        tableSheet = &candidateSheet; table = &candidate; break;
                    }
                }
                if (table) break;
            }
        } else if (const auto current = currentFormulaKey(); current && current->sheet == &context) {
            for (const auto& candidate : static_cast<const Worksheet&>(context).tables()) {
                try {
                    const auto colon = candidate.reference().find(':');
                    auto a = CellReference::parse(colon == std::string::npos ? candidate.reference() : candidate.reference().substr(0, colon));
                    auto b = CellReference::parse(colon == std::string::npos ? candidate.reference() : candidate.reference().substr(colon + 1));
                    const auto r1 = std::min(a.row, b.row), r2 = std::max(a.row, b.row);
                    const auto c1 = std::min(a.column, b.column), c2 = std::max(a.column, b.column);
                    if (current->row >= r1 && current->row <= r2 && current->column >= c1 && current->column <= c2) {
                        tableSheet = &context; table = &candidate; break;
                    }
                } catch (...) {}
            }
        }
        if (!table || !tableSheet) return EvalValue::fromScalar(CellError::Name);

        const auto colon = table->reference().find(':');
        try {
            auto first = CellReference::parse(colon == std::string::npos ? table->reference() : table->reference().substr(0, colon));
            auto last = CellReference::parse(colon == std::string::npos ? table->reference() : table->reference().substr(colon + 1));
            if (first.row > last.row) std::swap(first.row, last.row);
            if (first.column > last.column) std::swap(first.column, last.column);

            std::size_t rowFirst = first.row + (table->showHeaderRow() ? 1 : 0);
            std::size_t rowLast = last.row - (table->showTotalsRow() && last.row > first.row ? 1 : 0);
            const auto selectorUpper = upperAscii(selector);
            const bool currentRowSelector = selectorUpper.find("#THIS ROW") != std::string::npos || selectorUpper.find("@") != std::string::npos;
            if (selectorUpper.find("#ALL") != std::string::npos) { rowFirst = first.row; rowLast = last.row; }
            else if (selectorUpper.find("#HEADERS") != std::string::npos) { rowFirst = first.row; rowLast = first.row; }
            else if (selectorUpper.find("#TOTALS") != std::string::npos) {
                if (!table->showTotalsRow()) return EvalValue::fromScalar(CellError::Reference);
                rowFirst = rowLast = last.row;
            } else if (currentRowSelector) {
                const auto current = currentFormulaKey();
                if (!current || current->sheet != tableSheet || current->row < rowFirst || current->row > rowLast)
                    return EvalValue::fromScalar(CellError::Value);
                rowFirst = rowLast = current->row;
            }
            if (rowFirst > rowLast) return EvalValue::fromRange({}, 0, 0);

            std::size_t colFirst = first.column, colLast = last.column;
            std::vector<std::size_t> selectedColumns;
            for (std::size_t i = 0; i < table->columns().size(); ++i) {
                const auto marker = "[" + upperAscii(table->columns()[i].name()) + "]";
                const auto atMarker = "[@" + upperAscii(table->columns()[i].name()) + "]";
                if (selectorUpper.find(marker) != std::string::npos || selectorUpper.find(atMarker) != std::string::npos)
                    selectedColumns.push_back(i);
            }
            if (!selectedColumns.empty()) {
                const auto [minimum, maximum] = std::minmax_element(selectedColumns.begin(), selectedColumns.end());
                colFirst = first.column + *minimum; colLast = first.column + *maximum;
            } else if (selectorUpper.find("#") == std::string::npos && selectorUpper.find("@") == std::string::npos && selector != "[]")
                return EvalValue::fromScalar(CellError::Reference);

            ++report_.structuredReferencesResolved;
            return resolveReference(context, tableSheet->name(),
                                    CellReference(rowFirst, colFirst).address(),
                                    CellReference(rowLast, colLast).address(), depth + 1);
        } catch (...) {
            return EvalValue::fromScalar(CellError::Reference);
        }
    }

    EvalValue callFunction(std::string name, const std::vector<EvalValue>& args, Worksheet& context, std::size_t depth) override {
        return internal::formula::evaluateFormulaFunction(*this, std::move(name), args, context, depth);
    }

    bool compare(const Scalar& lhs, const Scalar& rhs, const std::string& op) const override {
        const auto ln = numberValue(lhs, workbook_.date1904());
        const auto rn = numberValue(rhs, workbook_.date1904());
        int cmp = 0;
        if (ln && rn) cmp = *ln < *rn ? -1 : (*ln > *rn ? 1 : 0);
        else {
            auto l = upperAscii(scalarText(lhs, workbook_.date1904()));
            auto r = upperAscii(scalarText(rhs, workbook_.date1904()));
            cmp = l < r ? -1 : (l > r ? 1 : 0);
        }
        if (op == "=") return cmp == 0;
        if (op == "<>") return cmp != 0;
        if (op == "<") return cmp < 0;
        if (op == "<=") return cmp <= 0;
        if (op == ">") return cmp > 0;
        if (op == ">=") return cmp >= 0;
        return false;
    }

    EvalValue unsupported(const std::string& name) {
        ++report_.unsupportedFormulas;
        report_.warnings.push_back("Unsupported formula function: " + name);
        return EvalValue::fromScalar(CellError::Name);
    }

    std::optional<FormulaKey> currentFormulaKey() const {
        if (evaluationStack_.empty()) return std::nullopt;
        return evaluationStack_.back();
    }

    const CalculationOptions& options() const noexcept { return options_; }
    bool evaluateVolatileFunctions() const noexcept override { return options_.evaluateVolatileFunctions; }
    EvalValue unsupportedFunction(const std::string& name) override { return unsupported(name); }
    bool date1904() const noexcept override { return workbook_.date1904(); }

private:
    Workbook& workbook_;
    CalculationOptions options_;
    CalculationReport report_;
    std::unordered_map<FormulaKey, Scalar, FormulaKeyHash> cache_;
    std::unordered_map<FormulaKey, Scalar, FormulaKeyHash> previousIteration_;
    std::unordered_set<FormulaKey, FormulaKeyHash> visiting_;
    std::vector<FormulaKey> evaluationStack_;
    bool iterativePass_{false};

    Scalar materializeDynamicArray(const FormulaKey& key, const EvalValue& evaluated) {
        if (iterativePass_ || !options_.spillDynamicArrays || !evaluated.isRange || evaluated.range.size() <= 1)
            return firstScalar(evaluated);
        if (!key.sheet || evaluated.rows == 0 || evaluated.columns == 0 ||
            evaluated.rows > 1048576 - key.row + 1 || evaluated.columns > 16384 - key.column + 1)
            return CellError::Spill;

        auto& formulaCell = key.sheet->cell(key.row, key.column);
        std::size_t oldR1=0,oldC1=0,oldR2=0,oldC2=0;
        bool hasOld=false;
        const auto oldRef=formulaCell.formulaMetadata().reference();
        if(!oldRef.empty()){
            try{
                const auto colon=oldRef.find(':'); auto a=CellReference::parse(colon==std::string::npos?oldRef:oldRef.substr(0,colon));
                auto b=CellReference::parse(colon==std::string::npos?oldRef:oldRef.substr(colon+1));
                oldR1=(std::min)(a.row,b.row);oldR2=(std::max)(a.row,b.row);oldC1=(std::min)(a.column,b.column);oldC2=(std::max)(a.column,b.column);hasOld=true;
            }catch(...){hasOld=false;}
        }
        auto inOld=[&](std::size_t r,std::size_t c){return hasOld&&r>=oldR1&&r<=oldR2&&c>=oldC1&&c<=oldC2;};
        const auto newR2=key.row+evaluated.rows-1,newC2=key.column+evaluated.columns-1;
        for(std::size_t r=key.row;r<=newR2;++r)for(std::size_t c=key.column;c<=newC2;++c){
            if(r==key.row&&c==key.column)continue;
            if(const auto* existing=key.sheet->tryCell(r,c);existing&&!existing->empty()&&!inOld(r,c)){
                ++report_.spillConflicts;
                report_.warnings.push_back(key.sheet->name()+"!"+formulaCell.address()+": dynamic array spill range is blocked at "+existing->address());
                formulaCell.formulaMetadata().setType(FormulaType::DynamicArray);
                formulaCell.formulaMetadata().setReference(formulaCell.address());
                formulaCell.formulaMetadata().setAlwaysCalculateArray(true);
                return CellError::Spill;
            }
        }
        if(hasOld){
            for(std::size_t r=oldR1;r<=oldR2;++r)for(std::size_t c=oldC1;c<=oldC2;++c){
                if(r==key.row&&c==key.column)continue;
                if(r<key.row||r>newR2||c<key.column||c>newC2){
                    if(auto* old=const_cast<Cell*>(key.sheet->tryCell(r,c));old&&!old->hasFormula())old->setValue(std::monostate{});
                }
            }
        }
        std::size_t i=0;
        for(std::size_t r=key.row;r<=newR2;++r)for(std::size_t c=key.column;c<=newC2;++c,++i){
            if(r==key.row&&c==key.column)continue;
            auto& target=key.sheet->cell(r,c); if(!sameScalar(target.value(),evaluated.range[i])){target.setValue(evaluated.range[i]);++report_.spillCellsUpdated;}
        }
        formulaCell.formulaMetadata().setType(FormulaType::DynamicArray);
        formulaCell.formulaMetadata().setReference(formulaCell.address()+":"+CellReference(newR2,newC2).address());
        formulaCell.formulaMetadata().setAlwaysCalculateArray(true);
        ++report_.dynamicArraysSpilled;
        return evaluated.range.empty()?Scalar{std::monostate{}}:evaluated.range.front();
    }

    Scalar evaluateCell(const FormulaKey& key, std::size_t depth) {
        if (const auto it = cache_.find(key); it != cache_.end()) return it->second;
        const auto* cell = key.sheet ? key.sheet->tryCell(key.row, key.column) : nullptr;
        if (!cell) return std::monostate{};
        if (!cell->hasFormula()) return cell->value();
        if (depth > options_.maxDepth) {
            ++report_.evaluationErrors; report_.warnings.push_back(key.sheet->name() + "!" + cell->address() + ": maximum calculation depth exceeded");
            return CellError::Number;
        }
        if (!visiting_.insert(key).second) {
            ++report_.circularReferences; report_.warnings.push_back(key.sheet->name() + "!" + cell->address() + ": circular reference");
            return CellError::Reference;
        }
        Scalar result;
        evaluationStack_.push_back(key);
        try {
            auto evaluated = internal::formula::parseFormula(*this, *key.sheet, cell->formula(), depth + 1);
            result = materializeDynamicArray(key, evaluated);
            ++report_.formulaCellsEvaluated;
        } catch (const std::exception& error) {
            ++report_.evaluationErrors;
            report_.warnings.push_back(key.sheet->name() + "!" + cell->address() + ": " + error.what());
            result = CellError::Value;
        }
        evaluationStack_.pop_back();
        visiting_.erase(key);
        cache_[key] = result;
        if (options_.updateCachedValues && !iterativePass_) {
            auto& mutableCell = key.sheet->cell(key.row, key.column);
            if (!sameScalar(mutableCell.value(), result)) { mutableCell.setValue(result); ++report_.cachedValuesUpdated; }
        }
        return result;
    }

};


} // namespace

CalculationReport Workbook::calculateFormulas(const CalculationOptions& options) {
    CalculationEngine engine(*this, options);
    return engine.run();
}

} // namespace xlpp
