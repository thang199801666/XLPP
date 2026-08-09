#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Formula/ReferenceTranslator.h>

#include <algorithm>
#include <cctype>
#include <stdexcept>

namespace xlpp {
namespace {

bool sameSheet(std::string_view a, std::string_view b) {
    if (a.size() != b.size()) return false;
    for (std::size_t i = 0; i < a.size(); ++i) {
        if (std::tolower(static_cast<unsigned char>(a[i])) !=
            std::tolower(static_cast<unsigned char>(b[i]))) return false;
    }
    return true;
}

Worksheet* findSheet(Workbook& workbook, std::string_view name) {
    for (auto& sheet : workbook.worksheets()) if (sameSheet(sheet.name(), name)) return &sheet;
    return nullptr;
}

void aggregate(StructuralEditReport& dst, const WorksheetStructuralEditReport& src) {
    ++dst.worksheetsVisited;
    dst.cellsMoved += src.cellsMoved;
    dst.cellsRemoved += src.cellsRemoved;
    dst.formulasUpdated += src.formulasUpdated;
    dst.formulaMetadataUpdated += src.formulaMetadataUpdated;
    dst.worksheetReferencesUpdated += src.worksheetReferencesUpdated;
    dst.chartReferencesUpdated += src.chartReferencesUpdated;
    dst.pivotReferencesUpdated += src.pivotReferencesUpdated;
    dst.drawingAnchorsUpdated += src.drawingAnchorsUpdated;
    dst.hyperlinksUpdated += src.hyperlinksUpdated;
    dst.referencesInvalidated += src.referencesInvalidated;
}

void validateBounds(const Worksheet& sheet, const StructuralEdit& edit) {
    if (edit.index == 0 || edit.amount == 0)
        throw std::invalid_argument("Structural edit index and amount must be positive");
    if (edit.kind == StructuralEditKind::InsertRows) {
        if (edit.index > MaxExcelRows || edit.amount > MaxExcelRows ||
            sheet.maxRow() > MaxExcelRows - edit.amount)
            throw std::out_of_range("Row insertion would exceed Excel's 1,048,576-row limit");
    } else if (edit.kind == StructuralEditKind::DeleteRows) {
        if (edit.index > MaxExcelRows)
            throw std::out_of_range("Row deletion index exceeds Excel's row limit");
    } else if (edit.kind == StructuralEditKind::InsertColumns) {
        if (edit.index > MaxExcelColumns || edit.amount > MaxExcelColumns ||
            sheet.maxColumn() > MaxExcelColumns - edit.amount)
            throw std::out_of_range("Column insertion would exceed Excel's 16,384-column limit");
    } else if (edit.index > MaxExcelColumns) {
        throw std::out_of_range("Column deletion index exceeds Excel's column limit");
    }
}

StructuralEditReport applyInPlace(Workbook& workbook, const StructuralEdit& edit,
                                  const StructuralEditOptions& options) {
    auto* target = findSheet(workbook, edit.sheetName);
    if (!target) throw std::invalid_argument("Structural edit target worksheet not found: " + edit.sheetName);
    validateBounds(*target, edit);

    StructuralEditReport report;
    for (auto& sheet : workbook.worksheets())
        aggregate(report, sheet.applyStructuralEdit(edit));

    if (options.updateDefinedNames) {
        for (auto& name : workbook.definedNames()) {
            std::string context;
            if (name.localSheetId() && *name.localSheetId() < workbook.sheetCount())
                context = workbook[*name.localSheetId()].name();
            auto tr = translateFormulaReferences(name.value(), context, edit);
            if (tr.changed()) {
                name.setValue(std::move(tr.value));
                ++report.definedNamesUpdated;
                report.referencesInvalidated += tr.referencesInvalidated;
            }
        }
    }

    workbook.resetChartCacheDependencyTracking();
    // A structural edit changes formula topology even if XLPP cannot evaluate
    // every formula. Ask the host to perform a full calculation on next open.
    workbook.calcProperties().setFullCalcOnLoad(true);
    workbook.calcProperties().setCalcOnSave(true);

    if (options.recalculateFormulas) {
        auto calculation = workbook.calculateFormulas();
        report.formulasCalculated = calculation.formulaCellsEvaluated;
        for (const auto& warning : calculation.warnings)
            report.warnings.push_back("formula: " + warning);
    }
    if (options.synchronizeChartCaches) {
        ChartCacheSyncOptions sync;
        sync.changedReferencesOnly = options.changedChartCachesOnly;
        auto charts = workbook.synchronizeChartCaches(sync);
        report.chartCachesUpdated = charts.cachesUpdated;
        for (const auto& warning : charts.warnings)
            report.warnings.push_back("chart cache: " + warning);
    }
    return report;
}
}

StructuralEditReport Workbook::applyStructuralEdit(const StructuralEdit& edit,
                                                    const StructuralEditOptions& options) {
    if (!options.transactional) {
        auto report = applyInPlace(*this, edit, options);
        if (options.failOnInvalidReference && report.referencesInvalidated != 0)
            throw std::runtime_error("Structural edit produced invalid #REF! dependencies");
        return report;
    }

    // Preflight the complete transaction on a private copy. This provides
    // rollback semantics without replacing the live Workbook on success,
    // which keeps Worksheet/Cell handles held by C/Python/C# callers valid.
    Workbook staged = *this;
    auto stagedOptions = options;
    stagedOptions.transactional = false;
    const auto preflight = applyInPlace(staged, edit, stagedOptions);
    if (options.failOnInvalidReference && preflight.referencesInvalidated != 0)
        throw std::runtime_error("Structural edit would produce invalid #REF! dependencies");

    auto liveOptions = options;
    liveOptions.transactional = false;
    return applyInPlace(*this, edit, liveOptions);
}

StructuralEditReport Workbook::insertRows(const std::string& sheetName, std::size_t index,
                                          std::size_t amount, const StructuralEditOptions& options) {
    return applyStructuralEdit({sheetName, StructuralEditKind::InsertRows, index, amount}, options);
}
StructuralEditReport Workbook::deleteRows(const std::string& sheetName, std::size_t index,
                                          std::size_t amount, const StructuralEditOptions& options) {
    return applyStructuralEdit({sheetName, StructuralEditKind::DeleteRows, index, amount}, options);
}
StructuralEditReport Workbook::insertColumns(const std::string& sheetName, std::size_t index,
                                             std::size_t amount, const StructuralEditOptions& options) {
    return applyStructuralEdit({sheetName, StructuralEditKind::InsertColumns, index, amount}, options);
}
StructuralEditReport Workbook::deleteColumns(const std::string& sheetName, std::size_t index,
                                             std::size_t amount, const StructuralEditOptions& options) {
    return applyStructuralEdit({sheetName, StructuralEditKind::DeleteColumns, index, amount}, options);
}

} // namespace xlpp
