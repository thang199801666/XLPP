#include <XLPP/Workbook/Workbook.h>
#include "XLPP/Formula/ReferenceTransformer.h"
#include "XLPP/Packaging/RelationshipGraph.h"
#include "../Internal/WorksheetName.h"
#include <algorithm>
#include <deque>
#include <stdexcept>
#include <unordered_map>
#include <unordered_set>

namespace xlpp {
namespace {

using internal::ReferenceRewriteResult;

std::string hyperlinkTargetWithFragment(bool leadingHash, const std::string& rewritten) {
    // A location that has become #REF! already owns its leading '#'. Avoid
    // producing the invalid "##REF!" fragment that older structural code did.
    if (leadingHash && !rewritten.empty() && rewritten.front() != '#') return "#" + rewritten;
    return rewritten;
}


std::unordered_set<std::string> exclusivePreservedClosure(
    const std::string& removedSheetPart,
    const std::vector<PreservedRelationship>& relationships) {
    std::unordered_map<std::string, std::vector<std::string>> outgoing;
    std::unordered_map<std::string, std::vector<std::string>> incoming;
    for (const auto& relationship : relationships) {
        if (relationship.targetMode == "External") continue;
        const auto target = internal::RelationshipGraph::resolveTarget(relationship.sourcePart, relationship.target);
        if (target.empty()) continue;
        outgoing[relationship.sourcePart].push_back(target);
        incoming[target].push_back(relationship.sourcePart);
    }

    std::unordered_set<std::string> removable{removedSheetPart};
    bool progress = true;
    while (progress) {
        progress = false;
        std::vector<std::string> candidates;
        for (const auto& source : removable) {
            const auto it = outgoing.find(source);
            if (it == outgoing.end()) continue;
            candidates.insert(candidates.end(), it->second.begin(), it->second.end());
        }
        for (const auto& candidate : candidates) {
            if (removable.contains(candidate)) continue;
            const auto in = incoming.find(candidate);
            if (in == incoming.end() || in->second.empty()) continue;
            const bool exclusivelyOwned = std::all_of(in->second.begin(), in->second.end(), [&](const auto& source) {
                return removable.contains(source);
            });
            if (exclusivelyOwned) {
                removable.insert(candidate);
                progress = true;
            }
        }
    }
    return removable;
}


} // namespace

namespace internal {
struct WorkbookSheetOperationsAccess {
template <class Rewrite>
    static bool rewriteWorksheetReferences(Worksheet& sheet, Rewrite&& rewrite, bool invalidateCaches) {
    bool changed = false;

    for (auto& [_, cell] : sheet.cells_) {
        if (cell.hasFormula()) {
            const auto result = rewrite(cell.formula());
            if (result.changed()) {
                cell.setFormulaTextPreservingMetadata(result.text);
                changed = true;
            }
        }
        const auto& formulaMetadata = static_cast<const Cell&>(cell).formulaMetadata();
        if (!formulaMetadata.reference().empty()) {
            const auto result = rewrite(formulaMetadata.reference());
            if (result.changed()) {
                cell.formulaMetadata().setReference(result.text);
                changed = true;
            }
        }
        if (cell.hyperlinkValue() && !cell.hyperlinkValue()->external()) {
            const auto& target = cell.hyperlinkValue()->target();
            const bool leadingHash = !target.empty() && target.front() == '#';
            const auto body = leadingHash ? std::string_view(target).substr(1) : std::string_view(target);
            const auto result = rewrite(body);
            if (result.changed()) {
                cell.hyperlink().setTarget(hyperlinkTargetWithFragment(leadingHash, result.text));
                changed = true;
            }
        }
    }

    for (auto& entry : sheet.conditionalFormatting_.entries()) {
        for (auto& rule : entry.rules()) {
            auto formulas = rule.formulas();
            bool formulasChanged = false;
            for (auto& formula : formulas) {
                const auto result = rewrite(formula);
                if (!result.changed()) continue;
                formula = result.text;
                formulasChanged = true;
            }
            if (formulasChanged) {
                rule.setFormulas(std::move(formulas));
                changed = true;
            }
            auto rewriteCfvo = [&](Cfvo& value) {
                if (value.type != "formula" || value.formula.empty()) return;
                const auto result = rewrite(value.formula);
                if (!result.changed()) return;
                value.formula = result.text;
                changed = true;
            };
            rewriteCfvo(rule.getDataBar().min);
            rewriteCfvo(rule.getDataBar().max);
            for (auto& stop : rule.getColorScale().stops) rewriteCfvo(stop);
            for (auto& stop : rule.getIconSet().thresholds) rewriteCfvo(stop);
        }
    }

    for (auto& validation : sheet.dataValidations_.items()) {
        if (!validation.formula1().empty()) {
            const auto result = rewrite(validation.formula1());
            if (result.changed()) { validation.setFormula1(result.text); changed = true; }
        }
        if (!validation.formula2().empty()) {
            const auto result = rewrite(validation.formula2());
            if (result.changed()) { validation.setFormula2(result.text); changed = true; }
        }
    }

    for (std::size_t chartIndex = 0; chartIndex < sheet.charts_.size(); ++chartIndex) {
        auto& chart = sheet.charts_[chartIndex];
        for (std::size_t seriesIndex = 0; seriesIndex < chart.series().size(); ++seriesIndex) {
            auto& series = chart.series()[seriesIndex];
            const auto categories = rewrite(series.categoriesReference());
            const auto values = rewrite(series.valuesReference());
            const auto title = rewrite(series.titleReference());
            const bool seriesRefsChanged = categories.changed() || values.changed() || title.changed();

            if (seriesRefsChanged) {
                if (chart.imported() && !categories.text.empty() && !values.text.empty()) {
                    sheet.setChartSeriesReferences(chart.stableId(), seriesIndex, categories.text, values.text);
                    if (title.changed()) {
                        // Imported title references do not yet have a dedicated
                        // lossless patch API. Regenerate this chart rather than
                        // silently retaining a stale worksheet qualifier.
                        series.setTitleReference(title.text);
                        sheet.drawingsDirty_ = true;
                        sheet.drawingAppendDirty_ = false;
                    }
                } else {
                    series.setCategoriesReference(categories.text);
                    series.setValuesReference(values.text);
                    series.setTitleReference(title.text);
                    sheet.drawingsDirty_ = true;
                    sheet.drawingAppendDirty_ = false;
                }
                if (invalidateCaches) {
                    series.setTitleCache({});
                    series.setCategoriesCache({});
                    series.setValuesCache({});
                }
                changed = true;
            }

            auto bars = series.errorBars();
            bool barsChanged = false;
            for (auto& errorBars : bars) {
                const auto plus = rewrite(errorBars.plusReference);
                const auto minus = rewrite(errorBars.minusReference);
                if (plus.changed()) { errorBars.plusReference = plus.text; barsChanged = true; }
                if (minus.changed()) { errorBars.minusReference = minus.text; barsChanged = true; }
            }
            if (barsChanged) {
                if (chart.imported()) {
                    for (const auto& value : bars) sheet.setChartSeriesErrorBars(chart.stableId(), seriesIndex, value);
                } else {
                    series.setErrorBars(std::move(bars));
                    sheet.drawingsDirty_ = true;
                }
                changed = true;
            }
        }
    }

    for (auto& pivot : sheet.pivotTables_) {
        const auto result = rewrite(pivot.cache().sourceData());
        if (!result.changed()) continue;
        pivot.cache().setSourceData(result.text);
        pivot.cache().setRefreshOnLoad(true);
        sheet.pivotsDirty_ = true;
        sheet.pivotAppendDirty_ = false;
        changed = true;
    }

    if (changed) sheet.dirty_ = true;
    return changed;
}

static void retireRemovedSourceSheet(Workbook& workbook,
                              const std::string& sheetName,
                              std::size_t currentSheetIndex) {
    const auto sourceIt = std::find(workbook.sourceSheetNames_.begin(), workbook.sourceSheetNames_.end(), sheetName);
    if (sourceIt != workbook.sourceSheetNames_.end()) {
        const auto sourceIndex = static_cast<std::size_t>(std::distance(workbook.sourceSheetNames_.begin(), sourceIt));
        if (sourceIndex < workbook.sourceSheetParts_.size()) {
            const auto removedPart = workbook.sourceSheetParts_[sourceIndex];
            const auto removable = exclusivePreservedClosure(removedPart, workbook.preservedRelationships_);
            std::unordered_set<std::string> removableRelationshipParts;
            removableRelationshipParts.reserve(removable.size());
            for (const auto& source : removable)
                removableRelationshipParts.insert(internal::RelationshipGraph::relationshipsPartForSource(source));
            workbook.preservedParts_.erase(std::remove_if(workbook.preservedParts_.begin(), workbook.preservedParts_.end(), [&](const auto& part) {
                return removable.contains(part.name) || removableRelationshipParts.contains(part.name);
            }), workbook.preservedParts_.end());
            workbook.preservedRelationships_.erase(std::remove_if(workbook.preservedRelationships_.begin(), workbook.preservedRelationships_.end(), [&](const auto& relationship) {
                if (removable.contains(relationship.sourcePart)) return true;
                if (relationship.targetMode == "External") return false;
                const auto target = internal::RelationshipGraph::resolveTarget(relationship.sourcePart, relationship.target);
                return removable.contains(target);
            }), workbook.preservedRelationships_.end());
        }
        workbook.sourceSheetNames_.erase(workbook.sourceSheetNames_.begin() + static_cast<std::ptrdiff_t>(sourceIndex));
        if (sourceIndex < workbook.sourceSheetParts_.size()) workbook.sourceSheetParts_.erase(workbook.sourceSheetParts_.begin() + static_cast<std::ptrdiff_t>(sourceIndex));
        if (sourceIndex < workbook.sourceSheetXml_.size()) workbook.sourceSheetXml_.erase(workbook.sourceSheetXml_.begin() + static_cast<std::ptrdiff_t>(sourceIndex));
    }

    if (currentSheetIndex < workbook.cachedSheetXml_.size())
        workbook.cachedSheetXml_.erase(workbook.cachedSheetXml_.begin() + static_cast<std::ptrdiff_t>(currentSheetIndex));
}
};
} // namespace internal

bool Workbook::renameWorksheet(const std::string& oldName, std::string newName) {
    internal::validateWorksheetName(newName);
    auto* target = worksheet(oldName);
    if (!target) return false;
    const std::string canonicalOldName = target->name();
    if (internal::worksheetNamesEquivalent(canonicalOldName, newName)) {
        // Case-only rename is valid and still needs reference qualifier repair.
        if (canonicalOldName == newName) return true;
    }
    if (std::any_of(sheets_.begin(), sheets_.end(), [&](const auto& sheet) {
            return &sheet != target && internal::worksheetNamesEquivalent(sheet.name(), newName);
        }) || std::any_of(chartsheets_.begin(), chartsheets_.end(), [&](const auto& sheet) {
            return internal::worksheetNamesEquivalent(sheet.name(), newName);
        }))
        throw std::invalid_argument("Duplicate workbook sheet name: " + newName);

    bool referencesChanged = false;
    const auto rewrite = [&](std::string_view value) {
        return internal::renameWorksheetReferences(value, canonicalOldName, newName);
    };
    for (auto& sheet : sheets_) referencesChanged = internal::WorkbookSheetOperationsAccess::rewriteWorksheetReferences(sheet, rewrite, false) || referencesChanged;
    for (auto& name : definedNames_) {
        const auto result = rewrite(name.value());
        if (!result.changed()) continue;
        name.setValue(result.text);
        referencesChanged = true;
    }

    for (auto& sourceName : sourceSheetNames_) if (internal::worksheetNamesEquivalent(sourceName, canonicalOldName)) sourceName = newName;
    target->rename(std::move(newName));

    if (referencesChanged) {
        calcProps_.setCalcOnSave(true);
        calcProps_.setFullCalcOnLoad(true);
    }
    return true;
}

bool Workbook::removeWorksheet(const std::string& name) {
    const auto it = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return internal::worksheetNamesEquivalent(sheet.name(), name); });
    if (it == sheets_.end()) return false;
    const std::string canonicalName = it->name();
    if (sheetOrder_.size() == 1) throw std::logic_error("Cannot remove the last sheet from a workbook");
    const auto removedIndex = static_cast<std::size_t>(std::distance(sheets_.begin(), it));
    const auto orderIt = std::find_if(sheetOrder_.begin(), sheetOrder_.end(), [&](const auto& entry) {
        return entry.kind == WorkbookSheetKind::Worksheet && entry.kindIndex == removedIndex;
    });
    const auto removedWorkbookIndex = orderIt == sheetOrder_.end()
        ? removedIndex : static_cast<std::size_t>(std::distance(sheetOrder_.begin(), orderIt));

    const auto invalidate = [&](std::string_view value) {
        return internal::invalidateWorksheetReferences(value, canonicalName);
    };
    bool referencesChanged = false;
    for (auto& sheet : sheets_) {
        if (&sheet == &*it) continue;
        referencesChanged = internal::WorkbookSheetOperationsAccess::rewriteWorksheetReferences(sheet, invalidate, true) || referencesChanged;
    }

    // Defined names scoped to the removed sheet cease to exist. Surviving
    // local scopes shift with the worksheet index, while explicit references
    // to the removed sheet become #REF! like Excel's structural semantics.
    definedNames_.erase(std::remove_if(definedNames_.begin(), definedNames_.end(), [&](DefinedName& defined) {
        if (defined.localSheetId() && *defined.localSheetId() == removedWorkbookIndex) return true;
        const auto result = invalidate(defined.value());
        if (result.changed()) {
            defined.setValue(result.text);
            referencesChanged = true;
        }
        if (defined.localSheetId() && *defined.localSheetId() > removedWorkbookIndex)
            defined.setLocalSheetId(*defined.localSheetId() - 1);
        return false;
    }), definedNames_.end());

    internal::WorkbookSheetOperationsAccess::retireRemovedSourceSheet(*this, canonicalName, removedIndex);
    const auto oldActive = activeWorkbookSheetIndex_;
    if (orderIt != sheetOrder_.end()) sheetOrder_.erase(orderIt);
    for (auto& entry : sheetOrder_)
        if (entry.kind == WorkbookSheetKind::Worksheet && entry.kindIndex > removedIndex) --entry.kindIndex;
    sheets_.erase(it);
    if (!sheetOrder_.empty()) {
        if (oldActive > removedWorkbookIndex) activeWorkbookSheetIndex_ = oldActive - 1;
        else if (oldActive == removedWorkbookIndex) activeWorkbookSheetIndex_ = std::min(removedWorkbookIndex, sheetOrder_.size() - 1);
        if (sheetOrder_[activeWorkbookSheetIndex_].visibility != WorkbookSheetVisibility::Visible)
            activeWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex();
        firstVisibleWorkbookSheetIndex_ = firstVisibleWorkbookSheetIndex();
    }

    calcProps_.setCalcOnSave(true);
    calcProps_.setFullCalcOnLoad(true);
    return true;
}

} // namespace xlpp
