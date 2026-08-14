#include <XLPP/Workbook/Workbook.h>
#include "XLPP/Formula/ReferenceTransformer.h"
#include "XLPP/Pivot/PivotStructuralEdit.h"
#include <algorithm>
#include <set>
#include <stdexcept>
#include <tuple>

namespace xlpp {
namespace {
internal::StructuralEditSpec makeEditSpec(const std::string& worksheetName,
                                          std::size_t index,
                                          std::size_t amount,
                                          StructuralEditKind kind) {
    internal::StructuralEditSpec edit;
    edit.targetSheetName = worksheetName;
    edit.index = index;
    edit.amount = amount;
    switch (kind) {
    case StructuralEditKind::InsertRows:
        edit.axis = internal::StructuralAxis::Row;
        edit.action = internal::StructuralAction::Insert;
        break;
    case StructuralEditKind::DeleteRows:
        edit.axis = internal::StructuralAxis::Row;
        edit.action = internal::StructuralAction::Delete;
        break;
    case StructuralEditKind::InsertColumns:
        edit.axis = internal::StructuralAxis::Column;
        edit.action = internal::StructuralAction::Insert;
        break;
    case StructuralEditKind::DeleteColumns:
        edit.axis = internal::StructuralAxis::Column;
        edit.action = internal::StructuralAction::Delete;
        break;
    }
    return edit;
}

void accumulateFormulaRewrite(StructuralEditReport& report,
                              const internal::ReferenceRewriteResult& rewritten) {
    report.formulasRewritten += rewritten.referencesRewritten;
    report.formulaReferencesInvalidated += rewritten.referencesInvalidated;
}

std::string validationIssueIdentity(const ModelValidationIssue& issue) {
    return issue.code + "\n" + issue.worksheetName + "\n" + issue.objectId + "\n" + issue.message;
}

std::set<std::string> validationErrorIdentities(const WorkbookModelValidationReport& report) {
    std::set<std::string> identities;
    for (const auto& issue : report.issues) {
        if (issue.severity == ModelValidationSeverity::Error)
            identities.insert(validationIssueIdentity(issue));
    }
    return identities;
}

void structuralCancellationPoint(const StructuralEditOptions& options) {
    if (options.cancel && options.cancel()) throw StructuralEditCancelled{};
}
} // namespace

StructuralEditReport Workbook::insertRows(const std::string& worksheetName,
                                          std::size_t index,
                                          std::size_t amount,
                                          const StructuralEditOptions& options) {
    return applyStructuralEdit(worksheetName, index, amount, StructuralEditKind::InsertRows, options);
}

StructuralEditReport Workbook::deleteRows(const std::string& worksheetName,
                                          std::size_t index,
                                          std::size_t amount,
                                          const StructuralEditOptions& options) {
    return applyStructuralEdit(worksheetName, index, amount, StructuralEditKind::DeleteRows, options);
}

StructuralEditReport Workbook::insertColumns(const std::string& worksheetName,
                                             std::size_t index,
                                             std::size_t amount,
                                             const StructuralEditOptions& options) {
    return applyStructuralEdit(worksheetName, index, amount, StructuralEditKind::InsertColumns, options);
}

StructuralEditReport Workbook::deleteColumns(const std::string& worksheetName,
                                             std::size_t index,
                                             std::size_t amount,
                                             const StructuralEditOptions& options) {
    return applyStructuralEdit(worksheetName, index, amount, StructuralEditKind::DeleteColumns, options);
}

StructuralEditReport Workbook::applyStructuralEdit(const std::string& worksheetName,
                                                   std::size_t index,
                                                   std::size_t amount,
                                                   StructuralEditKind kind,
                                                   const StructuralEditOptions& options) {
    if (!options.rollbackOnFailure)
        return applyStructuralEditImpl(worksheetName, index, amount, kind, options);

    // Snapshot only state that structural edits are allowed to mutate. Restoring
    // worksheet contents element-by-element preserves Worksheet object identity
    // and therefore keeps caller-held Worksheet& references valid on rollback.
    const auto sheetSnapshot = sheets_;
    const auto definedNamesSnapshot = definedNames_;
    const auto calcPropsSnapshot = calcProps_;
    const auto cachedSheetXmlSnapshot = cachedSheetXml_;
    const bool cachedStrictSnapshot = cachedSheetXmlStrict_;
    const bool cachedDateSnapshot = cachedSheetXmlDate1904_;

    try {
        return applyStructuralEditImpl(worksheetName, index, amount, kind, options);
    } catch (...) {
        if (sheets_.size() == sheetSnapshot.size()) {
            for (std::size_t i = 0; i < sheets_.size(); ++i) sheets_[i] = sheetSnapshot[i];
        } else {
            // Structural edits never intentionally change sheet count, but keep
            // a defensive fallback if a future implementation does.
            sheets_ = sheetSnapshot;
        }
        definedNames_ = definedNamesSnapshot;
        calcProps_ = calcPropsSnapshot;
        cachedSheetXml_ = cachedSheetXmlSnapshot;
        cachedSheetXmlStrict_ = cachedStrictSnapshot;
        cachedSheetXmlDate1904_ = cachedDateSnapshot;
        throw;
    }
}

StructuralEditReport Workbook::applyStructuralEditImpl(const std::string& worksheetName,
                                                   std::size_t index,
                                                   std::size_t amount,
                                                   StructuralEditKind kind,
                                                   const StructuralEditOptions& options) {
    if (index == 0 || amount == 0)
        throw std::invalid_argument("Structural edit index and amount must be positive");
    auto* target = worksheet(worksheetName);
    if (!target) throw std::out_of_range("Worksheet not found: " + worksheetName);

    structuralCancellationPoint(options);
    std::set<std::string> baselineValidationErrors;
    if (options.validateResult) baselineValidationErrors = validationErrorIdentities(validateModelIntegrity());

    StructuralEditReport report;
    report.kind = kind;
    report.worksheetName = worksheetName;
    report.index = index;
    report.amount = amount;
    const auto edit = makeEditSpec(worksheetName, index, amount, kind);

    // The worksheet-local pass rewrites cell geometry, local formulas/ranges,
    // tables, drawings and local chart/pivot references. Workbook then performs
    // the cross-sheet pass below.
    switch (kind) {
    case StructuralEditKind::InsertRows:
        target->shiftRows(index, amount, true, &options, &report);
        break;
    case StructuralEditKind::DeleteRows:
        target->shiftRows(index, amount, false, &options, &report);
        break;
    case StructuralEditKind::InsertColumns:
        target->shiftColumns(index, amount, true, &options, &report);
        break;
    case StructuralEditKind::DeleteColumns:
        target->shiftColumns(index, amount, false, &options, &report);
        break;
    }
    structuralCancellationPoint(options);

    // Rewrite formulas and formula-bearing worksheet features on every other
    // sheet when they explicitly point at the edited sheet.
    if (options.updateCellFormulas || options.updateHyperlinks || options.updateConditionalFormatting || options.updateDataValidations) {
        for (auto& sheet : sheets_) {
            if (&sheet == target) continue;
            bool sheetChanged = false;

            if (options.updateCellFormulas) {
                for (auto& [_, cell] : sheet.cells_) {
                    if (!cell.hasFormula()) continue;
                    const auto rewritten = internal::rewriteFormulaReferences(cell.formula(), sheet.name_, edit);
                    report.referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                    if (!rewritten.changed()) continue;
                    cell.setFormula(rewritten.text);
                    sheetChanged = true;
                    accumulateFormulaRewrite(report, rewritten);
                }
            }

            if (options.updateHyperlinks) {
                for (auto& [_, cell] : sheet.cells_) {
                    if (!cell.hyperlinkValue() || cell.hyperlinkValue()->external()) continue;
                    const auto& currentTarget = cell.hyperlinkValue()->target();
                    const bool leadingHash = !currentTarget.empty() && currentTarget.front() == '#';
                    const auto body = leadingHash ? std::string_view(currentTarget).substr(1) : std::string_view(currentTarget);
                    const auto rewritten = internal::rewriteFormulaReferences(body, sheet.name_, edit);
                    report.referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                    if (!rewritten.changed()) continue;
                    cell.hyperlink().setTarget(leadingHash && !rewritten.text.empty() && rewritten.text.front() != '#' ? "#" + rewritten.text : rewritten.text);
                    report.hyperlinksRewritten += rewritten.referencesRewritten;
                    report.formulaReferencesInvalidated += rewritten.referencesInvalidated;
                    sheetChanged = true;
                }
            }

            if (options.updateConditionalFormatting) {
                for (auto& entry : sheet.conditionalFormatting_.entries()) {
                    for (auto& rule : entry.rules()) {
                        auto formulas = rule.formulas();
                        bool changed = false;
                        for (auto& formula : formulas) {
                            const auto rewritten = internal::rewriteFormulaReferences(formula, sheet.name_, edit);
                            report.referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                            if (!rewritten.changed()) continue;
                            formula = rewritten.text;
                            changed = true;
                            accumulateFormulaRewrite(report, rewritten);
                        }
                        if (changed) { rule.setFormulas(std::move(formulas)); sheetChanged = true; }
                        auto rewriteCfvo = [&](Cfvo& value) {
                            if (value.type != "formula" || value.formula.empty()) return;
                            const auto rewritten = internal::rewriteFormulaReferences(value.formula, sheet.name_, edit);
                            report.referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                            if (!rewritten.changed()) return;
                            value.formula = rewritten.text;
                            sheetChanged = true;
                            accumulateFormulaRewrite(report, rewritten);
                        };
                        rewriteCfvo(rule.getDataBar().min);
                        rewriteCfvo(rule.getDataBar().max);
                        for (auto& stop : rule.getColorScale().stops) rewriteCfvo(stop);
                        for (auto& stop : rule.getIconSet().thresholds) rewriteCfvo(stop);
                    }
                }
            }

            if (options.updateDataValidations) {
                for (auto& validation : sheet.dataValidations_.items()) {
                    if (!validation.formula1().empty()) {
                        const auto rewritten = internal::rewriteFormulaReferences(validation.formula1(), sheet.name_, edit);
                        report.referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                        if (rewritten.changed()) {
                            validation.setFormula1(rewritten.text);
                            sheetChanged = true;
                            accumulateFormulaRewrite(report, rewritten);
                        }
                    }
                    if (!validation.formula2().empty()) {
                        const auto rewritten = internal::rewriteFormulaReferences(validation.formula2(), sheet.name_, edit);
                        report.referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                        if (rewritten.changed()) {
                            validation.setFormula2(rewritten.text);
                            sheetChanged = true;
                            accumulateFormulaRewrite(report, rewritten);
                        }
                    }
                }
            }

            if (sheetChanged) sheet.dirty_ = true;
        }
    }
    structuralCancellationPoint(options);

    if (options.updateDefinedNames) {
        for (auto& name : definedNames_) {
            std::string ownerSheet;
            if (name.localSheetId() && *name.localSheetId() < sheets_.size())
                ownerSheet = sheets_[*name.localSheetId()].name();
            const auto rewritten = internal::rewriteFormulaReferences(name.value(), ownerSheet, edit);
            report.referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
            if (!rewritten.changed()) continue;
            name.setValue(rewritten.text);
            ++report.definedNamesRewritten;
            report.formulaReferencesInvalidated += rewritten.referencesInvalidated;
        }
    }
    structuralCancellationPoint(options);

    if (options.updateCharts) {
        for (auto& sheet : sheets_) {
            if (&sheet == target) continue; // local chart references already handled by Worksheet.
            for (auto& chart : sheet.charts_) {
                for (std::size_t seriesIndex = 0; seriesIndex < chart.series().size(); ++seriesIndex) {
                    auto& series = chart.series()[seriesIndex];
                    const auto categories = internal::rewriteFormulaReferences(series.categoriesReference(), sheet.name_, edit);
                    const auto values = internal::rewriteFormulaReferences(series.valuesReference(), sheet.name_, edit);
                    const auto title = internal::rewriteFormulaReferences(series.titleReference(), sheet.name_, edit);
                    report.referencesSkippedUnsupported += categories.referencesSkippedUnsupported +
                                                          values.referencesSkippedUnsupported +
                                                          title.referencesSkippedUnsupported;
                    if (!categories.changed() && !values.changed() && !title.changed()) continue;

                    if (chart.imported() && !categories.text.empty() && !values.text.empty()) {
                        sheet.setChartSeriesReferences(chart.stableId(), seriesIndex, categories.text, values.text);
                        if (title.changed()) {
                            // There is no lossless imported-title-reference edit
                            // in P1J yet. Force full chart regeneration rather
                            // than silently retaining a stale title reference.
                            series.setTitleReference(title.text);
                            sheet.drawingsDirty_ = true;
                            sheet.drawingAppendDirty_ = false;
                            report.diagnostics.push_back("Imported chart title reference required full chart regeneration: " + chart.stableId());
                        }
                    } else {
                        series.setCategoriesReference(categories.text);
                        series.setValuesReference(values.text);
                        series.setTitleReference(title.text);
                        sheet.dirty_ = true;
                        sheet.drawingsDirty_ = true;
                        sheet.drawingAppendDirty_ = false;
                    }
                    report.chartReferencesRewritten += categories.referencesRewritten + values.referencesRewritten + title.referencesRewritten;

                    auto bars = series.errorBars();
                    bool barsChanged = false;
                    for (auto& errorBars : bars) {
                        const auto plus = internal::rewriteFormulaReferences(errorBars.plusReference, sheet.name_, edit);
                        const auto minus = internal::rewriteFormulaReferences(errorBars.minusReference, sheet.name_, edit);
                        if (plus.changed()) { errorBars.plusReference = plus.text; barsChanged = true; }
                        if (minus.changed()) { errorBars.minusReference = minus.text; barsChanged = true; }
                        report.chartReferencesRewritten += plus.referencesRewritten + minus.referencesRewritten;
                        report.referencesSkippedUnsupported += plus.referencesSkippedUnsupported + minus.referencesSkippedUnsupported;
                    }
                    if (barsChanged) {
                        if (chart.imported()) {
                            for (const auto& value : bars) sheet.setChartSeriesErrorBars(chart.stableId(), seriesIndex, value);
                        } else {
                            series.setErrorBars(std::move(bars));
                            sheet.dirty_ = true;
                            sheet.drawingsDirty_ = true;
                        }
                    }
                }
            }
        }
    }

    structuralCancellationPoint(options);

    if (options.updatePivotSources) {
        for (auto& sheet : sheets_) {
            if (&sheet == target) continue;
            for (auto& pivot : sheet.pivotTables_) {
                if (!internal::rewritePivotSourceForStructuralEdit(pivot, sheet.name_, edit, report)) continue;
                sheet.dirty_ = true;
                sheet.pivotsDirty_ = true;
                sheet.pivotAppendDirty_ = false;
                if (!pivot.cache().sharedCacheKey().empty())
                    report.diagnostics.push_back("Shared imported PivotCache source was rewritten; all owners should be validated after save: " + pivot.cache().sharedCacheKey());
            }
        }
    }

    structuralCancellationPoint(options);

    if (options.updateCharts && options.synchronizeChartCaches && report.chartReferencesRewritten != 0) {
        ChartCacheSyncOptions syncOptions;
        syncOptions.onlyChangedCells = false;
        syncOptions.preserveFormulaCachedValues = true;
        syncOptions.propagateFormulaDependencies = true;
        const auto sync = synchronizeChartCaches(syncOptions);
        report.chartCachesSynchronized += sync.cachesUpdated;
        for (const auto& warning : sync.warnings)
            report.diagnostics.push_back("Chart cache sync: " + warning);
        for (const auto& diagnostic : sync.formulaDependencyDiagnostics)
            report.diagnostics.push_back("Chart dependency: " + diagnostic);
    }

    structuralCancellationPoint(options);

    // Structural edits change formula dependency geometry. Ask Excel/Calc to
    // recalculate on open/save rather than trusting stale cached formula values.
    if (report.formulasRewritten != 0 || report.definedNamesRewritten != 0 ||
        report.pivotSourcesRewritten != 0 || report.chartReferencesRewritten != 0) {
        calcProps_.setCalcOnSave(true);
        calcProps_.setFullCalcOnLoad(true);
    }

    if (report.referencesSkippedUnsupported != 0)
        report.diagnostics.push_back("Structural edit preserved " + std::to_string(report.referencesSkippedUnsupported) +
                                     " unsupported 3-D worksheet reference(s) without guessing coordinate semantics");

    if (options.validateResult) {
        const auto validation = validateModelIntegrity();
        report.modelErrorsAfterEdit = validation.errorCount();
        report.modelWarningsAfterEdit = validation.warningCount();
        for (const auto& issue : validation.issues) {
            if (issue.severity != ModelValidationSeverity::Error) continue;
            const auto identity = validationIssueIdentity(issue);
            if (baselineValidationErrors.contains(identity)) continue;
            throw std::runtime_error("Structural edit introduced model-integrity error [" + issue.code + "]: " + issue.message);
        }
    }

    return report;
}

} // namespace xlpp
