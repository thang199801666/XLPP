#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/WorksheetName.h>
#include <XLPP/Formula/ReferenceTranslator.h>

#include <stdexcept>

namespace xlpp {

WorksheetRenameReport Workbook::renameWorksheet(const std::string& oldName,
                                                 const std::string& newName,
                                                 const WorksheetRenameOptions& options) {
    validateWorksheetName(newName);

    auto apply = [&](Workbook& workbook) {
        std::size_t targetIndex = workbook.sheetCount();
        for (std::size_t i = 0; i < workbook.sheetCount(); ++i) {
            if (worksheetNamesEqual(workbook[i].name(), oldName)) {
                targetIndex = i;
                break;
            }
        }
        if (targetIndex == workbook.sheetCount())
            throw std::invalid_argument("Worksheet not found: " + oldName);
        for (std::size_t i = 0; i < workbook.sheetCount(); ++i) {
            if (i != targetIndex && worksheetNamesEqual(workbook[i].name(), newName))
                throw std::invalid_argument("Worksheet already exists: " + newName);
        }

        WorksheetRenameReport report;
        const auto currentName = workbook[targetIndex].name();
        if (currentName == newName) return report;

        for (auto& sheet : workbook.sheets_) {
            const auto translated = sheet.translateWorksheetRenameReferences(currentName, newName);
            ++report.worksheetsVisited;
            report.formulasUpdated += translated.formulasUpdated;
            report.formulaMetadataUpdated += translated.formulaMetadataUpdated;
            report.chartReferencesUpdated += translated.chartReferencesUpdated;
            report.pivotReferencesUpdated += translated.pivotReferencesUpdated;
            report.hyperlinksUpdated += translated.hyperlinksUpdated;
            report.referencesUpdated += translated.worksheetReferencesUpdated;
        }
        for (auto& name : workbook.definedNames_) {
            auto translated = renameWorksheetReferences(name.value(), currentName, newName);
            if (translated.changed()) {
                name.setValue(std::move(translated.value));
                ++report.definedNamesUpdated;
                report.referencesUpdated += translated.referencesChanged;
            }
        }

        workbook[targetIndex].rename(newName);
        workbook.resetChartCacheDependencyTracking();
        workbook.calcProperties().setFullCalcOnLoad(true);
        workbook.calcProperties().setCalcOnSave(true);

        if (options.recalculateFormulas) {
            const auto calculation = workbook.calculateFormulas();
            report.formulasCalculated = calculation.formulaCellsEvaluated;
            for (const auto& warning : calculation.warnings)
                report.warnings.push_back("formula: " + warning);
        }
        if (options.synchronizeChartCaches) {
            ChartCacheSyncOptions sync;
            sync.changedReferencesOnly = options.changedChartCachesOnly;
            const auto charts = workbook.synchronizeChartCaches(sync);
            report.chartCachesUpdated = charts.cachesUpdated;
            for (const auto& warning : charts.warnings)
                report.warnings.push_back("chart cache: " + warning);
        }
        return report;
    };

    // Semantic preflight on a copy catches all modeled rewrite/calculation
    // failures before mutating live objects. The successful live pass remains
    // in-place so C/Python/C# handles to worksheets and cells stay valid.
    Workbook staged = *this;
    (void)apply(staged);
    return apply(*this);
}

} // namespace xlpp
