#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Cell/CellReference.h>
#include "../Internal/WorksheetName.h"
#include <algorithm>
#include <cctype>
#include <cstddef>
#include <optional>
#include <set>
#include <string>
#include <unordered_set>

namespace xlpp {
std::size_t WorkbookModelValidationReport::errorCount() const noexcept {
    return static_cast<std::size_t>(std::count_if(issues.begin(), issues.end(), [](const auto& issue) {
        return issue.severity == ModelValidationSeverity::Error;
    }));
}
std::size_t WorkbookModelValidationReport::warningCount() const noexcept {
    return static_cast<std::size_t>(std::count_if(issues.begin(), issues.end(), [](const auto& issue) {
        return issue.severity == ModelValidationSeverity::Warning;
    }));
}

namespace {
constexpr std::size_t kMaxExcelRow = 1048576;
constexpr std::size_t kMaxExcelColumn = 16384;

struct RectReference {
    std::string sheetName;
    CellReference first;
    CellReference last;
};

std::string unquoteSheet(std::string value) {
    if (value.size() >= 2 && value.front() == '\'' && value.back() == '\'') {
        value = value.substr(1, value.size() - 2);
        std::size_t at = 0;
        while ((at = value.find("''", at)) != std::string::npos) {
            value.replace(at, 2, "'");
            ++at;
        }
    }
    return value;
}

std::optional<RectReference> parseRect(std::string text, const std::string& defaultSheet) {
    try {
        if (!text.empty() && text.front() == '=') text.erase(text.begin());
        if (text.find("#REF!") != std::string::npos) return std::nullopt;
        std::string sheet = defaultSheet;
        const auto bang = text.rfind('!');
        if (bang != std::string::npos) {
            sheet = unquoteSheet(text.substr(0, bang));
            text = text.substr(bang + 1);
        }
        const auto colon = text.find(':');
        if (colon != std::string::npos && text.find(':', colon + 1) != std::string::npos) return std::nullopt;
        auto first = CellReference::parse(colon == std::string::npos ? text : text.substr(0, colon));
        auto last = CellReference::parse(colon == std::string::npos ? text : text.substr(colon + 1));
        if (first.row > last.row) std::swap(first.row, last.row);
        if (first.column > last.column) std::swap(first.column, last.column);
        if (first.row > kMaxExcelRow || last.row > kMaxExcelRow ||
            first.column > kMaxExcelColumn || last.column > kMaxExcelColumn) return std::nullopt;
        return RectReference{std::move(sheet), first, last};
    } catch (...) {
        return std::nullopt;
    }
}

void addIssue(WorkbookModelValidationReport& report,
              ModelValidationSeverity severity,
              std::string code,
              std::string sheet,
              std::string object,
              std::string message) {
    report.issues.push_back({severity, std::move(code), std::move(sheet), std::move(object), std::move(message)});
}

bool rectanglesOverlap(const RectReference& a, const RectReference& b) noexcept {
    return internal::worksheetNamesEquivalent(a.sheetName, b.sheetName) &&
           a.first.row <= b.last.row && b.first.row <= a.last.row &&
           a.first.column <= b.last.column && b.first.column <= a.last.column;
}

std::vector<std::string> splitReferenceAreas(std::string_view text) {
    std::vector<std::string> areas;
    std::string token;
    bool quoted = false;
    for (std::size_t i = 0; i < text.size(); ++i) {
        const char c = text[i];
        if (c == '\'') {
            token.push_back(c);
            if (quoted && i + 1 < text.size() && text[i + 1] == '\'') { token.push_back('\''); ++i; }
            else quoted = !quoted;
            continue;
        }
        if (!quoted && (c == ',' || std::isspace(static_cast<unsigned char>(c)))) {
            if (!token.empty()) { areas.push_back(std::move(token)); token.clear(); }
            continue;
        }
        token.push_back(c);
    }
    if (!token.empty()) areas.push_back(std::move(token));
    return areas;
}

bool validFieldIndex(int index, std::size_t count) noexcept {
    return index >= 0 && static_cast<std::size_t>(index) < count;
}

std::string asciiFold(std::string value) {
    for (char& c : value) {
        const unsigned char u = static_cast<unsigned char>(c);
        if (u >= 'A' && u <= 'Z') c = static_cast<char>(u - 'A' + 'a');
    }
    return value;
}

bool hasSheet(const std::deque<Worksheet>& sheets, std::string_view name) {
    return std::any_of(sheets.begin(), sheets.end(), [&](const auto& sheet) {
        return internal::worksheetNamesEquivalent(sheet.name(), name);
    });
}

void validateAnchor(const DrawingAnchorInfo& anchor,
                    const std::string& sheet,
                    const std::string& id,
                    WorkbookModelValidationReport& report) {
    if (anchor.type == DrawingAnchorType::Absolute) return;
    auto bad = [](const DrawingMarker& marker) {
        return marker.row == 0 || marker.column == 0 || marker.row > kMaxExcelRow || marker.column > kMaxExcelColumn;
    };
    if (bad(anchor.from) || (anchor.type == DrawingAnchorType::TwoCell && bad(anchor.to)))
        addIssue(report, ModelValidationSeverity::Error, "drawing.anchor.out_of_bounds", sheet, id,
                 "Drawing anchor lies outside the Excel worksheet grid");
}
} // namespace

WorkbookModelValidationReport Workbook::validateModelIntegrity() const {
    WorkbookModelValidationReport report;
    if (sheetOrder_.empty()) {
        addIssue(report, ModelValidationSeverity::Error, "workbook.no_sheets", {}, {},
                 "Workbook contains no worksheets or chartsheets");
        return report;
    }

    std::size_t visibleSheetCount = 0;
    for (const auto& entry : sheetOrder_)
        if (entry.visibility == WorkbookSheetVisibility::Visible) ++visibleSheetCount;
    if (visibleSheetCount == 0)
        addIssue(report, ModelValidationSeverity::Error, "workbook.no_visible_sheet", {}, {},
                 "Workbook must contain at least one visible worksheet or chartsheet");
    if (activeWorkbookSheetIndex_ >= sheetOrder_.size())
        addIssue(report, ModelValidationSeverity::Error, "workbook.active_sheet_out_of_range", {}, {},
                 "Workbook active sheet index is outside the mixed sheet order");
    else if (sheetOrder_[activeWorkbookSheetIndex_].visibility != WorkbookSheetVisibility::Visible)
        addIssue(report, ModelValidationSeverity::Error, "workbook.active_sheet_hidden", {}, {},
                 "Workbook active sheet must be visible");

    std::unordered_set<std::string> sheetNames;
    std::unordered_set<std::string> tableNames;
    std::unordered_set<std::string> pivotNames;
    for (const auto& chartSheet : chartsheets_) {
        try { internal::validateWorksheetName(chartSheet.name()); }
        catch (const std::exception& e) {
            addIssue(report, ModelValidationSeverity::Error, "chartsheet.invalid_name", chartSheet.name(), {}, e.what());
        }
        const auto foldedChartSheetName = asciiFold(chartSheet.name());
        if (!sheetNames.insert(foldedChartSheetName).second)
            addIssue(report, ModelValidationSeverity::Error, "workbook.duplicate_sheet_name", chartSheet.name(), {}, "Duplicate workbook sheet name (case-insensitive)");
        if (!chartSheet.hasChart() && !chartSheet.imported())
            addIssue(report, ModelValidationSeverity::Error, "chartsheet.missing_chart", chartSheet.name(), {}, "Generated chartsheet has no chart");
        std::unordered_set<std::string> customViewGuids;
        for (const auto& customView : chartSheet.customViews()) {
            if (customView.guid().empty()) {
                addIssue(report, ModelValidationSeverity::Error, "chartsheet.custom_view_missing_guid", chartSheet.name(), {},
                         "Custom Chartsheet view must have a GUID");
            } else if (!customViewGuids.insert(asciiFold(customView.guid())).second) {
                addIssue(report, ModelValidationSeverity::Error, "chartsheet.custom_view_duplicate_guid", chartSheet.name(), customView.guid(),
                         "Custom Chartsheet view GUID must be unique within the Chartsheet");
            }
        }
        const auto& chartProtection = chartSheet.protection();
        const bool hasModernProtection = chartProtection.algorithmName() || chartProtection.hashValue() ||
                                         chartProtection.saltValue() || chartProtection.spinCount();
        const bool completeModernProtection = chartProtection.algorithmName() && chartProtection.hashValue() &&
                                              chartProtection.saltValue() && chartProtection.spinCount();
        if (hasModernProtection && !completeModernProtection)
            addIssue(report, ModelValidationSeverity::Warning, "chartsheet.protection_incomplete_modern_hash", chartSheet.name(), {},
                     "Modern Chartsheet protection metadata should include algorithmName, hashValue, saltValue and spinCount together");
        if (!chartSheet.imported() && chartSheet.hasPageSetup() && chartSheet.pageSetup().relationshipId() && !chartSheet.hasPrinterSettings())
            addIssue(report, ModelValidationSeverity::Warning, "chartsheet.pagesetup_unresolved_printer_settings", chartSheet.name(), {},
                     "Generated Chartsheet pageSetup contains an r:id but no printerSettings payload is attached");
        if (chartSheet.hasPrinterSettings() && !chartSheet.hasPageSetup())
            addIssue(report, ModelValidationSeverity::Error, "chartsheet.printer_settings_without_pagesetup", chartSheet.name(), {},
                     "Chartsheet printerSettings payload requires a pageSetup owner");
    }
    for (const auto& sheet : sheets_) {
        try { internal::validateWorksheetName(sheet.name_); }
        catch (const std::exception& e) {
            addIssue(report, ModelValidationSeverity::Error, "worksheet.invalid_name", sheet.name_, {}, e.what());
        }
        const auto foldedSheetName = asciiFold(sheet.name_);
        if (!sheetNames.insert(foldedSheetName).second)
            addIssue(report, ModelValidationSeverity::Error, "workbook.duplicate_sheet_name", sheet.name_, {}, "Duplicate workbook sheet name (case-insensitive)");

        for (const auto& [_, cell] : sheet.cells_) {
            if (cell.row() == 0 || cell.column() == 0 || cell.row() > kMaxExcelRow || cell.column() > kMaxExcelColumn)
                addIssue(report, ModelValidationSeverity::Error, "cell.out_of_bounds", sheet.name_, cell.address(),
                         "Cell lies outside the Excel worksheet grid");
            if (cell.hasFormula() && cell.formula().find("#REF!") != std::string::npos)
                addIssue(report, ModelValidationSeverity::Warning, "formula.broken_reference", sheet.name_, cell.address(),
                         "Formula contains #REF!");
            if (cell.hasHyperlink() && !cell.hyperlinkValue()->external() &&
                cell.hyperlinkValue()->target().find("#REF!") != std::string::npos)
                addIssue(report, ModelValidationSeverity::Warning, "hyperlink.broken_internal_target", sheet.name_, cell.address(),
                         "Internal hyperlink target contains #REF!");
            if (!cell.formulaMetadata().reference().empty() &&
                cell.formulaMetadata().reference().find("#REF!") != std::string::npos)
                addIssue(report, ModelValidationSeverity::Warning, "formula.metadata_ref", sheet.name_, cell.address(),
                         "Formula metadata range contains #REF!");
        }

        std::vector<std::pair<std::string, RectReference>> mergedRects;
        for (const auto& range : sheet.mergedRanges_) {
            const auto parsed = parseRect(range, sheet.name_);
            if (!parsed || (parsed->first.row == parsed->last.row && parsed->first.column == parsed->last.column)) {
                addIssue(report, ModelValidationSeverity::Error, "merge.invalid_range", sheet.name_, range,
                         "Merged-cell range is invalid or collapsed to one cell");
                continue;
            }
            for (const auto& [otherName, other] : mergedRects) {
                if (rectanglesOverlap(*parsed, other))
                    addIssue(report, ModelValidationSeverity::Error, "merge.overlap", sheet.name_, range,
                             "Merged-cell range overlaps another merged range: " + otherName);
            }
            mergedRects.emplace_back(range, *parsed);
        }

        std::vector<std::pair<std::string, RectReference>> tableRects;
        for (const auto& table : sheet.tables_) {
            if (!tableNames.insert(asciiFold(table.name())).second)
                addIssue(report, ModelValidationSeverity::Error, "table.duplicate_name", sheet.name_, table.name(),
                         "Table names must be unique across the workbook (case-insensitive)");
            const auto parsed = parseRect(table.reference(), sheet.name_);
            if (!parsed) {
                addIssue(report, ModelValidationSeverity::Error, "table.invalid_range", sheet.name_, table.name(),
                         "Table range is invalid");
                continue;
            }
            if (!internal::worksheetNamesEquivalent(parsed->sheetName, sheet.name_))
                addIssue(report, ModelValidationSeverity::Error, "table.cross_sheet_range", sheet.name_, table.name(),
                         "Table range cannot belong to a different worksheet");
            for (const auto& [otherName, other] : tableRects) {
                if (rectanglesOverlap(*parsed, other))
                    addIssue(report, ModelValidationSeverity::Error, "table.overlap", sheet.name_, table.name(),
                             "Table range overlaps another table: " + otherName);
            }
            tableRects.emplace_back(table.name(), *parsed);

            const auto width = parsed->last.column - parsed->first.column + 1;
            if (!table.columns().empty() && table.columns().size() != width)
                addIssue(report, ModelValidationSeverity::Error, "table.column_width_mismatch", sheet.name_, table.name(),
                         "Table column model width does not match its worksheet range");
        }

        if (sheet.autoFilter_.enabled()) {
            const auto parsed = parseRect(sheet.autoFilter_.reference(), sheet.name_);
            if (!parsed || !internal::worksheetNamesEquivalent(parsed->sheetName, sheet.name_))
                addIssue(report, ModelValidationSeverity::Error, "autofilter.invalid_range", sheet.name_, sheet.autoFilter_.reference(),
                         "AutoFilter range is invalid or belongs to another worksheet");
        }

        auto validateAreaList = [&](std::string_view value, const char* code, const std::string& objectId) {
            const auto areas = splitReferenceAreas(value);
            if (areas.empty()) {
                addIssue(report, ModelValidationSeverity::Error, code, sheet.name_, objectId, "Reference list is empty");
                return;
            }
            for (const auto& area : areas) {
                const auto parsed = parseRect(area, sheet.name_);
                if (!parsed || !internal::worksheetNamesEquivalent(parsed->sheetName, sheet.name_)) {
                    addIssue(report, ModelValidationSeverity::Error, code, sheet.name_, objectId,
                             "Reference list contains an invalid or cross-sheet area: " + area);
                    break;
                }
            }
        };
        for (const auto& entry : sheet.conditionalFormatting_.entries())
            validateAreaList(entry.reference(), "conditional_formatting.invalid_range", entry.reference());
        for (const auto& validation : sheet.dataValidations_.items())
            validateAreaList(validation.reference(), "data_validation.invalid_range", validation.reference());

        if (sheet.freezePane_) {
            try { (void)CellReference::parse(*sheet.freezePane_); }
            catch (...) {
                addIssue(report, ModelValidationSeverity::Error, "worksheet.freeze_pane_invalid", sheet.name_, {},
                         "Freeze-pane top-left cell is invalid");
            }
        }

        for (const auto& image : sheet.images_) validateAnchor(image.anchorInfo(), sheet.name_, image.stableId(), report);
        for (const auto& chart : sheet.charts_) {
            validateAnchor(chart.anchorInfo(), sheet.name_, chart.stableId(), report);
            for (std::size_t i = 0; i < chart.series().size(); ++i) {
                const auto& series = chart.series()[i];
                const auto seriesId = chart.stableId() + ":series:" + std::to_string(i);
                if (!series.titleCache().valid(true) || !series.categoriesCache().valid(true) || !series.valuesCache().valid(true))
                    addIssue(report, ModelValidationSeverity::Error, "chart.cache_invalid", sheet.name_, seriesId,
                             "Chart cache contains duplicate/out-of-range point indexes");
                if (series.categoriesReference().find("#REF!") != std::string::npos ||
                    series.valuesReference().find("#REF!") != std::string::npos ||
                    series.titleReference().find("#REF!") != std::string::npos)
                    addIssue(report, ModelValidationSeverity::Warning, "chart.reference_broken", sheet.name_, seriesId,
                             "Chart series contains a #REF! reference");
                auto validateSimpleChartSource = [&](const std::string& reference) {
                    if (reference.empty() || reference.find("#REF!") != std::string::npos) return;
                    const auto parsed = parseRect(reference, sheet.name_);
                    if (parsed && !hasSheet(sheets_, parsed->sheetName))
                        addIssue(report, ModelValidationSeverity::Error, "chart.source_sheet_missing", sheet.name_, seriesId,
                                 "Chart series references a worksheet that does not exist: " + parsed->sheetName);
                };
                validateSimpleChartSource(series.categoriesReference());
                validateSimpleChartSource(series.valuesReference());
                validateSimpleChartSource(series.titleReference());
            }
        }

        for (const auto& pivot : sheet.pivotTables_) {
            if (!pivotNames.insert(asciiFold(pivot.name())).second)
                addIssue(report, ModelValidationSeverity::Error, "pivot.duplicate_name", sheet.name_, pivot.name(),
                         "PivotTable names must be unique across the workbook (case-insensitive)");
            const auto& cache = pivot.cache();
            for (const auto& record : cache.records()) {
                if (record.size() != cache.fields().size()) {
                    addIssue(report, ModelValidationSeverity::Error, "pivot.record_width_mismatch", sheet.name_, pivot.name(),
                             "Pivot cache record width does not match cache field count");
                    break;
                }
            }
            if (cache.sourceData().find("#REF!") != std::string::npos) {
                addIssue(report, ModelValidationSeverity::Warning, "pivot.source_broken", sheet.name_, pivot.name(),
                         "Pivot cache worksheet source contains #REF!");
            } else if (!cache.sourceData().empty() && cache.sourceName().empty()) {
                const auto source = parseRect(cache.sourceData(), sheet.name_);
                if (!source) {
                    addIssue(report, ModelValidationSeverity::Error, "pivot.source_invalid", sheet.name_, pivot.name(),
                             "Pivot cache worksheet source is not a valid rectangular A1 range");
                } else if (!hasSheet(sheets_, source->sheetName)) {
                    addIssue(report, ModelValidationSeverity::Error, "pivot.source_sheet_missing", sheet.name_, pivot.name(),
                             "Pivot cache source worksheet does not exist: " + source->sheetName);
                } else if (!cache.fields().empty()) {
                    const auto width = source->last.column - source->first.column + 1;
                    std::size_t databaseFields = 0;
                    for (std::size_t i = 0; i < cache.fields().size(); ++i)
                        if (cache.fieldDatabaseField(i)) ++databaseFields;
                    if (databaseFields != width)
                        addIssue(report, ModelValidationSeverity::Error, "pivot.source_width_mismatch", sheet.name_, pivot.name(),
                                 "Pivot source width does not match database cache-field count");
                }
            }

            auto validateFieldVector = [&](const auto& fields, const char* code) {
                for (const auto& field : fields)
                    if (!validFieldIndex(field.fieldIndex(), cache.fields().size()))
                        addIssue(report, ModelValidationSeverity::Error, code, sheet.name_, pivot.name(),
                                 "Pivot field index is outside the cache field schema");
            };
            validateFieldVector(pivot.rowFields(), "pivot.row_field_index");
            validateFieldVector(pivot.columnFields(), "pivot.column_field_index");
            validateFieldVector(pivot.pageFields(), "pivot.page_field_index");
            validateFieldVector(pivot.dataFields(), "pivot.data_field_index");
            validateFieldVector(pivot.pageFieldSettings(), "pivot.page_setting_index");
            for (const auto& filter : pivot.filters())
                if (!validFieldIndex(filter.fieldIndex, cache.fields().size()))
                    addIssue(report, ModelValidationSeverity::Error, "pivot.filter_field_index", sheet.name_, pivot.name(),
                             "Pivot filter references a cache field outside the schema");
        }
    }

    std::unordered_set<std::string> definedNameScopes;
    for (const auto& name : definedNames_) {
        if (name.localSheetId() && *name.localSheetId() >= sheets_.size())
            addIssue(report, ModelValidationSeverity::Error, "defined_name.local_scope_invalid", {}, name.name(),
                     "Defined name localSheetId is outside the worksheet collection");
        const auto scope = name.localSheetId() ? std::to_string(*name.localSheetId()) : std::string("global");
        if (!definedNameScopes.insert(scope + "\n" + asciiFold(name.name())).second)
            addIssue(report, ModelValidationSeverity::Error, "defined_name.duplicate_in_scope", {}, name.name(),
                     "Defined-name identifiers must be unique case-insensitively within a scope");
        if (name.value().find("#REF!") != std::string::npos)
            addIssue(report, ModelValidationSeverity::Warning, "defined_name.broken_reference", {}, name.name(),
                     "Defined name contains #REF!");
    }

    const auto pivotLinks = validatePivotChartLinks();
    for (const auto& issue : pivotLinks.issues)
        addIssue(report, ModelValidationSeverity::Error, "pivot_chart.invalid_link", issue.worksheetName, issue.chartId, issue.message);

    const auto designers = validateVbaDesignerProject();
    for (const auto& issue : designers.issues)
        addIssue(report, ModelValidationSeverity::Error, "vba.designer_invalid", {}, issue.designerName, issue.message);

    return report;
}

} // namespace xlpp
