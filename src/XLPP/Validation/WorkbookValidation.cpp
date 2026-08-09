#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/WorksheetName.h>
#include <XLPP/Cell/CellReference.h>

#include <algorithm>
#include <cctype>
#include <optional>
#include <string>
#include <sstream>
#include <iomanip>
#include <unordered_map>
#include <unordered_set>

namespace xlpp {
namespace {
std::string asciiFold(std::string_view value) {
    std::string result(value);
    for (auto& ch : result) {
        const auto c = static_cast<unsigned char>(ch);
        if (c < 0x80) ch = static_cast<char>(std::tolower(c));
    }
    return result;
}

void addIssue(WorkbookValidationReport& report, WorkbookValidationSeverity severity,
              std::string code, std::string message, std::string worksheet = {}) {
    report.issues.push_back({severity, std::move(code), std::move(message), std::move(worksheet)});
    if (severity == WorkbookValidationSeverity::Error) ++report.errorCount;
    else ++report.warningCount;
}

bool validRangeReference(std::string_view reference) {
    if (reference.empty()) return false;
    const auto colon = reference.find(':');
    try {
        if (colon == std::string_view::npos) {
            (void)CellReference::parse(reference);
            return true;
        }
        if (reference.find(':', colon + 1) != std::string_view::npos) return false;
        (void)CellReference::parse(reference.substr(0, colon));
        (void)CellReference::parse(reference.substr(colon + 1));
        return true;
    } catch (...) { return false; }
}

struct QualifiedRange {
    std::string worksheet;
    std::string range;
    bool external{false};
};

std::optional<QualifiedRange> parseQualifiedRange(std::string_view expression,
                                                   std::string_view defaultWorksheet = {}) {
    if (!expression.empty() && expression.front() == '=') expression.remove_prefix(1);
    if (expression.empty()) return std::nullopt;

    bool quoted = false;
    std::size_t bang = std::string_view::npos;
    for (std::size_t i = 0; i < expression.size(); ++i) {
        if (expression[i] == '\'') {
            if (quoted && i + 1 < expression.size() && expression[i + 1] == '\'') {
                ++i;
                continue;
            }
            quoted = !quoted;
        } else if (expression[i] == '!' && !quoted) {
            bang = i;
        }
    }

    QualifiedRange result;
    std::string_view range = expression;
    if (bang != std::string_view::npos) {
        auto qualifier = expression.substr(0, bang);
        range = expression.substr(bang + 1);
        if (qualifier.find('[') != std::string_view::npos || qualifier.find(']') != std::string_view::npos) {
            result.external = true;
            return result;
        }
        if (qualifier.size() >= 2 && qualifier.front() == '\'' && qualifier.back() == '\'') {
            qualifier.remove_prefix(1);
            qualifier.remove_suffix(1);
            for (std::size_t i = 0; i < qualifier.size(); ++i) {
                if (qualifier[i] == '\'' && i + 1 < qualifier.size() && qualifier[i + 1] == '\'') ++i;
                result.worksheet.push_back(qualifier[i]);
            }
        } else {
            result.worksheet.assign(qualifier);
        }
    } else {
        result.worksheet.assign(defaultWorksheet);
    }
    result.range.assign(range);
    return result;
}

bool worksheetExists(const Workbook& workbook, std::string_view name) {
    return !name.empty() && workbook.worksheet(std::string(name)) != nullptr;
}

std::string pivotValidationCellText(const Cell* cell) {
    if (!cell || !cell->hasValue()) return {};
    const auto& value = cell->value();
    if (const auto* text = std::get_if<std::string>(&value)) return *text;
    if (const auto* number = std::get_if<double>(&value)) {
        std::ostringstream out;
        out << std::setprecision(15) << *number;
        return out.str();
    }
    if (const auto* boolean = std::get_if<bool>(&value)) return *boolean ? "true" : "false";
    if (const auto* error = std::get_if<CellError>(&value)) return toString(*error);
    if (const auto* date = std::get_if<DateTime>(&value)) return toIso8601(*date);
    return {};
}

std::vector<std::string> inferredPivotFields(const Worksheet& sourceSheet, std::string_view range) {
    const auto colon = range.find(':');
    auto first = CellReference::parse(colon == std::string_view::npos ? range : range.substr(0, colon));
    auto last = CellReference::parse(colon == std::string_view::npos ? range : range.substr(colon + 1));
    if (first.row > last.row) std::swap(first.row, last.row);
    if (first.column > last.column) std::swap(first.column, last.column);

    std::vector<std::string> fields;
    fields.reserve(last.column - first.column + 1);
    for (std::size_t column = first.column; column <= last.column; ++column) {
        auto name = pivotValidationCellText(sourceSheet.tryCell(first.row, column));
        if (name.empty()) name = "Field" + std::to_string(column - first.column + 1);
        if (std::find(fields.begin(), fields.end(), name) != fields.end())
            name += "_" + std::to_string(column - first.column + 1);
        fields.push_back(std::move(name));
    }
    return fields;
}

int resolvedPivotFieldIndex(const std::vector<std::string>& fields, int fieldIndex, std::string_view name) {
    if (fieldIndex >= 0) return fieldIndex;
    const auto it = std::find(fields.begin(), fields.end(), name);
    return it == fields.end() ? -1 : static_cast<int>(std::distance(fields.begin(), it));
}

std::size_t uniquePivotMemberCount(const PivotCache& cache, int fieldIndex) {
    if (fieldIndex < 0 || static_cast<std::size_t>(fieldIndex) >= cache.fields().size()) return 0;
    std::unordered_set<std::string> values;
    const auto index = static_cast<std::size_t>(fieldIndex);
    for (const auto& record : cache.records())
        if (index < record.size()) values.insert(record[index]);
    return values.size();
}

void validatePivotField(WorkbookValidationReport& report, const PivotTable& pivot,
                        const PivotField& field, std::string_view expectedAxis,
                        const std::vector<std::string>& effectiveFields,
                        const std::string& worksheet) {
    const auto resolvedIndex = resolvedPivotFieldIndex(effectiveFields, field.fieldIndex(), field.name());
    if (resolvedIndex < 0 || static_cast<std::size_t>(resolvedIndex) >= effectiveFields.size())
        addIssue(report, WorkbookValidationSeverity::Error, "PV010",
                 "Pivot field '" + field.name() + "' cannot be resolved to a cache/source field", worksheet);
    if (field.axis() != expectedAxis)
        addIssue(report, WorkbookValidationSeverity::Error, "PV011",
                 "Pivot field '" + field.name() + "' has an axis inconsistent with its collection", worksheet);
    if (field.sortType() < 0 || field.sortType() > 2)
        addIssue(report, WorkbookValidationSeverity::Error, "PV012",
                 "Pivot field '" + field.name() + "' has an invalid sort type", worksheet);

    // Hidden/selected item indexes can be checked precisely only when the caller
    // supplied concrete cache records. Auto-cache records are materialized from
    // the source range later by effectivePivotTable().
    if (!pivot.cache().records().empty() && resolvedIndex >= 0) {
        const auto memberCount = uniquePivotMemberCount(pivot.cache(), resolvedIndex);
        for (const auto index : field.hiddenItemIndexes()) {
            if (index < 0 || static_cast<std::size_t>(index) >= memberCount)
                addIssue(report, WorkbookValidationSeverity::Error, "PV013",
                         "Pivot field '" + field.name() + "' hides an item outside the cache member set", worksheet);
        }
        if (field.selectedItemIndex() >= 0 && static_cast<std::size_t>(field.selectedItemIndex()) >= memberCount)
            addIssue(report, WorkbookValidationSeverity::Error, "PV014",
                     "Pivot field '" + field.name() + "' selects an item outside the cache member set", worksheet);
    }

    const auto& grouping = field.grouping();
    if (grouping.kind == PivotGrouping::Kind::Numeric) {
        if (!(grouping.interval > 0.0))
            addIssue(report, WorkbookValidationSeverity::Error, "PV015",
                     "Numeric pivot grouping interval must be greater than zero", worksheet);
        if (!grouping.autoStart && !grouping.autoEnd && grouping.start > grouping.end)
            addIssue(report, WorkbookValidationSeverity::Error, "PV016",
                     "Numeric pivot grouping start cannot exceed end", worksheet);
    }
}

void validateSeriesReference(WorkbookValidationReport& report, const Workbook& workbook,
                             const std::string& owningWorksheet, const std::string& reference,
                             std::string_view label) {
    if (reference.empty()) return;
    const auto parsed = parseQualifiedRange(reference, owningWorksheet);
    if (!parsed) {
        addIssue(report, WorkbookValidationSeverity::Error, "CH010",
                 "Chart " + std::string(label) + " reference is malformed: " + reference, owningWorksheet);
        return;
    }
    if (parsed->external) return; // preserved external references are intentionally allowed.
    if (!worksheetExists(workbook, parsed->worksheet))
        addIssue(report, WorkbookValidationSeverity::Error, "CH011",
                 "Chart " + std::string(label) + " reference targets a missing worksheet: " + parsed->worksheet,
                 owningWorksheet);
    if (!validRangeReference(parsed->range))
        addIssue(report, WorkbookValidationSeverity::Error, "CH012",
                 "Chart " + std::string(label) + " reference has an invalid A1 range: " + parsed->range,
                 owningWorksheet);
}
} // namespace

WorkbookValidationReport Workbook::validate(const WorkbookValidationOptions& options) const {
    WorkbookValidationReport report;
    if (sheets_.empty())
        addIssue(report, WorkbookValidationSeverity::Error, "WB001", "Workbook must contain at least one worksheet");

    if (options.validateWorksheetNames) {
        std::unordered_map<std::string, std::string> seen;
        for (const auto& sheet : sheets_) {
            if (!isValidWorksheetName(sheet.name()))
                addIssue(report, WorkbookValidationSeverity::Error, "WS001", "Invalid worksheet name", sheet.name());
            const auto key = asciiFold(sheet.name());
            if (const auto [it, inserted] = seen.emplace(key, sheet.name()); !inserted)
                addIssue(report, WorkbookValidationSeverity::Error, "WS002",
                         "Worksheet name conflicts case-insensitively with '" + it->second + "'", sheet.name());
        }
    }

    if (options.validateVba && hasVbaProject()) {
        std::unordered_set<std::string> codeNames;
        for (const auto& sheet : sheets_) {
            if (sheet.vbaCodeName().empty()) {
                addIssue(report, WorkbookValidationSeverity::Error, "VB001",
                         "Macro-enabled worksheet is missing a VBA code name", sheet.name());
                continue;
            }
            const auto key = asciiFold(sheet.vbaCodeName());
            if (key == "thisworkbook")
                addIssue(report, WorkbookValidationSeverity::Error, "VB002",
                         "Worksheet VBA code name conflicts with ThisWorkbook", sheet.name());
            if (!codeNames.insert(key).second)
                addIssue(report, WorkbookValidationSeverity::Error, "VB003",
                         "Worksheet VBA code names must be unique: " + sheet.vbaCodeName(), sheet.name());
        }
        if (generatedVbaProject_) {
            try {
                const auto modules = vbaModules();
                std::unordered_set<std::string> moduleNames;
                bool workbookDocumentFound = false;
                for (const auto& module : modules) {
                    const auto key = asciiFold(module.name);
                    if (!moduleNames.insert(key).second)
                        addIssue(report, WorkbookValidationSeverity::Error, "VB004",
                                 "Duplicate VBA module name: " + module.name);
                    if (module.type == VbaModuleType::Document && key == "thisworkbook")
                        workbookDocumentFound = true;
                }
                if (!workbookDocumentFound)
                    addIssue(report, WorkbookValidationSeverity::Error, "VB005",
                             "Generated VBA project is missing the ThisWorkbook document module");
            } catch (const std::exception& error) {
                addIssue(report, WorkbookValidationSeverity::Error, "VB006",
                         std::string("Generated VBA project cannot be parsed: ") + error.what());
            }
        }
    }

    if (options.validateDefinedNames) {
        std::unordered_set<std::string> names;
        for (const auto& name : definedNames_) {
            if (name.name().empty()) {
                addIssue(report, WorkbookValidationSeverity::Error, "DN001", "Defined name cannot be empty");
                continue;
            }
            if (name.localSheetId() && *name.localSheetId() >= sheets_.size())
                addIssue(report, WorkbookValidationSeverity::Error, "DN002",
                         "Defined name '" + name.name() + "' refers to a missing local worksheet scope");
            const auto scope = name.localSheetId() ? std::to_string(*name.localSheetId()) : std::string{"global"};
            const auto key = scope + "\n" + asciiFold(name.name());
            if (!names.insert(key).second)
                addIssue(report, WorkbookValidationSeverity::Error, "DN003",
                         "Duplicate defined name in the same scope: " + name.name());
        }
    }

    if (options.validateTables || options.validatePivots || options.validateCharts) {
        std::unordered_map<std::string, std::string> tableNames;
        std::unordered_map<std::string, std::string> pivotNames;
        for (const auto& sheet : sheets_) {
            if (options.validateTables) {
                for (const auto& table : static_cast<const Worksheet&>(sheet).tables()) {
                    if (table.name().empty())
                        addIssue(report, WorkbookValidationSeverity::Error, "TB001", "Table name cannot be empty", sheet.name());
                    else {
                        const auto key = asciiFold(table.name());
                        if (const auto [it, inserted] = tableNames.emplace(key, sheet.name()); !inserted)
                            addIssue(report, WorkbookValidationSeverity::Error, "TB002",
                                     "Table name must be unique across the workbook: " + table.name(), sheet.name());
                    }
                    if (!validRangeReference(table.reference()))
                        addIssue(report, WorkbookValidationSeverity::Error, "TB003",
                                 "Table has an invalid worksheet range: " + table.reference(), sheet.name());
                }
            }

            if (options.validatePivots) {
                for (const auto& pivot : static_cast<const Worksheet&>(sheet).pivotTables()) {
                    if (pivot.name().empty())
                        addIssue(report, WorkbookValidationSeverity::Error, "PV001", "Pivot table name cannot be empty", sheet.name());
                    else {
                        const auto key = asciiFold(pivot.name());
                        if (const auto [it, inserted] = pivotNames.emplace(key, sheet.name()); !inserted)
                            addIssue(report, WorkbookValidationSeverity::Error, "PV002",
                                     "Pivot table name must be unique across the workbook: " + pivot.name(), sheet.name());
                    }
                    // Empty location is valid: the writer uses Excel-compatible D2 as
                    // the default anchor and expands it to the occupied view rectangle.
                    if (!pivot.location().empty() && !validRangeReference(pivot.location()))
                        addIssue(report, WorkbookValidationSeverity::Error, "PV003",
                                 "Pivot table has an invalid output location: " + pivot.location(), sheet.name());

                    const auto& cache = pivot.cache();
                    for (const auto& record : cache.records())
                        if (record.size() != cache.fields().size()) {
                            addIssue(report, WorkbookValidationSeverity::Error, "PV005",
                                     "Pivot cache record width does not match cache field count", sheet.name());
                            break;
                        }

                    // The serializer intentionally supports lazy/automatic cache
                    // materialization. If sourceData is omitted it uses the owner
                    // worksheet's current dimensions; if fields/records are omitted it
                    // derives them from the source range. Validate that effective state
                    // instead of rejecting the pre-materialized model.
                    const auto sourceExpression = cache.sourceData().empty()
                        ? (sheet.name() + "!" + sheet.dimensions())
                        : cache.sourceData();
                    const auto source = parseQualifiedRange(sourceExpression, sheet.name());
                    std::vector<std::string> effectiveFields = cache.fields();
                    if (!source)
                        addIssue(report, WorkbookValidationSeverity::Error, "PV006",
                                 "Pivot cache source is malformed", sheet.name());
                    else if (!source->external) {
                        const auto* sourceSheet = worksheet(source->worksheet);
                        if (!sourceSheet)
                            addIssue(report, WorkbookValidationSeverity::Error, "PV007",
                                     "Pivot cache source targets a missing worksheet: " + source->worksheet, sheet.name());
                        const bool rangeValid = validRangeReference(source->range);
                        if (!rangeValid)
                            addIssue(report, WorkbookValidationSeverity::Error, "PV008",
                                     "Pivot cache source has an invalid A1 range: " + source->range, sheet.name());
                        else if (effectiveFields.empty() && sourceSheet) {
                            try { effectiveFields = inferredPivotFields(*sourceSheet, source->range); }
                            catch (...) {
                                addIssue(report, WorkbookValidationSeverity::Error, "PV008",
                                         "Pivot cache source cannot be materialized: " + source->range, sheet.name());
                            }
                        }
                    }
                    if (effectiveFields.empty())
                        addIssue(report, WorkbookValidationSeverity::Error, "PV004",
                                 "Pivot cache/source must resolve to at least one field", sheet.name());

                    for (const auto& field : pivot.rowFields()) validatePivotField(report, pivot, field, "axisRow", effectiveFields, sheet.name());
                    for (const auto& field : pivot.columnFields()) validatePivotField(report, pivot, field, "axisCol", effectiveFields, sheet.name());
                    for (const auto& field : pivot.pageFields()) validatePivotField(report, pivot, field, "axisPage", effectiveFields, sheet.name());
                    for (const auto& dataField : pivot.dataFields()) {
                        const auto index = resolvedPivotFieldIndex(effectiveFields, dataField.fieldIndex(), dataField.name());
                        if (index < 0 || static_cast<std::size_t>(index) >= effectiveFields.size())
                            addIssue(report, WorkbookValidationSeverity::Error, "PV020",
                                     "Pivot data field '" + dataField.name() + "' cannot be resolved to a cache/source field", sheet.name());
                        if (dataField.baseField() < 0 ||
                            (!effectiveFields.empty() && static_cast<std::size_t>(dataField.baseField()) >= effectiveFields.size()))
                            addIssue(report, WorkbookValidationSeverity::Error, "PV021",
                                     "Pivot data field has an invalid base field", sheet.name());
                    }
                    for (const auto& filter : pivot.filters()) {
                        if (filter.fieldIndex < 0 || static_cast<std::size_t>(filter.fieldIndex) >= effectiveFields.size())
                            addIssue(report, WorkbookValidationSeverity::Error, "PV022",
                                     "Pivot filter has an invalid cache/source field index", sheet.name());
                        if (filter.measureFieldIndex >= 0 &&
                            static_cast<std::size_t>(filter.measureFieldIndex) >= pivot.dataFields().size())
                            addIssue(report, WorkbookValidationSeverity::Error, "PV023",
                                     "Pivot filter has an invalid measure field index", sheet.name());
                        if ((filter.type == "count" || filter.type == "percent" || filter.type == "sum") &&
                            !(filter.top10Value > 0.0))
                            addIssue(report, WorkbookValidationSeverity::Error, "PV024",
                                     "Pivot Top10 filter value must be greater than zero", sheet.name());
                    }
                }
            }

            if (options.validateCharts) {
                for (const auto& chart : static_cast<const Worksheet&>(sheet).charts()) {
                    if (chart.width() <= 0 || chart.height() <= 0)
                        addIssue(report, WorkbookValidationSeverity::Error, "CH001",
                                 "Chart dimensions must be positive", sheet.name());
                    if (chart.series().empty())
                        addIssue(report, WorkbookValidationSeverity::Warning, "CH002",
                                 "Chart has no data series", sheet.name());
                    for (const auto& plot : chart.plots()) {
                        if (plot.firstSeries > chart.series().size() ||
                            plot.seriesCount > chart.series().size() - std::min(plot.firstSeries, chart.series().size()))
                            addIssue(report, WorkbookValidationSeverity::Error, "CH003",
                                     "Chart plot series span is outside the chart series collection", sheet.name());
                        if (plot.histogramBinWidth < 0.0 || plot.histogramBinCount < 0)
                            addIssue(report, WorkbookValidationSeverity::Error, "CH004",
                                     "Histogram bin width/count cannot be negative", sheet.name());
                        if (plot.bubbleScale < 0 || plot.bubbleScale > 300)
                            addIssue(report, WorkbookValidationSeverity::Error, "CH005",
                                     "Bubble scale must be in the Excel range 0..300", sheet.name());
                    }
                    for (const auto& series : chart.series()) {
                        validateSeriesReference(report, *this, sheet.name(), series.titleReference(), "title");
                        validateSeriesReference(report, *this, sheet.name(), series.categoriesReference(), "category/X");
                        validateSeriesReference(report, *this, sheet.name(), series.valuesReference(), "value/Y");
                        validateSeriesReference(report, *this, sheet.name(), series.bubbleSizeReference(), "bubble-size");
                        for (const auto& bars : series.errorBars()) {
                            validateSeriesReference(report, *this, sheet.name(), bars.plusReference, "error-bar plus");
                            validateSeriesReference(report, *this, sheet.name(), bars.minusReference, "error-bar minus");
                        }
                    }
                }
            }
        }
    }
    return report;
}

} // namespace xlpp
