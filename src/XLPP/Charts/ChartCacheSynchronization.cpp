#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Cell/CellReference.h>
#include <algorithm>
#include <iomanip>
#include <sstream>
#include <stdexcept>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

namespace {
struct ParsedChartReference {
    const xlpp::Worksheet* sheet{nullptr};
    xlpp::CellReference first{};
    xlpp::CellReference last{};
    std::string normalized;
};

std::string trimChartReference(std::string value) {
    const auto first = value.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) return {};
    const auto last = value.find_last_not_of(" \t\r\n");
    return value.substr(first, last - first + 1);
}

bool parseChartReference(const xlpp::Workbook& workbook, const xlpp::Worksheet& owner,
                         std::string reference, ParsedChartReference& parsed, std::string& reason) {
    reference = trimChartReference(std::move(reference));
    if (!reference.empty() && reference.front() == '=') reference.erase(reference.begin());
    if (reference.empty()) { reason = "empty reference"; return false; }
    if (reference.find('[') != std::string::npos || reference.find(']') != std::string::npos) {
        reason = "external workbook references are not synchronized"; return false;
    }
    std::size_t bang = std::string::npos;
    bool quoted = false;
    for (std::size_t i = 0; i < reference.size(); ++i) {
        if (reference[i] == '\'') {
            if (quoted && i + 1 < reference.size() && reference[i + 1] == '\'') { ++i; continue; }
            quoted = !quoted; continue;
        }
        if (!quoted && reference[i] == '!') { bang = i; break; }
    }
    if (quoted) { reason = "unterminated worksheet quote"; return false; }
    std::string sheetName = owner.name();
    std::string range = reference;
    if (bang != std::string::npos) {
        auto token = trimChartReference(reference.substr(0, bang));
        range = trimChartReference(reference.substr(bang + 1));
        if (token.size() >= 2 && token.front() == '\'' && token.back() == '\'') {
            std::string unquoted;
            for (std::size_t i = 1; i + 1 < token.size(); ++i) {
                if (token[i] == '\'' && i + 1 < token.size() - 1 && token[i + 1] == '\'') { unquoted.push_back('\''); ++i; }
                else unquoted.push_back(token[i]);
            }
            sheetName = std::move(unquoted);
        } else sheetName = std::move(token);
    }
    if (range.find(',') != std::string::npos || range.find(';') != std::string::npos) {
        reason = "union references are not synchronized"; return false;
    }
    const auto* source = workbook.worksheet(sheetName);
    if (!source) { reason = "worksheet not found: " + sheetName; return false; }
    try {
        const auto colon = range.find(':');
        if (colon != std::string::npos && range.find(':', colon + 1) != std::string::npos) {
            reason = "invalid range reference"; return false;
        }
        auto first = xlpp::CellReference::parse(colon == std::string::npos ? range : range.substr(0, colon));
        auto last = xlpp::CellReference::parse(colon == std::string::npos ? range : range.substr(colon + 1));
        if (first.row > last.row) std::swap(first.row, last.row);
        if (first.column > last.column) std::swap(first.column, last.column);
        if (first.row != last.row && first.column != last.column) {
            reason = "two-dimensional ranges are not synchronized"; return false;
        }
        parsed.sheet = source; parsed.first = first; parsed.last = last; parsed.normalized = reference; return true;
    } catch (const std::exception& ex) {
        reason = ex.what(); return false;
    }
}


void appendChartDependencyField(std::string& out, std::string_view value) {
    out += std::to_string(value.size());
    out.push_back(':');
    out.append(value.data(), value.size());
    out.push_back('|');
}

std::string chartDependencySnapshot(const ParsedChartReference& ref, bool date1904) {
    std::string snapshot;
    snapshot.reserve(128);
    appendChartDependencyField(snapshot, ref.sheet ? ref.sheet->name() : std::string_view{});
    appendChartDependencyField(snapshot, ref.normalized);
    appendChartDependencyField(snapshot, date1904 ? "1904" : "1900");
    const auto appendCell = [&](std::size_t row, std::size_t column) {
        snapshot += std::to_string(row); snapshot.push_back(',');
        snapshot += std::to_string(column); snapshot.push_back('=');
        const auto* cell = ref.sheet ? ref.sheet->tryCell(row, column) : nullptr;
        if (!cell) { snapshot += "M|"; return; }
        appendChartDependencyField(snapshot, cell->numberFormat());
        appendChartDependencyField(snapshot, cell->formula());
        const auto& value = cell->value();
        if (std::holds_alternative<std::monostate>(value)) snapshot += "E|";
        else if (const auto* number = std::get_if<double>(&value)) {
            snapshot += "N:";
            std::ostringstream text; text << std::setprecision(17) << *number;
            appendChartDependencyField(snapshot, text.str());
        } else if (const auto* boolean = std::get_if<bool>(&value)) {
            snapshot += *boolean ? "B:1|" : "B:0|";
        } else if (const auto* string = std::get_if<std::string>(&value)) {
            snapshot += "S:"; appendChartDependencyField(snapshot, *string);
        } else if (const auto* error = std::get_if<xlpp::CellError>(&value)) {
            snapshot += "X:"; appendChartDependencyField(snapshot, xlpp::toString(*error));
        } else if (const auto* date = std::get_if<xlpp::DateTime>(&value)) {
            snapshot += "D:";
            std::ostringstream text; text << std::setprecision(17) << xlpp::toExcelSerial(*date, date1904);
            appendChartDependencyField(snapshot, text.str());
        }
    };
    if (ref.first.row == ref.last.row) {
        for (std::size_t column = ref.first.column; column <= ref.last.column; ++column)
            appendCell(ref.first.row, column);
    } else {
        for (std::size_t row = ref.first.row; row <= ref.last.row; ++row)
            appendCell(row, ref.first.column);
    }
    return snapshot;
}

std::string chartDependencyKey(const xlpp::Worksheet& owner, const xlpp::Chart& chart,
                               std::size_t chartIndex, std::size_t seriesIndex,
                               const char* label, const ParsedChartReference& parsed) {
    std::string key;
    key.reserve(128);
    appendChartDependencyField(key, owner.name());
    appendChartDependencyField(key, chart.stableId().empty() ? std::string("#generated-") + std::to_string(chartIndex) : chart.stableId());
    appendChartDependencyField(key, std::to_string(seriesIndex));
    appendChartDependencyField(key, label);
    appendChartDependencyField(key, parsed.normalized);
    return key;
}

std::string chartCacheNumber(double value) {
    if (value == 0.0) value = 0.0; // normalize negative zero
    std::ostringstream out; out << std::setprecision(15) << value; return out.str();
}

enum class ChartCacheKind { String, Numeric, Automatic };

xlpp::ChartSeriesCache buildChartCache(const ParsedChartReference& ref, ChartCacheKind requested,
                                       bool date1904, const xlpp::ChartSeriesCache& existing,
                                       std::vector<std::string>* warnings = nullptr) {
    struct SourceValue { std::size_t index; const xlpp::Cell* cell; };
    std::vector<SourceValue> cells;
    if (ref.first.row == ref.last.row) {
        cells.reserve(ref.last.column - ref.first.column + 1);
        for (std::size_t col = ref.first.column, index = 0; col <= ref.last.column; ++col, ++index)
            cells.push_back({index, ref.sheet->tryCell(ref.first.row, col)});
    } else {
        cells.reserve(ref.last.row - ref.first.row + 1);
        for (std::size_t row = ref.first.row, index = 0; row <= ref.last.row; ++row, ++index)
            cells.push_back({index, ref.sheet->tryCell(row, ref.first.column)});
    }
    bool numeric = requested == ChartCacheKind::Numeric;
    if (requested == ChartCacheKind::Automatic) {
        numeric = true;
        bool sawValue = false;
        for (const auto& source : cells) {
            if (!source.cell || !source.cell->hasValue()) continue;
            sawValue = true;
            const auto& value = source.cell->value();
            if (!(std::holds_alternative<double>(value) || std::holds_alternative<xlpp::DateTime>(value) || std::holds_alternative<bool>(value))) {
                numeric = false; break;
            }
        }
        if (!sawValue) numeric = existing.present ? existing.numeric : false;
    }
    xlpp::ChartSeriesCache cache; cache.present = true; cache.numeric = numeric; cache.pointCount = cells.size();
    if (numeric) {
        cache.formatCode = existing.numeric && !existing.formatCode.empty() ? existing.formatCode : "General";
        if (cache.formatCode == "General") {
            for (const auto& source : cells) if (source.cell && source.cell->hasValue() && source.cell->numberFormat() != "General") {
                cache.formatCode = source.cell->numberFormat(); break;
            }
        }
    }
    for (const auto& source : cells) {
        if (!source.cell || !source.cell->hasValue()) continue; // sparse cache: preserve index, omit blank point
        const auto& value = source.cell->value();
        std::string text;
        if (numeric) {
            if (const auto* number = std::get_if<double>(&value)) text = chartCacheNumber(*number);
            else if (const auto* date = std::get_if<xlpp::DateTime>(&value)) text = chartCacheNumber(xlpp::toExcelSerial(*date, date1904));
            else if (const auto* boolean = std::get_if<bool>(&value)) text = *boolean ? "1" : "0";
            else {
                if (warnings) warnings->push_back("Skipped non-numeric cache point at " + source.cell->address() + " in " + ref.sheet->name());
                continue;
            }
        } else {
            if (const auto* string = std::get_if<std::string>(&value)) text = *string;
            else if (const auto* number = std::get_if<double>(&value)) text = chartCacheNumber(*number);
            else if (const auto* date = std::get_if<xlpp::DateTime>(&value)) text = chartCacheNumber(xlpp::toExcelSerial(*date, date1904));
            else if (const auto* boolean = std::get_if<bool>(&value)) text = *boolean ? "TRUE" : "FALSE";
            else if (const auto* error = std::get_if<xlpp::CellError>(&value)) text = xlpp::toString(*error);
        }
        cache.points.push_back({source.index, std::move(text)});
    }
    return cache;
}

} // namespace

namespace xlpp {
ChartCacheSyncReport Workbook::synchronizeChartCaches(const ChartCacheSyncOptions& options) {
    ChartCacheSyncReport report;
    const auto cacheEqual = [](const ChartSeriesCache& a, const ChartSeriesCache& b) {
        if (a.present != b.present || a.numeric != b.numeric || a.formatCode != b.formatCode || a.pointCount != b.pointCount || a.points.size() != b.points.size()) return false;
        for (std::size_t i = 0; i < a.points.size(); ++i)
            if (a.points[i].index != b.points[i].index || a.points[i].value != b.points[i].value) return false;
        return true;
    };
    for (auto& sheet : sheets_) {
        for (std::size_t chartIndex = 0; chartIndex < sheet.charts_.size(); ++chartIndex) {
            auto& chart = sheet.charts_[chartIndex];
            ++report.chartsVisited;
            for (std::size_t seriesIndex = 0; seriesIndex < chart.series().size(); ++seriesIndex) {
                ++report.seriesVisited;
                auto& series = chart.series()[seriesIndex];
                bool unsupportedSelectedReference = false;
                auto synchronize = [&](const std::string& reference, ChartCacheKind kind, const ChartSeriesCache& existing,
                                       const char* label, bool enabled) {
                    if (!enabled || reference.empty()) return;
                    ParsedChartReference parsed; std::string reason;
                    if (!parseChartReference(*this, sheet, reference, parsed, reason)) {
                        ++report.referencesSkipped; unsupportedSelectedReference = true;
                        report.warnings.push_back(sheet.name() + ": chart " + (chart.stableId().empty() ? std::string("<generated>") : chart.stableId()) +
                                                  ", series " + std::to_string(seriesIndex) + ", " + label + " reference '" + reference + "': " + reason);
                        return;
                    }
                    ++report.referencesChecked;
                    if (kind == ChartCacheKind::String && (parsed.first.row != parsed.last.row || parsed.first.column != parsed.last.column)) {
                        ++report.referencesSkipped; unsupportedSelectedReference = true;
                        report.warnings.push_back(sheet.name() + ": chart series title reference must resolve to one cell: " + reference);
                        return;
                    }

                    const auto dependencyKey = chartDependencyKey(sheet, chart, chartIndex, seriesIndex, label, parsed);
                    const auto snapshot = chartDependencySnapshot(parsed, date1904_);
                    const auto tracked = chartCacheDependencySnapshots_.find(dependencyKey);
                    const bool newDependency = tracked == chartCacheDependencySnapshots_.end();
                    if (!newDependency && tracked->second == snapshot) {
                        ++report.referencesUnchanged;
                        if (options.changedReferencesOnly) return;
                    } else if (!newDependency) {
                        ++report.dependenciesChanged;
                    }

                    auto rebuilt = buildChartCache(parsed, kind, date1904_, existing, &report.warnings);
                    if (!rebuilt.valid(true)) {
                        ++report.referencesSkipped; unsupportedSelectedReference = true;
                        report.warnings.push_back(sheet.name() + ": rebuilt " + label + " cache failed validation for " + reference);
                        return;
                    }
                    if (cacheEqual(rebuilt, existing)) {
                        chartCacheDependencySnapshots_[dependencyKey] = snapshot;
                        if (newDependency) ++report.dependenciesRegistered;
                        return;
                    }
                    bool accepted = true;
                    if (chart.imported()) {
                        if (kind == ChartCacheKind::String) accepted = sheet.setChartSeriesTitleCache(chart.stableId(), seriesIndex, rebuilt);
                        else if (kind == ChartCacheKind::Numeric) accepted = sheet.setChartSeriesValueCache(chart.stableId(), seriesIndex, rebuilt);
                        else accepted = sheet.setChartSeriesCategoryCache(chart.stableId(), seriesIndex, rebuilt);
                    } else {
                        if (kind == ChartCacheKind::String) series.setTitleCache(rebuilt);
                        else if (kind == ChartCacheKind::Numeric) series.setValuesCache(rebuilt);
                        else series.setCategoriesCache(rebuilt);
                        sheet.dirty_ = true;
                        sheet.drawingAppendDirty_ = true;
                    }
                    if (accepted) {
                        ++report.cachesUpdated;
                        chartCacheDependencySnapshots_[dependencyKey] = snapshot;
                        if (newDependency) ++report.dependenciesRegistered;
                    } else {
                        report.warnings.push_back(sheet.name() + ": failed to apply rebuilt " + label + " cache for series " + std::to_string(seriesIndex));
                    }
                };
                synchronize(series.titleReference(), ChartCacheKind::String, series.titleCache(), "title", options.synchronizeTitles);
                synchronize(series.categoriesReference(), ChartCacheKind::Automatic, series.categoriesCache(), "category", options.synchronizeCategories);
                synchronize(series.valuesReference(), ChartCacheKind::Numeric, series.valuesCache(), "value", options.synchronizeValues);
                if (options.clearUnsupportedReferences && unsupportedSelectedReference) {
                    if (chart.imported()) {
                        if (sheet.clearChartSeriesCaches(chart.stableId(), seriesIndex)) ++report.cachesCleared;
                    } else {
                        const bool hadAny = series.titleCache().present || series.categoriesCache().present || series.valuesCache().present;
                        series.setTitleCache({}); series.setCategoriesCache({}); series.setValuesCache({});
                        if (hadAny) ++report.cachesCleared;
                        sheet.dirty_ = true; sheet.drawingAppendDirty_ = true;
                    }
                }
            }
        }
    }
    return report;
}


} // namespace xlpp
