#include "WorkbookSheetWriter.h"
#include "WorkbookDrawingIO.h"
#include "WorkbookNamespaces.h"
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Cell/Cell.h>
#include <XLPP/Cell/CellReference.h>
#include <XLPP/Cell/DateTime.h>
#include <XLPP/Pivot/PivotTable.h>
#include "../XML/XmlUtilities.h"
#include "../XML/NumericParsing.h"
#include <algorithm>
#include <cmath>
#include <limits>
#include <optional>
#include <sstream>
#include <stdexcept>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>
#include <cstdint>

namespace xlpp {
namespace internal {
std::string filterOperatorName(xlpp::FilterOperator op) {
    switch (op) {
    case xlpp::FilterOperator::NotEqual: return "notEqual";
    case xlpp::FilterOperator::LessThan: return "lessThan";
    case xlpp::FilterOperator::LessThanOrEqual: return "lessThanOrEqual";
    case xlpp::FilterOperator::GreaterThan: return "greaterThan";
    case xlpp::FilterOperator::GreaterThanOrEqual: return "greaterThanOrEqual";
    default: return "equal";
    }
}

xlpp::FilterOperator parseFilterOperator(const std::string& value) {
    if (value == "notEqual") return xlpp::FilterOperator::NotEqual;
    if (value == "lessThan") return xlpp::FilterOperator::LessThan;
    if (value == "lessThanOrEqual") return xlpp::FilterOperator::LessThanOrEqual;
    if (value == "greaterThan") return xlpp::FilterOperator::GreaterThan;
    if (value == "greaterThanOrEqual") return xlpp::FilterOperator::GreaterThanOrEqual;
    return xlpp::FilterOperator::Equal;
}

std::string conditionalOperatorName(xlpp::ConditionalOperator op) {
    switch (op) {
    case xlpp::ConditionalOperator::NotEqual: return "notEqual";
    case xlpp::ConditionalOperator::LessThan: return "lessThan";
    case xlpp::ConditionalOperator::LessThanOrEqual: return "lessThanOrEqual";
    case xlpp::ConditionalOperator::GreaterThan: return "greaterThan";
    case xlpp::ConditionalOperator::GreaterThanOrEqual: return "greaterThanOrEqual";
    case xlpp::ConditionalOperator::Between: return "between";
    case xlpp::ConditionalOperator::NotBetween: return "notBetween";
    default: return "equal";
    }
}

xlpp::ConditionalOperator parseConditionalOperator(const std::string& value) {
    if (value == "notEqual") return xlpp::ConditionalOperator::NotEqual;
    if (value == "lessThan") return xlpp::ConditionalOperator::LessThan;
    if (value == "lessThanOrEqual") return xlpp::ConditionalOperator::LessThanOrEqual;
    if (value == "greaterThan") return xlpp::ConditionalOperator::GreaterThan;
    if (value == "greaterThanOrEqual") return xlpp::ConditionalOperator::GreaterThanOrEqual;
    if (value == "between") return xlpp::ConditionalOperator::Between;
    if (value == "notBetween") return xlpp::ConditionalOperator::NotBetween;
    return xlpp::ConditionalOperator::Equal;
}

// Write a <cfvo> element for data bars / color scales / icon sets.
// Formula thresholds are represented by the cfvo val attribute; a sibling
// <f> element is not part of CT_Cfvo and is rejected by strict consumers.
void writeCfvo(std::ostringstream& xml, const xlpp::Cfvo& cfvo) {
    xml << "<cfvo type=\"" << xmlEscape(cfvo.type) << "\"";
    if (!cfvo.formula.empty())
        xml << " val=\"" << xmlEscape(cfvo.formula) << "\"";
    else if (cfvo.hasValue)
        xml << " val=\"" << cfvo.value << "\"";
    xml << "/>";
}

xlpp::Cfvo parseCfvo(const std::string& tag) {
    xlpp::Cfvo result;
    result.type = xlpp::internal::attribute(tag, "type");
    const auto value = xlpp::internal::attribute(tag, "val");
    if (!value.empty()) {
        if (result.type == "formula") {
            result.formula = value;
            result.hasValue = true;
        } else {
            if (!xlpp::internal::tryParseDoubleExact(value, result.value)) return result;
            result.hasValue = true;
        }
    }
    return result;
}


std::string dataValidationTypeName(xlpp::DataValidationType type) {
    switch (type) {
    case xlpp::DataValidationType::Whole: return "whole";
    case xlpp::DataValidationType::Decimal: return "decimal";
    case xlpp::DataValidationType::List: return "list";
    case xlpp::DataValidationType::Date: return "date";
    case xlpp::DataValidationType::Time: return "time";
    case xlpp::DataValidationType::TextLength: return "textLength";
    case xlpp::DataValidationType::Custom: return "custom";
    default: return "none";
    }
}
xlpp::DataValidationType parseDataValidationType(const std::string& value) {
    if (value == "whole") return xlpp::DataValidationType::Whole;
    if (value == "decimal") return xlpp::DataValidationType::Decimal;
    if (value == "list") return xlpp::DataValidationType::List;
    if (value == "date") return xlpp::DataValidationType::Date;
    if (value == "time") return xlpp::DataValidationType::Time;
    if (value == "textLength") return xlpp::DataValidationType::TextLength;
    if (value == "custom") return xlpp::DataValidationType::Custom;
    return xlpp::DataValidationType::None;
}
std::string dataValidationOperatorName(xlpp::DataValidationOperator op) {
    switch (op) {
    case xlpp::DataValidationOperator::NotBetween: return "notBetween";
    case xlpp::DataValidationOperator::Equal: return "equal";
    case xlpp::DataValidationOperator::NotEqual: return "notEqual";
    case xlpp::DataValidationOperator::LessThan: return "lessThan";
    case xlpp::DataValidationOperator::LessThanOrEqual: return "lessThanOrEqual";
    case xlpp::DataValidationOperator::GreaterThan: return "greaterThan";
    case xlpp::DataValidationOperator::GreaterThanOrEqual: return "greaterThanOrEqual";
    default: return "between";
    }
}
xlpp::DataValidationOperator parseDataValidationOperator(const std::string& value) {
    if (value == "notBetween") return xlpp::DataValidationOperator::NotBetween;
    if (value == "equal") return xlpp::DataValidationOperator::Equal;
    if (value == "notEqual") return xlpp::DataValidationOperator::NotEqual;
    if (value == "lessThan") return xlpp::DataValidationOperator::LessThan;
    if (value == "lessThanOrEqual") return xlpp::DataValidationOperator::LessThanOrEqual;
    if (value == "greaterThan") return xlpp::DataValidationOperator::GreaterThan;
    if (value == "greaterThanOrEqual") return xlpp::DataValidationOperator::GreaterThanOrEqual;
    return xlpp::DataValidationOperator::Between;
}
std::string dataValidationErrorStyleName(xlpp::DataValidationErrorStyle style) {
    switch (style) {
    case xlpp::DataValidationErrorStyle::Warning: return "warning";
    case xlpp::DataValidationErrorStyle::Information: return "information";
    default: return "stop";
    }
}
xlpp::DataValidationErrorStyle parseDataValidationErrorStyle(const std::string& value) {
    if (value == "warning") return xlpp::DataValidationErrorStyle::Warning;
    if (value == "information") return xlpp::DataValidationErrorStyle::Information;
    return xlpp::DataValidationErrorStyle::Stop;
}


std::string serializedFormula(const xlpp::Cell& cell) {
    std::string formula = cell.formula();
    if (cell.formulaMetadata().type() != xlpp::FormulaType::DynamicArray) return formula;
    // Office stores dynamic-array functions with compatibility prefixes.
    // _xlws is required for worksheet dynamic-array functions such as SORT,
    // FILTER, UNIQUE, SEQUENCE and SORTBY. Preserve an explicit prefix.
    if (formula.rfind("_xlfn.", 0) == 0 || formula.rfind("_xlws.", 0) == 0) return formula;
    const auto paren = formula.find('(');
    if (paren == std::string::npos) return formula;
    const std::string functionName = formula.substr(0, paren);
    static const std::unordered_set<std::string> worksheetDynamicFunctions{
        "FILTER", "RANDARRAY", "SEQUENCE", "SORT", "SORTBY", "UNIQUE"
    };
    if (worksheetDynamicFunctions.count(functionName) != 0)
        return "_xlfn._xlws." + formula;
    return "_xlfn." + formula;
}

bool richTextBooleanProperty(const std::string& properties, std::string_view name) {
    const auto nodes = xlpp::internal::tags(properties, name);
    if (nodes.empty()) return false;
    const auto value = xlpp::internal::attribute(nodes.front(), "val");
    return value.empty() || (value != "0" && value != "false" && value != "off");
}

std::optional<xlpp::RichText> parseRichTextRuns(std::string_view container) {
    const auto runNodes = xlpp::internal::tags(container, "r");
    if (runNodes.empty()) return std::nullopt;

    xlpp::RichText result;
    for (const auto& runNode : runNodes) {
        xlpp::RichTextRun run(xlpp::internal::tagText(runNode, "t"));
        const auto properties = xlpp::internal::tags(runNode, "rPr");
        if (!properties.empty()) {
            const auto& rPr = properties.front();
            run.setBold(richTextBooleanProperty(rPr, "b"));
            run.setItalic(richTextBooleanProperty(rPr, "i"));
            run.setUnderline(richTextBooleanProperty(rPr, "u"));
            run.setStrike(richTextBooleanProperty(rPr, "strike"));

            const auto colors = xlpp::internal::tags(rPr, "color");
            if (!colors.empty()) run.setColor(xlpp::internal::attribute(colors.front(), "rgb"));
            const auto fonts = xlpp::internal::tags(rPr, "rFont");
            if (!fonts.empty()) run.setFontName(xlpp::internal::attribute(fonts.front(), "val"));
            const auto sizes = xlpp::internal::tags(rPr, "sz");
            if (!sizes.empty()) {
                const auto value = xlpp::internal::attribute(sizes.front(), "val");
                if (!value.empty()) run.setSize(xlpp::internal::parseDoubleExact(value, "rich-text font size"));
            }
        }
        result.addRun(std::move(run));
    }
    return result;
}

void writeRichTextRuns(std::ostringstream& xml, const xlpp::RichText& richText) {
    xml << "<is>";
    for (const auto& run : richText.runs()) {
        xml << "<r>";
        const bool hasProperties = run.bold() || run.italic() || run.underline() || run.strike()
            || !run.color().empty() || !run.fontName().empty() || run.size() > 0.0;
        if (hasProperties) {
            xml << "<rPr>";
            if (!run.fontName().empty()) {
                xml << "<rFont val=\"";
                writeXmlEscaped(xml, run.fontName());
                xml << "\"/>";
            }
            if (run.bold()) xml << "<b val=\"1\"/>";
            if (run.italic()) xml << "<i val=\"1\"/>";
            if (run.strike()) xml << "<strike val=\"1\"/>";
            if (!run.color().empty()) {
                xml << "<color rgb=\"";
                writeXmlEscaped(xml, run.color());
                xml << "\"/>";
            }
            if (run.size() > 0.0) xml << "<sz val=\"" << run.size() << "\"/>";
            // SpreadsheetML CT_RPrElt requires underline after size.
            if (run.underline()) xml << "<u val=\"single\"/>";
            xml << "</rPr>";
        }
        xml << "<t xml:space=\"preserve\">";
        writeXmlEscaped(xml, run.text());
        xml << "</t></r>";
    }
    xml << "</is>";
}

void writeCell(std::ostringstream& xml, const xlpp::Cell& cell, const StyleCatalog& styles, bool date1904,
                const std::unordered_map<std::string, std::size_t>* sstIndex) {
    // Preserve cells that carry a style even when they hold no value or
    // formula (e.g. a highlighted empty range).
    if (cell.empty() && !cell.hasNonDefaultStyle() && !cell.styleIndex()) return;
    xml << "<c r=\"" << cell.address() << "\"";
    if (cell.styleIndex()) xml << " s=\"" << *cell.styleIndex() << "\"";
    else if (cell.hasNonDefaultStyle()) {
        const auto styleId = styles.find(cell.style());
        if (styleId != 0) xml << " s=\"" << styleId << "\"";
    }
    if (cell.empty()) { xml << "/>"; return; }
    if (cell.hasRichText()) {
        xml << " t=\"inlineStr\">";
        if (cell.hasFormula()) {
            xml << "<f>";
            writeXmlEscaped(xml, serializedFormula(cell));
            xml << "</f>";
        }
        writeRichTextRuns(xml, *cell.richTextValue());
        xml << "</c>";
        return;
    }
    if (const auto* stringValue = std::get_if<std::string>(&cell.value())) {
        if (sstIndex) {
            const auto it = sstIndex->find(*stringValue);
            if (it != sstIndex->end()) {
                xml << " t=\"s\">";
                if (cell.hasFormula()) {
                    xml << "<f";
                    const auto& metadata = cell.formulaMetadata();
                    if (metadata.type() == xlpp::FormulaType::Shared) xml << " t=\"shared\"";
                    else if (metadata.type() == xlpp::FormulaType::Array) xml << " t=\"array\"";
                    else if (metadata.type() == xlpp::FormulaType::DynamicArray) xml << " t=\"array\"";
                    else if (metadata.type() == xlpp::FormulaType::DataTable) xml << " t=\"dataTable\"";
                    if (!metadata.reference().empty()) { xml << " ref=\""; writeXmlEscaped(xml, metadata.reference()); xml << "\""; }
                    if (metadata.sharedIndex()) xml << " si=\"" << *metadata.sharedIndex() << "\"";
                    if (metadata.alwaysCalculateArray()) xml << " aca=\"1\"";
                    if (metadata.calculateOnLoad()) xml << " ca=\"1\"";
                    xml << ">"; writeXmlEscaped(xml, serializedFormula(cell)); xml << "</f>";
                }
                xml << "<v>" << it->second << "</v></c>";
                return;
            }
        }
        xml << " t=\"inlineStr\">";
    }
    else if (std::holds_alternative<bool>(cell.value())) xml << " t=\"b\">";
    else if (std::holds_alternative<xlpp::CellError>(cell.value())) xml << " t=\"e\">";
    else xml << ">";
    if (cell.hasFormula()) {
        xml << "<f";
        const auto& metadata = cell.formulaMetadata();
        if (metadata.type() == xlpp::FormulaType::Shared) xml << " t=\"shared\"";
        else if (metadata.type() == xlpp::FormulaType::Array) xml << " t=\"array\"";
        else if (metadata.type() == xlpp::FormulaType::DynamicArray) xml << " t=\"array\"";
        else if (metadata.type() == xlpp::FormulaType::DataTable) xml << " t=\"dataTable\"";
        if (!metadata.reference().empty()) { xml << " ref=\""; writeXmlEscaped(xml, metadata.reference()); xml << "\""; }
        if (metadata.sharedIndex()) xml << " si=\"" << *metadata.sharedIndex() << "\"";
        if (metadata.alwaysCalculateArray()) xml << " aca=\"1\"";
        if (metadata.calculateOnLoad()) xml << " ca=\"1\"";
        xml << ">"; writeXmlEscaped(xml, serializedFormula(cell)); xml << "</f>";
    }
    if (const auto* stringValue = std::get_if<std::string>(&cell.value())) {
        xml << "<is><t xml:space=\"preserve\">"; writeXmlEscaped(xml, *stringValue); xml << "</t></is>";
    }
    else if (const auto* numberValue = std::get_if<double>(&cell.value()))
        xml << "<v>" << *numberValue << "</v>";
    else if (const auto* booleanValue = std::get_if<bool>(&cell.value()))
        xml << "<v>" << (*booleanValue ? 1 : 0) << "</v>";
    else if (const auto* errorValue = std::get_if<xlpp::CellError>(&cell.value())) {
        xml << "<v>"; writeXmlEscaped(xml, xlpp::toString(*errorValue)); xml << "</v>";
    }
    else if (const auto* dateValue = std::get_if<xlpp::DateTime>(&cell.value()))
        xml << "<v>" << xlpp::toExcelSerial(*dateValue, date1904) << "</v>";
    xml << "</c>";
}

void writeColAttrs(std::ostringstream& xml, const xlpp::ColumnDimension& dim) {
    if (dim.width) xml << " width=\"" << *dim.width << "\" customWidth=\"1\"";
    if (dim.hidden) xml << " hidden=\"1\"";
    if (dim.bestFit) xml << " bestFit=\"1\"";
    if (dim.outlineLevel) xml << " outlineLevel=\"" << dim.outlineLevel << "\"";
    if (dim.collapsed) xml << " collapsed=\"1\"";
}

std::string sheetXml(const xlpp::Worksheet& sheet, const StyleCatalog& styles, const DxfCatalog& dxfs, bool date1904, bool strict,
                     const std::unordered_map<std::string, std::size_t>* sstIndex,
                     std::size_t rowWorkers,
                     std::string_view vbaCodeName) {
    std::ostringstream xml;
    xml.precision(17);
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"" << nsMain(strict) << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">";
    // CT_Worksheet requires sheetPr before dimension/sheetViews. Fit-to-page
    // is activated by pageSetUpPr, not by fitToWidth/fitToHeight alone.
    if (!vbaCodeName.empty() || sheet.sheetView().tabColor() || sheet.pageSetup().fitToPage()) {
        xml << "<sheetPr";
        if (!vbaCodeName.empty()) xml << " codeName=\"" << xmlEscape(vbaCodeName) << "\"";
        xml << ">";
        if (sheet.sheetView().tabColor())
            xml << "<tabColor rgb=\"" << xmlEscape(*sheet.sheetView().tabColor()) << "\"/>";
        if (sheet.pageSetup().fitToPage()) xml << "<pageSetUpPr fitToPage=\"1\"/>";
        xml << "</sheetPr>";
    }
    xml << "<dimension ref=\"" << sheet.dimensions() << "\"/><sheetViews><sheetView workbookViewId=\"" << sheet.sheetView().workbookViewId() << "\"";
    if (sheet.sheetView().tabSelected()) xml << " tabSelected=\"1\"";
    if (!sheet.sheetView().showGridLines()) xml << " showGridLines=\"0\"";
    if (!sheet.sheetView().showRowColHeaders()) xml << " showRowColHeaders=\"0\"";
    if (sheet.sheetView().rightToLeft()) xml << " rightToLeft=\"1\"";
    xml << " zoomScale=\"" << sheet.sheetView().zoomScale() << "\" zoomScaleNormal=\"" << sheet.sheetView().zoomScaleNormal() << "\"";
    if (!sheet.sheetView().showOutlineSymbols()) xml << " showOutlineSymbols=\"0\"";
    xml << ">";
    if (sheet.frozenPane()) {
        const auto pane = xlpp::CellReference::parse(*sheet.frozenPane());
        const auto xSplit = pane.column > 1 ? pane.column - 1 : 0;
        const auto ySplit = pane.row > 1 ? pane.row - 1 : 0;
        xml << "<pane";
        if (xSplit) xml << " xSplit=\"" << xSplit << "\"";
        if (ySplit) xml << " ySplit=\"" << ySplit << "\"";
        xml << " topLeftCell=\"" << *sheet.frozenPane() << "\" activePane=\"bottomRight\" state=\"frozen\"/>";
    } else if (sheet.sheetView().xSplit() != 0 || sheet.sheetView().ySplit() != 0 ||
               !sheet.sheetView().pane().empty() || !sheet.sheetView().topLeftCell().empty()) {
        xml << "<pane";
        if (sheet.sheetView().xSplit() != 0) xml << " xSplit=\"" << sheet.sheetView().xSplit() << "\"";
        if (sheet.sheetView().ySplit() != 0) xml << " ySplit=\"" << sheet.sheetView().ySplit() << "\"";
        if (!sheet.sheetView().topLeftCell().empty())
            xml << " topLeftCell=\"" << xmlEscape(sheet.sheetView().topLeftCell()) << "\"";
        if (!sheet.sheetView().pane().empty())
            xml << " activePane=\"" << xmlEscape(sheet.sheetView().pane()) << "\"";
        xml << " state=\"split\"/>";
    }
    xml << "</sheetView></sheetViews>";
    if (!sheet.preservedSheetFormatPrXml().empty()) xml << sheet.preservedSheetFormatPrXml();
    else xml << "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>";
    if (!sheet.columnDimensions().empty()) {
        xml << "<cols>";
        std::size_t rangeStart = 0, rangeEnd = 0;
        const xlpp::ColumnDimension* lastDim = nullptr;
        for (const auto& [column, dimension] : sheet.columnDimensions()) {
            if (rangeStart == 0) {
                rangeStart = rangeEnd = column;
                lastDim = &dimension;
            } else if (lastDim &&
                       lastDim->width == dimension.width &&
                       lastDim->hidden == dimension.hidden &&
                       lastDim->bestFit == dimension.bestFit &&
                       lastDim->outlineLevel == dimension.outlineLevel &&
                       lastDim->collapsed == dimension.collapsed) {
                rangeEnd = column;
            } else {
                xml << "<col min=\"" << rangeStart << "\" max=\"" << rangeEnd << "\"";
                writeColAttrs(xml, *lastDim);
                xml << "/>";
                rangeStart = rangeEnd = column;
                lastDim = &dimension;
            }
        }
        if (rangeStart > 0) {
            xml << "<col min=\"" << rangeStart << "\" max=\"" << rangeEnd << "\"";
            writeColAttrs(xml, *lastDim);
            xml << "/>";
        }
        xml << "</cols>";
    }
    xml << "<sheetData>";

    auto writeRowStart = [&](std::ostringstream& out, std::size_t row) {
        out << "<row r=\"" << row << "\"";
        if (const auto* dim = sheet.tryRowDimension(row)) {
            if (dim->height) out << " ht=\"" << *dim->height << "\" customHeight=\"1\"";
            if (dim->hidden) out << " hidden=\"1\"";
            if (dim->outlineLevel) out << " outlineLevel=\"" << dim->outlineLevel << "\"";
            if (dim->collapsed) out << " collapsed=\"1\"";
        }
        out << ">";
    };

    if (rowWorkers <= 1) {
        // std::map is already row-major by makeCellKey().  Stream the common
        // sequential path directly instead of allocating an O(cell-count)
        // pointer array plus a second row-span array before emitting XML.
        std::size_t currentRow = 0;
        for (const auto& [_, cell] : sheet.cells()) {
            if (cell.empty() && !cell.hasNonDefaultStyle() && !cell.styleIndex()) continue;
            if (cell.row() != currentRow) {
                if (currentRow != 0) xml << "</row>";
                currentRow = cell.row();
                writeRowStart(xml, currentRow);
            }
            writeCell(xml, cell, styles, date1904, sstIndex);
        }
        if (currentRow != 0) xml << "</row>";
    } else {
        // Parallel row serialization needs random-access row partitions, so
        // materialize lightweight cell pointers only for this opt-in path.
        std::vector<const xlpp::Cell*> ordered;
        ordered.reserve(sheet.cells().size());
        for (const auto& [_, cell] : sheet.cells())
            if (!cell.empty() || cell.hasNonDefaultStyle() || cell.styleIndex()) ordered.push_back(&cell);

        if (!ordered.empty()) {
            struct RowSpan { std::size_t row; std::size_t begin; std::size_t end; };
            std::vector<RowSpan> rowSpans;
            rowSpans.reserve(std::min(ordered.size(), sheet.maxRow()));
            for (std::size_t i = 0; i < ordered.size(); ) {
                const std::size_t r = ordered[i]->row();
                std::size_t j = i + 1;
                while (j < ordered.size() && ordered[j]->row() == r) ++j;
                rowSpans.push_back({r, i, j});
                i = j;
            }

            auto writeRows = [&](std::size_t rowBegin, std::size_t rowEnd, std::ostringstream& out) {
                for (std::size_t ri = rowBegin; ri < rowEnd; ++ri) {
                    const auto& span = rowSpans[ri];
                    writeRowStart(out, span.row);
                    for (std::size_t ci = span.begin; ci < span.end; ++ci)
                        writeCell(out, *ordered[ci], styles, date1904, sstIndex);
                    out << "</row>";
                }
            };

            const std::size_t threads = std::min(rowWorkers, rowSpans.size());
            const std::size_t chunk = (rowSpans.size() + threads - 1) / threads;
            std::vector<std::ostringstream> chunks(threads);
            std::vector<std::thread> workers;
            workers.reserve(threads);
            for (std::size_t t = 0; t < threads; ++t) {
                const std::size_t begin = t * chunk;
                const std::size_t end = std::min(begin + chunk, rowSpans.size());
                if (begin >= end) break;
                workers.emplace_back([&, t, begin, end] {
                    chunks[t].precision(17);
                    writeRows(begin, end, chunks[t]);
                });
            }
            for (auto& w : workers) w.join();
            // Use the rvalue str() overload so implementations can transfer
            // the underlying string buffer without keeping an extra full copy.
            for (auto& c : chunks) xml << std::move(c).str();
        }
    }
    xml << "</sheetData>";

    // CT_Worksheet requires sheetProtection immediately after sheetData /
    // sheetCalcPr and before mergeCells, autoFilter and formatting blocks.
    const auto& protection = sheet.protection();
    if (protection.enabled()) {
        xml << "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"";
        if (!protection.passwordHash().empty()) xml << " password=\"" << xmlEscape(protection.passwordHash()) << "\"";
        xml << " selectLockedCells=\"" << (protection.selectLockedCells()?0:1) << "\" selectUnlockedCells=\"" << (protection.selectUnlockedCells()?0:1)
            << "\" formatCells=\"" << (protection.formatCells()?0:1) << "\" formatColumns=\"" << (protection.formatColumns()?0:1)
            << "\" formatRows=\"" << (protection.formatRows()?0:1) << "\" insertRows=\"" << (protection.insertRows()?0:1)
            << "\" insertColumns=\"" << (protection.insertColumns()?0:1) << "\" deleteRows=\"" << (protection.deleteRows()?0:1)
            << "\" deleteColumns=\"" << (protection.deleteColumns()?0:1) << "\" sort=\"" << (protection.sort()?0:1)
            << "\" autoFilter=\"" << (protection.autoFilter()?0:1) << "\"/>";
    }

    if (!sheet.mergedRanges().empty()) {
        xml << "<mergeCells count=\"" << sheet.mergedRanges().size() << "\">";
        for (const auto& range : sheet.mergedRanges()) xml << "<mergeCell ref=\"" << range << "\"/>";
        xml << "</mergeCells>";
    }
    if (sheet.autoFilter().enabled()) {
        const auto& autoFilter = sheet.autoFilter();
        xml << "<autoFilter ref=\"" << xmlEscape(autoFilter.reference()) << "\">";
        for (const auto& [columnId, column] : autoFilter.columns()) {
            xml << "<filterColumn colId=\"" << columnId << "\">";
            if (!column.values().empty() || column.includeBlank()) {
                xml << "<filters";
                if (column.includeBlank()) xml << " blank=\"1\"";
                xml << ">";
                for (const auto& value : column.values())
                    xml << "<filter val=\"" << xmlEscape(value) << "\"/>";
                xml << "</filters>";
            }
            if (!column.customFilters().empty()) {
                xml << "<customFilters" << (column.andMode() ? " and=\"1\"" : "") << ">";
                for (const auto& filter : column.customFilters())
                    xml << "<customFilter operator=\"" << filterOperatorName(filter.op)
                        << "\" val=\"" << xmlEscape(filter.value) << "\"/>";
                xml << "</customFilters>";
            }
            if (const auto& top10 = column.top10(); top10) {
                xml << "<top10 top=\"" << (top10->top ? 1 : 0) << "\" percent=\""
                    << (top10->percent ? 1 : 0) << "\" val=\"" << top10->value << "\"/>";
            }
            if (const auto& dynamic = column.dynamicFilter(); dynamic) {
                xml << "<dynamicFilter type=\"" << xmlEscape(dynamic->type) << "\"";
                if (dynamic->value) xml << " val=\"" << *dynamic->value << "\"";
                xml << "/>";
            }
            if (const auto& extension = column.filterExtension(); extension && !extension->rawXml.empty()) {
                xml << extension->rawXml;
            }
            xml << "</filterColumn>";
        }
        if (autoFilter.sortStateValue()) {
            const auto& sort = *autoFilter.sortStateValue();
            xml << "<sortState ref=\"" << xmlEscape(sort.reference()) << "\"";
            if (sort.caseSensitive()) xml << " caseSensitive=\"1\"";
            xml << ">";
            for (const auto& condition : sort.conditions())
                xml << "<sortCondition ref=\"" << xmlEscape(condition.reference)
                    << "\" descending=\"" << (condition.descending ? 1 : 0) << "\"/>";
            xml << "</sortState>";
        }
        xml << "</autoFilter>";
    }
    std::set<std::size_t> emittedConditionalPriorities;
    std::size_t nextConditionalPriority = 1;
    for (const auto& entry : sheet.conditionalFormatting().entries()) {
        if (entry.rules().empty()) continue;
        xml << "<conditionalFormatting sqref=\"" << xmlEscape(entry.reference()) << "\">";
        for (const auto& rule : entry.rules()) {
            std::size_t emittedPriority = rule.priority();
            if (emittedPriority == 0 || emittedConditionalPriorities.count(emittedPriority) != 0) {
                while (emittedConditionalPriorities.count(nextConditionalPriority) != 0) ++nextConditionalPriority;
                emittedPriority = nextConditionalPriority++;
            }
            emittedConditionalPriorities.insert(emittedPriority);
            switch (rule.type()) {
                case xlpp::ConditionalRuleType::Formula:
                case xlpp::ConditionalRuleType::CellIs: {
                    xml << "<cfRule type=\"" << (rule.type() == xlpp::ConditionalRuleType::Formula ? "expression" : "cellIs") << "\"";
                    if (rule.type() == xlpp::ConditionalRuleType::CellIs)
                        xml << " operator=\"" << conditionalOperatorName(rule.op()) << "\"";
                    if (rule.hasDifferentialStyle()) xml << " dxfId=\"" << dxfs.find(rule.differentialStyle()) << "\"";
                    xml << " priority=\"" << emittedPriority << "\"";
                    if (rule.stopIfTrue()) xml << " stopIfTrue=\"1\"";
                    xml << ">";
                    for (const auto& formula : rule.formulas()) xml << "<formula>" << xmlEscape(formula) << "</formula>";
                    xml << "</cfRule>";
                    break;
                }
                case xlpp::ConditionalRuleType::DataBar: {
                    const auto& db = rule.getDataBar();
                    xml << "<cfRule type=\"dataBar\" priority=\"" << emittedPriority << "\"><dataBar";
                    if (!db.direction.empty() && db.direction != "leftToRight") xml << " direction=\"" << db.direction << "\"";
                    if (!db.showValue) xml << " showValue=\"0\"";
                    xml << ">";
                    writeCfvo(xml, db.min);
                    writeCfvo(xml, db.max);
                    xml << "<color rgb=\"" << xmlEscape(db.color) << "\"/>";
                    if (db.axisPosition) {
                        xml << "<cfvo type=\"autoMin\"";
                        if (*db.axisPosition != 0.0) xml << " val=\"" << *db.axisPosition << "\"";
                        xml << "/>";
                    }
                    xml << "</dataBar></cfRule>";
                    break;
                }
                case xlpp::ConditionalRuleType::ColorScale: {
                    const auto& cs = rule.getColorScale();
                    xml << "<cfRule type=\"colorScale\" priority=\"" << emittedPriority << "\">";
                    xml << "<colorScale>";
                    if (!cs.stops.empty()) {
                        for (const auto& stop : cs.stops) writeCfvo(xml, stop);
                        for (const auto& stop : cs.stops)
                            if (stop.color) xml << "<color rgb=\"" << xmlEscape(*stop.color) << "\"/>";
                    }
                    xml << "</colorScale></cfRule>";
                    break;
                }
                case xlpp::ConditionalRuleType::IconSet: {
                    const auto& is = rule.getIconSet();
                    xml << "<cfRule type=\"iconSet\" priority=\"" << emittedPriority << "\"><iconSet";
                    if (is.reverse) xml << " reverse=\"1\"";
                    if (!is.showValue) xml << " showValue=\"0\"";
                    xml << " iconSet=\"" << xmlEscape(is.icons) << "\">";
                    for (const auto& t : is.thresholds) writeCfvo(xml, t);
                    xml << "</iconSet></cfRule>";
                    break;
                }
                case xlpp::ConditionalRuleType::ContainsText:
                case xlpp::ConditionalRuleType::NotContainsText:
                case xlpp::ConditionalRuleType::BeginsWith:
                case xlpp::ConditionalRuleType::EndsWith: {
                    const char* cfType = rule.type() == xlpp::ConditionalRuleType::ContainsText ? "containsText"
                        : rule.type() == xlpp::ConditionalRuleType::NotContainsText ? "notContainsText"
                        : rule.type() == xlpp::ConditionalRuleType::BeginsWith ? "beginsWith" : "endsWith";
                    xml << "<cfRule type=\"" << cfType << "\"";
                    if (rule.hasDifferentialStyle()) xml << " dxfId=\"" << dxfs.find(rule.differentialStyle()) << "\"";
                    xml << " priority=\"" << emittedPriority << "\"";
                    if (rule.stopIfTrue()) xml << " stopIfTrue=\"1\"";
                    std::string formula;
                    if (rule.type() == xlpp::ConditionalRuleType::ContainsText)
                        formula = "NOT(ISERROR(SEARCH(\"" + rule.text() + "\",A1)))";
                    else if (rule.type() == xlpp::ConditionalRuleType::NotContainsText)
                        formula = "ISERROR(SEARCH(\"" + rule.text() + "\",A1))";
                    else if (rule.type() == xlpp::ConditionalRuleType::BeginsWith)
                        formula = "LEFT(A1,LEN(\"" + rule.text() + "\"))=\"" + rule.text() + "\"";
                    else
                        formula = "RIGHT(A1,LEN(\"" + rule.text() + "\"))=\"" + rule.text() + "\"";
                    xml << "><formula>" << xmlEscape(formula) << "</formula></cfRule>";
                    break;
                }
                case xlpp::ConditionalRuleType::AboveAverage:
                case xlpp::ConditionalRuleType::BelowAverage:
                case xlpp::ConditionalRuleType::AboveOrEqualAverage:
                case xlpp::ConditionalRuleType::BelowOrEqualAverage: {
                    const bool above = rule.type() == xlpp::ConditionalRuleType::AboveAverage
                        || rule.type() == xlpp::ConditionalRuleType::AboveOrEqualAverage;
                    xml << "<cfRule type=\"aboveAverage\"";
                    if (rule.type() == xlpp::ConditionalRuleType::BelowAverage
                        || rule.type() == xlpp::ConditionalRuleType::BelowOrEqualAverage) xml << " aboveAverage=\"0\"";
                    if (rule.equalAverage()) xml << " equalAverage=\"1\"";
                    if (rule.stdDev()) xml << " stdDev=\"1\"";
                    xml << " priority=\"" << emittedPriority << "\"";
                    if (rule.hasDifferentialStyle()) xml << " dxfId=\"" << dxfs.find(rule.differentialStyle()) << "\"";
                    if (rule.stopIfTrue()) xml << " stopIfTrue=\"1\"";
                    xml << "/></cfRule>";
                    (void)above;
                    break;
                }
                case xlpp::ConditionalRuleType::Top10:
                case xlpp::ConditionalRuleType::Bottom10: {
                    xml << "<cfRule type=\"top10\"";
                    if (rule.type() == xlpp::ConditionalRuleType::Bottom10) xml << " bottom=\"1\"";
                    if (rule.top10Config().percent) xml << " percent=\"1\"";
                    xml << " rank=\"" << rule.top10Config().rank << "\"";
                    xml << " priority=\"" << emittedPriority << "\"";
                    if (rule.hasDifferentialStyle()) xml << " dxfId=\"" << dxfs.find(rule.differentialStyle()) << "\"";
                    if (rule.stopIfTrue()) xml << " stopIfTrue=\"1\"";
                    xml << "/></cfRule>";
                    break;
                }
                case xlpp::ConditionalRuleType::DuplicateValues:
                case xlpp::ConditionalRuleType::UniqueValues: {
                    xml << "<cfRule type=\"" << (rule.type() == xlpp::ConditionalRuleType::DuplicateValues ? "duplicateValues" : "uniqueValues") << "\"";
                    xml << " priority=\"" << emittedPriority << "\"";
                    if (rule.hasDifferentialStyle()) xml << " dxfId=\"" << dxfs.find(rule.differentialStyle()) << "\"";
                    if (rule.stopIfTrue()) xml << " stopIfTrue=\"1\"";
                    xml << "/></cfRule>";
                    break;
                }
            }
        }
        xml << "</conditionalFormatting>";
    }
    if (!sheet.dataValidations().empty()) {
        xml << "<dataValidations count=\"" << sheet.dataValidations().items().size() << "\">";
        for (const auto& validation : sheet.dataValidations().items()) {
            xml << "<dataValidation type=\"" << dataValidationTypeName(validation.type()) << "\"";
            if (validation.type() != xlpp::DataValidationType::None &&
                validation.type() != xlpp::DataValidationType::List &&
                validation.type() != xlpp::DataValidationType::Custom)
                xml << " operator=\"" << dataValidationOperatorName(validation.op()) << "\"";
            xml << " errorStyle=\"" << dataValidationErrorStyleName(validation.errorStyle())
                << "\" sqref=\"" << xmlEscape(validation.reference()) << "\"";
            if (validation.allowBlank()) xml << " allowBlank=\"1\"";
            if (validation.showDropDown()) xml << " showDropDown=\"1\"";
            if (validation.showInputMessage()) xml << " showInputMessage=\"1\"";
            if (validation.showErrorMessage()) xml << " showErrorMessage=\"1\"";
            if (!validation.promptTitle().empty()) xml << " promptTitle=\"" << xmlEscape(validation.promptTitle()) << "\"";
            if (!validation.prompt().empty()) xml << " prompt=\"" << xmlEscape(validation.prompt()) << "\"";
            if (!validation.errorTitle().empty()) xml << " errorTitle=\"" << xmlEscape(validation.errorTitle()) << "\"";
            if (!validation.error().empty()) xml << " error=\"" << xmlEscape(validation.error()) << "\"";
            xml << ">";
            if (!validation.formula1().empty()) xml << "<formula1>" << xmlEscape(validation.formula1()) << "</formula1>";
            if (!validation.formula2().empty()) xml << "<formula2>" << xmlEscape(validation.formula2()) << "</formula2>";
            xml << "</dataValidation>";
        }
        xml << "</dataValidations>";
    }
    { std::size_t hyperlinkId=1; bool any=false; for(const auto& pair:sheet.cells()) if(pair.second.hasHyperlink()){ if(!any){xml<<"<hyperlinks>";any=true;} const auto& h=*pair.second.hyperlinkValue(); xml<<"<hyperlink ref=\""<<pair.second.address()<<"\""; if(h.external()) xml<<" r:id=\"rIdHyperlink"<<hyperlinkId++<<"\""; else xml<<" location=\""<<xmlEscape(h.target())<<"\""; if(!h.display().empty())xml<<" display=\""<<xmlEscape(h.display())<<"\"";if(!h.tooltip().empty())xml<<" tooltip=\""<<xmlEscape(h.tooltip())<<"\"";xml<<"/>";} if(any)xml<<"</hyperlinks>"; }
    bool hasComments = false; for (const auto& pair : sheet.cells()) if (pair.second.hasComment()) { hasComments = true; break; }
    const auto& options = sheet.printOptions();
    if (options.horizontalCentered() || options.verticalCentered() || options.headings() || options.gridLines())
        xml << "<printOptions horizontalCentered=\"" << (options.horizontalCentered()?1:0) << "\" verticalCentered=\"" << (options.verticalCentered()?1:0)
            << "\" headings=\"" << (options.headings()?1:0) << "\" gridLines=\"" << (options.gridLines()?1:0) << "\"/>";
    const auto& margins = sheet.pageMargins();
    xml << "<pageMargins left=\"" << margins.left() << "\" right=\"" << margins.right() << "\" top=\"" << margins.top()
        << "\" bottom=\"" << margins.bottom() << "\" header=\"" << margins.header() << "\" footer=\"" << margins.footer() << "\"/>";
    const auto& setup = sheet.pageSetup();
    if (setup.orientation()!=xlpp::PageOrientation::Default || setup.paperSize()!=xlpp::PaperSize::Default || setup.scale()!=100 || setup.fitToPage() || setup.blackAndWhite() || setup.draft() || setup.useFirstPageNumber() || setup.hasExtendedSettings()) {
        xml << "<pageSetup";
        if (setup.orientation()==xlpp::PageOrientation::Portrait) xml << " orientation=\"portrait\"";
        else if (setup.orientation()==xlpp::PageOrientation::Landscape) xml << " orientation=\"landscape\"";
        if (setup.paperSize()!=xlpp::PaperSize::Default) xml << " paperSize=\"" << static_cast<unsigned>(setup.paperSize()) << "\"";
        xml << " scale=\"" << setup.scale() << "\"";
        if (setup.fitToPage()) xml << " fitToWidth=\"" << setup.fitToWidth() << "\" fitToHeight=\"" << setup.fitToHeight() << "\"";
        if (setup.blackAndWhite()) xml << " blackAndWhite=\"1\"";
        if (setup.draft()) xml << " draft=\"1\"";
        if (setup.useFirstPageNumber()) xml << " firstPageNumber=\"" << setup.firstPageNumber() << "\" useFirstPageNumber=\"1\"";
        if (setup.paperHeight()) xml << " paperHeight=\"" << xmlEscape(*setup.paperHeight()) << "\"";
        if (setup.paperWidth()) xml << " paperWidth=\"" << xmlEscape(*setup.paperWidth()) << "\"";
        if (setup.pageOrder()==xlpp::PageOrder::DownThenOver) xml << " pageOrder=\"downThenOver\"";
        else if (setup.pageOrder()==xlpp::PageOrder::OverThenDown) xml << " pageOrder=\"overThenDown\"";
        if (setup.usePrinterDefaults()) xml << " usePrinterDefaults=\"" << (*setup.usePrinterDefaults()?1:0) << "\"";
        if (setup.cellComments()==xlpp::PageCellComments::AsDisplayed) xml << " cellComments=\"asDisplayed\"";
        else if (setup.cellComments()==xlpp::PageCellComments::AtEnd) xml << " cellComments=\"atEnd\"";
        if (setup.errors()==xlpp::PageErrorDisplay::Displayed) xml << " errors=\"displayed\"";
        else if (setup.errors()==xlpp::PageErrorDisplay::Blank) xml << " errors=\"blank\"";
        else if (setup.errors()==xlpp::PageErrorDisplay::Dash) xml << " errors=\"dash\"";
        else if (setup.errors()==xlpp::PageErrorDisplay::NA) xml << " errors=\"NA\"";
        if (setup.horizontalDpi()) xml << " horizontalDpi=\"" << *setup.horizontalDpi() << "\"";
        if (setup.verticalDpi()) xml << " verticalDpi=\"" << *setup.verticalDpi() << "\"";
        if (setup.copies()) xml << " copies=\"" << *setup.copies() << "\"";
        if (setup.relationshipId()) xml << " r:id=\"" << xmlEscape(*setup.relationshipId()) << "\"";
        xml << "/>";
    }
    const auto& hf = sheet.headerFooter();
    if (!hf.oddHeader().empty() || !hf.oddFooter().empty() || !hf.evenHeader().empty() || !hf.evenFooter().empty()) {
        xml << "<headerFooter differentOddEven=\"" << (hf.differentOddEven()?1:0) << "\" differentFirst=\"" << (hf.differentFirst()?1:0) << "\">";
        if (!hf.oddHeader().empty()) xml << "<oddHeader>" << xmlEscape(hf.oddHeader()) << "</oddHeader>";
        if (!hf.oddFooter().empty()) xml << "<oddFooter>" << xmlEscape(hf.oddFooter()) << "</oddFooter>";
        if (!hf.evenHeader().empty()) xml << "<evenHeader>" << xmlEscape(hf.evenHeader()) << "</evenHeader>";
        if (!hf.evenFooter().empty()) xml << "<evenFooter>" << xmlEscape(hf.evenFooter()) << "</evenFooter>";
        if (!hf.firstHeader().empty()) xml << "<firstHeader>" << xmlEscape(hf.firstHeader()) << "</firstHeader>";
        if (!hf.firstFooter().empty()) xml << "<firstFooter>" << xmlEscape(hf.firstFooter()) << "</firstFooter>";
        xml << "</headerFooter>";
    }
    // Manual page breaks (SpreadsheetML <rowBreaks>/<colBreaks>). Row breaks
    // use the 0-based row index above which a break occurs.
    const auto& rowBreaks = sheet.rowBreaks();
    if (!rowBreaks.empty()) {
        xml << "<rowBreaks count=\"" << rowBreaks.size() << "\" manualBreakCount=\"" << rowBreaks.size() << "\">";
        for (const auto row : rowBreaks)
            xml << "<brk id=\"" << row << "\" max=\"16383\" man=\"1\"/>";
        xml << "</rowBreaks>";
    }
    const auto& columnBreaks = sheet.columnBreaks();
    if (!columnBreaks.empty()) {
        xml << "<colBreaks count=\"" << columnBreaks.size() << "\" manualBreakCount=\"" << columnBreaks.size() << "\">";
        for (const auto column : columnBreaks)
            xml << "<brk id=\"" << column << "\" max=\"1048575\" man=\"1\"/>";
        xml << "</colBreaks>";
    }
    if (!sheet.images().empty() || sheet.chartCount() > 0)
        xml << "<drawing r:id=\"rIdDrawing\"/>";
    // CT_Worksheet requires legacyDrawing after page settings and drawing.
    // Placing it before printOptions/pageMargins causes desktop Excel to repair
    // or discard otherwise valid legacy note/comment parts.
    if (hasComments) xml << "<legacyDrawing r:id=\"rIdCommentsVml\"/>";
    if (!sheet.tables().empty()) {
        xml << "<tableParts count=\"" << sheet.tables().size() << "\">";
        for (std::size_t i = 0; i < sheet.tables().size(); ++i) xml << "<tablePart r:id=\"rId" << i + 1 << "\"/>";
        xml << "</tableParts>";
    }
    const auto generatedPivotOwnerCount = sheet.pivotsDirty()
        ? sheet.pivotTables().size()
        : (sheet.pivotAppendDirty() ? sheet.appendedPivotCount() : std::size_t{0});
    if (generatedPivotOwnerCount > 0) {
        xml << "<pivotTableParts count=\"" << generatedPivotOwnerCount << "\">";
        for (std::size_t i = 0; i < generatedPivotOwnerCount; ++i)
            xml << "<pivotTablePart r:id=\"rIdPivot" << i + 1 << "\"/>";
        xml << "</pivotTableParts>";
    }
    // Sparklines are emitted through the x14 worksheet extension list. When the
    // worksheet carries an untouched imported sparklines block it is preserved
    // verbatim; otherwise the modeled groups are serialized from scratch.
    if (sheet.hasSparklines()) {
        xml << "<extLst><ext uri=\"{05C60535-1F16-4FD2-B633-F4F36F0B64E0}\" xmlns:x14=\""
            << (strict ? "http://purl.oclc.org/ooxml/spreadsheetml/2009/9/main" : "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
            << "\"><x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">";
        if (sheet.sparklineGroups().empty()) {
            // Preserve the original imported block byte-for-byte.
            const auto imported = sheet.sparklineGroupsRawXml();
            if (!imported.empty()) xml << imported;
        } else {
            for (const auto& group : sheet.sparklineGroups()) {
                if (!group.rawXml.empty()) { xml << group.rawXml; continue; }
                xml << "<x14:sparklineGroup";
                if (group.type == "column") xml << " type=\"column\"";
                else if (group.type == "stacked") xml << " type=\"stacked\"";
                if (group.lineStyle == "smooth") xml << " lineWeight=\"1.5\"";
                if (group.displayHidden) xml << " displayHidden=\"1\"";
                if (group.displayXAxis) xml << " displayXAxis=\"1\"";
                if (group.displayMarkers) xml << " markers=\"1\"";
                if (group.high) xml << " high=\"1\"";
                if (group.low) xml << " low=\"1\"";
                if (group.first) xml << " first=\"1\"";
                if (group.last) xml << " last=\"1\"";
                if (group.negative) xml << " negative=\"1\"";
                if (group.colorSeries) xml << " colorSeries=\"1\"";
                if (group.colorAxis) xml << " colorAxis=\"1\"";
                if (group.colorMarkers) xml << " colorMarkers=\"1\"";
                if (group.colorFirst) xml << " colorFirst=\"1\"";
                if (group.colorLast) xml << " colorLast=\"1\"";
                if (group.colorHigh) xml << " colorHigh=\"1\"";
                if (group.colorLow) xml << " colorLow=\"1\"";
                if (group.rightToLeft) xml << " rightToLeft=\"1\"";
                xml << ">";
                if (!group.markersColor.empty())
                    xml << "<x14:colorSeries><x14:rgb value=\"" << xmlEscape(group.markersColor) << "\"/></x14:colorSeries>";
                if (!group.negativeColor.empty())
                    xml << "<x14:colorNegative><x14:rgb value=\"" << xmlEscape(group.negativeColor) << "\"/></x14:colorNegative>";
                if (!group.axisColor.empty())
                    xml << "<x14:colorAxis><x14:rgb value=\"" << xmlEscape(group.axisColor) << "\"/></x14:colorAxis>";
                if (!group.dateAxis.empty())
                    xml << "<xm:f>" << xmlEscape(group.dateAxis) << "</xm:f>";
                xml << "<x14:sparklines>";
                for (const auto& sparkline : group.sparklines) {
                    xml << "<x14:sparkline><xm:f>" << xmlEscape(sparkline.reference)
                        << "</xm:f><xm:sqref>" << xmlEscape(sparkline.location) << "</xm:sqref></x14:sparkline>";
                }
                xml << "</x14:sparklines></x14:sparklineGroup>";
            }
        }
        xml << "</x14:sparklineGroups></ext></extLst>";
    }
    xml << "</worksheet>";
    return xml.str();
}

} // namespace internal
} // namespace xlpp


