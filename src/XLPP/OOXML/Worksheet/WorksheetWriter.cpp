#include "OOXML/Worksheet/WorksheetWriter.h"
#include "OOXML/Worksheet/WorksheetFeatureCodec.h"
#include "OOXML/Common/RichTextCodec.h"
#include "OOXML/Common/Namespaces.h"
#include "Package/Xml/XmlUtilities.h"
#include "Core/Threading/ThreadPool.h"

#include <algorithm>
#include <sstream>
#include <unordered_set>
#include <vector>

using xlpp::internal::xmlEscape;
using xlpp::internal::writeXmlEscaped;

namespace xlpp::internal::ooxml {


namespace {
const char* dynamicFilterTypeName(xlpp::DynamicFilterType type) {
    using T = xlpp::DynamicFilterType;
    switch (type) {
        case T::AboveAverage: return "aboveAverage"; case T::BelowAverage: return "belowAverage";
        case T::Tomorrow: return "tomorrow"; case T::Today: return "today"; case T::Yesterday: return "yesterday";
        case T::NextWeek: return "nextWeek"; case T::ThisWeek: return "thisWeek"; case T::LastWeek: return "lastWeek";
        case T::NextMonth: return "nextMonth"; case T::ThisMonth: return "thisMonth"; case T::LastMonth: return "lastMonth";
        case T::NextQuarter: return "nextQuarter"; case T::ThisQuarter: return "thisQuarter"; case T::LastQuarter: return "lastQuarter";
        case T::NextYear: return "nextYear"; case T::ThisYear: return "thisYear"; case T::LastYear: return "lastYear";
        case T::YearToDate: return "yearToDate";
        case T::Quarter1: return "Q1"; case T::Quarter2: return "Q2"; case T::Quarter3: return "Q3"; case T::Quarter4: return "Q4";
        case T::Month1: return "M1"; case T::Month2: return "M2"; case T::Month3: return "M3"; case T::Month4: return "M4";
        case T::Month5: return "M5"; case T::Month6: return "M6"; case T::Month7: return "M7"; case T::Month8: return "M8";
        case T::Month9: return "M9"; case T::Month10: return "M10"; case T::Month11: return "M11"; case T::Month12: return "M12";
    }
    return "today";
}
const char* dateGroupingName(xlpp::DateTimeGrouping grouping) {
    using T = xlpp::DateTimeGrouping;
    switch (grouping) {
        case T::Year: return "year"; case T::Month: return "month"; case T::Day: return "day";
        case T::Hour: return "hour"; case T::Minute: return "minute"; case T::Second: return "second";
    }
    return "year";
}
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

void writeCell(std::ostringstream& xml, const xlpp::Cell& cell, const StyleCatalog& styles, bool date1904,
                const std::unordered_map<std::string, std::size_t>* sstIndex) {
    // Preserve cells that carry a style even when they hold no value or
    // formula (e.g. a highlighted empty range).
    if (cell.empty() && cell.style().isDefault() && !cell.styleIndex()) return;
    xml << "<c r=\"" << cell.address() << "\"";
    if (cell.styleIndex()) xml << " s=\"" << *cell.styleIndex() << "\"";
    else if (!cell.style().isDefault()) {
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

    // Collect non-empty cells (already row-major ordered by std::map key)
    std::vector<const xlpp::Cell*> ordered;
    ordered.reserve(sheet.cells().size());
    for (const auto& [_, cell] : sheet.cells())
        if (!cell.empty() || !cell.style().isDefault() || cell.styleIndex()) ordered.push_back(&cell);

    if (ordered.empty()) {
        xml << "</sheetData>";
    } else {
        // Build row boundaries: for each unique row, record [start, end) index into ordered
        struct RowSpan { std::size_t row; std::size_t begin; std::size_t end; };
        std::vector<RowSpan> rowSpans;
        rowSpans.reserve(ordered.size());
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
                out << "<row r=\"" << span.row << "\"";
                if (const auto* dim = sheet.tryRowDimension(span.row)) {
                    if (dim->height) out << " ht=\"" << *dim->height << "\" customHeight=\"1\"";
                    if (dim->hidden) out << " hidden=\"1\"";
                    if (dim->outlineLevel) out << " outlineLevel=\"" << dim->outlineLevel << "\"";
                    if (dim->collapsed) out << " collapsed=\"1\"";
                }
                out << ">";
                for (std::size_t ci = span.begin; ci < span.end; ++ci)
                    writeCell(out, *ordered[ci], styles, date1904, sstIndex);
                out << "</row>";
            }
        };

        if (rowWorkers > 1 && rowSpans.size() > 1) {
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
            for (auto& c : chunks) xml << c.str();
        } else {
            writeRows(0, rowSpans.size(), xml);
        }
        xml << "</sheetData>";
    }

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
            if (!column.values().empty() || !column.dateGroups().empty() || column.includeBlank()) {
                xml << "<filters";
                if (column.includeBlank()) xml << " blank=\"1\"";
                xml << ">";
                for (const auto& value : column.values())
                    xml << "<filter val=\"" << xmlEscape(value) << "\"/>";
                for (const auto& item : column.dateGroups()) {
                    xml << "<dateGroupItem year=\"" << item.year << "\"";
                    if (item.month) xml << " month=\"" << *item.month << "\"";
                    if (item.day) xml << " day=\"" << *item.day << "\"";
                    if (item.hour) xml << " hour=\"" << *item.hour << "\"";
                    if (item.minute) xml << " minute=\"" << *item.minute << "\"";
                    if (item.second) xml << " second=\"" << *item.second << "\"";
                    xml << " dateTimeGrouping=\"" << dateGroupingName(item.grouping) << "\"/>";
                }
                xml << "</filters>";
            }
            if (!column.customFilters().empty()) {
                xml << "<customFilters" << (column.andMode() ? " and=\"1\"" : "") << ">";
                for (const auto& filter : column.customFilters())
                    xml << "<customFilter operator=\"" << filterOperatorName(filter.op)
                        << "\" val=\"" << xmlEscape(filter.value) << "\"/>";
                xml << "</customFilters>";
            }
            if (column.dynamicFilter()) {
                const auto& filter = *column.dynamicFilter();
                xml << "<dynamicFilter type=\"" << dynamicFilterTypeName(filter.type) << "\"";
                if (filter.value) xml << " val=\"" << *filter.value << "\"";
                if (filter.maxValue) xml << " maxVal=\"" << *filter.maxValue << "\"";
                xml << "/>";
            }
            if (column.top10Filter()) {
                const auto& filter = *column.top10Filter();
                xml << "<top10 top=\"" << (filter.top ? 1 : 0) << "\" percent=\"" << (filter.percent ? 1 : 0)
                    << "\" val=\"" << filter.value << "\"";
                if (filter.filterValue) xml << " filterVal=\"" << *filter.filterValue << "\"";
                xml << "/>";
            }
            if (column.colorFilter()) {
                const auto& filter = *column.colorFilter();
                xml << "<colorFilter dxfId=\"" << filter.dxfId << "\" cellColor=\"" << (filter.cellColor ? 1 : 0) << "\"/>";
            }
            if (column.iconFilter()) {
                const auto& filter = *column.iconFilter();
                xml << "<iconFilter iconSet=\"" << xmlEscape(filter.iconSet) << "\" iconId=\"" << filter.iconId << "\"/>";
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
    if (setup.orientation()!=xlpp::PageOrientation::Default || setup.paperSize()!=xlpp::PaperSize::Default || setup.scale()!=100 || setup.fitToPage() || setup.blackAndWhite() || setup.draft() || setup.useFirstPageNumber()) {
        xml << "<pageSetup";
        if (setup.orientation()==xlpp::PageOrientation::Portrait) xml << " orientation=\"portrait\"";
        else if (setup.orientation()==xlpp::PageOrientation::Landscape) xml << " orientation=\"landscape\"";
        if (setup.paperSize()!=xlpp::PaperSize::Default) xml << " paperSize=\"" << static_cast<unsigned>(setup.paperSize()) << "\"";
        xml << " scale=\"" << setup.scale() << "\"";
        if (setup.fitToPage()) xml << " fitToWidth=\"" << setup.fitToWidth() << "\" fitToHeight=\"" << setup.fitToHeight() << "\"";
        if (setup.blackAndWhite()) xml << " blackAndWhite=\"1\"";
        if (setup.draft()) xml << " draft=\"1\"";
        if (setup.useFirstPageNumber()) xml << " firstPageNumber=\"" << setup.firstPageNumber() << "\" useFirstPageNumber=\"1\"";
        xml << "/>";
    }
    const auto& hf = sheet.headerFooter();
    if (!hf.oddHeader().empty() || !hf.oddFooter().empty() || !hf.evenHeader().empty() || !hf.evenFooter().empty()) {
        xml << "<headerFooter differentOddEven=\"" << (hf.differentOddEven()?1:0) << "\" differentFirst=\"" << (hf.differentFirst()?1:0) << "\">";
        if (!hf.oddHeader().empty()) xml << "<oddHeader>" << xmlEscape(hf.oddHeader()) << "</oddHeader>";
        if (!hf.oddFooter().empty()) xml << "<oddFooter>" << xmlEscape(hf.oddFooter()) << "</oddFooter>";
        if (!hf.evenHeader().empty()) xml << "<evenHeader>" << xmlEscape(hf.evenHeader()) << "</evenHeader>";
        if (!hf.evenFooter().empty()) xml << "<evenFooter>" << xmlEscape(hf.evenFooter()) << "</evenFooter>";
        xml << "</headerFooter>";
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
    const auto pivotStart = sheet.generatedPivotStart();
    const auto generatedPivotCount = sheet.pivotTables().size() - (std::min)(pivotStart, sheet.pivotTables().size());
    if (generatedPivotCount != 0) {
        xml << "<pivotTableParts count=\"" << generatedPivotCount << "\">";
        for (std::size_t i = pivotStart; i < sheet.pivotTables().size(); ++i)
            xml << "<pivotTablePart r:id=\"rIdPivot" << i + 1 << "\"/>";
        xml << "</pivotTableParts>";
    }
    xml << "</worksheet>";
    return xml.str();
}


} // namespace xlpp::internal::ooxml
