#include "OOXML/Worksheet/WorksheetReader.h"
#include "OOXML/Worksheet/WorksheetFeatureCodec.h"
#include "OOXML/Drawings/DrawingReader.h"
#include "OOXML/Charts/ChartReader.h"
#include "Package/Xml/XmlUtilities.h"
#include "Package/Zip/ZipArchive.h"
#include <XLPP/Worksheet/Worksheet.h>

#include <algorithm>
#include <cstddef>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

namespace xlpp::internal::ooxml {


namespace {
xlpp::DynamicFilterType parseDynamicFilterType(const std::string& value) {
    using T = xlpp::DynamicFilterType;
    static const std::unordered_map<std::string, T> values{
        {"aboveAverage",T::AboveAverage},{"belowAverage",T::BelowAverage},{"tomorrow",T::Tomorrow},{"today",T::Today},{"yesterday",T::Yesterday},
        {"nextWeek",T::NextWeek},{"thisWeek",T::ThisWeek},{"lastWeek",T::LastWeek},{"nextMonth",T::NextMonth},{"thisMonth",T::ThisMonth},{"lastMonth",T::LastMonth},
        {"nextQuarter",T::NextQuarter},{"thisQuarter",T::ThisQuarter},{"lastQuarter",T::LastQuarter},{"nextYear",T::NextYear},{"thisYear",T::ThisYear},{"lastYear",T::LastYear},
        {"yearToDate",T::YearToDate},{"Q1",T::Quarter1},{"Q2",T::Quarter2},{"Q3",T::Quarter3},{"Q4",T::Quarter4},
        {"M1",T::Month1},{"M2",T::Month2},{"M3",T::Month3},{"M4",T::Month4},{"M5",T::Month5},{"M6",T::Month6},
        {"M7",T::Month7},{"M8",T::Month8},{"M9",T::Month9},{"M10",T::Month10},{"M11",T::Month11},{"M12",T::Month12}
    };
    const auto it = values.find(value); return it == values.end() ? T::Today : it->second;
}
xlpp::DateTimeGrouping parseDateGrouping(const std::string& value) {
    using T = xlpp::DateTimeGrouping;
    if (value == "month") return T::Month;
    if (value == "day") return T::Day;
    if (value == "hour") return T::Hour;
    if (value == "minute") return T::Minute;
    if (value == "second") return T::Second;
    return T::Year;
}
std::optional<double> parseOptionalDoubleAttribute(const std::string& tag, const char* name) {
    const auto text = xlpp::internal::attribute(tag, name); if (text.empty()) return std::nullopt;
    try { return std::stod(text); } catch (...) { return std::nullopt; }
}
std::optional<int> parseOptionalIntAttribute(const std::string& tag, const char* name) {
    const auto text = xlpp::internal::attribute(tag, name); if (text.empty()) return std::nullopt;
    try { return std::stoi(text); } catch (...) { return std::nullopt; }
}
}

void loadWorksheetModel(xlpp::Worksheet& ws, const std::string& xml, const xlpp::internal::ZipArchive& z, const std::string& target, const StyleCatalog& styleCatalog, const std::vector<xlpp::Style>& dxfStyles, const std::vector<LoadedSharedString>& shared, bool date1904) {
    using namespace xlpp;
    const auto sheetProperties = internal::tags(xml, "sheetPr");
    if (!sheetProperties.empty()) {
        const auto codeName = internal::attribute(sheetProperties.front(), "codeName");
        if (!codeName.empty()) ws.setVbaCodeName(codeName);
    }
    const auto sheetFormats = internal::tags(xml, "sheetFormatPr");
    if (!sheetFormats.empty()) ws.setLoadedSheetFormatPrXml(sheetFormats.front());
for (const auto& margin : internal::tags(xml, "pageMargins")) {
    const auto setDouble=[&](const char* attributeName, auto setter){const auto attributeValue=internal::attribute(margin,attributeName);if(!attributeValue.empty()) setter(std::stod(attributeValue));};
    setDouble("left",[&](double v){ws.pageMargins().setLeft(v);}); setDouble("right",[&](double v){ws.pageMargins().setRight(v);});
    setDouble("top",[&](double v){ws.pageMargins().setTop(v);}); setDouble("bottom",[&](double v){ws.pageMargins().setBottom(v);});
    setDouble("header",[&](double v){ws.pageMargins().setHeader(v);}); setDouble("footer",[&](double v){ws.pageMargins().setFooter(v);});
}
for (const auto& setup : internal::tags(xml, "pageSetup")) {
    const auto orientation=internal::attribute(setup,"orientation");
    if(orientation=="portrait") ws.pageSetup().setOrientation(PageOrientation::Portrait); else if(orientation=="landscape") ws.pageSetup().setOrientation(PageOrientation::Landscape);
    const auto paper=internal::attribute(setup,"paperSize"); if(!paper.empty()) ws.pageSetup().setPaperSize(static_cast<PaperSize>(std::stoul(paper)));
    const auto scale=internal::attribute(setup,"scale"); if(!scale.empty()) ws.pageSetup().setScale(static_cast<unsigned>(std::stoul(scale)));
    const auto fw=internal::attribute(setup,"fitToWidth"), fh=internal::attribute(setup,"fitToHeight"); if(!fw.empty()){ws.pageSetup().setFitToPage(true);ws.pageSetup().setFitToWidth(static_cast<unsigned>(std::stoul(fw)));} if(!fh.empty())ws.pageSetup().setFitToHeight(static_cast<unsigned>(std::stoul(fh)));
    ws.pageSetup().setBlackAndWhite(internal::attribute(setup,"blackAndWhite")=="1"); ws.pageSetup().setDraft(internal::attribute(setup,"draft")=="1");
    const auto first=internal::attribute(setup,"firstPageNumber"); if(!first.empty())ws.pageSetup().setFirstPageNumber(static_cast<unsigned>(std::stoul(first))); ws.pageSetup().setUseFirstPageNumber(internal::attribute(setup,"useFirstPageNumber")=="1");
}
for (const auto& printOptions : internal::tags(xml, "printOptions")) { ws.printOptions().setHorizontalCentered(internal::attribute(printOptions,"horizontalCentered")=="1"); ws.printOptions().setVerticalCentered(internal::attribute(printOptions,"verticalCentered")=="1"); ws.printOptions().setHeadings(internal::attribute(printOptions,"headings")=="1"); ws.printOptions().setGridLines(internal::attribute(printOptions,"gridLines")=="1"); }
for (const auto& hf : internal::tags(xml, "headerFooter")) { ws.headerFooter().setDifferentOddEven(internal::attribute(hf,"differentOddEven")=="1"); ws.headerFooter().setDifferentFirst(internal::attribute(hf,"differentFirst")=="1"); ws.headerFooter().setOddHeader(internal::tagText(hf,"oddHeader")); ws.headerFooter().setOddFooter(internal::tagText(hf,"oddFooter")); ws.headerFooter().setEvenHeader(internal::tagText(hf,"evenHeader")); ws.headerFooter().setEvenFooter(internal::tagText(hf,"evenFooter")); }
for (const auto& protectionNode : internal::tags(xml, "sheetProtection")) { ws.protection().setEnabled(true); ws.protection().setPasswordHash(internal::attribute(protectionNode,"password")); ws.protection().setSelectLockedCells(internal::attribute(protectionNode,"selectLockedCells")!="1"); ws.protection().setSelectUnlockedCells(internal::attribute(protectionNode,"selectUnlockedCells")!="1"); ws.protection().setFormatCells(internal::attribute(protectionNode,"formatCells")!="1"); ws.protection().setFormatColumns(internal::attribute(protectionNode,"formatColumns")!="1"); ws.protection().setFormatRows(internal::attribute(protectionNode,"formatRows")!="1"); ws.protection().setInsertRows(internal::attribute(protectionNode,"insertRows")!="1"); ws.protection().setInsertColumns(internal::attribute(protectionNode,"insertColumns")!="1"); ws.protection().setDeleteRows(internal::attribute(protectionNode,"deleteRows")!="1"); ws.protection().setDeleteColumns(internal::attribute(protectionNode,"deleteColumns")!="1"); ws.protection().setSort(internal::attribute(protectionNode,"sort")!="1"); ws.protection().setAutoFilter(internal::attribute(protectionNode,"autoFilter")!="1"); }
for (const auto& sv : internal::tags(xml, "sheetView")) {
    auto& view = ws.sheetView();
    const auto workbookViewId = internal::attribute(sv, "workbookViewId");
    if (!workbookViewId.empty()) view.setWorkbookViewId(static_cast<int>(std::stoul(workbookViewId)));
    const auto zoom = internal::attribute(sv, "zoomScale");
    if (!zoom.empty()) view.setZoomScale(static_cast<int>(std::stoul(zoom)));
    const auto normalZoom = internal::attribute(sv, "zoomScaleNormal");
    if (!normalZoom.empty()) view.setZoomScaleNormal(static_cast<int>(std::stoul(normalZoom)));
    view.setShowGridLines(internal::attribute(sv, "showGridLines") != "0");
    view.setTabSelected(internal::attribute(sv, "tabSelected") == "1");
    view.setRightToLeft(internal::attribute(sv, "rightToLeft") == "1");
    view.setShowOutlineSymbols(internal::attribute(sv, "showOutlineSymbols") != "0");
}
for (const auto& tc : internal::tags(xml, "tabColor")) {
    const auto rgb = internal::attribute(tc, "rgb");
    if (!rgb.empty()) ws.sheetView().setTabColor(std::string(rgb));
}
for (auto& pane : internal::tags(xml, "pane")) {
    const auto state = internal::attribute(pane, "state");
    const auto topLeft = internal::attribute(pane, "topLeftCell");
    if (state == "frozen" && !topLeft.empty()) ws.freezePanes(topLeft);
    auto& view = ws.sheetView();
    const auto activePane = internal::attribute(pane, "activePane");
    if (!activePane.empty()) view.setPane(activePane);
    if (!topLeft.empty()) view.setTopLeftCell(topLeft);
    const auto xSplit = internal::attribute(pane, "xSplit");
    if (!xSplit.empty()) view.setXSplit(static_cast<int>(std::stod(xSplit)));
    const auto ySplit = internal::attribute(pane, "ySplit");
    if (!ySplit.empty()) view.setYSplit(static_cast<int>(std::stod(ySplit)));
}
for (auto& col : internal::tags(xml, "col")) {
    const auto minText = internal::attribute(col, "min");
    if (minText.empty()) continue;
    const auto minColumn = static_cast<std::size_t>(std::stoul(minText));
    const auto maxText = internal::attribute(col, "max");
    const auto maxColumn = maxText.empty() ? minColumn : static_cast<std::size_t>(std::stoul(maxText));
    if (minColumn == 0 || minColumn > maxColumn) throw std::runtime_error("Malformed sheet: invalid column range");
    if (maxColumn - minColumn + 1 > 1048576u) throw std::runtime_error("Malformed sheet: column range too large");
    for (std::size_t column = minColumn; column <= maxColumn; ++column) {
        auto& dimension = ws.columnDimension(column);
        const auto width = internal::attribute(col, "width");
        if (!width.empty()) dimension.width = std::stod(width);
        dimension.hidden = internal::attribute(col, "hidden") == "1";
        dimension.bestFit = internal::attribute(col, "bestFit") == "1";
        const auto outline = internal::attribute(col, "outlineLevel");
        if (!outline.empty()) dimension.outlineLevel = static_cast<int>(std::stoul(outline));
        dimension.collapsed = internal::attribute(col, "collapsed") == "1";
    }
}
for (auto& row : internal::tags(xml, "row")) {
    const auto indexText = internal::attribute(row, "r");
    if (indexText.empty()) continue;
    auto& dimension = ws.rowDimension(static_cast<std::size_t>(std::stoul(indexText)));
    const auto height = internal::attribute(row, "ht");
    if (!height.empty()) dimension.height = std::stod(height);
    dimension.hidden = internal::attribute(row, "hidden") == "1";
    const auto outline = internal::attribute(row, "outlineLevel");
    if (!outline.empty()) dimension.outlineLevel = static_cast<int>(std::stoul(outline));
    dimension.collapsed = internal::attribute(row, "collapsed") == "1";
}

for (auto& autoFilterTag : internal::tags(xml, "autoFilter")) {
    auto& autoFilter = ws.autoFilter();
    autoFilter.setReference(internal::attribute(autoFilterTag, "ref"));
    for (auto& columnTag : internal::tags(autoFilterTag, "filterColumn")) {
        const auto columnIdText = internal::attribute(columnTag, "colId");
        if (columnIdText.empty()) continue;
        auto& column = autoFilter.column(static_cast<std::size_t>(std::stoul(columnIdText)));
        for (auto& filtersTag : internal::tags(columnTag, "filters")) {
            column.setIncludeBlank(internal::attribute(filtersTag, "blank") == "1");
            for (auto& filterTag : internal::tags(filtersTag, "filter"))
                column.addValue(internal::attribute(filterTag, "val"));
            for (auto& dateTag : internal::tags(filtersTag, "dateGroupItem")) {
                DateGroupItem item;
                const auto year = internal::attribute(dateTag, "year"); if (!year.empty()) item.year = std::stoi(year);
                item.month = parseOptionalIntAttribute(dateTag, "month"); item.day = parseOptionalIntAttribute(dateTag, "day");
                item.hour = parseOptionalIntAttribute(dateTag, "hour"); item.minute = parseOptionalIntAttribute(dateTag, "minute");
                item.second = parseOptionalIntAttribute(dateTag, "second"); item.grouping = parseDateGrouping(internal::attribute(dateTag, "dateTimeGrouping"));
                column.addDateGroup(std::move(item));
            }
        }
        for (auto& customFiltersTag : internal::tags(columnTag, "customFilters")) {
            column.setAndMode(internal::attribute(customFiltersTag, "and") == "1");
            for (auto& customFilterTag : internal::tags(customFiltersTag, "customFilter"))
                column.addCustomFilter(parseFilterOperator(internal::attribute(customFilterTag, "operator")),
                                       internal::attribute(customFilterTag, "val"));
        }
        for (auto& tag : internal::tags(columnTag, "dynamicFilter")) {
            DynamicFilter filter; filter.type = parseDynamicFilterType(internal::attribute(tag, "type"));
            filter.value = parseOptionalDoubleAttribute(tag, "val"); filter.maxValue = parseOptionalDoubleAttribute(tag, "maxVal");
            column.setDynamicFilter(std::move(filter));
        }
        for (auto& tag : internal::tags(columnTag, "top10")) {
            Top10Filter filter; filter.top = internal::attribute(tag, "top") != "0"; filter.percent = internal::attribute(tag, "percent") == "1";
            if (const auto value = parseOptionalDoubleAttribute(tag, "val")) filter.value = *value;
            filter.filterValue = parseOptionalDoubleAttribute(tag, "filterVal"); column.setTop10Filter(std::move(filter));
        }
        for (auto& tag : internal::tags(columnTag, "colorFilter")) {
            ColorFilter filter; const auto dxf = internal::attribute(tag, "dxfId"); if (!dxf.empty()) filter.dxfId = static_cast<std::size_t>(std::stoull(dxf));
            filter.cellColor = internal::attribute(tag, "cellColor") != "0"; column.setColorFilter(std::move(filter));
        }
        for (auto& tag : internal::tags(columnTag, "iconFilter")) {
            IconFilter filter; const auto iconSet = internal::attribute(tag, "iconSet"); if (!iconSet.empty()) filter.iconSet = iconSet;
            const auto id = internal::attribute(tag, "iconId"); if (!id.empty()) filter.iconId = static_cast<std::size_t>(std::stoull(id));
            column.setIconFilter(std::move(filter));
        }
    }
    for (auto& sortStateTag : internal::tags(autoFilterTag, "sortState")) {
        auto& sortState = autoFilter.sortState();
        sortState.setReference(internal::attribute(sortStateTag, "ref"));
        sortState.setCaseSensitive(internal::attribute(sortStateTag, "caseSensitive") == "1");
        for (auto& conditionTag : internal::tags(sortStateTag, "sortCondition"))
            sortState.addCondition(internal::attribute(conditionTag, "ref"),
                                   internal::attribute(conditionTag, "descending") == "1");
    }
}

for (auto& formattingTag : internal::tags(xml, "conditionalFormatting")) {
    const auto reference = internal::attribute(formattingTag, "sqref");
    if (reference.empty()) continue;
    auto& formatting = ws.conditionalFormatting().add(reference);
    for (auto& ruleTag : internal::tags(formattingTag, "cfRule")) {
        const auto type = internal::attribute(ruleTag, "type");
        const auto formulaNodes = internal::tags(ruleTag, "formula");
        std::vector<std::string> formulas;
        for (const auto& formulaNode : formulaNodes) formulas.push_back(internal::tagText(formulaNode, "formula"));
        ConditionalRule rule;
        if (type == "cellIs") {
            rule = ConditionalRule::cellIs(parseConditionalOperator(internal::attribute(ruleTag, "operator")), formulas.empty() ? std::string{} : formulas.front());
            rule.setFormulas(std::move(formulas));
        } else if (type == "dataBar") {
            rule = ConditionalRule::dataBar();
            for (const auto& dbTag : internal::tags(ruleTag, "dataBar")) {
                const auto dir = internal::attribute(dbTag, "direction");
                if (!dir.empty()) rule.getDataBar().direction = dir;
                rule.getDataBar().showValue = internal::attribute(dbTag, "showValue") != "0";
                const auto cfvoTags = internal::tags(dbTag, "cfvo");
                if (cfvoTags.size() >= 1) rule.getDataBar().min = parseCfvo(cfvoTags[0]);
                if (cfvoTags.size() >= 2) rule.getDataBar().max = parseCfvo(cfvoTags[1]);
                const auto colorTags = internal::tags(dbTag, "color");
                if (!colorTags.empty()) rule.getDataBar().color = internal::attribute(colorTags.front(), "rgb");
            }
        } else if (type == "colorScale") {
            rule = ConditionalRule::colorScale();
            for (const auto& csTag : internal::tags(ruleTag, "colorScale")) {
                const auto cfvoTags = internal::tags(csTag, "cfvo");
                const auto colorTags = internal::tags(csTag, "color");
                std::size_t idx = 0;
                for (const auto& cfvoTag : cfvoTags) {
                    auto stop = parseCfvo(cfvoTag);
                    if (idx < colorTags.size()) stop.color = internal::attribute(colorTags[idx], "rgb");
                    rule.getColorScale().addStop(std::move(stop));
                    ++idx;
                }
            }
        } else if (type == "iconSet") {
            rule = ConditionalRule::iconSet();
            for (const auto& isTag : internal::tags(ruleTag, "iconSet")) {
                rule.getIconSet().icons = internal::attribute(isTag, "iconSet");
                rule.getIconSet().reverse = internal::attribute(isTag, "reverse") == "1";
                rule.getIconSet().showValue = internal::attribute(isTag, "showValue") != "0";
                for (const auto& cfvoTag : internal::tags(isTag, "cfvo"))
                    rule.getIconSet().addThreshold(parseCfvo(cfvoTag));
            }
        } else {
            rule = ConditionalRule::formula(formulas.empty() ? std::string{} : formulas.front());
            rule.setFormulas(std::move(formulas));
        }
        const auto priority = internal::attribute(ruleTag, "priority");
        if (!priority.empty()) rule.setPriority(static_cast<std::size_t>(std::stoul(priority)));
        rule.setStopIfTrue(internal::attribute(ruleTag, "stopIfTrue") == "1");
        const auto dxfId = internal::attribute(ruleTag, "dxfId");
        if (!dxfId.empty()) {
            const auto id = static_cast<std::size_t>(std::stoul(dxfId));
            if (id < dxfStyles.size()) rule.setDifferentialStyle(dxfStyles[id]);
        }
        formatting.addRule(std::move(rule));
    }
}

for (auto& validationTag : internal::tags(xml, "dataValidation")) {
    const auto reference = internal::attribute(validationTag, "sqref");
    if (reference.empty()) continue;
    DataValidation validation(parseDataValidationType(internal::attribute(validationTag, "type")));
    validation.setReference(reference);
    validation.setOperator(parseDataValidationOperator(internal::attribute(validationTag, "operator")));
    validation.setErrorStyle(parseDataValidationErrorStyle(internal::attribute(validationTag, "errorStyle")));
    validation.setAllowBlank(internal::attribute(validationTag, "allowBlank") == "1");
    validation.setShowDropDown(internal::attribute(validationTag, "showDropDown") == "1");
    validation.setShowInputMessage(internal::attribute(validationTag, "showInputMessage") == "1");
    validation.setShowErrorMessage(internal::attribute(validationTag, "showErrorMessage") == "1");
    validation.setPromptTitle(internal::attribute(validationTag, "promptTitle"));
    validation.setPrompt(internal::attribute(validationTag, "prompt"));
    validation.setErrorTitle(internal::attribute(validationTag, "errorTitle"));
    validation.setError(internal::attribute(validationTag, "error"));
    validation.setFormula1(internal::tagText(validationTag, "formula1"));
    validation.setFormula2(internal::tagText(validationTag, "formula2"));
    ws.dataValidations().add(std::move(validation));
}

{
    const auto slash = target.find_last_of('/');
    const auto fileName = slash == std::string::npos ? target : target.substr(slash + 1);
    const auto relPath = "xl/worksheets/_rels/" + fileName + ".rels";
    std::unordered_map<std::string,std::string> tableTargets;
    if (z.contains(relPath)) {
        for (const auto& rel : internal::tags(z.get(relPath), "Relationship"))
            tableTargets[internal::attribute(rel,"Id")] = internal::attribute(rel,"Target");
        for (const auto& rel : internal::tags(z.get(relPath), "Relationship")) {
            if (internal::attribute(rel, "Type").find("/comments") == std::string::npos) continue;
            auto commentsTarget = internal::attribute(rel, "Target");
            if (commentsTarget.rfind("/", 0) == 0) commentsTarget = commentsTarget.substr(1);            // "/xl/..." absolute
            else if (commentsTarget.rfind("../", 0) == 0) commentsTarget = "xl/" + commentsTarget.substr(3); // relative to worksheets/
            else commentsTarget = "xl/worksheets/" + commentsTarget;
            if (!z.contains(commentsTarget)) continue;
            const auto commentsText = z.get(commentsTarget);
            std::vector<std::string> authors;
            for (const auto& authorNode : internal::tags(commentsText, "author")) authors.push_back(internal::tagText(authorNode, "author"));
            for (const auto& commentNode : internal::tags(commentsText, "comment")) {
                const auto ref = internal::attribute(commentNode, "ref");
                if (ref.empty()) continue;
                std::string author;
                const auto authorIdText = internal::attribute(commentNode, "authorId");
                if (!authorIdText.empty()) {
                    const auto authorId = static_cast<std::size_t>(std::stoul(authorIdText));
                    if (authorId < authors.size()) author = authors[authorId];
                }
                const auto textNodes = internal::tags(commentNode, "t");
                std::string text;
                for (const auto& textNode : textNodes) text += internal::tagText(textNode, "t");
                ws.cell(ref).setComment(Comment(std::move(text), author));
            }
        }
        for (const auto& part : internal::tags(xml, "tablePart")) {
            auto tableTarget = tableTargets[internal::attribute(part,"r:id")];
            if (tableTarget.rfind("/",0)==0) tableTarget = tableTarget.substr(1);
            else if (tableTarget.rfind("../",0)==0) tableTarget = "xl/" + tableTarget.substr(3);
            else tableTarget = "xl/worksheets/" + tableTarget;
            if (!z.contains(tableTarget)) continue;
            const auto tableText = z.get(tableTarget);
            const auto tableNodes = internal::tags(tableText,"table");
            if (tableNodes.empty()) continue;
            const auto& tableNode = tableNodes.front();
            auto& table = ws.addTable(internal::attribute(tableNode,"name"), internal::attribute(tableNode,"ref"));
            const auto displayName = internal::attribute(tableNode,"displayName"); if(!displayName.empty()) table.setDisplayName(displayName);
            table.setShowHeaderRow(internal::attribute(tableNode,"headerRowCount") != "0");
            table.setShowTotalsRow(internal::attribute(tableNode,"totalsRowShown") == "1");
            for (const auto& columnNode : internal::tags(tableText,"tableColumn")) table.addColumn(internal::attribute(columnNode,"name"));
            const auto styleNodes = internal::tags(tableText,"tableStyleInfo");
            if(!styleNodes.empty()) { const auto& style=styleNodes.front(); table.styleInfo().setName(internal::attribute(style,"name")); table.styleInfo().setShowFirstColumn(internal::attribute(style,"showFirstColumn")=="1"); table.styleInfo().setShowLastColumn(internal::attribute(style,"showLastColumn")=="1"); table.styleInfo().setShowRowStripes(internal::attribute(style,"showRowStripes")!="0"); table.styleInfo().setShowColumnStripes(internal::attribute(style,"showColumnStripes")=="1"); }
        }
    }
    // Hyperlinks are parsed unconditionally: external links carry an r:id into
    // the sheet relationships, while internal links use the location attribute
    // and need no relationships part at all.
    for (const auto& linkNode : internal::tags(xml, "hyperlink")) {
        const auto ref=internal::attribute(linkNode,"ref"); if(ref.empty()) continue;
        Hyperlink link; const auto hyperlinkRelationshipId=internal::attribute(linkNode,"r:id");
        if(!hyperlinkRelationshipId.empty()){link.setTarget(tableTargets[hyperlinkRelationshipId]);link.setExternal(true);} else {link.setTarget(internal::attribute(linkNode,"location"));link.setExternal(false);}
        link.setDisplay(internal::attribute(linkNode,"display")); link.setTooltip(internal::attribute(linkNode,"tooltip")); ws.cell(ref).setHyperlink(std::move(link));
    }
}

xlpp::internal::ooxml::loadWorksheetImages(ws, xml, z, target);
xlpp::internal::ooxml::loadWorksheetCharts(ws, xml, z, target);

for (auto& merge : internal::tags(xml, "mergeCell")) {
    const auto ref = internal::attribute(merge, "ref");
    if (!ref.empty()) ws.mergeCells(ref);
}
for (auto& c : internal::tags(xml, "c")) {
    const auto a = internal::attribute(c, "r");
    const auto t = internal::attribute(c, "t");
    auto& cell = ws.cell(a);
    const auto styleText = internal::attribute(c, "s");
    if (!styleText.empty()) {
        const auto styleId = static_cast<std::size_t>(std::stoul(styleText));
        cell.setRawStyleIndex(styleId);
        if (styleId < styleCatalog.items.size()) cell.style() = styleCatalog.items[styleId];
    }
    const auto formulaTags = internal::tags(c, "f");
    if (!formulaTags.empty()) {
        const auto& formulaTag = formulaTags.front();
        auto formulaText = internal::tagText(c, "f");
        cell.setFormula(formulaText);
        const auto formulaType = internal::attribute(formulaTag, "t");
        // Detect Excel 365 dynamic array formulas: _xlfn. prefix + aca="1"
        if (formulaText.rfind("_xlfn.", 0) == 0 &&
            internal::attribute(formulaTag, "aca") == "1")
            cell.formulaMetadata().setType(FormulaType::DynamicArray);
        else if (formulaType == "shared") cell.formulaMetadata().setType(FormulaType::Shared);
        else if (formulaType == "array") cell.formulaMetadata().setType(FormulaType::Array);
        else if (formulaType == "dataTable") cell.formulaMetadata().setType(FormulaType::DataTable);
        const auto reference = internal::attribute(formulaTag, "ref");
        if (!reference.empty()) cell.formulaMetadata().setReference(reference);
        const auto sharedIndex = internal::attribute(formulaTag, "si");
        if (!sharedIndex.empty()) cell.formulaMetadata().setSharedIndex(static_cast<unsigned>(std::stoul(sharedIndex)));
        cell.formulaMetadata().setAlwaysCalculateArray(internal::attribute(formulaTag, "aca") == "1");
        cell.formulaMetadata().setCalculateOnLoad(internal::attribute(formulaTag, "ca") == "1");
    }
    if (t == "inlineStr") {
        if (auto richText = parseRichTextRuns(c)) cell.setRichText(std::move(*richText));
        else cell.setValue(internal::tagText(c, "t"));
    }
    else {
        const auto v = internal::tagText(c, "v");
        if (t == "s" && !v.empty()) {
            const auto i = std::stoul(v);
            if (i < shared.size()) {
                if (shared[i].richText) cell.setRichText(*shared[i].richText);
                else cell.setValue(shared[i].plainText);
            }
        }
        else if (t == "b") cell.setValue(v == "1");
        else if (t == "e") cell.setError(cellErrorFromString(v));
        else if (!v.empty()) {
            const auto number = std::stod(v);
            if (xlpp::isDateFormatCode(cell.style().numberFormat(), cell.style().numFmtId()))
                cell.setValue(xlpp::fromExcelSerial(number, date1904));
            else
                cell.setValue(number);
        }
    }
    }
}


} // namespace xlpp::internal::ooxml
