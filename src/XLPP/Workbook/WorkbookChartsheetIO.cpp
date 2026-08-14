#include "WorkbookChartsheetIO.h"

#include <XLPP/Chart/Chartsheet.h>
#include "../XML/XmlUtilities.h"
#include "../XML/NumericParsing.h"

#include <sstream>
#include <stdexcept>
#include <utility>
#include <unordered_set>
#include <cctype>

namespace xlpp::internal {

struct WorkbookChartsheetIOAccess {
    static bool customViewsDirty(const Chartsheet& sheet) noexcept { return sheet.customViewsDirty_; }
    static const std::string& customViewsRawXml(const Chartsheet& sheet) noexcept { return sheet.customViewsRawXml_; }
    static void setCustomViewsRawXml(Chartsheet& sheet, std::string value) { sheet.customViewsRawXml_ = std::move(value); }
    static void setCustomViewsDirty(Chartsheet& sheet, bool value) noexcept { sheet.customViewsDirty_ = value; }
    static const std::string& sourceXml(const Chartsheet& sheet) noexcept { return sheet.sourceXml_; }
};

namespace {
std::string nsMain(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/spreadsheetml/main"
                  : "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
}
std::string nsRelsPkg(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/package/relationships"
                  : "http://schemas.openxmlformats.org/package/2006/relationships";
}
std::string nsRelsDoc(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
                  : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}
bool parseBool(std::string_view value) {
    return value == "1" || value == "true" || value == "TRUE";
}
std::string asciiFold(std::string_view value) {
    std::string folded; folded.reserve(value.size());
    for (unsigned char ch : value) folded.push_back(static_cast<char>(std::tolower(ch)));
    return folded;
}

const char* pageOrderToken(PageOrder value) {
    switch (value) {
        case PageOrder::DownThenOver: return "downThenOver";
        case PageOrder::OverThenDown: return "overThenDown";
        case PageOrder::Default: break;
    }
    return "";
}
const char* commentsToken(PageCellComments value) {
    switch (value) {
        case PageCellComments::AsDisplayed: return "asDisplayed";
        case PageCellComments::AtEnd: return "atEnd";
        case PageCellComments::Default: break;
    }
    return "";
}
const char* errorsToken(PageErrorDisplay value) {
    switch (value) {
        case PageErrorDisplay::Displayed: return "displayed";
        case PageErrorDisplay::Blank: return "blank";
        case PageErrorDisplay::Dash: return "dash";
        case PageErrorDisplay::NA: return "NA";
        case PageErrorDisplay::Default: break;
    }
    return "";
}
const char* customViewStateToken(CustomChartsheetViewState value) {
    switch (value) {
        case CustomChartsheetViewState::Hidden: return "hidden";
        case CustomChartsheetViewState::VeryHidden: return "veryHidden";
        case CustomChartsheetViewState::Visible: break;
    }
    return "visible";
}

void writePageMargins(std::ostringstream& xml, const PageMargins& margins) {
    xml << "<pageMargins left=\"" << margins.left() << "\" right=\"" << margins.right()
        << "\" top=\"" << margins.top() << "\" bottom=\"" << margins.bottom()
        << "\" header=\"" << margins.header() << "\" footer=\"" << margins.footer() << "\"/>";
}

void writePageSetup(std::ostringstream& xml, const PageSetup& setup,
                    std::string_view relationshipIdOverride = {}) {
    xml << "<pageSetup";
    if (setup.orientation() == PageOrientation::Portrait) xml << " orientation=\"portrait\"";
    else if (setup.orientation() == PageOrientation::Landscape) xml << " orientation=\"landscape\"";
    if (setup.paperSize() != PaperSize::Default) xml << " paperSize=\"" << static_cast<unsigned>(setup.paperSize()) << "\"";
    if (setup.scale() != 100) xml << " scale=\"" << setup.scale() << "\"";
    if (setup.fitToWidth() != 0) xml << " fitToWidth=\"" << setup.fitToWidth() << "\"";
    if (setup.fitToHeight() != 0) xml << " fitToHeight=\"" << setup.fitToHeight() << "\"";
    if (setup.firstPageNumber() != 1) xml << " firstPageNumber=\"" << setup.firstPageNumber() << "\"";
    if (setup.useFirstPageNumber()) xml << " useFirstPageNumber=\"1\"";
    if (setup.paperHeight()) xml << " paperHeight=\"" << xmlEscape(*setup.paperHeight()) << "\"";
    if (setup.paperWidth()) xml << " paperWidth=\"" << xmlEscape(*setup.paperWidth()) << "\"";
    if (setup.pageOrder() != PageOrder::Default) xml << " pageOrder=\"" << pageOrderToken(setup.pageOrder()) << "\"";
    if (setup.usePrinterDefaults()) xml << " usePrinterDefaults=\"" << (*setup.usePrinterDefaults() ? 1 : 0) << "\"";
    if (setup.blackAndWhite()) xml << " blackAndWhite=\"1\"";
    if (setup.draft()) xml << " draft=\"1\"";
    if (setup.cellComments() != PageCellComments::Default) xml << " cellComments=\"" << commentsToken(setup.cellComments()) << "\"";
    if (setup.errors() != PageErrorDisplay::Default) xml << " errors=\"" << errorsToken(setup.errors()) << "\"";
    if (setup.horizontalDpi()) xml << " horizontalDpi=\"" << *setup.horizontalDpi() << "\"";
    if (setup.verticalDpi()) xml << " verticalDpi=\"" << *setup.verticalDpi() << "\"";
    if (setup.copies()) xml << " copies=\"" << *setup.copies() << "\"";
    if (!relationshipIdOverride.empty()) xml << " r:id=\"" << xmlEscape(std::string(relationshipIdOverride)) << "\"";
    else if (setup.relationshipId()) xml << " r:id=\"" << xmlEscape(*setup.relationshipId()) << "\"";
    xml << "/>";
}

void writeHeaderFooter(std::ostringstream& xml, const HeaderFooter& hf) {
    xml << "<headerFooter";
    if (hf.differentOddEven()) xml << " differentOddEven=\"1\"";
    if (hf.differentFirst()) xml << " differentFirst=\"1\"";
    xml << ">";
    if (!hf.oddHeader().empty()) xml << "<oddHeader>" << xmlEscape(hf.oddHeader()) << "</oddHeader>";
    if (!hf.oddFooter().empty()) xml << "<oddFooter>" << xmlEscape(hf.oddFooter()) << "</oddFooter>";
    if (!hf.evenHeader().empty()) xml << "<evenHeader>" << xmlEscape(hf.evenHeader()) << "</evenHeader>";
    if (!hf.evenFooter().empty()) xml << "<evenFooter>" << xmlEscape(hf.evenFooter()) << "</evenFooter>";
    if (!hf.firstHeader().empty()) xml << "<firstHeader>" << xmlEscape(hf.firstHeader()) << "</firstHeader>";
    if (!hf.firstFooter().empty()) xml << "<firstFooter>" << xmlEscape(hf.firstFooter()) << "</firstFooter>";
    xml << "</headerFooter>";
}

void parsePageMargins(PageMargins& margins, const std::string& node, std::string_view context) {
    const auto set = [&](const char* name, auto fn) {
        const auto value = attribute(node, name);
        if (!value.empty()) fn(parseDoubleExact(value, std::string(context) + " " + name));
    };
    set("left", [&](double v){ margins.setLeft(v); });
    set("right", [&](double v){ margins.setRight(v); });
    set("top", [&](double v){ margins.setTop(v); });
    set("bottom", [&](double v){ margins.setBottom(v); });
    set("header", [&](double v){ margins.setHeader(v); });
    set("footer", [&](double v){ margins.setFooter(v); });
}

void parsePageSetup(PageSetup& setup, const std::string& node, std::string_view context) {
    const auto orientation = attribute(node, "orientation");
    if (orientation == "portrait") setup.setOrientation(PageOrientation::Portrait);
    else if (orientation == "landscape") setup.setOrientation(PageOrientation::Landscape);
    const auto paper = attribute(node, "paperSize");
    if (!paper.empty()) setup.setPaperSize(static_cast<PaperSize>(parseIntegerExact<unsigned>(paper, std::string(context) + " paperSize")));
    const auto scale = attribute(node, "scale");
    if (!scale.empty()) setup.setScale(parseIntegerExact<unsigned>(scale, std::string(context) + " scale"));
    const auto fw = attribute(node, "fitToWidth");
    if (!fw.empty()) setup.setFitToWidth(parseIntegerExact<unsigned>(fw, std::string(context) + " fitToWidth"));
    const auto fh = attribute(node, "fitToHeight");
    if (!fh.empty()) setup.setFitToHeight(parseIntegerExact<unsigned>(fh, std::string(context) + " fitToHeight"));
    const auto first = attribute(node, "firstPageNumber");
    if (!first.empty()) setup.setFirstPageNumber(parseIntegerExact<unsigned>(first, std::string(context) + " firstPageNumber"));
    const auto useFirst = attribute(node, "useFirstPageNumber");
    if (!useFirst.empty()) setup.setUseFirstPageNumber(parseBool(useFirst));
    const auto paperHeight = attribute(node, "paperHeight"); if (!paperHeight.empty()) setup.setPaperHeight(paperHeight);
    const auto paperWidth = attribute(node, "paperWidth"); if (!paperWidth.empty()) setup.setPaperWidth(paperWidth);
    const auto order = attribute(node, "pageOrder");
    if (order == "downThenOver") setup.setPageOrder(PageOrder::DownThenOver);
    else if (order == "overThenDown") setup.setPageOrder(PageOrder::OverThenDown);
    const auto usePrinter = attribute(node, "usePrinterDefaults"); if (!usePrinter.empty()) setup.setUsePrinterDefaults(parseBool(usePrinter));
    const auto bw = attribute(node, "blackAndWhite"); if (!bw.empty()) setup.setBlackAndWhite(parseBool(bw));
    const auto draft = attribute(node, "draft"); if (!draft.empty()) setup.setDraft(parseBool(draft));
    const auto comments = attribute(node, "cellComments");
    if (comments == "asDisplayed") setup.setCellComments(PageCellComments::AsDisplayed);
    else if (comments == "atEnd") setup.setCellComments(PageCellComments::AtEnd);
    const auto errors = attribute(node, "errors");
    if (errors == "displayed") setup.setErrors(PageErrorDisplay::Displayed);
    else if (errors == "blank") setup.setErrors(PageErrorDisplay::Blank);
    else if (errors == "dash") setup.setErrors(PageErrorDisplay::Dash);
    else if (errors == "NA") setup.setErrors(PageErrorDisplay::NA);
    const auto hdpi = attribute(node, "horizontalDpi"); if (!hdpi.empty()) setup.setHorizontalDpi(parseIntegerExact<unsigned>(hdpi, std::string(context) + " horizontalDpi"));
    const auto vdpi = attribute(node, "verticalDpi"); if (!vdpi.empty()) setup.setVerticalDpi(parseIntegerExact<unsigned>(vdpi, std::string(context) + " verticalDpi"));
    const auto copies = attribute(node, "copies"); if (!copies.empty()) setup.setCopies(parseIntegerExact<unsigned>(copies, std::string(context) + " copies"));
    const auto rid = attribute(node, "r:id"); if (!rid.empty()) setup.setRelationshipId(rid);
}

void parseHeaderFooter(HeaderFooter& hf, const std::string& node) {
    const auto oddEven = attribute(node, "differentOddEven"); if (!oddEven.empty()) hf.setDifferentOddEven(parseBool(oddEven));
    const auto first = attribute(node, "differentFirst"); if (!first.empty()) hf.setDifferentFirst(parseBool(first));
    hf.setOddHeader(tagText(node, "oddHeader"));
    hf.setOddFooter(tagText(node, "oddFooter"));
    hf.setEvenHeader(tagText(node, "evenHeader"));
    hf.setEvenFooter(tagText(node, "evenFooter"));
    hf.setFirstHeader(tagText(node, "firstHeader"));
    hf.setFirstFooter(tagText(node, "firstFooter"));
}

std::string serializeCustomViews(const Chartsheet& sheet) {
    if (!WorkbookChartsheetIOAccess::customViewsDirty(sheet) && !WorkbookChartsheetIOAccess::customViewsRawXml(sheet).empty())
        return WorkbookChartsheetIOAccess::customViewsRawXml(sheet);
    if (sheet.customViews().empty()) return {};
    std::ostringstream xml;
    xml.precision(17);
    xml << "<customSheetViews>";
    std::unordered_set<std::string> seenGuids;
    for (const auto& view : sheet.customViews()) {
        if (view.guid().empty()) throw std::invalid_argument("Custom Chartsheet view GUID cannot be empty");
        if (!seenGuids.insert(asciiFold(view.guid())).second)
            throw std::invalid_argument("Duplicate custom Chartsheet view GUID: " + view.guid());
        xml << "<customSheetView guid=\"" << xmlEscape(view.guid()) << "\"";
        if (view.scale()) xml << " scale=\"" << *view.scale() << "\"";
        if (view.state() != CustomChartsheetViewState::Visible) xml << " state=\"" << customViewStateToken(view.state()) << "\"";
        if (view.zoomToFit()) xml << " zoomToFit=\"" << (*view.zoomToFit() ? 1 : 0) << "\"";
        const bool hasChildren = view.pageMarginsValue() || view.pageSetupValue() || view.headerFooterValue();
        if (!hasChildren) { xml << "/>"; continue; }
        xml << ">";
        if (view.pageMarginsValue()) writePageMargins(xml, *view.pageMarginsValue());
        if (view.pageSetupValue()) writePageSetup(xml, *view.pageSetupValue());
        if (view.headerFooterValue()) writeHeaderFooter(xml, *view.headerFooterValue());
        xml << "</customSheetView>";
    }
    xml << "</customSheetViews>";
    return xml.str();
}

} // namespace


std::string serializeChartsheetXml(const Chartsheet& sheet, bool strict, std::string_view drawingRelationshipId,
                                   std::string_view printerSettingsRelationshipId) {
    std::ostringstream xml;
    xml.precision(17);
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><chartsheet xmlns=\"" << nsMain(strict)
        << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">";

    const auto& properties = sheet.properties();
    if (!properties.empty()) {
        xml << "<sheetPr";
        if (!properties.codeName().empty()) xml << " codeName=\"" << xmlEscape(properties.codeName()) << "\"";
        if (properties.published()) xml << " published=\"" << (*properties.published() ? 1 : 0) << "\"";
        if (properties.tabColor()) xml << "><tabColor rgb=\"" << xmlEscape(*properties.tabColor()) << "\"/></sheetPr>";
        else xml << "/>";
    }

    const auto& view = sheet.view();
    xml << "<sheetViews><sheetView workbookViewId=\"" << view.workbookViewId() << "\"";
    if (view.tabSelected()) xml << " tabSelected=\"1\"";
    if (view.zoomScale()) xml << " zoomScale=\"" << *view.zoomScale() << "\"";
    xml << " zoomToFit=\"" << (view.zoomToFit() ? 1 : 0) << "\"/></sheetViews>";

    const auto& protection = sheet.protection();
    if (protection.enabled()) {
        xml << "<sheetProtection content=\"" << (protection.content() ? 1 : 0)
            << "\" objects=\"" << (protection.objects() ? 1 : 0) << "\"";
        if (!protection.passwordHash().empty()) xml << " password=\"" << xmlEscape(protection.passwordHash()) << "\"";
        if (protection.algorithmName()) xml << " algorithmName=\"" << xmlEscape(*protection.algorithmName()) << "\"";
        if (protection.hashValue()) xml << " hashValue=\"" << xmlEscape(*protection.hashValue()) << "\"";
        if (protection.saltValue()) xml << " saltValue=\"" << xmlEscape(*protection.saltValue()) << "\"";
        if (protection.spinCount()) xml << " spinCount=\"" << *protection.spinCount() << "\"";
        xml << "/>";
    }

    const auto customViews = serializeCustomViews(sheet);
    if (!customViews.empty()) xml << customViews;
    if (sheet.hasPageMargins()) writePageMargins(xml, sheet.pageMargins());
    if (sheet.hasPageSetup()) writePageSetup(xml, sheet.pageSetup(), printerSettingsRelationshipId);
    if (sheet.hasHeaderFooter()) writeHeaderFooter(xml, sheet.headerFooter());
    xml << "<drawing r:id=\"" << xmlEscape(std::string(drawingRelationshipId)) << "\"/>";

    // Unsupported/opaque chart-sheet extension nodes remain preservation-first.
    const auto& source = WorkbookChartsheetIOAccess::sourceXml(sheet);
    if (!source.empty()) {
        for (const auto* tag : {"legacyDrawingHF", "picture", "webPublishItems", "extLst"})
            for (const auto& block : tags(source, tag)) xml << block;
    }
    xml << "</chartsheet>";
    return xml.str();
}

std::string serializeChartsheetRelationshipsXml(std::size_t drawingId, bool strict) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + nsRelsPkg(strict)
        + "\"><Relationship Id=\"rId1\" Type=\"" + nsRelsDoc(strict) + "/drawing\" Target=\"../drawings/drawing"
        + std::to_string(drawingId) + ".xml\"/></Relationships>";
}

std::string serializeChartsheetDrawingXml(bool strict) {
    const auto drawingNs = strict ? "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" : "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const auto drawingMainNs = strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main";
    const auto chartNs = strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart";
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"" << drawingNs
        << "\" xmlns:a=\"" << drawingMainNs << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">"
        << "<xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/><xdr:ext cx=\"0\" cy=\"0\"/>"
        << "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"1\" name=\"Chart 1\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
        << "<xdr:xfrm/><a:graphic><a:graphicData uri=\"" << chartNs << "\"><c:chart xmlns:c=\"" << chartNs
        << "\" r:id=\"rId1\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>";
    return xml.str();
}

std::string serializeChartsheetDrawingRelationshipsXml(std::size_t chartId, bool strict) {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + nsRelsPkg(strict)
        + "\"><Relationship Id=\"rId1\" Type=\"" + nsRelsDoc(strict) + "/chart\" Target=\"../charts/chart"
        + std::to_string(chartId) + ".xml\"/></Relationships>";
}

void parseChartsheetModel(Chartsheet& sheet, const std::string& xml) {
    for (const auto& pr : tags(xml, "sheetPr")) {
        auto& properties = sheet.properties();
        properties.setCodeName(attribute(pr, "codeName"));
        const auto published = attribute(pr, "published");
        if (!published.empty()) properties.setPublished(parseBool(published));
        const auto colors = tags(pr, "tabColor");
        if (!colors.empty()) {
            const auto rgb = attribute(colors.front(), "rgb");
            if (!rgb.empty()) properties.setTabColor(rgb);
        }
        break;
    }
    for (const auto& node : tags(xml, "sheetView")) {
        auto& view = sheet.view();
        const auto workbookViewId = attribute(node, "workbookViewId");
        if (!workbookViewId.empty()) view.setWorkbookViewId(parseIntegerExact<int>(workbookViewId, "chartsheet workbookViewId"));
        const auto tabSelected = attribute(node, "tabSelected"); if (!tabSelected.empty()) view.setTabSelected(parseBool(tabSelected));
        const auto zoomScale = attribute(node, "zoomScale"); if (!zoomScale.empty()) view.setZoomScale(parseIntegerExact<unsigned>(zoomScale, "chartsheet zoomScale"));
        const auto zoomToFit = attribute(node, "zoomToFit"); if (!zoomToFit.empty()) view.setZoomToFit(parseBool(zoomToFit));
        break;
    }
    for (const auto& node : tags(xml, "sheetProtection")) {
        auto& protection = sheet.protection();
        protection.setEnabled(true);
        const auto content = attribute(node, "content"); if (!content.empty()) protection.setContent(parseBool(content));
        const auto objects = attribute(node, "objects"); if (!objects.empty()) protection.setObjects(parseBool(objects));
        const auto password = attribute(node, "password"); if (!password.empty()) protection.setPasswordHash(password);
        const auto algorithm = attribute(node, "algorithmName"); if (!algorithm.empty()) protection.setAlgorithmName(algorithm);
        const auto hash = attribute(node, "hashValue"); if (!hash.empty()) protection.setHashValue(hash);
        const auto salt = attribute(node, "saltValue"); if (!salt.empty()) protection.setSaltValue(salt);
        const auto spins = attribute(node, "spinCount"); if (!spins.empty()) protection.setSpinCount(parseIntegerExact<unsigned>(spins, "chartsheet protection spinCount"));
        break;
    }
    for (const auto& block : tags(xml, "customSheetViews")) {
        WorkbookChartsheetIOAccess::setCustomViewsRawXml(sheet, block);
        auto& views = sheet.customViews();
        views.clear();
        std::unordered_set<std::string> seenGuids;
        for (const auto& node : tags(block, "customSheetView")) {
            CustomChartsheetView view;
            view.setGuid(attribute(node, "guid"));
            if (view.guid().empty()) throw std::runtime_error("Custom Chartsheet view has an empty GUID");
            if (!seenGuids.insert(asciiFold(view.guid())).second)
                throw std::runtime_error("Duplicate custom Chartsheet view GUID: " + view.guid());
            const auto scale = attribute(node, "scale"); if (!scale.empty()) view.setScale(parseIntegerExact<unsigned>(scale, "custom chartsheet view scale"));
            const auto state = attribute(node, "state");
            if (state == "hidden") view.setState(CustomChartsheetViewState::Hidden);
            else if (state == "veryHidden") view.setState(CustomChartsheetViewState::VeryHidden);
            const auto ztf = attribute(node, "zoomToFit"); if (!ztf.empty()) view.setZoomToFit(parseBool(ztf));
            for (const auto& child : tags(node, "pageMargins")) { parsePageMargins(view.pageMargins(), child, "custom chartsheet view"); break; }
            for (const auto& child : tags(node, "pageSetup")) { parsePageSetup(view.pageSetup(), child, "custom chartsheet view"); break; }
            for (const auto& child : tags(node, "headerFooter")) { parseHeaderFooter(view.headerFooter(), child); break; }
            views.push_back(std::move(view));
        }
        // Loading semantic views must not opt imported chartsheets into regeneration.
        WorkbookChartsheetIOAccess::setCustomViewsDirty(sheet, false);
        break;
    }
    std::string topLevelXml = xml;
    for (const auto& block : tags(xml, "customSheetViews")) {
        const auto pos = topLevelXml.find(block);
        if (pos != std::string::npos) topLevelXml.replace(pos, block.size(), "");
    }
    for (const auto& node : tags(topLevelXml, "pageMargins")) { parsePageMargins(sheet.pageMargins(), node, "chartsheet"); break; }
    for (const auto& node : tags(topLevelXml, "pageSetup")) { parsePageSetup(sheet.pageSetup(), node, "chartsheet"); break; }
    for (const auto& node : tags(topLevelXml, "headerFooter")) { parseHeaderFooter(sheet.headerFooter(), node); break; }
}

} // namespace xlpp::internal
