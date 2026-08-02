#include <XLPP/Workbook/Workbook.h>
#include "../XML/XmlUtilities.h"
#include "../Packaging/ZipArchive.h"
#include "../Threading/ThreadPool.h"
#include <algorithm>
#include <fstream>
#include <sstream>
#include <set>
#include <stdexcept>
#include <thread>
#include <unordered_map>
#include <vector>
#include <iomanip>

using xlpp::internal::xmlEscape;
namespace {

// Map public compression enums onto zlib values without exposing zlib.h.
int zlibLevel(xlpp::CompressionLevel level) {
    switch (level) {
        case xlpp::CompressionLevel::Store: return 0;
        case xlpp::CompressionLevel::Fastest: return 1;
        case xlpp::CompressionLevel::Default: return -1;
        case xlpp::CompressionLevel::Best: return 9;
    }
    return -1;
}

int zlibStrategy(xlpp::CompressionStrategy strategy) {
    switch (strategy) {
        case xlpp::CompressionStrategy::Default: return 0;
        case xlpp::CompressionStrategy::Filtered: return 1;
        case xlpp::CompressionStrategy::HuffmanOnly: return 2;
        case xlpp::CompressionStrategy::Rle: return 3;
        case xlpp::CompressionStrategy::Fixed: return 4;
    }
    return 0;
}
// Namespace URI pairs for transitional (default) and strict OOXML (ISO 29500
// strict). Strict uses the purl.oclc.org base URIs; transitional uses the
// schemas.openxmlformats.org URIs.
std::string nsMain(bool strict){return strict?std::string("http://purl.oclc.org/ooxml/spreadsheetml/main"):std::string("http://schemas.openxmlformats.org/spreadsheetml/2006/main");}
std::string nsRelsPkg(bool strict){return strict?std::string("http://purl.oclc.org/ooxml/package/relationships"):std::string("http://schemas.openxmlformats.org/package/2006/relationships");}
std::string nsRelsDoc(bool strict){return strict?std::string("http://purl.oclc.org/ooxml/officeDocument/relationships"):std::string("http://schemas.openxmlformats.org/officeDocument/2006/relationships");}
std::string nsCtPkg(bool strict){return strict?std::string("http://purl.oclc.org/ooxml/package/content-types"):std::string("http://schemas.openxmlformats.org/package/2006/content-types");}
std::string nsCoreProps(bool strict){return strict?std::string("http://purl.oclc.org/ooxml/package/metadata/core-properties"):std::string("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");}
std::string nsExtProps(bool strict){return strict?std::string("http://purl.oclc.org/ooxml/officeDocument/extended-properties"):std::string("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");}
std::string nsVTypes(bool strict){return strict?std::string("http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"):std::string("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");}

std::string contentTypes(std::size_t sheetCount, std::size_t tableCount, std::size_t commentCount, std::size_t drawingCount, const std::vector<xlpp::PreservedPart>& preserved, bool strict, bool hasSst = false, std::size_t chartCount = 0, std::size_t pivotCount = 0) {
    std::ostringstream xml;
    xml << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns=")" << nsCtPkg(strict) << R"("><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>)";
    for (std::size_t index = 1; index <= sheetCount; ++index)
        xml << "<Override PartName=\"/xl/worksheets/sheet" << index
            << ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
    for (std::size_t index = 1; index <= tableCount; ++index)
        xml << "<Override PartName=\"/xl/tables/table" << index << ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>";
    for (std::size_t index = 1; index <= commentCount; ++index)
        xml << "<Override PartName=\"/xl/comments" << index << ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>";
    for (std::size_t index = 1; index <= drawingCount; ++index)
        xml << "<Override PartName=\"/xl/drawings/drawing" << index << ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>";
    if (hasSst)
        xml << "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
    for (std::size_t i = 1; i <= chartCount; ++i)
        xml << "<Override PartName=\"/xl/charts/chart" << i << ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>";
    for (std::size_t i = 1; i <= pivotCount; ++i) {
        xml << "<Override PartName=\"/xl/pivotTables/pivotTable" << i << ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"/>";
        xml << "<Override PartName=\"/xl/pivotCache/pivotCacheDefinition" << i << ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/>";
    }
    const std::set<std::string> builtInDefaults{"rels", "xml", "vml", "png", "jpg", "jpeg"};
    for (const auto& part : preserved) {
        if (!part.overrideType.empty())
            xml << "<Override PartName=\"/" << xmlEscape(part.name) << "\" ContentType=\"" << xmlEscape(part.overrideType) << "\"/>";
        else if (!part.defaultType.empty() && !builtInDefaults.count(part.extension))
            xml << "<Default Extension=\"" << xmlEscape(part.extension) << "\" ContentType=\"" << xmlEscape(part.defaultType) << "\"/>";
    }
    xml << "</Types>";
    return xml.str();
}

// Package parts that `save` regenerates from the model; everything else is
// preserved verbatim across a load/save round trip.
bool isRegeneratedPart(const std::string& name) {
    return name == "[Content_Types].xml" || name == "_rels/.rels"
        || name == "docProps/core.xml" || name == "docProps/app.xml" || name == "docProps/custom.xml"
        || name == "xl/workbook.xml" || name == "xl/_rels/workbook.xml.rels"
        || name == "xl/styles.xml" || name == "xl/sharedStrings.xml"
        || name.rfind("xl/worksheets/", 0) == 0 || name.rfind("xl/tables/", 0) == 0
        || name.rfind("xl/charts/", 0) == 0 || name.rfind("xl/pivotTables/", 0) == 0
        || name.rfind("xl/pivotCache/", 0) == 0
        || name.rfind("xl/comments", 0) == 0 || name.rfind("xl/drawings/", 0) == 0
        || name.rfind("xl/media/", 0) == 0;
}

// Effective content types for preserved parts: <Override> rules win; parts
// without one fall back to the matching <Default> rule for their extension.
std::string extensionOf(const std::string& name) {
    const auto dot = name.rfind('.');
    if (dot == std::string::npos || dot + 1 == name.size()) return {};
    return name.substr(dot + 1);
}



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

std::string rootrels(bool strict){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\""<<nsRelsPkg(strict)<<"\"><Relationship Id=\"rId1\" Type=\""<<nsRelsDoc(strict)<<"/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\""<<nsRelsPkg(strict)<<"/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\""<<nsRelsDoc(strict)<<"/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>";return x.str();}
std::string corePropertiesXml(const xlpp::DocumentProperties& p, bool strict){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\""<<nsCoreProps(strict)<<"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"; if(!p.title().empty())x<<"<dc:title>"<<xmlEscape(p.title())<<"</dc:title>";if(!p.subject().empty())x<<"<dc:subject>"<<xmlEscape(p.subject())<<"</dc:subject>";if(!p.creator().empty())x<<"<dc:creator>"<<xmlEscape(p.creator())<<"</dc:creator>";if(!p.description().empty())x<<"<dc:description>"<<xmlEscape(p.description())<<"</dc:description>";if(!p.keywords().empty())x<<"<cp:keywords>"<<xmlEscape(p.keywords())<<"</cp:keywords>";if(!p.category().empty())x<<"<cp:category>"<<xmlEscape(p.category())<<"</cp:category>";if(!p.lastModifiedBy().empty())x<<"<cp:lastModifiedBy>"<<xmlEscape(p.lastModifiedBy())<<"</cp:lastModifiedBy>";x<<"</cp:coreProperties>";return x.str();}
std::string appPropertiesXml(bool strict){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\""<<nsExtProps(strict)<<"\" xmlns:vt=\""<<nsVTypes(strict)<<"\"><Application>XL++</Application><AppVersion>1.0</AppVersion></Properties>";return x.str();}
std::string customPropertiesXml(const xlpp::CustomProperties& props){std::ostringstream x;x<<R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">)";for(std::size_t i=0;i<props.items().size();++i){const auto&p=props.items()[i];x<<"<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\""<<(i+2)<<"\" name=\""<<xmlEscape(p.name())<<"\"><vt:"<<xmlEscape(p.type())<<">"<<xmlEscape(p.value())<<"</vt:"<<xmlEscape(p.type())<<"></property>";}x<<"</Properties>";return x.str();}

struct DxfCatalog {
    std::vector<xlpp::Style> items;
    std::unordered_map<std::size_t, std::size_t> index_;

    std::size_t id(const xlpp::Style& style) {
        const auto h = style.hash();
        const auto it = index_.find(h);
        if (it != index_.end() && items[it->second] == style) return it->second;
        const auto pos = items.size();
        items.push_back(style);
        index_[h] = pos;
        return pos;
    }

    std::size_t find(const xlpp::Style& style) const {
        const auto h = style.hash();
        const auto it = index_.find(h);
        if (it != index_.end() && items[it->second] == style) return it->second;
        return 0;
    }
};

struct StyleCatalog {
    std::vector<xlpp::Style> items{xlpp::Style{}};
    std::unordered_map<std::size_t, std::size_t> index_;

    StyleCatalog() { index_[xlpp::Style{}.hash()] = 0; }

    std::size_t id(const xlpp::Style& style) {
        const auto h = style.hash();
        const auto it = index_.find(h);
        if (it != index_.end() && items[it->second] == style) return it->second;
        const auto pos = items.size();
        items.push_back(style);
        index_[h] = pos;
        return pos;
    }

    std::size_t find(const xlpp::Style& style) const {
        const auto h = style.hash();
        const auto it = index_.find(h);
        if (it != index_.end() && items[it->second] == style) return it->second;
        return 0;
    }
};

void writeColor(std::ostringstream& xml, const xlpp::Color& color) {
    if (!color.empty()) xml << "<color rgb=\"" << xmlEscape(color.argb()) << "\"/>";
}

void writeBorderSide(std::ostringstream& xml, const char* name, const xlpp::BorderSide& side) {
    xml << '<' << name;
    if (!side.style().empty()) xml << " style=\"" << xmlEscape(side.style()) << "\"";
    if (side.color().empty()) xml << "/>";
    else { xml << '>'; writeColor(xml, side.color()); xml << "</" << name << '>'; }
}

std::string stylesXml(const StyleCatalog& catalog, const std::vector<xlpp::NamedStyle>& namedStyles, const DxfCatalog& dxfs, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"" << nsMain(strict) << "\">";

    std::unordered_map<std::string, std::size_t> customFormatIds;
    for (const auto& style : catalog.items) {
        const auto& fmt = style.numberFormat();
        if (fmt != "General" && customFormatIds.find(fmt) == customFormatIds.end())
            customFormatIds[fmt] = 164 + customFormatIds.size();
    }
    if (!customFormatIds.empty()) {
        xml << "<numFmts count=\"" << customFormatIds.size() << "\">";
        for (const auto& [format, id] : customFormatIds)
            xml << "<numFmt numFmtId=\"" << id << "\" formatCode=\"" << xmlEscape(format) << "\"/>";
        xml << "</numFmts>";
    }

    xml << "<fonts count=\"" << catalog.items.size() << "\">";
    for (const auto& style : catalog.items) {
        const auto& font = style.font();
        xml << "<font>";
        if (font.bold()) xml << "<b/>";
        if (font.italic()) xml << "<i/>";
        if (font.underline()) xml << "<u/>";
        if (font.strike()) xml << "<strike/>";
        xml << "<sz val=\"" << font.size() << "\"/><name val=\"" << xmlEscape(font.name()) << "\"/>";
        writeColor(xml, font.color());
        xml << "</font>";
    }
    xml << "</fonts>";

    xml << "<fills count=\"" << catalog.items.size() + 1 << "\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>";
    for (std::size_t i = 1; i < catalog.items.size(); ++i) {
        const auto& fill = catalog.items[i].fill();
        xml << "<fill><patternFill patternType=\"" << xmlEscape(fill.patternType()) << "\">";
        if (!fill.foregroundColor().empty()) xml << "<fgColor rgb=\"" << xmlEscape(fill.foregroundColor().argb()) << "\"/>";
        if (!fill.backgroundColor().empty()) xml << "<bgColor rgb=\"" << xmlEscape(fill.backgroundColor().argb()) << "\"/>";
        xml << "</patternFill></fill>";
    }
    xml << "</fills>";

    xml << "<borders count=\"" << catalog.items.size() << "\">";
    for (const auto& style : catalog.items) {
        xml << "<border>";
        writeBorderSide(xml, "left", style.border().left());
        writeBorderSide(xml, "right", style.border().right());
        writeBorderSide(xml, "top", style.border().top());
        writeBorderSide(xml, "bottom", style.border().bottom());
        writeBorderSide(xml, "diagonal", style.border().diagonal());
        xml << "</border>";
    }
    xml << "</borders>";
    xml << "<cellStyleXfs count=\"" << catalog.items.size() << "\">";
    for (std::size_t i = 0; i < catalog.items.size(); ++i)
        xml << "<xf numFmtId=\"0\" fontId=\"" << i << "\" fillId=\"" << (i == 0 ? 0 : i + 1)
            << "\" borderId=\"" << i << "\"/>";
    xml << "</cellStyleXfs>";
    xml << "<cellXfs count=\"" << catalog.items.size() << "\">";
    for (std::size_t i = 0; i < catalog.items.size(); ++i) {
        const auto& style = catalog.items[i];
        std::size_t numFmtId = 0;
        const auto fmtIt = customFormatIds.find(style.numberFormat());
        if (fmtIt != customFormatIds.end()) numFmtId = fmtIt->second;
        xml << "<xf numFmtId=\"" << numFmtId << "\" fontId=\"" << i << "\" fillId=\"" << (i == 0 ? 0 : i + 1)
            << "\" borderId=\"" << i << "\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyNumberFormat=\"1\">";
        const auto& alignment = style.alignment();
        if (!alignment.horizontal().empty() || !alignment.vertical().empty() || alignment.wrapText() || alignment.shrinkToFit() || alignment.textRotation() || alignment.indent()) {
            xml << "<alignment";
            if (!alignment.horizontal().empty()) xml << " horizontal=\"" << xmlEscape(alignment.horizontal()) << "\"";
            if (!alignment.vertical().empty()) xml << " vertical=\"" << xmlEscape(alignment.vertical()) << "\"";
            if (alignment.wrapText()) xml << " wrapText=\"1\"";
            if (alignment.shrinkToFit()) xml << " shrinkToFit=\"1\"";
            if (alignment.textRotation()) xml << " textRotation=\"" << alignment.textRotation() << "\"";
            if (alignment.indent()) xml << " indent=\"" << alignment.indent() << "\"";
            xml << "/>";
        }
        if (!style.locked() || style.hidden())
            xml << "<protection locked=\"" << (style.locked() ? 1 : 0) << "\" hidden=\"" << (style.hidden() ? 1 : 0) << "\"/>";
        xml << "</xf>";
    }
    xml << "</cellXfs><cellStyles count=\"" << (namedStyles.size() + 1) << "\">";
    xml << "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>";
    for (const auto& named : namedStyles)
        xml << "<cellStyle name=\"" << xmlEscape(named.name()) << "\" xfId=\"" << catalog.find(named.style()) << "\"/>";
    xml << "</cellStyles>";
    xml << "<dxfs count=\"" << dxfs.items.size() << "\">";
    for (const auto& style : dxfs.items) {
        xml << "<dxf>";
        const auto& font = style.font();
        if (!(font == xlpp::Font{})) {
            xml << "<font>";
            if (font.bold()) xml << "<b/>";
            if (font.italic()) xml << "<i/>";
            if (font.underline()) xml << "<u/>";
            if (font.strike()) xml << "<strike/>";
            if (!font.color().empty()) writeColor(xml, font.color());
            xml << "</font>";
        }
        const auto& fill = style.fill();
        if (!(fill == xlpp::Fill{})) {
            xml << "<fill><patternFill patternType=\"" << xmlEscape(fill.patternType()) << "\">";
            if (!fill.foregroundColor().empty()) xml << "<fgColor rgb=\"" << xmlEscape(fill.foregroundColor().argb()) << "\"/>";
            if (!fill.backgroundColor().empty()) xml << "<bgColor rgb=\"" << xmlEscape(fill.backgroundColor().argb()) << "\"/>";
            xml << "</patternFill></fill>";
        }
        if (!(style.border() == xlpp::Border{})) {
            xml << "<border>";
            writeBorderSide(xml, "left", style.border().left());
            writeBorderSide(xml, "right", style.border().right());
            writeBorderSide(xml, "top", style.border().top());
            writeBorderSide(xml, "bottom", style.border().bottom());
            writeBorderSide(xml, "diagonal", style.border().diagonal());
            xml << "</border>";
        }
        if (style.numberFormat() != "General")
            xml << "<numFmt numFmtId=\"0\" formatCode=\"" << xmlEscape(style.numberFormat()) << "\"/>";
        xml << "</dxf>";
    }
    xml << "</dxfs></styleSheet>";
    return xml.str();
}


xlpp::Color parseColor(const std::string& container) {
    auto colors = xlpp::internal::tags(container, "color");
    if (colors.empty()) colors = xlpp::internal::tags(container, "fgColor");
    return colors.empty() ? xlpp::Color{} : xlpp::Color(xlpp::internal::attribute(colors.front(), "rgb"));
}

xlpp::BorderSide parseBorderSide(const std::string& borderXml, const char* name) {
    xlpp::BorderSide side;
    const auto nodes = xlpp::internal::tags(borderXml, name);
    if (!nodes.empty()) {
        side.setStyle(xlpp::internal::attribute(nodes.front(), "style"));
        side.color() = parseColor(nodes.front());
    }
    return side;
}

StyleCatalog parseStyleCatalog(const std::string& xml) {
    StyleCatalog catalog;
    catalog.items.clear();
    catalog.index_.clear();
    std::unordered_map<std::size_t, std::string> formats;
    for (const auto& node : xlpp::internal::tags(xml, "numFmt")) {
        const auto id = xlpp::internal::attribute(node, "numFmtId");
        if (!id.empty()) formats[static_cast<std::size_t>(std::stoul(id))] = xlpp::internal::attribute(node, "formatCode");
    }

    std::vector<xlpp::Font> fonts;
    const auto fontContainers = xlpp::internal::tags(xml, "fonts");
    if (!fontContainers.empty()) for (const auto& node : xlpp::internal::tags(fontContainers.front(), "font")) {
        xlpp::Font font;
        const auto names = xlpp::internal::tags(node, "name");
        if (!names.empty()) font.setName(xlpp::internal::attribute(names.front(), "val"));
        const auto sizes = xlpp::internal::tags(node, "sz");
        if (!sizes.empty()) { const auto value = xlpp::internal::attribute(sizes.front(), "val"); if (!value.empty()) font.setSize(std::stod(value)); }
        font.setBold(!xlpp::internal::tags(node, "b").empty());
        font.setItalic(!xlpp::internal::tags(node, "i").empty());
        font.setUnderline(!xlpp::internal::tags(node, "u").empty());
        font.setStrike(!xlpp::internal::tags(node, "strike").empty());
        font.color() = parseColor(node);
        fonts.push_back(font);
    }

    std::vector<xlpp::Fill> fills;
    const auto fillContainers = xlpp::internal::tags(xml, "fills");
    if (!fillContainers.empty()) for (const auto& node : xlpp::internal::tags(fillContainers.front(), "fill")) {
        xlpp::Fill fill;
        const auto patterns = xlpp::internal::tags(node, "patternFill");
        if (!patterns.empty()) {
            fill.setPatternType(xlpp::internal::attribute(patterns.front(), "patternType"));
            const auto fg = xlpp::internal::tags(patterns.front(), "fgColor");
            if (!fg.empty()) fill.foregroundColor().setArgb(xlpp::internal::attribute(fg.front(), "rgb"));
            const auto bg = xlpp::internal::tags(patterns.front(), "bgColor");
            if (!bg.empty()) fill.backgroundColor().setArgb(xlpp::internal::attribute(bg.front(), "rgb"));
        }
        fills.push_back(fill);
    }

    std::vector<xlpp::Border> borders;
    const auto borderContainers = xlpp::internal::tags(xml, "borders");
    if (!borderContainers.empty()) for (const auto& node : xlpp::internal::tags(borderContainers.front(), "border")) {
        xlpp::Border border;
        border.left() = parseBorderSide(node, "left");
        border.right() = parseBorderSide(node, "right");
        border.top() = parseBorderSide(node, "top");
        border.bottom() = parseBorderSide(node, "bottom");
        border.diagonal() = parseBorderSide(node, "diagonal");
        borders.push_back(border);
    }

    const auto xfContainers = xlpp::internal::tags(xml, "cellXfs");
    if (!xfContainers.empty()) for (const auto& node : xlpp::internal::tags(xfContainers.front(), "xf")) {
        xlpp::Style style;
        const auto fontIdText = xlpp::internal::attribute(node, "fontId");
        const auto fillIdText = xlpp::internal::attribute(node, "fillId");
        const auto borderIdText = xlpp::internal::attribute(node, "borderId");
        const auto numFmtIdText = xlpp::internal::attribute(node, "numFmtId");
        if (!fontIdText.empty()) { const auto id = static_cast<std::size_t>(std::stoul(fontIdText)); if (id < fonts.size()) style.font() = fonts[id]; }
        if (!fillIdText.empty()) { const auto id = static_cast<std::size_t>(std::stoul(fillIdText)); if (id < fills.size()) style.fill() = fills[id]; }
        if (!borderIdText.empty()) { const auto id = static_cast<std::size_t>(std::stoul(borderIdText)); if (id < borders.size()) style.border() = borders[id]; }
        if (!numFmtIdText.empty()) { const auto id = static_cast<int>(std::stoul(numFmtIdText)); style.setNumFmtId(id); const auto it = formats.find(id); if (it != formats.end()) style.setNumberFormat(it->second); }
        const auto alignments = xlpp::internal::tags(node, "alignment");
        if (!alignments.empty()) {
            const auto& a = alignments.front();
            style.alignment().setHorizontal(xlpp::internal::attribute(a, "horizontal"));
            style.alignment().setVertical(xlpp::internal::attribute(a, "vertical"));
            style.alignment().setWrapText(xlpp::internal::attribute(a, "wrapText") == "1");
            style.alignment().setShrinkToFit(xlpp::internal::attribute(a, "shrinkToFit") == "1");
            const auto rotation = xlpp::internal::attribute(a, "textRotation"); if (!rotation.empty()) style.alignment().setTextRotation(std::stoi(rotation));
            const auto indent = xlpp::internal::attribute(a, "indent"); if (!indent.empty()) style.alignment().setIndent(std::stoi(indent));
        }
        const auto protections = xlpp::internal::tags(node, "protection");
        if (!protections.empty()) {
            style.setLocked(xlpp::internal::attribute(protections.front(), "locked") != "0");
            style.setHidden(xlpp::internal::attribute(protections.front(), "hidden") == "1");
        }
        catalog.items.push_back(style);
    }
    if (catalog.items.empty()) catalog.items.push_back(xlpp::Style{});
    return catalog;
}


std::vector<xlpp::Style> parseDifferentialStyles(const std::string& xml) {
    std::vector<xlpp::Style> result;
    const auto containers = xlpp::internal::tags(xml, "dxfs");
    if (containers.empty()) return result;
    for (const auto& node : xlpp::internal::tags(containers.front(), "dxf")) {
        xlpp::Style style;
        const auto fonts = xlpp::internal::tags(node, "font");
        if (!fonts.empty()) {
            const auto& fontNode = fonts.front();
            style.font().setBold(!xlpp::internal::tags(fontNode, "b").empty());
            style.font().setItalic(!xlpp::internal::tags(fontNode, "i").empty());
            style.font().setUnderline(!xlpp::internal::tags(fontNode, "u").empty());
            style.font().setStrike(!xlpp::internal::tags(fontNode, "strike").empty());
            style.font().color() = parseColor(fontNode);
        }
        const auto fills = xlpp::internal::tags(node, "fill");
        if (!fills.empty()) {
            const auto patterns = xlpp::internal::tags(fills.front(), "patternFill");
            if (!patterns.empty()) {
                style.fill().setPatternType(xlpp::internal::attribute(patterns.front(), "patternType"));
                const auto fg = xlpp::internal::tags(patterns.front(), "fgColor");
                if (!fg.empty()) style.fill().foregroundColor().setArgb(xlpp::internal::attribute(fg.front(), "rgb"));
                const auto bg = xlpp::internal::tags(patterns.front(), "bgColor");
                if (!bg.empty()) style.fill().backgroundColor().setArgb(xlpp::internal::attribute(bg.front(), "rgb"));
            }
        }
        const auto borders = xlpp::internal::tags(node, "border");
        if (!borders.empty()) {
            style.border().left() = parseBorderSide(borders.front(), "left");
            style.border().right() = parseBorderSide(borders.front(), "right");
            style.border().top() = parseBorderSide(borders.front(), "top");
            style.border().bottom() = parseBorderSide(borders.front(), "bottom");
            style.border().diagonal() = parseBorderSide(borders.front(), "diagonal");
        }
        const auto numFmts = xlpp::internal::tags(node, "numFmt");
        if (!numFmts.empty()) style.setNumberFormat(xlpp::internal::attribute(numFmts.front(), "formatCode"));
        result.push_back(style);
    }
    return result;
}


void writeCell(std::ostringstream& xml, const xlpp::Cell& cell, const StyleCatalog& styles, bool date1904,
                const std::unordered_map<std::string, std::size_t>* sstIndex) {
    if (cell.empty()) return;
    xml << "<c r=\"" << cell.address() << "\"";
    if (cell.styleIndex()) xml << " s=\"" << *cell.styleIndex() << "\"";
    else if (!cell.style().isDefault()) {
        const auto styleId = styles.find(cell.style());
        if (styleId != 0) xml << " s=\"" << styleId << "\"";
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
                    else if (metadata.type() == xlpp::FormulaType::DataTable) xml << " t=\"dataTable\"";
                    if (!metadata.reference().empty()) xml << " ref=\"" << xmlEscape(metadata.reference()) << "\"";
                    if (metadata.sharedIndex()) xml << " si=\"" << *metadata.sharedIndex() << "\"";
                    if (metadata.alwaysCalculateArray()) xml << " aca=\"1\"";
                    if (metadata.calculateOnLoad()) xml << " ca=\"1\"";
                    xml << ">" << xmlEscape(cell.formula()) << "</f>";
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
        else if (metadata.type() == xlpp::FormulaType::DataTable) xml << " t=\"dataTable\"";
        if (!metadata.reference().empty()) xml << " ref=\"" << xmlEscape(metadata.reference()) << "\"";
        if (metadata.sharedIndex()) xml << " si=\"" << *metadata.sharedIndex() << "\"";
        if (metadata.alwaysCalculateArray()) xml << " aca=\"1\"";
        if (metadata.calculateOnLoad()) xml << " ca=\"1\"";
        xml << ">" << xmlEscape(cell.formula()) << "</f>";
    }
    if (const auto* stringValue = std::get_if<std::string>(&cell.value()))
        xml << "<is><t xml:space=\"preserve\">" << xmlEscape(*stringValue) << "</t></is>";
    else if (const auto* numberValue = std::get_if<double>(&cell.value()))
        xml << "<v>" << *numberValue << "</v>";
    else if (const auto* booleanValue = std::get_if<bool>(&cell.value()))
        xml << "<v>" << (*booleanValue ? 1 : 0) << "</v>";
    else if (const auto* errorValue = std::get_if<xlpp::CellError>(&cell.value()))
        xml << "<v>" << xmlEscape(xlpp::toString(*errorValue)) << "</v>";
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
                     const std::unordered_map<std::string, std::size_t>* sstIndex = nullptr,
                     std::size_t rowWorkers = 0) {
    std::ostringstream xml;
    xml.precision(17);
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"" << nsMain(strict) << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">";
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
    }
    xml << "</sheetView></sheetViews><sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>";
    if (sheet.sheetView().tabColor())
        xml << "<sheetPr><tabColor rgb=\"" << xmlEscape(*sheet.sheetView().tabColor()) << "\"/></sheetPr>";
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
        if (!cell.empty()) ordered.push_back(&cell);

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
    for (const auto& entry : sheet.conditionalFormatting().entries()) {
        if (entry.rules().empty()) continue;
        xml << "<conditionalFormatting sqref=\"" << xmlEscape(entry.reference()) << "\">";
        for (const auto& rule : entry.rules()) {
            xml << "<cfRule type=\"" << (rule.type() == xlpp::ConditionalRuleType::Formula ? "expression" : "cellIs") << "\"";
            if (rule.type() == xlpp::ConditionalRuleType::CellIs)
                xml << " operator=\"" << conditionalOperatorName(rule.op()) << "\"";
            if (rule.hasDifferentialStyle()) xml << " dxfId=\"" << dxfs.find(rule.differentialStyle()) << "\"";
            xml << " priority=\"" << rule.priority() << "\"";
            if (rule.stopIfTrue()) xml << " stopIfTrue=\"1\"";
            xml << ">";
            for (const auto& formula : rule.formulas()) xml << "<formula>" << xmlEscape(formula) << "</formula>";
            xml << "</cfRule>";
        }
        xml << "</conditionalFormatting>";
    }
    if (!sheet.dataValidations().empty()) {
        xml << "<dataValidations count=\"" << sheet.dataValidations().items().size() << "\">";
        for (const auto& validation : sheet.dataValidations().items()) {
            xml << "<dataValidation type=\"" << dataValidationTypeName(validation.type())
                << "\" operator=\"" << dataValidationOperatorName(validation.op())
                << "\" errorStyle=\"" << dataValidationErrorStyleName(validation.errorStyle())
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
    if (!sheet.tables().empty()) {
        xml << "<tableParts count=\"" << sheet.tables().size() << "\">";
        for (std::size_t i = 0; i < sheet.tables().size(); ++i) xml << "<tablePart r:id=\"rId" << i + 1 << "\"/>";
        xml << "</tableParts>";
    }
    bool hasComments = false; for (const auto& pair : sheet.cells()) if (pair.second.hasComment()) { hasComments = true; break; }
    if (hasComments) xml << "<legacyDrawing r:id=\"rIdCommentsVml\"/>";
    if (!sheet.images().empty()) xml << "<drawing r:id=\"rIdDrawing\"/>";
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
    if (!sheet.printArea().empty())
        xml << "<pageSetup><printArea>" << xmlEscape(sheet.printArea()) << "</printArea></pageSetup>";
    if (!sheet.printTitlesRows().empty() || !sheet.printTitlesCols().empty())
        xml << "<rowBreaks count=\"0\" manualBreakCount=\"0\"/>";
    xml << "</worksheet>";
    return xml.str();
}

std::string chartXml(const xlpp::Chart& chart, const std::string& sheetName, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    xml << "<c:chartSpace xmlns:c=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart") << "\"";
    xml << " xmlns:a=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main") << "\">";
    xml << "<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>" << xmlEscape(chart.title()) << "</a:t></a:r></a:p></c:rich></c:tx></c:title>";
    xml << "<c:plotArea><c:layout/>";
    xml << "<c:" << xlpp::Chart::typeName(chart.type(), chart.grouping()) << ">";
    for (std::size_t s = 0; s < chart.series().size(); ++s) {
        const auto& series = chart.series()[s];
        xml << "<c:ser><c:idx val=\"" << s << "\"/><c:order val=\"" << s << "\"/>";
        if (!series.title().empty())
            xml << "<c:tx><c:strRef><c:f>" << xmlEscape("'" + sheetName + "'!$A$1") << "</c:f></c:strRef></c:tx>";
        xml << "<c:spPr><a:ln><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln></c:spPr>";
        if (!series.valuesReference().empty())
            xml << "<c:val><c:numRef><c:f>" << xmlEscape(series.valuesReference()) << "</c:f></c:numRef></c:val>";
        if (!series.categoriesReference().empty())
            xml << "<c:cat><c:strRef><c:f>" << xmlEscape(series.categoriesReference()) << "</c:f></c:strRef></c:cat>";
        xml << "</c:ser>";
    }
    xml << "<c:axId val=\"1\"/><c:axId val=\"2\"/>";
    if (chart.type() == xlpp::Chart::Type::Pie) xml << "<c:firstSliceAng val=\"0\"/>";
    xml << "</c:" << xlpp::Chart::typeName(chart.type(), chart.grouping()) << ">";
    xml << "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/>";
    if (!chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx></c:title>";
    xml << "</c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/>";
    if (!chart.yAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.yAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx></c:title>";
    xml << "</c:valAx></c:plotArea>";
    if (chart.showLegend()) xml << "<c:legend><c:legendPos val=\"" << xmlEscape(chart.legendPosition()) << "\"/></c:legend>";
    xml << "</c:chart></c:chartSpace>";
    return xml.str();
}

std::string chartDrawingXml(const xlpp::Worksheet& sheet, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    xml << "<xdr:wsDr xmlns:xdr=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" : "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing") << "\"";
    xml << " xmlns:a=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main") << "\">";
    for (std::size_t i = 0; i < sheet.chartCount(); ++i) {
        const auto& chartRef = sheet.chart(i);
        xml << "<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>";
        xml << "<xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>";
        xml << "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" << (i + 10) << "\" name=\"Chart " << (i + 1) << "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>";
        xml << "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" << chartRef.width() << "\" cy=\"" << chartRef.height() << "\"/></xdr:xfrm>";
        xml << "<a:graphic><a:graphicData uri=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart") << "\">";
        xml << "<c:chart xmlns:c=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart") << "\" r:id=\"rIdChart" << (i + 1) << "\"/>";
        xml << "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>";
    }
    xml << "</xdr:wsDr>";
    return xml.str();
}

std::string pivotTableXml(const xlpp::PivotTable& pt, std::size_t id, bool strict) {
    std::ostringstream xml;
    xml << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name=")"
        << xmlEscape(pt.name()) << "\" cacheId=\"" << pt.cache().cacheId() << "\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\""
        << " applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" dataCaption=\"Values\" updatedVersion=\"6\" minRefreshableVersion=\"3\""
        << " showCalcMbrs=\"0\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" createdVersion=\"6\" indent=\"0\" compact=\"0\" compactData=\"1\"";
    if (!pt.location().empty()) xml << " ref=\"" << xmlEscape(pt.location()) << "\"";
    xml << " gridDropZones=\"1\" multipleFieldFilters=\"0\"><location firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\"";
    xml << " ref=\"" << xmlEscape(pt.location().empty() ? "A3" : pt.location()) << "\"/>";
    xml << "<pivotFields count=\"" << (pt.rowFields().size() + pt.columnFields().size() + pt.pageFields().size() + pt.dataFields().size()) << "\">";
    for (const auto& f : pt.rowFields())
        xml << "<pivotField axis=\"" << xmlEscape(f.axis()) << "\" showAll=\"" << (f.showAll() ? 1 : 0) << "\"/>";
    for (const auto& f : pt.columnFields())
        xml << "<pivotField axis=\"" << xmlEscape(f.axis()) << "\" showAll=\"" << (f.showAll() ? 1 : 0) << "\"/>";
    for (const auto& f : pt.pageFields())
        xml << "<pivotField axis=\"" << xmlEscape(f.axis()) << "\" showAll=\"" << (f.showAll() ? 1 : 0) << "\"/>";
    for (const auto& f : pt.dataFields())
        xml << "<pivotField dataField=\"1\" showAll=\"0\"/>";
    xml << "</pivotFields>";
    xml << "<rowFields count=\"" << pt.rowFields().size() << "\">";
    for (std::size_t i = 0; i < pt.rowFields().size(); ++i) xml << "<field x=\"" << i << "\"/>";
    xml << "</rowFields>";
    xml << "<colFields count=\"" << pt.columnFields().size() << "\">";
    for (std::size_t i = 0; i < pt.columnFields().size(); ++i) xml << "<field x=\"" << (pt.rowFields().size() + i) << "\"/>";
    xml << "</colFields>";
    xml << "<pageFields count=\"" << pt.pageFields().size() << "\">";
    for (std::size_t i = 0; i < pt.pageFields().size(); ++i) xml << "<pageField hier=\"-1\"/>";
    xml << "</pageFields>";
    xml << "<dataFields count=\"" << pt.dataFields().size() << "\">";
    for (std::size_t i = 0; i < pt.dataFields().size(); ++i)
        xml << "<dataField name=\"Values\" fld=\"" << (pt.rowFields().size() + pt.columnFields().size() + pt.pageFields().size() + i) << "\" baseField=\"0\" baseItem=\"0\"/>";
    xml << "</dataFields><pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\"/>";
    xml << "</pivotTableDefinition>";
    return xml.str();
}

std::string pivotCacheXml(const xlpp::PivotTable& pt, bool strict) {
    std::ostringstream xml;
    xml << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshOnLoad="1" recordCount="0">)";
    xml << "<cacheSource type=\"worksheet\"><worksheetSource ref=\"" << xmlEscape(pt.cache().sourceData()) << "\"/></cacheSource>";
    xml << "<cacheFields count=\"0\"/></pivotCacheDefinition>";
    return xml.str();
}


std::string commentsXml(const xlpp::Worksheet& sheet, bool strict) {
    std::vector<std::string> authors;
    for (const auto& pair : sheet.cells()) {
        if (!pair.second.hasComment()) continue;
        const auto& author = pair.second.commentValue()->author();
        if (std::find(authors.begin(), authors.end(), author) == authors.end()) authors.push_back(author);
    }
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><comments xmlns=\"" << nsMain(strict) << "\"><authors>";
    for (const auto& author : authors) xml << "<author>" << xmlEscape(author) << "</author>";
    xml << "</authors><commentList>";
    for (const auto& pair : sheet.cells()) {
        if (!pair.second.hasComment()) continue;
        const auto& comment = *pair.second.commentValue();
        const auto it = std::find(authors.begin(), authors.end(), comment.author());
        const auto authorId = static_cast<std::size_t>(std::distance(authors.begin(), it));
        xml << "<comment ref=\"" << pair.second.address() << "\" authorId=\"" << authorId << "\"><text><t xml:space=\"preserve\">"
            << xmlEscape(comment.text()) << "</t></text></comment>";
    }
    xml << "</commentList></comments>";
    return xml.str();
}

std::string commentsVml(const xlpp::Worksheet& sheet) {
    std::ostringstream xml;
    xml << R"(<?xml version="1.0" encoding="UTF-8"?><xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>)";
    std::size_t shapeId = 1025;
    for (const auto& pair : sheet.cells()) {
        if (!pair.second.hasComment()) continue;
        xml << "<v:shape id=\"_x0000_s" << shapeId++ << "\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>"
            << (pair.second.row() - 1) << "</x:Row><x:Column>" << (pair.second.column() - 1) << "</x:Column></x:ClientData></v:shape>";
    }
    xml << "</xml>";
    return xml.str();
}


std::string drawingXml(const xlpp::Worksheet& sheet, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" : "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
        << "\" xmlns:a=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main")
        << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">";
    std::size_t id = 1;
    for (const auto& image : sheet.images()) {
        const auto ref = xlpp::CellReference::parse(image.anchor());
        const auto cx = static_cast<long long>(image.widthPixels() * 9525.0);
        const auto cy = static_cast<long long>(image.heightPixels() * 9525.0);
        xml << "<xdr:oneCellAnchor><xdr:from><xdr:col>" << (ref.column-1) << "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << (ref.row-1) << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            << "<xdr:ext cx=\"" << cx << "\" cy=\"" << cy << "\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" << id << "\" name=\"" << xmlEscape(image.name()) << "\"/><xdr:cNvPicPr/></xdr:nvPicPr>"
            << "<xdr:blipFill><a:blip r:embed=\"rIdImage" << id << "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";
        ++id;
    }
    xml << "</xdr:wsDr>";
    return xml.str();
}
std::string drawingRelationshipsXml(const xlpp::Worksheet& sheet, std::size_t firstMediaId, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
    for (std::size_t i=0;i<sheet.images().size();++i)
        xml << "<Relationship Id=\"rIdImage" << i+1 << "\" Type=\"" << nsRelsDoc(strict) << "/image\" Target=\"../media/image" << firstMediaId+i << '.' << sheet.images()[i].extension() << "\"/>";
    xml << "</Relationships>";
    return xml.str();
}

std::string tableXml(const xlpp::Table& table, std::size_t id, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><table xmlns=\"" << nsMain(strict) << "\""
        << " id=\"" << id << "\" name=\"" << xmlEscape(table.name())
        << "\" displayName=\"" << xmlEscape(table.displayName()) << "\" ref=\""
        << xmlEscape(table.reference()) << "\" headerRowCount=\"" << (table.showHeaderRow() ? 1 : 0)
        << "\" totalsRowShown=\"" << (table.showTotalsRow() ? 1 : 0) << "\">";
    xml << "<autoFilter ref=\"" << xmlEscape(table.reference()) << "\"/>";
    xml << "<tableColumns count=\"" << table.columns().size() << "\">";
    for (const auto& column : table.columns())
        xml << "<tableColumn id=\"" << column.id() << "\" name=\"" << xmlEscape(column.name()) << "\"/>";
    xml << "</tableColumns>";
    const auto& style = table.styleInfo();
    xml << "<tableStyleInfo name=\"" << xmlEscape(style.name())
        << "\" showFirstColumn=\"" << (style.showFirstColumn() ? 1 : 0)
        << "\" showLastColumn=\"" << (style.showLastColumn() ? 1 : 0)
        << "\" showRowStripes=\"" << (style.showRowStripes() ? 1 : 0)
        << "\" showColumnStripes=\"" << (style.showColumnStripes() ? 1 : 0) << "\"/>";
    xml << "</table>";
    return xml.str();
}

// Serialize every worksheet to XML, using a ThreadPool when workers > 1.
// Output is indexed by worksheet order and is identical to the sequential result.
std::vector<std::string> serializeSheets(const std::vector<xlpp::Worksheet>& sheets,
                                         const StyleCatalog& styles, const DxfCatalog& dxfs,
                                         bool date1904, bool strict, std::size_t workers,
                                         bool parallelRows,
                                         const std::unordered_map<std::string, std::size_t>* sstIndex) {
    std::vector<std::string> result(sheets.size());
    if (workers > 1 && sheets.size() > 1) {
        xlpp::internal::ThreadPool pool(std::min(workers, sheets.size()));
        pool.parallelFor(0, sheets.size(), [&](std::size_t i) {
            result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, 0);
        });
    } else if (parallelRows && workers > 1 && sheets.size() == 1) {
        // Single large sheet: parallelize across rows within the sheet
        for (std::size_t i = 0; i < sheets.size(); ++i)
            result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, workers);
    } else {
        for (std::size_t i = 0; i < sheets.size(); ++i)
            result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, 0);
    }
    return result;
}
void parseSheet(xlpp::Worksheet& ws, const std::string& xml, const xlpp::internal::ZipArchive& z, const std::string& target, const StyleCatalog& styleCatalog, const std::vector<xlpp::Style>& dxfStyles, const std::vector<std::string>& shared, bool date1904) {
    using namespace xlpp;
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
    const auto zoom = internal::attribute(sv, "zoomScale");
    if (!zoom.empty()) ws.sheetView().setZoomScale(static_cast<int>(std::stoul(zoom)));
    ws.sheetView().setShowGridLines(internal::attribute(sv, "showGridLines") != "0");
    ws.sheetView().setTabSelected(internal::attribute(sv, "tabSelected") == "1");
    ws.sheetView().setRightToLeft(internal::attribute(sv, "rightToLeft") == "1");
}
for (const auto& tc : internal::tags(xml, "tabColor")) {
    const auto rgb = internal::attribute(tc, "rgb");
    if (!rgb.empty()) ws.sheetView().setTabColor(std::string(rgb));
}
for (auto& pane : internal::tags(xml, "pane")) {
    if (internal::attribute(pane, "state") == "frozen") {
        const auto topLeft = internal::attribute(pane, "topLeftCell");
        if (!topLeft.empty()) ws.freezePanes(topLeft);
    }
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
        }
        for (auto& customFiltersTag : internal::tags(columnTag, "customFilters")) {
            column.setAndMode(internal::attribute(customFiltersTag, "and") == "1");
            for (auto& customFilterTag : internal::tags(customFiltersTag, "customFilter"))
                column.addCustomFilter(parseFilterOperator(internal::attribute(customFilterTag, "operator")),
                                       internal::attribute(customFilterTag, "val"));
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
        ConditionalRule rule = type == "cellIs"
            ? ConditionalRule::cellIs(parseConditionalOperator(internal::attribute(ruleTag, "operator")), formulas.empty() ? std::string{} : formulas.front())
            : ConditionalRule::formula(formulas.empty() ? std::string{} : formulas.front());
        rule.setFormulas(std::move(formulas));
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
    if (z.contains(relPath)) {
        std::unordered_map<std::string,std::string> tableTargets;
        for (const auto& rel : internal::tags(z.get(relPath), "Relationship"))
            tableTargets[internal::attribute(rel,"Id")] = internal::attribute(rel,"Target");
        for (const auto& rel : internal::tags(z.get(relPath), "Relationship")) {
            if (internal::attribute(rel, "Type").find("/comments") == std::string::npos) continue;
            auto commentsTarget = internal::attribute(rel, "Target");
            if (commentsTarget.rfind("../", 0) == 0) commentsTarget = "xl/" + commentsTarget.substr(3);
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
                const auto text = textNodes.empty() ? std::string{} : internal::tagText(textNodes.front(), "t");
                ws.cell(ref).setComment(Comment(text, author));
            }
        }
        for (const auto& linkNode : internal::tags(xml, "hyperlink")) {
            const auto ref=internal::attribute(linkNode,"ref"); if(ref.empty()) continue;
            Hyperlink link; const auto hyperlinkRelationshipId=internal::attribute(linkNode,"r:id");
            if(!hyperlinkRelationshipId.empty()){link.setTarget(tableTargets[hyperlinkRelationshipId]);link.setExternal(true);} else {link.setTarget(internal::attribute(linkNode,"location"));link.setExternal(false);}
            link.setDisplay(internal::attribute(linkNode,"display")); link.setTooltip(internal::attribute(linkNode,"tooltip")); ws.cell(ref).setHyperlink(std::move(link));
        }
        for (const auto& part : internal::tags(xml, "tablePart")) {
            auto tableTarget = tableTargets[internal::attribute(part,"r:id")];
            if (tableTarget.rfind("../",0)==0) tableTarget = "xl/" + tableTarget.substr(3);
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
}

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
        cell.setFormula(internal::tagText(c, "f"));
        const auto formulaType = internal::attribute(formulaTag, "t");
        if (formulaType == "shared") cell.formulaMetadata().setType(FormulaType::Shared);
        else if (formulaType == "array") cell.formulaMetadata().setType(FormulaType::Array);
        else if (formulaType == "dataTable") cell.formulaMetadata().setType(FormulaType::DataTable);
        const auto reference = internal::attribute(formulaTag, "ref");
        if (!reference.empty()) cell.formulaMetadata().setReference(reference);
        const auto sharedIndex = internal::attribute(formulaTag, "si");
        if (!sharedIndex.empty()) cell.formulaMetadata().setSharedIndex(static_cast<unsigned>(std::stoul(sharedIndex)));
        cell.formulaMetadata().setAlwaysCalculateArray(internal::attribute(formulaTag, "aca") == "1");
        cell.formulaMetadata().setCalculateOnLoad(internal::attribute(formulaTag, "ca") == "1");
    }
    if (t == "inlineStr") cell.setValue(internal::tagText(c, "t"));
    else {
        const auto v = internal::tagText(c, "v");
        if (t == "s" && !v.empty()) { const auto i = std::stoul(v); if (i < shared.size()) cell.setValue(shared[i]); }
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
}namespace xlpp {
NamedStyle& Workbook::addNamedStyle(NamedStyle style){if(style.name().empty())throw std::invalid_argument("Named style name cannot be empty");if(namedStyle(style.name()))throw std::invalid_argument("Named style already exists: "+style.name());namedStyles_.push_back(std::move(style));return namedStyles_.back();}
NamedStyle* Workbook::namedStyle(const std::string& name) noexcept{for(auto& style:namedStyles_)if(style.name()==name)return &style;return nullptr;}
const NamedStyle* Workbook::namedStyle(const std::string& name) const noexcept{for(const auto& style:namedStyles_)if(style.name()==name)return &style;return nullptr;}
void Workbook::applyNamedStyle(Cell& cell,const std::string& name) const{const auto* style=namedStyle(name);if(!style)throw std::out_of_range("Unknown named style: "+name);cell.style()=style->style();cell.setNamedStyle(name);}
DefinedName& Workbook::addDefinedName(DefinedName name){if(definedName(name.name()))throw std::invalid_argument("Defined name already exists: "+name.name());definedNames_.push_back(std::move(name));return definedNames_.back();}
DefinedName* Workbook::definedName(const std::string& name) noexcept{for(auto& item:definedNames_)if(item.name()==name)return &item;return nullptr;}
const DefinedName* Workbook::definedName(const std::string& name) const noexcept{for(const auto& item:definedNames_)if(item.name()==name)return &item;return nullptr;}
Worksheet& Workbook::addWorksheet(std::string name){if(name.empty())throw std::invalid_argument("Worksheet name cannot be empty");if(worksheet(name))throw std::invalid_argument("Duplicate worksheet name");sheets_.emplace_back(std::move(name));return sheets_.back();}
bool Workbook::removeWorksheet(const std::string& name){const auto it=std::find_if(sheets_.begin(),sheets_.end(),[&](auto&s){return s.name()==name;});if(it==sheets_.end())return false;sheets_.erase(it);return true;}
Worksheet* Workbook::worksheet(const std::string& n)noexcept{auto i=std::find_if(sheets_.begin(),sheets_.end(),[&](auto&s){return s.name()==n;});return i==sheets_.end()?nullptr:&*i;}const Worksheet* Workbook::worksheet(const std::string& n)const noexcept{auto i=std::find_if(sheets_.begin(),sheets_.end(),[&](auto&s){return s.name()==n;});return i==sheets_.end()?nullptr:&*i;}
Worksheet& Workbook::operator[](std::size_t index){return sheets_.at(index);}
const Worksheet& Workbook::operator[](std::size_t index) const{return sheets_.at(index);}
std::size_t Workbook::index(const Worksheet& sheet) const{const auto it=std::find_if(sheets_.begin(),sheets_.end(),[&](auto&s){return &s==&sheet;});if(it==sheets_.end())throw std::out_of_range("Worksheet not in this workbook");return static_cast<std::size_t>(std::distance(sheets_.begin(),it));}
std::vector<std::string> Workbook::sheetNames() const{std::vector<std::string> names;names.reserve(sheets_.size());for(const auto& sheet:sheets_)names.push_back(sheet.name());return names;}
Worksheet& Workbook::copyWorksheet(const Worksheet& source, std::string newName){if(newName.empty())throw std::invalid_argument("Worksheet name cannot be empty");if(worksheet(newName))throw std::invalid_argument("Duplicate worksheet name");sheets_.push_back(source);sheets_.back().rename(newName);return sheets_.back();}
void Workbook::save(const std::filesystem::path& p) const { save(p, SaveOptions{}); }
void Workbook::save(const std::filesystem::path& p, const SaveOptions& options) const {
    if (sheets_.empty()) throw std::runtime_error("Workbook needs at least one worksheet");
    const bool strict = options.strictNamespace;
    StyleCatalog styleCatalog; DxfCatalog dxfCatalog;     std::size_t tableCount = 0; std::size_t commentCount = 0; std::size_t drawingCount = 0;
    std::size_t chartCount = 0; std::size_t pivotCount = 0;
    std::unordered_map<std::string, std::size_t> sstIndex;
    std::vector<std::string> sstStrings;
    std::size_t sstOccurrences = 0;
    for (const auto& sheet : sheets_) {
        tableCount += sheet.tables().size();
        chartCount += sheet.chartCount();
        pivotCount += sheet.pivotTables().size();
        bool hasComments = false; for (const auto& pair : sheet.cells()) if (pair.second.hasComment()) { hasComments = true; break; }
        if (hasComments) ++commentCount;
        if (!sheet.images().empty()) ++drawingCount;
        if (sheet.chartCount() > 0) ++drawingCount;
        for (const auto& entry : sheet.cells()) {
            if (!entry.second.style().isDefault())
                styleCatalog.id(entry.second.style());
            if (const auto* sv = std::get_if<std::string>(&entry.second.value())) {
                ++sstOccurrences;
                const auto [it, inserted] = sstIndex.try_emplace(*sv, sstStrings.size());
                if (inserted) sstStrings.push_back(*sv);
            }
        }
        for (const auto& formatting : sheet.conditionalFormatting().entries())
            for (const auto& rule : formatting.rules()) if (rule.hasDifferentialStyle()) dxfCatalog.id(rule.differentialStyle());
    }
    for (const auto& named : namedStyles_) styleCatalog.id(named.style());
    const auto sheetXmls = serializeSheets(sheets_, styleCatalog, dxfCatalog, date1904_, strict, options.parallelSheets ? options.parallelWorkers : 0, options.parallelRows, &sstIndex);
    internal::ZipArchive z;
    z.setCompressionLevel(zlibLevel(options.compressionLevel));
    z.setCompressionStrategy(zlibStrategy(options.compressionStrategy));
    z.setParallelWorkers(options.parallelWorkers);
    z.add("[Content_Types].xml", contentTypes(sheets_.size(), tableCount, commentCount, drawingCount, preservedParts_, strict, !sstStrings.empty(), chartCount, pivotCount));
    z.add("_rels/.rels", rootrels(strict));
    z.add("docProps/core.xml", corePropertiesXml(properties_, strict));
    z.add("docProps/app.xml", appPropertiesXml(strict));
    if (!customProps_.empty()) z.add("docProps/custom.xml", customPropertiesXml(customProps_));
    z.add("xl/styles.xml", stylesXml(styleCatalog, namedStyles_, dxfCatalog, strict));
    if (!sstStrings.empty()) {
        std::ostringstream sstXml;
        sstXml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"" << nsMain(strict)
               << "\" count=\"" << sstOccurrences << "\" uniqueCount=\"" << sstStrings.size() << "\">";
        for (const auto& text : sstStrings)
            sstXml << "<si><t xml:space=\"preserve\">" << xmlEscape(text) << "</t></si>";
        sstXml << "</sst>";
        z.add("xl/sharedStrings.xml", sstXml.str());
    }
    std::ostringstream wb, rels;
    wb << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"" << nsMain(strict) << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">";
    {
        const auto& cp = calcProps_;
        wb << "<workbookPr date1904=\"" << (date1904_ ? 1 : 0) << "\"";
        if (cp.calcOnSave()) wb << " calcOnSave=\"1\"";
        if (cp.fullCalcOnLoad()) wb << " fullCalcOnLoad=\"1\"";
        wb << "/>";
        if (cp.calcMode() != "auto" || cp.calcId() != 191029 || !cp.fullPrecision() || cp.iterate())
            wb << "<calcPr calcId=\"" << cp.calcId() << "\" calcMode=\"" << xmlEscape(cp.calcMode()) << "\""
               << " fullPrecision=\"" << (cp.fullPrecision() ? 1 : 0) << "\""
               << " iterate=\"" << (cp.iterate() ? 1 : 0) << "\""
               << " iterateCount=\"" << cp.iterateCount() << "\" iterateDelta=\"" << cp.iterateDelta() << "\"/>";
    }
    const auto& wp = protection_;
    if (wp.lockStructure() || wp.lockWindows() || wp.lockRevision() || !wp.workbookPasswordHash().empty()) {
        wb << "<workbookProtection lockStructure=\"" << (wp.lockStructure()?1:0) << "\" lockWindows=\"" << (wp.lockWindows()?1:0) << "\" lockRevision=\"" << (wp.lockRevision()?1:0) << "\"";
        if (!wp.workbookPasswordHash().empty()) wb << " workbookPassword=\"" << xmlEscape(wp.workbookPasswordHash()) << "\"";
        wb << "/>";
    }
    wb << "<sheets>";
    rels << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
    std::size_t globalTableId = 1;
    std::size_t globalCommentId = 1;
    std::size_t globalDrawingId = 1;
    std::size_t globalMediaId = 1;
    std::size_t globalChartId = 1;
    std::size_t globalPivotId = 1;
    bool hasCharts = false, hasPivots = false;
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        const auto& sheet = sheets_[i];
        wb << "<sheet name=\"" << xmlEscape(sheet.name()) << "\" sheetId=\"" << i+1 << "\" r:id=\"rId" << i+1 << "\"/>";
        rels << "<Relationship Id=\"rId" << i+1 << "\" Type=\"" << nsRelsDoc(strict) << "/worksheet\" Target=\"worksheets/sheet" << i+1 << ".xml\"/>";
        z.add("xl/worksheets/sheet"+std::to_string(i+1)+".xml", sheetXmls[i]);
        bool hasLinks=false; for(const auto& pair:sheet.cells()) if(pair.second.hasHyperlink() && pair.second.hyperlinkValue()->external()) hasLinks=true;
        bool hasComments=false; for(const auto& pair:sheet.cells()) if(pair.second.hasComment()) { hasComments=true; break; }
        bool hasSheetCharts = sheet.chartCount() > 0;
        bool hasSheetPivots = !sheet.pivotTables().empty();
        if (hasSheetCharts) hasCharts = true;
        if (hasSheetPivots) hasPivots = true;
        if (!sheet.tables().empty() || hasLinks || hasComments || !sheet.images().empty() || hasSheetCharts || hasSheetPivots) {
            std::ostringstream sheetRels;
            sheetRels << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
            for (std::size_t t = 0; t < sheet.tables().size(); ++t, ++globalTableId) {
                sheetRels << "<Relationship Id=\"rId" << t+1 << "\" Type=\"" << nsRelsDoc(strict) << "/table\" Target=\"../tables/table" << globalTableId << ".xml\"/>";
                z.add("xl/tables/table"+std::to_string(globalTableId)+".xml", tableXml(sheet.tables()[t], globalTableId, strict));
            }
            std::size_t hid=1; for(const auto& pair:sheet.cells()) if(pair.second.hasHyperlink() && pair.second.hyperlinkValue()->external()) { const auto& h=*pair.second.hyperlinkValue(); sheetRels<<"<Relationship Id=\"rIdHyperlink"<<hid++<<"\" Type=\""<<nsRelsDoc(strict)<<"/hyperlink\" Target=\""<<xmlEscape(h.target())<<"\" TargetMode=\"External\"/>"; }
            if (hasComments) {
                sheetRels << "<Relationship Id=\"rIdComments\" Type=\"" << nsRelsDoc(strict) << "/comments\" Target=\"../comments" << globalCommentId << ".xml\"/>";
                sheetRels << "<Relationship Id=\"rIdCommentsVml\" Type=\"" << nsRelsDoc(strict) << "/vmlDrawing\" Target=\"../drawings/commentsDrawing" << globalCommentId << ".vml\"/>";
                z.add("xl/comments" + std::to_string(globalCommentId) + ".xml", commentsXml(sheet, strict));
                z.add("xl/drawings/commentsDrawing" + std::to_string(globalCommentId) + ".vml", commentsVml(sheet));
                ++globalCommentId;
            }
            if (!sheet.images().empty()) {
                sheetRels << "<Relationship Id=\"rIdDrawing\" Type=\"" << nsRelsDoc(strict) << "/drawing\" Target=\"../drawings/drawing" << globalDrawingId << ".xml\"/>";
                z.add("xl/drawings/drawing" + std::to_string(globalDrawingId) + ".xml", drawingXml(sheet, strict));
                z.add("xl/drawings/_rels/drawing" + std::to_string(globalDrawingId) + ".xml.rels", drawingRelationshipsXml(sheet, globalMediaId, strict));
                for (const auto& image : sheet.images()) {
                    const std::string bytes(reinterpret_cast<const char*>(image.bytes().data()), image.bytes().size());
                    z.add("xl/media/image" + std::to_string(globalMediaId++) + "." + image.extension(), bytes, false);
                }
                ++globalDrawingId;
            }
            if (hasSheetCharts) {
                for (std::size_t ci = 0; ci < sheet.chartCount(); ++ci, ++globalChartId) {
                    sheetRels << "<Relationship Id=\"rIdChart" << (ci + 1) << "\" Type=\"" << nsRelsDoc(strict) << "/chart\" Target=\"../charts/chart" << globalChartId << ".xml\"/>";
                    z.add("xl/charts/chart" + std::to_string(globalChartId) + ".xml", chartXml(sheet.chart(ci), sheet.name(), strict));
                }
                sheetRels << "<Relationship Id=\"rIdDrawing\" Type=\"" << nsRelsDoc(strict) << "/drawing\" Target=\"../drawings/drawing" << globalDrawingId << ".xml\"/>";
                z.add("xl/drawings/drawing" + std::to_string(globalDrawingId) + ".xml", chartDrawingXml(sheet, strict));
                ++globalDrawingId;
            }
            if (hasSheetPivots) {
                for (std::size_t pi = 0; pi < sheet.pivotTables().size(); ++pi, ++globalPivotId) {
                    sheetRels << "<Relationship Id=\"rIdPivot" << (pi + 1) << "\" Type=\"" << nsRelsDoc(strict) << "/pivotTable\" Target=\"../pivotTables/pivotTable" << globalPivotId << ".xml\"/>";
                    sheetRels << "<Relationship Id=\"rIdPivotCache" << (pi + 1) << "\" Type=\"" << nsRelsDoc(strict) << "/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition" << globalPivotId << ".xml\"/>";
                    z.add("xl/pivotTables/pivotTable" + std::to_string(globalPivotId) + ".xml", pivotTableXml(sheet.pivotTables()[pi], globalPivotId, strict));
                    z.add("xl/pivotCache/pivotCacheDefinition" + std::to_string(globalPivotId) + ".xml", pivotCacheXml(sheet.pivotTables()[pi], strict));
                }
            }
            sheetRels << "</Relationships>";
            z.add("xl/worksheets/_rels/sheet"+std::to_string(i+1)+".xml.rels", sheetRels.str());
        }
    }
    if (!sstStrings.empty()) rels << "<Relationship Id=\"rId" << sheets_.size()+1 << "\" Type=\"" << nsRelsDoc(strict) << "/sharedStrings\" Target=\"sharedStrings.xml\"/>";
    rels << "<Relationship Id=\"rId" << sheets_.size()+(sstStrings.empty()?1:2) << "\" Type=\"" << nsRelsDoc(strict) << "/styles\" Target=\"styles.xml\"/></Relationships>";
    wb << "</sheets>";
    if (!definedNames_.empty()) {
        wb << "<definedNames>";
        for (const auto& item : definedNames_) {
            wb << "<definedName name=\"" << xmlEscape(item.name()) << "\"";
            if (item.localSheetId()) wb << " localSheetId=\"" << *item.localSheetId() << "\"";
            if (item.hidden()) wb << " hidden=\"1\"";
            if (!item.comment().empty()) wb << " comment=\"" << xmlEscape(item.comment()) << "\"";
            wb << ">" << xmlEscape(item.value()) << "</definedName>";
        }
        wb << "</definedNames>";
    }
    wb << "</workbook>";
    z.add("xl/workbook.xml", wb.str()); z.add("xl/_rels/workbook.xml.rels", rels.str());
    for (const auto& part : preservedParts_) z.add(part.name, part.data, part.compress);
    z.save(p);
}
void Workbook::load(const std::filesystem::path& p) { load(p, LoadOptions{}); }
void Workbook::load(const std::filesystem::path& p, const LoadOptions& options) { clear(); diagnostics_ = LoadDiagnostics{}; internal::ZipOpenLimits limits; limits.maxEntries = options.maxEntries; limits.maxEntryBytes = options.maxEntryBytes; limits.maxTotalBytes = options.maxTotalBytes; limits.maxFileBytes = options.maxFileBytes; limits.cancel = options.cancel; limits.progress = options.progress; auto z = internal::ZipArchive::open(p, limits); if(z.contains("docProps/core.xml")){auto cp=z.get("docProps/core.xml");properties_.setTitle(internal::tagText(cp,"dc:title"));properties_.setSubject(internal::tagText(cp,"dc:subject"));properties_.setCreator(internal::tagText(cp,"dc:creator"));properties_.setDescription(internal::tagText(cp,"dc:description"));properties_.setKeywords(internal::tagText(cp,"cp:keywords"));properties_.setCategory(internal::tagText(cp,"cp:category"));properties_.setLastModifiedBy(internal::tagText(cp,"cp:lastModifiedBy"));}if(z.contains("docProps/custom.xml")){auto cust=z.get("docProps/custom.xml");for(const auto& p:internal::tags(cust,"property")){const auto n=internal::attribute(p,"name");const auto vt=internal::attribute(p,internal::attribute(p,"vt:lpwstr")=="vt:lpwstr"?"vt:lpwstr":"");if(!n.empty())customProps_.add(CustomProperty(std::string(n),internal::tagText(p,"vt:lpwstr")));}}StyleCatalog styleCatalog;std::vector<Style> dxfStyles;if(z.contains("xl/styles.xml")){const auto stylesText=z.get("xl/styles.xml");styleCatalog=parseStyleCatalog(stylesText);dxfStyles=parseDifferentialStyles(stylesText);for(const auto& node:internal::tags(stylesText,"cellStyle")){const auto name=internal::attribute(node,"name");if(name.empty()||name=="Normal")continue;const auto xf=internal::attribute(node,"xfId");if(xf.empty())continue;const auto id=static_cast<std::size_t>(std::stoul(xf));if(id<styleCatalog.items.size())namedStyles_.emplace_back(name,styleCatalog.items[id]);}}std::vector<std::string> shared;if(z.contains("xl/sharedStrings.xml")){const auto sstXml = z.get("xl/sharedStrings.xml"); for(auto&si:internal::tags(sstXml,"si")){const auto rElements = internal::tags(si, "r"); if(!rElements.empty()){std::string richText; for(const auto& r : rElements) richText += internal::tagText(r, "t"); shared.push_back(std::move(richText));} else shared.push_back(internal::tagText(si,"t"));}}auto wb=z.get("xl/workbook.xml");strictNamespaces_ = wb.find("http://purl.oclc.org/ooxml/spreadsheetml/main") != std::string::npos;
for(const auto& protectionNode:internal::tags(wb,"workbookProtection")){
    protection_.setLockStructure(internal::attribute(protectionNode,"lockStructure")=="1");
    protection_.setLockWindows(internal::attribute(protectionNode,"lockWindows")=="1");
    protection_.setLockRevision(internal::attribute(protectionNode,"lockRevision")=="1");
    protection_.setWorkbookPasswordHash(internal::attribute(protectionNode,"workbookPassword"));
}
for(const auto& prNode:internal::tags(wb,"workbookPr"))
    date1904_ = internal::attribute(prNode,"date1904")=="1";
for(const auto& calcNode:internal::tags(wb,"calcPr")){
    const auto cm = internal::attribute(calcNode,"calcMode"); if(!cm.empty()) calcProps_.setCalcMode(std::string(cm));
    const auto ci = internal::attribute(calcNode,"calcId"); if(!ci.empty()) calcProps_.setCalcId(static_cast<int>(std::stoul(ci)));
    calcProps_.setCalcOnSave(internal::attribute(calcNode,"calcOnSave") == "1");
    calcProps_.setFullCalcOnLoad(internal::attribute(calcNode,"fullCalcOnLoad") == "1");
    calcProps_.setFullPrecision(internal::attribute(calcNode,"fullPrecision") != "0");
    calcProps_.setIterate(internal::attribute(calcNode,"iterate") == "1");
    const auto ic = internal::attribute(calcNode,"iterateCount"); if(!ic.empty()) calcProps_.setIterateCount(static_cast<int>(std::stoul(ic)));
    const auto id = internal::attribute(calcNode,"iterateDelta"); if(!id.empty()) calcProps_.setIterateDelta(std::stod(id));
}
for(const auto& node:internal::tags(wb,"definedName")){DefinedName item(internal::attribute(node,"name"),internal::tagText(node,"definedName"));const auto local=internal::attribute(node,"localSheetId");if(!local.empty())item.setLocalSheetId(static_cast<std::size_t>(std::stoul(local)));item.setHidden(internal::attribute(node,"hidden")=="1");item.setComment(internal::attribute(node,"comment"));definedNames_.push_back(std::move(item));}auto relxml=z.get("xl/_rels/workbook.xml.rels");std::unordered_map<std::string,std::string> targets;for(auto&r:internal::tags(relxml,"Relationship"))targets[internal::attribute(r,"Id")]=internal::attribute(r,"Target");for(auto&s:internal::tags(wb,"sheet")){try {auto name=internal::attribute(s,"name"),rid=internal::attribute(s,"r:id");auto& ws=addWorksheet(name);std::string target=targets[rid];if(target.rfind("/",0)==0)target.erase(0,1);else target="xl/"+target;parseSheet(ws, z.get(target), z, target, styleCatalog, dxfStyles, shared, date1904_);
    } catch (const std::exception& e) {
        diagnostics_.errors.push_back("Sheet '" + internal::attribute(s, "name") + "': " + e.what());
        if (!options.lenient) throw;
        continue;
}
}
    preservedParts_.clear();
    std::unordered_map<std::string, std::string> overrides, defaults;
    if (z.contains("[Content_Types].xml")) {
        const auto& ctText = z.get("[Content_Types].xml");
        for (const auto& node : internal::tags(ctText, "Override")) {
            const auto part = internal::attribute(node, "PartName");
            const auto ct = internal::attribute(node, "ContentType");
            if (!part.empty() && !ct.empty()) overrides[part] = ct;
        }
        for (const auto& node : internal::tags(ctText, "Default")) {
            const auto ext = internal::attribute(node, "Extension");
            const auto ct = internal::attribute(node, "ContentType");
            if (!ext.empty() && !ct.empty()) defaults[ext] = ct;
        }
    }
    for (const auto& partName : z.entryNames()) {
        if (isRegeneratedPart(partName)) continue;
        PreservedPart part;
        part.name = partName;
        part.data = z.get(partName);
        const auto overrideIt = overrides.find("/" + partName);
        if (overrideIt != overrides.end()) {
            part.overrideType = overrideIt->second;
        } else {
            const auto ext = extensionOf(partName);
            const auto defaultIt = defaults.find(ext);
            if (defaultIt != defaults.end()) { part.extension = ext; part.defaultType = defaultIt->second; }
        }
        part.compress = !(part.extension == "png" || part.extension == "jpg" || part.extension == "jpeg"
            || part.extension == "gif" || part.extension == "bmp" || part.extension == "bin");
        preservedParts_.push_back(std::move(part));
    }
}

void Workbook::load(std::istream& stream) { load(stream, LoadOptions{}); }
void Workbook::load(std::istream& stream, const LoadOptions& options) {
    const auto tmpPath = std::filesystem::temp_directory_path() / "xlpp_stream_load.tmp";
    {
        std::ofstream tmp(tmpPath, std::ios::binary);
        tmp << stream.rdbuf();
    }
    load(tmpPath, options);
    std::filesystem::remove(tmpPath);
}
void Workbook::save(std::ostream& stream) const { save(stream, SaveOptions{}); }
void Workbook::save(std::ostream& stream, const SaveOptions& options) const {
    const auto tmpPath = std::filesystem::temp_directory_path() / "xlpp_stream_save.tmp";
    save(tmpPath, options);
    {
        std::ifstream tmp(tmpPath, std::ios::binary);
        stream << tmp.rdbuf();
    }
    std::filesystem::remove(tmpPath);
}
}
