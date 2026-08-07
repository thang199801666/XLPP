#include <unordered_set>
#include <XLPP/Workbook/Workbook.h>
#include "../XML/XmlUtilities.h"
#include "../Packaging/ZipArchive.h"
#include "../Packaging/RelationshipGraph.h"
#include "../Threading/ThreadPool.h"
#include "../Vba/VbaProjectBinary.h"
#include <algorithm>
#include <array>
#include <cctype>
#include <cmath>
#include <cstdlib>
#include <fstream>
#include <functional>
#include <sstream>
#include <set>
#include <stdexcept>
#include <thread>
#include <type_traits>
#include <tuple>
#include <unordered_map>
#include <vector>
#include <iomanip>
#include <limits>

using xlpp::internal::xmlEscape;
using xlpp::internal::writeXmlEscaped;

namespace xlpp::internal {
struct WorkbookDrawingAccess {
    static const auto& imageEdits(const xlpp::Worksheet& sheet) noexcept { return sheet.importedImageEdits_; }
    static const auto& chartEdits(const xlpp::Worksheet& sheet) noexcept { return sheet.importedChartEdits_; }
};
}

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

std::string quoteSheetName(const std::string& name) {
    std::string escaped;
    escaped.reserve(name.size() + 2);
    escaped.push_back('\'');
    for (const char ch : name) {
        escaped.push_back(ch);
        if (ch == '\'') escaped.push_back('\'');
    }
    escaped.push_back('\'');
    return escaped;
}

std::string absoluteReferenceToken(std::string token) {
    token.erase(std::remove(token.begin(), token.end(), '$'), token.end());
    if (token.empty()) return token;
    std::size_t letters = 0;
    while (letters < token.size() && std::isalpha(static_cast<unsigned char>(token[letters]))) ++letters;
    if (letters == token.size()) return "$" + token;
    if (letters == 0) return "$" + token;
    return "$" + token.substr(0, letters) + "$" + token.substr(letters);
}

std::string absoluteReference(std::string reference) {
    if (!reference.empty() && reference.front() == '=') reference.erase(reference.begin());
    const auto bang = reference.rfind('!');
    if (bang != std::string::npos) reference = reference.substr(bang + 1);
    std::ostringstream result;
    std::size_t start = 0;
    bool firstArea = true;
    while (start <= reference.size()) {
        const auto comma = reference.find(',', start);
        const auto area = reference.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        const auto colon = area.find(':');
        if (!firstArea) result << ',';
        firstArea = false;
        if (colon == std::string::npos) result << absoluteReferenceToken(area);
        else result << absoluteReferenceToken(area.substr(0, colon)) << ':' << absoluteReferenceToken(area.substr(colon + 1));
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}

std::string qualifiedPrintReference(const std::string& sheetName, const std::string& reference) {
    std::ostringstream result;
    const auto absolute = absoluteReference(reference);
    std::size_t start = 0;
    bool first = true;
    while (start <= absolute.size()) {
        const auto comma = absolute.find(',', start);
        if (!first) result << ',';
        first = false;
        result << quoteSheetName(sheetName) << '!' << absolute.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}

std::string localPrintReference(std::string value) {
    std::ostringstream result;
    std::size_t start = 0;
    bool first = true;
    while (start <= value.size()) {
        const auto comma = value.find(',', start);
        auto area = value.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        const auto bang = area.rfind('!');
        if (bang != std::string::npos) area = area.substr(bang + 1);
        area.erase(std::remove(area.begin(), area.end(), '$'), area.end());
        if (!first) result << ',';
        first = false;
        result << area;
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}

std::string contentTypes(std::size_t sheetCount,
                         std::size_t tableCount,
                         std::size_t commentCount,
                         const std::vector<std::size_t>& drawingIds,
                         const std::vector<xlpp::PreservedPart>& preserved,
                         bool strict,
                         bool hasSst,
                         const std::vector<std::size_t>& chartIds,
                         const std::vector<std::size_t>& pivotIds,
                         bool hasCustomProperties,
                         bool macroEnabled) {
    std::ostringstream xml;
    std::set<std::string> emittedDefaults;
    std::set<std::string> emittedOverrides;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\""
        << nsCtPkg(strict) << "\">";

    const auto addDefault = [&](const std::string& extension, const std::string& contentType) {
        if (extension.empty() || contentType.empty() || !emittedDefaults.insert(extension).second) return;
        xml << "<Default Extension=\"" << xmlEscape(extension) << "\" ContentType=\""
            << xmlEscape(contentType) << "\"/>";
    };
    const auto addOverride = [&](const std::string& partName, const std::string& contentType) {
        if (partName.empty() || contentType.empty() || !emittedOverrides.insert(partName).second) return;
        xml << "<Override PartName=\"/" << xmlEscape(partName) << "\" ContentType=\""
            << xmlEscape(contentType) << "\"/>";
    };

    addDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
    addDefault("xml", "application/xml");
    addDefault("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing");
    addDefault("png", "image/png");
    addDefault("jpg", "image/jpeg");
    addDefault("jpeg", "image/jpeg");

    addOverride("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
    addOverride("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
    if (hasCustomProperties)
        addOverride("docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml");
    addOverride("xl/workbook.xml", macroEnabled
        ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    addOverride("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");

    for (std::size_t index = 1; index <= sheetCount; ++index)
        addOverride("xl/worksheets/sheet" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    for (std::size_t index = 1; index <= tableCount; ++index)
        addOverride("xl/tables/table" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml");
    for (std::size_t index = 1; index <= commentCount; ++index)
        addOverride("xl/comments" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
    for (const auto index : drawingIds)
        addOverride("xl/drawings/drawing" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.drawing+xml");
    if (hasSst)
        addOverride("xl/sharedStrings.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    for (const auto index : chartIds)
        addOverride("xl/charts/chart" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    for (const auto index : pivotIds) {
        addOverride("xl/pivotTables/pivotTable" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
        addOverride("xl/pivotCache/pivotCacheDefinition" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
        addOverride("xl/pivotCache/pivotCacheRecords" + std::to_string(index) + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
    }
    for (const auto& part : preserved) {
        if (!part.overrideType.empty()) addOverride(part.name, part.overrideType);
        else if (!part.defaultType.empty()) addDefault(part.extension, part.defaultType);
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
        || name.rfind("xl/worksheets/", 0) == 0;
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
            result.value = std::stod(value);
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

std::string rootrels(bool strict, bool hasCustomProperties){std::ostringstream x;x<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\""<<nsRelsPkg(strict)<<"\"><Relationship Id=\"rId1\" Type=\""<<nsRelsDoc(strict)<<"/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\""<<nsRelsPkg(strict)<<"/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\""<<nsRelsDoc(strict)<<"/extended-properties\" Target=\"docProps/app.xml\"/>";if(hasCustomProperties)x<<"<Relationship Id=\"rId4\" Type=\""<<nsRelsDoc(strict)<<"/custom-properties\" Target=\"docProps/custom.xml\"/>";x<<"</Relationships>";return x.str();}
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
        const std::string fillPattern = fill.patternType().empty() ? "none" : fill.patternType();
        xml << "<fill><patternFill patternType=\"" << xmlEscape(fillPattern) << "\">";
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
            const std::string fillPattern = fill.patternType().empty() ? "none" : fill.patternType();
            xml << "<fill><patternFill patternType=\"" << xmlEscape(fillPattern) << "\">";
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

// Formats for the built-in (ECMA-376 §18.8.30) numFmtIds 0-49 that files can
// reference without declaring a <numFmt> element. Used as a fallback when a
// cell's numFmtId is not present in the custom formats table.
std::string builtinNumFmt(int id) {
    switch (id) {
    case 0: return "General";
    case 1: return "0";
    case 2: return "0.00";
    case 3: return "#,##0";
    case 4: return "#,##0.00";
    case 9: return "0%";
    case 10: return "0.00%";
    case 11: return "0.00E+00";
    case 12: return "# ?/?";
    case 13: return "# ?" "?/??";
    case 14: return "mm-dd-yy";
    case 15: return "d-mmm-yy";
    case 16: return "d-mmm";
    case 17: return "mmm-yy";
    case 18: return "h:mm AM/PM";
    case 19: return "h:mm:ss AM/PM";
    case 20: return "h:mm";
    case 21: return "h:mm:ss";
    case 22: return "m/d/yy h:mm";
    case 37: return "#,##0 ;(#,##0)";
    case 38: return "#,##0 ;[Red](#,##0)";
    case 39: return "#,##0.00;(#,##0.00)";
    case 40: return "#,##0.00;[Red](#,##0.00)";
    case 45: return "mm:ss";
    case 46: return "[h]:mm:ss";
    case 47: return "mmss.0";
    case 48: return "##0.0E+0";
    case 49: return "@";
    default: return "General";
    }
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
        if (!numFmtIdText.empty()) { const auto id = static_cast<int>(std::stoul(numFmtIdText)); style.setNumFmtId(id); const auto it = formats.find(id); style.setNumberFormat(it != formats.end() ? it->second : builtinNumFmt(id)); }
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
                if (!value.empty()) run.setSize(std::stod(value));
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

struct LoadedSharedString {
    std::string plainText;
    std::optional<xlpp::RichText> richText;
};

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
                     const std::unordered_map<std::string, std::size_t>* sstIndex = nullptr,
                     std::size_t rowWorkers = 0,
                     std::string_view vbaCodeName = {}) {
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
    if (!sheet.pivotTables().empty()) {
        xml << "<pivotTableParts count=\"" << sheet.pivotTables().size() << "\">";
        for (std::size_t i = 0; i < sheet.pivotTables().size(); ++i)
            xml << "<pivotTablePart r:id=\"rIdPivot" << i + 1 << "\"/>";
        xml << "</pivotTableParts>";
    }
    xml << "</worksheet>";
    return xml.str();
}

std::string chartSeriesCacheXml(const xlpp::ChartSeriesCache& cache, bool prefixed = true) {
    if (!cache.present) return {};
    const auto c = prefixed ? "c:" : "";
    const auto local = cache.numeric ? "numCache" : "strCache";
    std::ostringstream xml;
    xml << "<" << c << local << ">";
    if (cache.numeric) xml << "<" << c << "formatCode>" << xmlEscape(cache.formatCode.empty() ? "General" : cache.formatCode) << "</" << c << "formatCode>";
    xml << "<" << c << "ptCount val=\"" << cache.effectivePointCount() << "\"/>";
    auto points = cache.points;
    std::sort(points.begin(), points.end(), [](const auto& a, const auto& b){ return a.index < b.index; });
    for (const auto& point : points)
        xml << "<" << c << "pt idx=\"" << point.index << "\"><" << c << "v>" << xmlEscape(point.value) << "</" << c << "v></" << c << "pt>";
    xml << "</" << c << local << ">";
    return xml.str();
}

std::string generatedPlotAuxiliaryXml(const xlpp::Chart::Plot& plot, bool strict);
std::string generatedDataTableXml(const xlpp::ChartDataTable& table, bool strict);
std::string chartView3DXml(const xlpp::ChartView3D& view, bool prefixed);
std::string generatedChartWallXml(const char* localName, const xlpp::ChartWallFormat& format, bool prefixed);
bool patchMarkerFormatInOwner(std::string& owner, const xlpp::ChartMarkerFormat& format);
bool patchNestedLineFormat(std::string& owner, const xlpp::ChartLineFormat& format);
std::vector<std::string> drawingTags(const std::string& xml, const char* prefixed, const char* local);

std::string chartXml(const xlpp::Chart& chart, bool strict) {
    std::ostringstream xml;
    const auto type = chart.type();
    const bool threeAxis = type == xlpp::Chart::Type::Bar3D || type == xlpp::Chart::Type::Line3D ||
                           type == xlpp::Chart::Type::Area3D || type == xlpp::Chart::Type::Surface ||
                           type == xlpp::Chart::Type::Surface3D;
    const bool projectedPie = type == xlpp::Chart::Type::PieOfPie || type == xlpp::Chart::Type::BarOfPie;
    const bool hasAxes = type != xlpp::Chart::Type::Pie && type != xlpp::Chart::Type::Pie3D && type != xlpp::Chart::Type::Doughnut && !projectedPie;
    const bool barLike = type == xlpp::Chart::Type::Bar || type == xlpp::Chart::Type::Bar3D;
    const bool lineOrArea3D = type == xlpp::Chart::Type::Line3D || type == xlpp::Chart::Type::Area3D;
    const bool surface = type == xlpp::Chart::Type::Surface || type == xlpp::Chart::Type::Surface3D;

    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    xml << "<c:chartSpace xmlns:c=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart") << "\"";
    xml << " xmlns:a=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main") << "\">";
    if (!chart.style().empty()) xml << "<c:style val=\"" << xmlEscape(chart.style()) << "\"/>";
    xml << "<c:chart>";
    if (!chart.title().empty())
        xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>" << xmlEscape(chart.title()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
    if (chart.view3D().present) xml << chartView3DXml(chart.view3D(), true);
    if (chart.floorFormat().present) xml << generatedChartWallXml("floor", chart.floorFormat(), true);
    if (chart.sideWallFormat().present) xml << generatedChartWallXml("sideWall", chart.sideWallFormat(), true);
    if (chart.backWallFormat().present) xml << generatedChartWallXml("backWall", chart.backWallFormat(), true);
    if (type == xlpp::Chart::Type::Stock && chart.series().size() != 3 && chart.series().size() != 4)
        throw std::invalid_argument("Stock charts require exactly 3 (high-low-close) or 4 (open-high-low-close) series");

    xml << "<c:plotArea><c:layout/>";
    xml << "<c:" << xlpp::Chart::typeName(type, chart.grouping()) << ">";
    if (barLike) {
        xml << "<c:barDir val=\"col\"/>";
        const char* grouping = chart.grouping() == xlpp::Chart::Grouping::Stacked ? "stacked"
            : (chart.grouping() == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "clustered");
        xml << "<c:grouping val=\"" << grouping << "\"/><c:varyColors val=\"0\"/>";
    } else if (lineOrArea3D) {
        const char* grouping = chart.grouping() == xlpp::Chart::Grouping::Stacked ? "stacked"
            : (chart.grouping() == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "standard");
        xml << "<c:grouping val=\"" << grouping << "\"/>";
    } else if (type == xlpp::Chart::Type::Pie3D || type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut || projectedPie) {
        xml << "<c:varyColors val=\"1\"/>";
    } else if (type == xlpp::Chart::Type::Radar) {
        const auto* plot = chart.primaryPlotOrNull();
        xml << "<c:radarStyle val=\"" << xmlEscape(plot && !plot->radarStyle.empty() ? plot->radarStyle : "standard") << "\"/>";
    } else if (surface) {
        const auto* plot = chart.primaryPlotOrNull();
        if (plot && plot->hasWireframe) xml << "<c:wireframe val=\"" << (plot->wireframe ? "1" : "0") << "\"/>";
    }

    for (std::size_t s = 0; s < chart.series().size(); ++s) {
        const auto& series = chart.series()[s];
        std::ostringstream seriesXml;
        seriesXml << "<c:ser><c:idx val=\"" << s << "\"/><c:order val=\"" << s << "\"/>";
        if (!series.titleReference().empty()) {
            seriesXml << "<c:tx><c:strRef><c:f>" << xmlEscape(series.titleReference()) << "</c:f>"
                      << chartSeriesCacheXml(series.titleCache()) << "</c:strRef></c:tx>";
        } else if (!series.title().empty())
            seriesXml << "<c:tx><c:v>" << xmlEscape(series.title()) << "</c:v></c:tx>";
        if (type == xlpp::Chart::Type::Radar && series.markerFormat().present) {
            std::string markerOwner = "<c:owner></c:owner>";
            if (!patchMarkerFormatInOwner(markerOwner, series.markerFormat())) throw std::runtime_error("Failed to serialize radar marker format");
            const auto markerNodes = drawingTags(markerOwner, "c:marker", "marker");
            if (!markerNodes.empty()) seriesXml << markerNodes.front();
        }
        if (!series.categoriesReference().empty()) {
            const bool numericCategories = type == xlpp::Chart::Type::Stock || (series.categoriesCache().present && series.categoriesCache().numeric);
            const auto refTag = numericCategories ? "numRef" : "strRef";
            seriesXml << "<c:cat><c:" << refTag << "><c:f>" << xmlEscape(series.categoriesReference()) << "</c:f>"
                      << chartSeriesCacheXml(series.categoriesCache()) << "</c:" << refTag << "></c:cat>";
        }
        if (!series.valuesReference().empty())
            seriesXml << "<c:val><c:numRef><c:f>" << xmlEscape(series.valuesReference()) << "</c:f>"
                      << chartSeriesCacheXml(series.valuesCache()) << "</c:numRef></c:val>";
        seriesXml << "</c:ser>";
        xml << seriesXml.str();
    }
    if (const auto* generatedPlot = chart.primaryPlotOrNull()) {
        xml << generatedPlotAuxiliaryXml(*generatedPlot, strict);
        if (generatedPlot->hasGapDepth && (type == xlpp::Chart::Type::Bar3D || type == xlpp::Chart::Type::Line3D || type == xlpp::Chart::Type::Area3D))
            xml << "<c:gapDepth val=\"" << generatedPlot->gapDepth << "\"/>";
        if (!generatedPlot->shape.empty() && type == xlpp::Chart::Type::Bar3D)
            xml << "<c:shape val=\"" << xmlEscape(generatedPlot->shape) << "\"/>";
        if ((type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut) && generatedPlot->hasFirstSliceAngle) {
            if (generatedPlot->firstSliceAngle < 0 || generatedPlot->firstSliceAngle > 360) throw std::invalid_argument("Chart first-slice angle must be between 0 and 360 degrees");
            xml << "<c:firstSliceAng val=\"" << generatedPlot->firstSliceAngle << "\"/>";
        }
        if (type == xlpp::Chart::Type::Doughnut && generatedPlot->hasHoleSize) {
            if (generatedPlot->holeSize < 10 || generatedPlot->holeSize > 90) throw std::invalid_argument("Doughnut hole size must be between 10 and 90 percent");
            xml << "<c:holeSize val=\"" << generatedPlot->holeSize << "\"/>";
        }
        if (projectedPie) {
            auto options = generatedPlot->projectedPie;
            options.present = true; options.ofPieType = type == xlpp::Chart::Type::BarOfPie ? "bar" : "pie";
            if (options.gapWidth < 0 || options.gapWidth > 500 || options.secondPlotSize < 5 || options.secondPlotSize > 200)
                throw std::invalid_argument("Projected-pie gap width or second-plot size is outside the supported OOXML range");
            if (options.splitType != "auto" && options.splitType != "cust" && options.splitType != "percent" && options.splitType != "pos" && options.splitType != "val")
                throw std::invalid_argument("Unsupported projected-pie split type");
            xml << "<c:ofPieType val=\"" << options.ofPieType << "\"/>";
            xml << "<c:gapWidth val=\"" << options.gapWidth << "\"/><c:splitType val=\"" << options.splitType << "\"/>";
            if (options.hasSplitPosition) xml << "<c:splitPos val=\"" << options.splitPosition << "\"/>";
            if (!options.customSplitPoints.empty()) { xml << "<c:custSplit>"; for (const auto point : options.customSplitPoints) { if (point < 0) throw std::invalid_argument("Projected-pie custom split indices cannot be negative"); xml << "<c:secondPiePt val=\"" << point << "\"/>"; } xml << "</c:custSplit>"; }
            xml << "<c:secondPieSize val=\"" << options.secondPlotSize << "\"/>";
            if (options.hasSeriesLines) {
                std::string lines = "<c:serLines></c:serLines>";
                if (options.seriesLinesFormat.present && !patchNestedLineFormat(lines, options.seriesLinesFormat)) throw std::runtime_error("Failed to serialize projected-pie series lines");
                xml << lines;
            }
        }
    } else if (projectedPie) {
        xlpp::ChartProjectedPieOptions options; options.present=true; options.ofPieType=type==xlpp::Chart::Type::BarOfPie?"bar":"pie";
        xml << "<c:ofPieType val=\"" << options.ofPieType << "\"/><c:gapWidth val=\"150\"/><c:splitType val=\"auto\"/><c:secondPieSize val=\"75\"/>";
    }
    if (hasAxes) {
        xml << "<c:axId val=\"1\"/><c:axId val=\"2\"/>";
        if (threeAxis) xml << "<c:axId val=\"3\"/>";
    }
    if ((type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut) && (!chart.primaryPlotOrNull() || !chart.primaryPlotOrNull()->hasFirstSliceAngle))
        xml << "<c:firstSliceAng val=\"0\"/>";
    if (type == xlpp::Chart::Type::Doughnut && (!chart.primaryPlotOrNull() || !chart.primaryPlotOrNull()->hasHoleSize))
        xml << "<c:holeSize val=\"10\"/>";
    xml << "</c:" << xlpp::Chart::typeName(type, chart.grouping()) << ">";

    if (hasAxes) {
        xml << "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/>";
        if (!chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
        xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>";
        xml << "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/>";
        if (!chart.yAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.yAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
        xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorGridlines/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>";
        if (threeAxis)
            xml << "<c:serAx><c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/></c:serAx>";
    }
    if (chart.dataTable().present) xml << generatedDataTableXml(chart.dataTable(), strict);
    xml << "</c:plotArea>";
    if (chart.showLegend()) xml << "<c:legend><c:legendPos val=\"" << xmlEscape(chart.legendPosition()) << "\"/></c:legend>";
    xml << "</c:chart></c:chartSpace>";
    return xml.str();
}

int resolvedPivotFieldIndex(const xlpp::PivotCache& cache, int index, const std::string& name) {
    if (index >= 0 && static_cast<std::size_t>(index) < cache.fields().size()) return index;
    if (!name.empty()) {
        const auto resolved = cache.fieldIndex(name);
        if (resolved >= 0) return resolved;
    }
    throw std::invalid_argument("Pivot field '" + name + "' is not present in the cache fields");
}

std::string pivotTableXml(const xlpp::PivotTable& pt, std::size_t id, bool strict) {
    std::ostringstream xml;
    const auto fieldCount = pt.cache().fields().size();
    if (pt.name().empty()) throw std::invalid_argument("Pivot table name cannot be empty");
    if (fieldCount == 0) throw std::invalid_argument("Pivot table cache must contain at least one field");

    std::vector<int> rowIndexes, columnIndexes, pageIndexes, dataIndexes;
    rowIndexes.reserve(pt.rowFields().size());
    columnIndexes.reserve(pt.columnFields().size());
    pageIndexes.reserve(pt.pageFields().size());
    dataIndexes.reserve(pt.dataFields().size());
    for (const auto& field : pt.rowFields()) rowIndexes.push_back(resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()));
    for (const auto& field : pt.columnFields()) columnIndexes.push_back(resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()));
    for (const auto& field : pt.pageFields()) pageIndexes.push_back(resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()));
    for (const auto& field : pt.dataFields()) dataIndexes.push_back(resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()));

    std::vector<std::vector<std::string>> sharedItems(fieldCount);
    for (const auto& record : pt.cache().records()) {
        if (record.size() != fieldCount)
            throw std::invalid_argument("Pivot cache record width must match field count");
        for (std::size_t i = 0; i < fieldCount; ++i) {
            if (std::find(sharedItems[i].begin(), sharedItems[i].end(), record[i]) == sharedItems[i].end())
                sharedItems[i].push_back(record[i]);
        }
    }

    // A PivotTable location is the complete occupied view range, not merely the
    // requested top-left cell.  When the caller supplies an anchor, derive the
    // smallest view rectangle from the emitted row-item matrix.  This keeps the
    // location, rowItems and colItems mutually consistent for desktop Excel.
    std::string location = pt.location().empty() ? "D2" : pt.location();
    if (location.find(':') == std::string::npos) {
        const auto first = xlpp::CellReference::parse(location);
        const auto rowItemCount = rowIndexes.empty()
            ? std::size_t{1}
            : sharedItems[static_cast<std::size_t>(rowIndexes.front())].size() + 1; // grand total
        const auto outputRows = std::max<std::size_t>(2, rowItemCount + 1);          // header row
        const auto outputColumns = std::max<std::size_t>(2, 1 + dataIndexes.size());
        xlpp::CellReference last{first.row + outputRows - 1, first.column + outputColumns - 1};
        location += ":" + last.address();
    }

    // Match the conservative view metadata emitted by desktop Excel.  The
    // extension list and revision UID are optional, but the layout/version flags
    // below make the intended non-OLAP compact view unambiguous.
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotTableDefinition xmlns=\""
        << nsMain(strict) << "\" name=\"" << xmlEscape(pt.name()) << "\" cacheId=\"" << id
        << "\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\""
        << " applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\""
        << " dataCaption=\"Values\" updatedVersion=\"8\" minRefreshableVersion=\"3\""
        << " useAutoFormatting=\"1\" itemPrintTitles=\"1\" createdVersion=\"8\""
        << " indent=\"0\" outline=\"1\" outlineData=\"1\" multipleFieldFilters=\"0\">";
    xml << "<location ref=\"" << xmlEscape(location) << "\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>";

    xml << "<pivotFields count=\"" << fieldCount << "\">";
    for (std::size_t fieldIndex = 0; fieldIndex < fieldCount; ++fieldIndex) {
        const auto contains = [fieldIndex](const std::vector<int>& values) {
            return std::find(values.begin(), values.end(), static_cast<int>(fieldIndex)) != values.end();
        };
        const auto isRow = contains(rowIndexes);
        const auto isCol = contains(columnIndexes);
        const auto isPage = contains(pageIndexes);
        const auto isData = contains(dataIndexes);
        const xlpp::PivotField* modelField = nullptr;
        if (isRow) modelField = &pt.rowFields()[static_cast<std::size_t>(std::distance(rowIndexes.begin(), std::find(rowIndexes.begin(), rowIndexes.end(), static_cast<int>(fieldIndex))))];
        else if (isCol) modelField = &pt.columnFields()[static_cast<std::size_t>(std::distance(columnIndexes.begin(), std::find(columnIndexes.begin(), columnIndexes.end(), static_cast<int>(fieldIndex))))];
        else if (isPage) modelField = &pt.pageFields()[static_cast<std::size_t>(std::distance(pageIndexes.begin(), std::find(pageIndexes.begin(), pageIndexes.end(), static_cast<int>(fieldIndex))))];

        xml << "<pivotField";
        if (isRow) xml << " axis=\"axisRow\"";
        else if (isCol) xml << " axis=\"axisCol\"";
        else if (isPage) xml << " axis=\"axisPage\"";
        if (isData) xml << " dataField=\"1\"";
        xml << " showAll=\"" << ((modelField && modelField->showAll()) ? 1 : 0) << "\"";
        if (isRow || isCol || isPage) xml << " defaultSubtotal=\"0\"";
        if (isRow || isCol || isPage) {
            // The axis item list contains only concrete cache items.  A synthetic
            // default item is not needed for an unfiltered field and can make the
            // row/column view indexes ambiguous to Excel.
            xml << "><items count=\"" << sharedItems[fieldIndex].size() << "\">";
            for (std::size_t itemIndex = 0; itemIndex < sharedItems[fieldIndex].size(); ++itemIndex)
                xml << "<item x=\"" << itemIndex << "\"/>";
            xml << "</items></pivotField>";
        } else {
            xml << "/>";
        }
    }
    xml << "</pivotFields>";

    if (!rowIndexes.empty()) {
        xml << "<rowFields count=\"" << rowIndexes.size() << "\">";
        for (const auto index : rowIndexes) xml << "<field x=\"" << index << "\"/>";
        xml << "</rowFields>";
        // A complete multi-axis item matrix is significantly more complex.  The
        // first row field is sufficient for the common one-row-axis pivot and is
        // accepted by Excel/LibreOffice; additional row fields are refreshed from
        // the cache when the workbook is opened.
        const auto itemCount = sharedItems[static_cast<std::size_t>(rowIndexes.front())].size();
        xml << "<rowItems count=\"" << (itemCount + 1) << "\">";
        for (std::size_t itemIndex = 0; itemIndex < itemCount; ++itemIndex)
            xml << "<i><x v=\"" << itemIndex << "\"/></i>";
        xml << "<i t=\"grand\"><x/></i></rowItems>";
    }
    if (!columnIndexes.empty()) {
        xml << "<colFields count=\"" << columnIndexes.size() << "\">";
        for (const auto index : columnIndexes) xml << "<field x=\"" << index << "\"/>";
        xml << "</colFields>";
    }
    if (columnIndexes.empty()) {
        // With no explicit column field and a single values field, the canonical
        // column item is an empty data item.  Do not invent a field/item index.
        xml << "<colItems count=\"1\"><i/></colItems>";
    } else {
        xml << "<colItems count=\"1\"><i t=\"grand\"><x/></i></colItems>";
    }
    if (!pageIndexes.empty()) {
        xml << "<pageFields count=\"" << pageIndexes.size() << "\">";
        for (const auto index : pageIndexes) xml << "<pageField fld=\"" << index << "\" hier=\"-1\"/>";
        xml << "</pageFields>";
    }
    if (!dataIndexes.empty()) {
        xml << "<dataFields count=\"" << dataIndexes.size() << "\">";
        for (std::size_t i = 0; i < dataIndexes.size(); ++i) {
            const auto index = dataIndexes[i];
            const auto& field = pt.dataFields()[i];
            const auto& cacheName = pt.cache().fields()[static_cast<std::size_t>(index)];
            const auto displayName = field.name().empty() ? cacheName : field.name();
            const auto subtotal = field.subtotal().empty() ? std::string("sum") : field.subtotal();
            xml << "<dataField name=\"" << xmlEscape(subtotal == "sum" ? "Sum of " + displayName : displayName)
                << "\" fld=\"" << index << "\" baseField=\"0\" baseItem=\"0\"";
            if (subtotal != "sum") xml << " subtotal=\"" << xmlEscape(subtotal) << "\"";
            xml << "/>";
        }
        xml << "</dataFields>";
    }
    xml << "<pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\"/>";
    xml << "</pivotTableDefinition>";
    return xml.str();
}

enum class PivotValueKind { Blank, Number, Boolean, Error, String };

struct ParsedPivotValue {
    PivotValueKind kind{PivotValueKind::String};
    double number{0.0};
};

ParsedPivotValue parsePivotValue(const std::string& value) {
    if (value.empty()) return {PivotValueKind::Blank, 0.0};
    if (value == "true") return {PivotValueKind::Boolean, 1.0};
    if (value == "false") return {PivotValueKind::Boolean, 0.0};
    if (value.front() == '#') return {PivotValueKind::Error, 0.0};
    char* end = nullptr;
    const auto number = std::strtod(value.c_str(), &end);
    if (end && end != value.c_str() && *end == '\0' && std::isfinite(number))
        return {PivotValueKind::Number, number};
    return {PivotValueKind::String, 0.0};
}

std::vector<std::vector<std::string>> pivotSharedItems(const xlpp::PivotCache& cache) {
    std::vector<std::vector<std::string>> result(cache.fields().size());
    for (const auto& record : cache.records()) {
        if (record.size() != cache.fields().size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        for (std::size_t i = 0; i < record.size(); ++i) {
            auto& items = result[i];
            if (std::find(items.begin(), items.end(), record[i]) == items.end())
                items.push_back(record[i]);
        }
    }
    return result;
}

bool pivotFieldIsPureData(const xlpp::PivotTable& pt, std::size_t fieldIndex) {
    const auto matchesPivotField = [&](const auto& fields) {
        return std::any_of(fields.begin(), fields.end(), [&](const auto& field) {
            return resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()) == static_cast<int>(fieldIndex);
        });
    };
    const bool onAxis = matchesPivotField(pt.rowFields()) || matchesPivotField(pt.columnFields())
        || matchesPivotField(pt.pageFields());
    const bool isData = std::any_of(pt.dataFields().begin(), pt.dataFields().end(), [&](const auto& field) {
        return resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()) == static_cast<int>(fieldIndex);
    });
    return isData && !onAxis;
}

void writePivotValue(std::ostringstream& xml, const std::string& value) {
    const auto parsed = parsePivotValue(value);
    switch (parsed.kind) {
        case PivotValueKind::Blank: xml << "<m/>"; break;
        case PivotValueKind::Number: xml << "<n v=\"" << std::setprecision(15) << parsed.number << "\"/>"; break;
        case PivotValueKind::Boolean: xml << "<b v=\"" << (parsed.number != 0.0 ? 1 : 0) << "\"/>"; break;
        case PivotValueKind::Error: xml << "<e v=\"" << xmlEscape(value) << "\"/>"; break;
        case PivotValueKind::String: xml << "<s v=\"" << xmlEscape(value) << "\"/>"; break;
    }
}

std::string pivotCacheXml(const xlpp::PivotTable& pt, bool strict) {
    std::ostringstream xml;
    const auto& cache = pt.cache();
    const auto fieldCount = cache.fields().size();
    const auto sharedItems = pivotSharedItems(cache);
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotCacheDefinition xmlns=\""
        << nsMain(strict) << "\" xmlns:r=\"" << nsRelsDoc(strict)
        << "\" r:id=\"rId1\" saveData=\"1\" refreshOnLoad=\"1\" enableRefresh=\"1\""
        << " createdVersion=\"3\" refreshedVersion=\"3\" minRefreshableVersion=\"3\" recordCount=\""
        << cache.records().size() << "\">";
    std::string sourceSheet;
    std::string sourceRef = cache.sourceData();
    const auto bang = sourceRef.find('!');
    if (bang != std::string::npos) {
        sourceSheet = sourceRef.substr(0, bang);
        sourceRef = sourceRef.substr(bang + 1);
        if (sourceSheet.size() >= 2 && sourceSheet.front() == '\'' && sourceSheet.back() == '\'') {
            sourceSheet = sourceSheet.substr(1, sourceSheet.size() - 2);
            std::size_t pos = 0;
            while ((pos = sourceSheet.find("''", pos)) != std::string::npos) sourceSheet.replace(pos, 2, "'");
        }
    }
    sourceRef.erase(std::remove(sourceRef.begin(), sourceRef.end(), '$'), sourceRef.end());
    xml << "<cacheSource type=\"worksheet\"><worksheetSource ref=\"" << xmlEscape(sourceRef) << "\"";
    if (!sourceSheet.empty()) xml << " sheet=\"" << xmlEscape(sourceSheet) << "\"";
    xml << "/></cacheSource><cacheFields count=\"" << fieldCount << "\">";

    for (std::size_t i = 0; i < fieldCount; ++i) {
        const auto& items = sharedItems[i];
        bool containsString = false, containsNumber = false, containsBoolean = false;
        bool containsError = false, containsBlank = false, allNumbersInteger = true;
        double minValue = 0.0, maxValue = 0.0;
        bool hasNumericRange = false;
        std::set<PivotValueKind> nonBlankKinds;
        for (const auto& value : items) {
            const auto parsed = parsePivotValue(value);
            if (parsed.kind != PivotValueKind::Blank) nonBlankKinds.insert(parsed.kind);
            switch (parsed.kind) {
                case PivotValueKind::Blank: containsBlank = true; break;
                case PivotValueKind::Number:
                    containsNumber = true;
                    allNumbersInteger = allNumbersInteger && std::floor(parsed.number) == parsed.number;
                    if (!hasNumericRange) { minValue = maxValue = parsed.number; hasNumericRange = true; }
                    else { minValue = std::min(minValue, parsed.number); maxValue = std::max(maxValue, parsed.number); }
                    break;
                case PivotValueKind::Boolean: containsBoolean = true; break;
                case PivotValueKind::Error: containsError = true; break;
                case PivotValueKind::String: containsString = true; break;
            }
        }

        xml << "<cacheField name=\"" << xmlEscape(cache.fields()[i]) << "\" numFmtId=\"0\"><sharedItems";
        if (!containsString) xml << " containsString=\"0\"";
        if (containsNumber) xml << " containsNumber=\"1\"";
        if (containsNumber && allNumbersInteger) xml << " containsInteger=\"1\"";
        if (containsBoolean) xml << " containsBoolean=\"1\"";
        if (containsError) xml << " containsError=\"1\"";
        if (containsBlank) xml << " containsBlank=\"1\"";
        if (nonBlankKinds.size() <= 1 && (containsNumber || containsBoolean || containsError))
            xml << " containsSemiMixedTypes=\"0\"";
        if (hasNumericRange) xml << " minValue=\"" << std::setprecision(15) << minValue
                                 << "\" maxValue=\"" << std::setprecision(15) << maxValue << "\"";
        const bool writeChildren = !pivotFieldIsPureData(pt, i);
        if (writeChildren) xml << " count=\"" << items.size() << "\"";

        if (!writeChildren || items.empty()) {
            xml << "/>";
        } else {
            xml << ">";
            for (const auto& value : items) writePivotValue(xml, value);
            xml << "</sharedItems>";
        }
        xml << "</cacheField>";
    }
    xml << "</cacheFields></pivotCacheDefinition>";
    return xml.str();
}

std::string pivotCacheRecordsXml(const xlpp::PivotTable& pt, bool strict) {
    std::ostringstream xml;
    const auto& cache = pt.cache();
    const auto sharedItems = pivotSharedItems(cache);
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotCacheRecords xmlns=\""
        << nsMain(strict) << "\" count=\"" << cache.records().size() << "\">";
    for (const auto& record : cache.records()) {
        if (record.size() != cache.fields().size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        xml << "<r>";
        for (std::size_t fieldIndex = 0; fieldIndex < record.size(); ++fieldIndex) {
            if (pivotFieldIsPureData(pt, fieldIndex)) {
                writePivotValue(xml, record[fieldIndex]);
            } else {
                const auto& items = sharedItems[fieldIndex];
                const auto item = std::find(items.begin(), items.end(), record[fieldIndex]);
                if (item == items.end()) throw std::logic_error("Pivot shared item lookup failed");
                xml << "<x v=\"" << std::distance(items.begin(), item) << "\"/>";
            }
        }
        xml << "</r>";
    }
    xml << "</pivotCacheRecords>";
    return xml.str();
}

std::string quotePivotSheetName(const std::string& name) {
    std::string escaped = name;
    std::size_t pos = 0;
    while ((pos = escaped.find('\'', pos)) != std::string::npos) {
        escaped.insert(pos, 1, '\'');
        pos += 2;
    }
    return "'" + escaped + "'";
}

std::string pivotCellText(const xlpp::Cell* cell) {
    if (!cell || !cell->hasValue()) return {};
    const auto& value = cell->value();
    if (const auto* text = std::get_if<std::string>(&value)) return *text;
    if (const auto* number = std::get_if<double>(&value)) {
        std::ostringstream out;
        out << std::setprecision(15) << *number;
        return out.str();
    }
    if (const auto* boolean = std::get_if<bool>(&value)) return *boolean ? "true" : "false";
    if (const auto* error = std::get_if<xlpp::CellError>(&value)) return xlpp::toString(*error);
    if (const auto* date = std::get_if<xlpp::DateTime>(&value)) return xlpp::toIso8601(*date);
    return {};
}

struct PivotSourceReference {
    std::string sheetName;
    xlpp::CellReference first;
    xlpp::CellReference last;
};

PivotSourceReference parsePivotSourceReference(const std::string& sourceData, const std::string& defaultSheet) {
    std::string sheetName = defaultSheet;
    std::string rangeText = sourceData;
    const auto bang = rangeText.find('!');
    if (bang != std::string::npos) {
        sheetName = rangeText.substr(0, bang);
        rangeText = rangeText.substr(bang + 1);
        if (sheetName.size() >= 2 && sheetName.front() == '\'' && sheetName.back() == '\'') {
            sheetName = sheetName.substr(1, sheetName.size() - 2);
            std::size_t pos = 0;
            while ((pos = sheetName.find("''", pos)) != std::string::npos) sheetName.replace(pos, 2, "'");
        }
    }
    const auto colon = rangeText.find(':');
    auto first = xlpp::CellReference::parse(colon == std::string::npos ? rangeText : rangeText.substr(0, colon));
    auto last = xlpp::CellReference::parse(colon == std::string::npos ? rangeText : rangeText.substr(colon + 1));
    if (first.row > last.row) std::swap(first.row, last.row);
    if (first.column > last.column) std::swap(first.column, last.column);
    return {std::move(sheetName), first, last};
}

xlpp::PivotTable effectivePivotTable(const xlpp::PivotTable& source,
                                     const std::deque<xlpp::Worksheet>& sheets,
                                     const xlpp::Worksheet& owner,
                                     std::size_t cacheId) {
    auto result = source;
    result.cache().setCacheId(static_cast<int>(cacheId));
    if (result.cache().sourceData().empty()) {
        result.cache().setSourceData(quotePivotSheetName(owner.name()) + "!" + owner.dimensions());
    }

    if (result.cache().fields().empty() || result.cache().records().empty()) {
        const auto parsed = parsePivotSourceReference(result.cache().sourceData(), owner.name());
        const auto sourceSheet = std::find_if(sheets.begin(), sheets.end(), [&](const auto& sheet) {
            return sheet.name() == parsed.sheetName;
        });
        if (sourceSheet == sheets.end())
            throw std::invalid_argument("Pivot source worksheet not found: " + parsed.sheetName);

        const auto width = parsed.last.column - parsed.first.column + 1;
        if (result.cache().fields().empty()) {
            std::vector<std::string> fields;
            fields.reserve(width);
            for (std::size_t column = parsed.first.column; column <= parsed.last.column; ++column) {
                auto name = pivotCellText(sourceSheet->tryCell(parsed.first.row, column));
                if (name.empty()) name = "Field" + std::to_string(column - parsed.first.column + 1);
                if (std::find(fields.begin(), fields.end(), name) != fields.end())
                    name += "_" + std::to_string(column - parsed.first.column + 1);
                fields.push_back(std::move(name));
            }
            result.cache().setFields(std::move(fields));
        }
        if (result.cache().fields().size() != width)
            throw std::invalid_argument("Pivot source width does not match cache field count");

        if (result.cache().records().empty() && parsed.first.row < parsed.last.row) {
            for (std::size_t row = parsed.first.row + 1; row <= parsed.last.row; ++row) {
                std::vector<std::string> record;
                record.reserve(width);
                for (std::size_t column = parsed.first.column; column <= parsed.last.column; ++column)
                    record.push_back(pivotCellText(sourceSheet->tryCell(row, column)));
                result.cache().addRecord(std::move(record));
            }
        }
    }
    return result;
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
        xml << "<comment ref=\"" << pair.second.address() << "\" authorId=\"" << authorId << "\" shapeId=\"0\"><text><t xml:space=\"preserve\">"
            << xmlEscape(comment.text()) << "</t></text></comment>";
    }
    xml << "</commentList></comments>";
    return xml.str();
}

std::string commentsVml(const xlpp::Worksheet& sheet) {
    std::ostringstream xml;
    xml << R"(<?xml version="1.0" encoding="UTF-8"?><xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>)";
    std::size_t shapeId = 1026;
    for (const auto& pair : sheet.cells()) {
        if (!pair.second.hasComment()) continue;
        xml << "<v:shape id=\"_x0000_s" << shapeId++ << "\" type=\"#_x0000_t202\" style=\"position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:144px;height:79px;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>"
            << (pair.second.row() - 1) << "</x:Row><x:Column>" << (pair.second.column() - 1) << "</x:Column></x:ClientData></v:shape>";
    }
    xml << "</xml>";
    return xml.str();
}


std::string drawingXml(const xlpp::Worksheet& sheet, bool strict) {
    std::ostringstream xml;
    const auto drawingNs = strict ? "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" : "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const auto drawingMainNs = strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main";
    const auto chartNs = strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart";
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"" << drawingNs
        << "\" xmlns:a=\"" << drawingMainNs << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">";
    std::size_t objectId = 1;
    for (std::size_t imageIndex = 0; imageIndex < sheet.images().size(); ++imageIndex) {
        const auto& image = sheet.images()[imageIndex];
        const auto ref = xlpp::CellReference::parse(image.anchor());
        const auto cx = static_cast<long long>(image.widthPixels() * 9525.0);
        const auto cy = static_cast<long long>(image.heightPixels() * 9525.0);
        xml << "<xdr:oneCellAnchor><xdr:from><xdr:col>" << (ref.column-1) << "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << (ref.row-1) << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            << "<xdr:ext cx=\"" << cx << "\" cy=\"" << cy << "\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" << objectId++ << "\" name=\"" << xmlEscape(image.name()) << "\"/><xdr:cNvPicPr/></xdr:nvPicPr>"
            << "<xdr:blipFill><a:blip r:embed=\"rIdImage" << imageIndex + 1 << "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";
    }
    for (std::size_t chartIndex = 0; chartIndex < sheet.chartCount(); ++chartIndex) {
        const auto& chart = sheet.chart(chartIndex);
        const auto widthEmu = static_cast<long long>(chart.width()) * 9525LL;
        const auto heightEmu = static_cast<long long>(chart.height()) * 9525LL;
        xml << "<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << (chartIndex * 20)
            << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"" << widthEmu << "\" cy=\"" << heightEmu << "\"/>"
            << "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" << objectId++ << "\" name=\"Chart " << chartIndex + 1
            << "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" << widthEmu << "\" cy=\"" << heightEmu
            << "\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"" << chartNs << "\"><c:chart xmlns:c=\"" << chartNs
            << "\" r:id=\"rIdChart" << chartIndex + 1 << "\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>";
    }
    xml << "</xdr:wsDr>";
    return xml.str();
}
std::string drawingRelationshipsXml(const xlpp::Worksheet& sheet, std::size_t firstMediaId, std::size_t firstChartId, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
    for (std::size_t i = 0; i < sheet.images().size(); ++i)
        xml << "<Relationship Id=\"rIdImage" << i+1 << "\" Type=\"" << nsRelsDoc(strict) << "/image\" Target=\"../media/image" << firstMediaId+i << '.' << sheet.images()[i].extension() << "\"/>";
    for (std::size_t i = 0; i < sheet.chartCount(); ++i)
        xml << "<Relationship Id=\"rIdChart" << i+1 << "\" Type=\"" << nsRelsDoc(strict) << "/chart\" Target=\"../charts/chart" << firstChartId+i << ".xml\"/>";
    xml << "</Relationships>";
    return xml.str();
}

std::string tableXml(const xlpp::Table& table, const xlpp::Worksheet& sheet, std::size_t id, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><table xmlns=\"" << nsMain(strict) << "\""
        << " id=\"" << id << "\" name=\"" << xmlEscape(table.name())
        << "\" displayName=\"" << xmlEscape(table.displayName()) << "\" ref=\""
        << xmlEscape(table.reference()) << "\" headerRowCount=\"" << (table.showHeaderRow() ? 1 : 0)
        << "\" totalsRowShown=\"" << (table.showTotalsRow() ? 1 : 0) << "\">";
    xml << "<autoFilter ref=\"" << xmlEscape(table.reference()) << "\"/>";
    std::vector<std::string> generatedColumns;
    if (table.columns().empty()) {
        const auto separator = table.reference().find(':');
        const auto first = xlpp::CellReference::parse(separator == std::string::npos ? table.reference() : table.reference().substr(0, separator));
        const auto last = xlpp::CellReference::parse(separator == std::string::npos ? table.reference() : table.reference().substr(separator + 1));
        if (first.column > last.column) throw std::invalid_argument("Invalid table range: " + table.reference());
        for (std::size_t column = first.column; column <= last.column; ++column) {
            std::string header = "Column" + std::to_string(column - first.column + 1);
            if (const auto* cell = sheet.tryCell(first.row, column)) {
                if (const auto* text = std::get_if<std::string>(&cell->value()); text && !text->empty()) header = *text;
            }
            const auto base = header;
            std::size_t suffix = 2;
            while (std::find(generatedColumns.begin(), generatedColumns.end(), header) != generatedColumns.end())
                header = base + "_" + std::to_string(suffix++);
            generatedColumns.push_back(std::move(header));
        }
    }
    const auto columnCount = table.columns().empty() ? generatedColumns.size() : table.columns().size();
    xml << "<tableColumns count=\"" << columnCount << "\">";
    if (table.columns().empty()) {
        for (std::size_t i = 0; i < generatedColumns.size(); ++i)
            xml << "<tableColumn id=\"" << i + 1 << "\" name=\"" << xmlEscape(generatedColumns[i]) << "\"/>";
    } else {
        for (const auto& column : table.columns())
            xml << "<tableColumn id=\"" << column.id() << "\" name=\"" << xmlEscape(column.name()) << "\"/>";
    }
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
// A cache snapshot is retained for deterministic output, but mutable-reference APIs require re-serialization for correctness.
std::vector<std::string> serializeSheets(const std::deque<xlpp::Worksheet>& sheets,
                                         const StyleCatalog& styles, const DxfCatalog& dxfs,
                                         bool date1904, bool strict, bool macroEnabled, std::size_t workers,
                                         bool parallelRows,
                                         const std::unordered_map<std::string, std::size_t>* sstIndex,
                                         std::vector<std::string>* cache,
                                         bool& cacheStrict, bool& cacheDate1904) {
    std::vector<std::string> result(sheets.size());
    std::vector<char> needsSerialize(sheets.size(), true);
    if (cache) {
        // Worksheet exposes mutable references (Cell&, Style&, AutoFilter&, ...).
        // A caller may keep one of those references and mutate it after a save,
        // bypassing Worksheet::dirty_. Re-serializing is therefore required for
        // correctness until mutations are tracked by owning proxy objects.
        if (cacheStrict != strict || cacheDate1904 != date1904) cache->clear();
    }

    // Serialize dirty sheets in parallel
    std::vector<std::size_t> dirtyIndexes;
    for (std::size_t i = 0; i < sheets.size(); ++i)
        if (needsSerialize[i]) dirtyIndexes.push_back(i);

    auto serializeOne = [&](std::size_t i) {
        const auto codeName = macroEnabled ? "Sheet" + std::to_string(i + 1) : std::string{};
        result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, 0, codeName);
    };

    if (workers > 1 && dirtyIndexes.size() > 1) {
        xlpp::internal::ThreadPool pool(std::min(workers, dirtyIndexes.size()));
        pool.parallelFor(0, dirtyIndexes.size(), [&](std::size_t j) {
            serializeOne(dirtyIndexes[j]);
        });
    } else if (parallelRows && workers > 1 && dirtyIndexes.size() == 1) {
        for (auto i : dirtyIndexes) {
            const auto codeName = macroEnabled ? "Sheet" + std::to_string(i + 1) : std::string{};
            result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, workers, codeName);
        }
    } else {
        for (auto i : dirtyIndexes) serializeOne(i);
    }

    // Update cache
    if (cache) {
        if (cache->size() < sheets.size()) cache->resize(sheets.size());
        for (std::size_t i = 0; i < sheets.size(); ++i)
            (*cache)[i] = result[i];
        cacheStrict = strict;
        cacheDate1904 = date1904;
    }
    return result;
}

std::string resolvePackagePart(const std::string& basePart, std::string relativeTarget) {
    if (relativeTarget.empty()) return {};
    const bool absoluteTarget = relativeTarget.front() == '/';
    if (absoluteTarget) relativeTarget.erase(relativeTarget.begin());
    std::vector<std::string> segments;
    const auto slash = basePart.find_last_of('/');
    std::string combined = absoluteTarget
        ? relativeTarget
        : (slash == std::string::npos ? std::string{} : basePart.substr(0, slash + 1)) + relativeTarget;
    std::size_t begin = 0;
    while (begin <= combined.size()) {
        const auto end = combined.find('/', begin);
        const auto segment = combined.substr(begin, end == std::string::npos ? std::string::npos : end - begin);
        if (segment == "..") { if (!segments.empty()) segments.pop_back(); }
        else if (!segment.empty() && segment != ".") segments.push_back(segment);
        if (end == std::string::npos) break;
        begin = end + 1;
    }
    std::ostringstream result;
    for (std::size_t i = 0; i < segments.size(); ++i) { if (i) result << '/'; result << segments[i]; }
    return result.str();
}


std::string relationshipKind(const xlpp::PreservedRelationship& relationship) {
    const auto slash = relationship.type.find_last_of('/');
    if (slash != std::string::npos && slash + 1 < relationship.type.size())
        return relationship.type.substr(slash + 1);
    const auto& target = relationship.target;
    if (target.find("/drawings/") != std::string::npos || target.rfind("../drawings/", 0) == 0) return "drawing";
    if (target.find("/pivotTables/") != std::string::npos || target.rfind("../pivotTables/", 0) == 0) return "pivotTable";
    if (target.find("/pivotCache/") != std::string::npos || target.rfind("pivotCache/", 0) == 0) return "pivotCacheDefinition";
    return {};
}

std::vector<xlpp::PreservedRelationship> relationshipsForSource(
    const std::vector<xlpp::PreservedRelationship>& relationships,
    const std::string& sourcePart) {
    std::vector<xlpp::PreservedRelationship> result;
    for (const auto& relationship : relationships)
        if (relationship.sourcePart == sourcePart) result.push_back(relationship);
    return result;
}

bool sameRelationship(const xlpp::PreservedRelationship& lhs,
                      const xlpp::PreservedRelationship& rhs) {
    return lhs.type == rhs.type && lhs.target == rhs.target && lhs.targetMode == rhs.targetMode;
}

std::string allocateRelationshipId(const std::set<std::string>& used) {
    for (std::size_t index = 1;; ++index) {
        const auto candidate = "rIdXLPP" + std::to_string(index);
        if (!used.count(candidate)) return candidate;
    }
}

void replaceRelationshipReference(std::string& ownerXml,
                                  const std::string& oldId,
                                  const std::string& newId) {
    if (oldId == newId || oldId.empty()) return;
    const std::array<std::string, 2> patterns{
        "r:id=\"" + oldId + "\"",
        "r:id='" + oldId + "'"
    };
    const std::array<std::string, 2> replacements{
        "r:id=\"" + newId + "\"",
        "r:id='" + newId + "'"
    };
    for (std::size_t p = 0; p < patterns.size(); ++p) {
        std::size_t position = 0;
        while ((position = ownerXml.find(patterns[p], position)) != std::string::npos) {
            ownerXml.replace(position, patterns[p].size(), replacements[p]);
            position += replacements[p].size();
        }
    }
}

std::string mergeRelationshipsXml(
    const std::string& generatedXml,
    const std::vector<xlpp::PreservedRelationship>& original,
    const std::function<bool(const xlpp::PreservedRelationship&)>& preserve,
    bool strict,
    std::string* generatedOwnerXml) {
    auto generated = xlpp::internal::RelationshipGraph::parseRelationshipsXml({}, generatedXml);
    std::vector<xlpp::PreservedRelationship> selected;
    for (const auto& relationship : original)
        if (preserve(relationship)) selected.push_back(relationship);

    std::set<std::string> originalIds;
    for (const auto& relationship : selected) originalIds.insert(relationship.id);
    std::set<std::string> used = originalIds;
    for (const auto& relationship : generated) used.insert(relationship.id);

    for (auto& relationship : generated) {
        const auto collision = std::find_if(selected.begin(), selected.end(), [&](const auto& candidate) {
            return candidate.id == relationship.id && !sameRelationship(candidate, relationship);
        });
        if (collision == selected.end()) continue;
        const auto oldId = relationship.id;
        const auto newId = allocateRelationshipId(used);
        used.insert(newId);
        relationship.id = newId;
        if (generatedOwnerXml) replaceRelationshipReference(*generatedOwnerXml, oldId, newId);
    }

    std::vector<xlpp::PreservedRelationship> merged = std::move(generated);
    for (const auto& relationship : selected) {
        const auto duplicate = std::find_if(merged.begin(), merged.end(), [&](const auto& candidate) {
            return sameRelationship(candidate, relationship);
        });
        if (duplicate == merged.end()) merged.push_back(relationship);
    }
    return xlpp::internal::RelationshipGraph::serializeRelationships(merged, strict);
}

std::vector<std::string> extractTagBlocks(const std::string& xml, const std::string& tag) {
    return xlpp::internal::tags(xml, tag);
}

void eraseTagBlocks(std::string& xml, const std::string& tag) {
    for (const auto& block : extractTagBlocks(xml, tag)) {
        std::size_t position = 0;
        while ((position = xml.find(block, position)) != std::string::npos)
            xml.erase(position, block.size());
    }
}

std::string joinBlocks(const std::vector<std::string>& blocks) {
    std::string result;
    for (const auto& block : blocks) result += block;
    return result;
}

void insertBefore(std::string& xml, const std::string& marker, const std::string& content) {
    if (content.empty()) return;
    const auto position = xml.find(marker);
    if (position == std::string::npos) return;
    xml.insert(position, content);
}


const xlpp::PreservedPart* findPreservedPart(const std::vector<xlpp::PreservedPart>& parts,
                                             const std::string& name) {
    const auto it = std::find_if(parts.begin(), parts.end(), [&](const auto& part) { return part.name == name; });
    return it == parts.end() ? nullptr : &*it;
}

std::size_t maximumDrawingObjectId(const std::string& drawingXmlText) {
    std::size_t maximum = 0;
    const auto inspect = [&](const std::vector<std::string>& nodes) {
        for (const auto& node : nodes) {
            const auto value = xlpp::internal::attribute(node, "id");
            if (value.empty() || !std::all_of(value.begin(), value.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
            maximum = std::max(maximum, static_cast<std::size_t>(std::stoull(value)));
        }
    };
    inspect(xlpp::internal::tags(drawingXmlText, "xdr:cNvPr"));
    inspect(xlpp::internal::tags(drawingXmlText, "cNvPr"));
    return maximum;
}

std::string appendedImageAnchorXml(const xlpp::Image& image,
                                   const std::string& relationshipId,
                                   std::size_t objectId,
                                   bool strict) {
    const auto ref = xlpp::CellReference::parse(image.anchor());
    const auto cx = static_cast<long long>(image.widthPixels() * 9525.0);
    const auto cy = static_cast<long long>(image.heightPixels() * 9525.0);
    const auto drawingNs = strict ? "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                                  : "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const auto drawingMainNs = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                      : "http://schemas.openxmlformats.org/drawingml/2006/main";
    std::ostringstream xml;
    xml << "<xdr:oneCellAnchor xmlns:xdr=\"" << drawingNs << "\" xmlns:a=\"" << drawingMainNs
        << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\"><xdr:from><xdr:col>" << (ref.column - 1)
        << "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << (ref.row - 1)
        << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"" << cx << "\" cy=\"" << cy
        << "\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" << objectId << "\" name=\"" << xmlEscape(image.name())
        << "\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"" << relationshipId
        << "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/>"
        << "</a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";
    return xml.str();
}

std::string appendedChartAnchorXml(const xlpp::Chart& chart,
                                   const std::string& relationshipId,
                                   std::size_t objectId,
                                   std::size_t placementIndex,
                                   bool strict) {
    const auto widthEmu = static_cast<long long>(chart.width()) * 9525LL;
    const auto heightEmu = static_cast<long long>(chart.height()) * 9525LL;
    const auto drawingNs = strict ? "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                                  : "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const auto drawingMainNs = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                      : "http://schemas.openxmlformats.org/drawingml/2006/main";
    const auto chartNs = strict ? "http://purl.oclc.org/ooxml/drawingml/chart"
                                : "http://schemas.openxmlformats.org/drawingml/2006/chart";
    std::ostringstream xml;
    xml << "<xdr:oneCellAnchor xmlns:xdr=\"" << drawingNs << "\" xmlns:a=\"" << drawingMainNs
        << "\" xmlns:c=\"" << chartNs << "\" xmlns:r=\"" << nsRelsDoc(strict)
        << "\"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" << (placementIndex * 20)
        << "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"" << widthEmu << "\" cy=\"" << heightEmu
        << "\"/><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" << objectId
        << "\" name=\"Chart " << objectId << "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
        << "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" << widthEmu << "\" cy=\"" << heightEmu
        << "\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"" << chartNs << "\"><c:chart r:id=\""
        << relationshipId << "\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>";
    return xml.str();
}

std::vector<std::string> drawingTags(const std::string& xml, const char* prefixed, const char* local);
std::string seriesDirectSpPr(const std::string& seriesXml);
std::string partExtension(const std::string& part);

bool replaceSimpleDrawingText(std::string& xml, const char* prefixed, const char* local, long long value) {
    const std::array<std::string, 2> names{prefixed, local};
    for (const auto& name : names) {
        if (name.empty()) continue;
        const auto open = "<" + name + ">";
        const auto close = "</" + name + ">";
        const auto begin = xml.find(open);
        if (begin == std::string::npos) continue;
        const auto textBegin = begin + open.size();
        const auto finish = xml.find(close, textBegin);
        if (finish == std::string::npos) continue;
        xml.replace(textBegin, finish - textBegin, std::to_string(value));
        return true;
    }
    return false;
}

bool replaceAttributeInNode(std::string& node, const std::string& attributeName, long long value) {
    const auto key = attributeName + "=\"";
    const auto begin = node.find(key);
    if (begin == std::string::npos) return false;
    const auto valueBegin = begin + key.size();
    const auto valueEnd = node.find('"', valueBegin);
    if (valueEnd == std::string::npos) return false;
    node.replace(valueBegin, valueEnd - valueBegin, std::to_string(value));
    return true;
}

bool patchFirstDrawingNodeAttributes(std::string& anchor,
                                     const char* prefixed,
                                     const char* local,
                                     long long first,
                                     long long second,
                                     const char* firstAttribute,
                                     const char* secondAttribute) {
    auto nodes = drawingTags(anchor, prefixed, local);
    if (nodes.empty()) return false;
    auto node = nodes.front();
    if (!replaceAttributeInNode(node, firstAttribute, first) ||
        !replaceAttributeInNode(node, secondAttribute, second)) return false;
    const auto position = anchor.find(nodes.front());
    if (position == std::string::npos) return false;
    anchor.replace(position, nodes.front().size(), node);
    return true;
}

bool patchDrawingMarker(std::string& anchor, const char* prefixed, const char* local,
                        const xlpp::DrawingMarker& marker, bool updateOffsets = false) {
    auto markers = drawingTags(anchor, prefixed, local);
    if (markers.empty()) return false;
    auto markerXml = markers.front();
    if (!replaceSimpleDrawingText(markerXml, "xdr:col", "col", static_cast<long long>(marker.column) - 1)) return false;
    if (!replaceSimpleDrawingText(markerXml, "xdr:row", "row", static_cast<long long>(marker.row) - 1)) return false;
    if (updateOffsets) {
        if (!replaceSimpleDrawingText(markerXml, "xdr:colOff", "colOff", marker.columnOffsetEmu)) return false;
        if (!replaceSimpleDrawingText(markerXml, "xdr:rowOff", "rowOff", marker.rowOffsetEmu)) return false;
    }
    // Normal moves deliberately retain producer-native sub-cell offsets. A
    // two-cell resize, however, must update the terminal offsets because those
    // offsets participate directly in the visible width/height.
    const auto position = anchor.find(markers.front());
    if (position == std::string::npos) return false;
    anchor.replace(position, markers.front().size(), markerXml);
    return true;
}

std::string drawingObjectIdFromStableId(const std::string& stableId) {
    const auto hash = stableId.rfind('#');
    return hash == std::string::npos ? std::string{} : stableId.substr(hash + 1);
}

bool anchorMatchesStableId(const std::string& anchor, const std::string& stableId) {
    const auto objectId = drawingObjectIdFromStableId(stableId);
    if (objectId.empty()) return false;
    const auto pictures = drawingTags(anchor, "xdr:pic", "pic");
    for (const auto& picture : pictures) {
        const auto properties = drawingTags(picture, "xdr:cNvPr", "cNvPr");
        for (const auto& property : properties)
            if (xlpp::internal::attribute(property, "id") == objectId) return true;
    }
    return false;
}

bool anchorReferencesRelationship(const std::string& anchor, const std::string& relationshipId) {
    const auto pictures = drawingTags(anchor, "xdr:pic", "pic");
    for (const auto& picture : pictures) {
        const auto blips = drawingTags(picture, "a:blip", "blip");
        for (const auto& blip : blips) {
            if (xlpp::internal::attribute(blip, "r:embed") == relationshipId ||
                xlpp::internal::attribute(blip, "r:link") == relationshipId) return true;
        }
    }
    return false;
}

std::string* findImageAnchorBlock(std::vector<std::string>& anchors,
                                  const std::string& stableId,
                                  const std::string& relationshipId) {
    for (auto& anchor : anchors)
        if (anchorMatchesStableId(anchor, stableId)) return &anchor;
    for (auto& anchor : anchors)
        if (anchorReferencesRelationship(anchor, relationshipId)) return &anchor;
    return nullptr;
}

bool drawingReferencesRelationship(const std::string& drawingXmlText, const std::string& relationshipId) {
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        for (const auto& anchor : drawingTags(drawingXmlText, prefixed, local))
            if (anchorReferencesRelationship(anchor, relationshipId)) return true;
    }
    return false;
}

bool patchImportedImageAnchor(std::string& drawingXmlText,
                              const std::string& stableId,
                              const std::string& relationshipId,
                              const xlpp::DrawingAnchorInfo& anchorInfo,
                              bool moved,
                              bool resized,
                              bool remove) {
    std::vector<std::string> anchors;
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        auto family = drawingTags(drawingXmlText, prefixed, local);
        anchors.insert(anchors.end(), std::make_move_iterator(family.begin()), std::make_move_iterator(family.end()));
    }
    auto* found = findImageAnchorBlock(anchors, stableId, relationshipId);
    if (!found) return false;
    const auto original = *found;
    auto patched = original;
    if (remove) {
        const auto position = drawingXmlText.find(original);
        if (position == std::string::npos) return false;
        drawingXmlText.erase(position, original.size());
        return true;
    }

    if (moved) {
        if (anchorInfo.type == xlpp::DrawingAnchorType::Absolute) {
            if (!patchFirstDrawingNodeAttributes(patched, "xdr:pos", "pos",
                                                 anchorInfo.xEmu, anchorInfo.yEmu, "x", "y")) return false;
        } else {
            if (!patchDrawingMarker(patched, "xdr:from", "from", anchorInfo.from)) return false;
            if (anchorInfo.type == xlpp::DrawingAnchorType::TwoCell &&
                !patchDrawingMarker(patched, "xdr:to", "to", anchorInfo.to)) return false;
        }
    }
    if (resized) {
        if (anchorInfo.type == xlpp::DrawingAnchorType::TwoCell) {
            // Two-cell anchors derive their visible geometry primarily from
            // from/to. Patch the terminal marker as well as a:xfrm/a:ext so
            // Excel/LibreOffice do not normalize the requested size back to
            // the old cell span on the next save.
            if (!patchDrawingMarker(patched, "xdr:to", "to", anchorInfo.to, true)) return false;
            auto pictures = drawingTags(patched, "xdr:pic", "pic");
            if (pictures.empty()) return false;
            auto picture = pictures.front();
            if (!patchFirstDrawingNodeAttributes(picture, "a:ext", "ext",
                                                 anchorInfo.widthEmu, anchorInfo.heightEmu, "cx", "cy")) return false;
            const auto picturePosition = patched.find(pictures.front());
            if (picturePosition == std::string::npos) return false;
            patched.replace(picturePosition, pictures.front().size(), picture);
        } else {
            if (!patchFirstDrawingNodeAttributes(patched, "xdr:ext", "ext",
                                                 anchorInfo.widthEmu, anchorInfo.heightEmu, "cx", "cy")) return false;
        }
    }
    const auto position = drawingXmlText.find(original);
    if (position == std::string::npos) return false;
    drawingXmlText.replace(position, original.size(), patched);
    return true;
}


bool chartAnchorMatchesStableId(const std::string& anchor, const std::string& stableId) {
    const auto objectId = drawingObjectIdFromStableId(stableId);
    if (objectId.empty()) return false;
    const auto frames = drawingTags(anchor, "xdr:graphicFrame", "graphicFrame");
    for (const auto& frame : frames) {
        const auto properties = drawingTags(frame, "xdr:cNvPr", "cNvPr");
        for (const auto& property : properties)
            if (xlpp::internal::attribute(property, "id") == objectId) return true;
    }
    return false;
}

bool chartAnchorReferencesRelationship(const std::string& anchor, const std::string& relationshipId) {
    const auto frames = drawingTags(anchor, "xdr:graphicFrame", "graphicFrame");
    for (const auto& frame : frames) {
        const auto charts = drawingTags(frame, "c:chart", "chart");
        for (const auto& chart : charts)
            if (xlpp::internal::attribute(chart, "r:id") == relationshipId) return true;
    }
    return false;
}

bool drawingReferencesChartRelationship(const std::string& drawingXmlText, const std::string& relationshipId) {
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        for (const auto& anchor : drawingTags(drawingXmlText, prefixed, local))
            if (chartAnchorReferencesRelationship(anchor, relationshipId)) return true;
    }
    return false;
}

bool removeImportedChartAnchor(std::string& drawingXmlText,
                               const std::string& stableId,
                               const std::string& relationshipId) {
    std::vector<std::string> anchors;
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        auto family = drawingTags(drawingXmlText, prefixed, local);
        anchors.insert(anchors.end(), std::make_move_iterator(family.begin()), std::make_move_iterator(family.end()));
    }
    for (const auto& candidate : anchors) {
        if (!chartAnchorMatchesStableId(candidate, stableId) &&
            !chartAnchorReferencesRelationship(candidate, relationshipId)) continue;
        const auto position = drawingXmlText.find(candidate);
        if (position == std::string::npos) return false;
        drawingXmlText.erase(position, candidate.size());
        return true;
    }
    return false;
}

bool patchImportedChartAnchor(std::string& drawingXmlText,
                              const std::string& stableId,
                              const std::string& relationshipId,
                              const xlpp::DrawingAnchorInfo& anchorInfo,
                              bool moved,
                              bool resized) {
    std::vector<std::string> anchors;
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 3>{
            std::pair{"xdr:oneCellAnchor", "oneCellAnchor"},
            std::pair{"xdr:twoCellAnchor", "twoCellAnchor"},
            std::pair{"xdr:absoluteAnchor", "absoluteAnchor"}}) {
        auto family = drawingTags(drawingXmlText, prefixed, local);
        anchors.insert(anchors.end(), std::make_move_iterator(family.begin()), std::make_move_iterator(family.end()));
    }
    std::string* found = nullptr;
    for (auto& candidate : anchors) {
        if (chartAnchorMatchesStableId(candidate, stableId)) { found = &candidate; break; }
    }
    if (!found) {
        for (auto& candidate : anchors) {
            if (chartAnchorReferencesRelationship(candidate, relationshipId)) { found = &candidate; break; }
        }
    }
    if (!found) return false;

    const auto original = *found;
    auto patched = original;
    if (moved) {
        if (anchorInfo.type == xlpp::DrawingAnchorType::Absolute) {
            if (!patchFirstDrawingNodeAttributes(patched, "xdr:pos", "pos",
                                                 anchorInfo.xEmu, anchorInfo.yEmu, "x", "y")) return false;
        } else {
            if (!patchDrawingMarker(patched, "xdr:from", "from", anchorInfo.from)) return false;
            if (anchorInfo.type == xlpp::DrawingAnchorType::TwoCell &&
                !patchDrawingMarker(patched, "xdr:to", "to", anchorInfo.to)) return false;
        }
    }
    if (resized) {
        if (anchorInfo.type == xlpp::DrawingAnchorType::TwoCell) {
            if (!patchDrawingMarker(patched, "xdr:to", "to", anchorInfo.to, true)) return false;
            auto frames = drawingTags(patched, "xdr:graphicFrame", "graphicFrame");
            if (!frames.empty()) {
                auto frame = frames.front();
                const auto extNodes = drawingTags(frame, "a:ext", "ext");
                if (!extNodes.empty()) {
                    if (!patchFirstDrawingNodeAttributes(frame, "a:ext", "ext",
                                                         anchorInfo.widthEmu, anchorInfo.heightEmu, "cx", "cy")) return false;
                    const auto framePosition = patched.find(frames.front());
                    if (framePosition == std::string::npos) return false;
                    patched.replace(framePosition, frames.front().size(), frame);
                }
            }
        } else {
            if (!patchFirstDrawingNodeAttributes(patched, "xdr:ext", "ext",
                                                 anchorInfo.widthEmu, anchorInfo.heightEmu, "cx", "cy")) return false;
        }
    }
    const auto position = drawingXmlText.find(original);
    if (position == std::string::npos) return false;
    drawingXmlText.replace(position, original.size(), patched);
    return true;
}

bool replaceSimpleElementText(std::string& xml, const char* prefixed, const char* local, const std::string& value) {
    const auto nodes = drawingTags(xml, prefixed, local);
    if (nodes.empty()) return false;
    const auto& original = nodes.front();
    const auto openEnd = original.find('>');
    const auto closeBegin = original.rfind("</");
    if (openEnd == std::string::npos || closeBegin == std::string::npos || closeBegin < openEnd) return false;
    auto patched = original;
    patched.replace(openEnd + 1, closeBegin - openEnd - 1, xmlEscape(value));
    const auto position = xml.find(original);
    if (position == std::string::npos) return false;
    xml.replace(position, original.size(), patched);
    return true;
}

void eraseChartCacheBlocks(std::string& xml) {
    for (const auto& [prefixed, local] : std::array<std::pair<const char*, const char*>, 2>{
            std::pair{"c:strCache", "strCache"}, std::pair{"c:numCache", "numCache"}}) {
        for (;;) {
            const auto nodes = drawingTags(xml, prefixed, local);
            if (nodes.empty()) break;
            const auto position = xml.find(nodes.front());
            if (position == std::string::npos) break;
            xml.erase(position, nodes.front().size());
        }
    }
}

bool patchSeriesReferenceContainer(std::string& seriesXml,
                                   const char* prefixedContainer,
                                   const char* localContainer,
                                   const std::string& reference) {
    const auto containers = drawingTags(seriesXml, prefixedContainer, localContainer);
    if (containers.empty()) return false;
    const auto original = containers.front();
    auto patched = original;
    if (!replaceSimpleElementText(patched, "c:f", "f", reference)) return false;
    eraseChartCacheBlocks(patched);
    const auto position = seriesXml.find(original);
    if (position == std::string::npos) return false;
    seriesXml.replace(position, original.size(), patched);
    return true;
}

std::string generatedChartTitleXml(const std::string& title, bool prefixed, bool strict) {
    const auto c = prefixed ? "c:" : "";
    const auto drawingMain = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                    : "http://schemas.openxmlformats.org/drawingml/2006/main";
    std::ostringstream xml;
    xml << '<' << c << "title><" << c << "tx><" << c << "rich xmlns:a=\"" << drawingMain
        << "\"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(title)
        << "</a:t></a:r></a:p></" << c << "rich></" << c << "tx><" << c
        << "overlay val=\"0\"/></" << c << "title>";
    return xml.str();
}

std::string chartSolidFillXml(const xlpp::ChartColor& color, bool declareNamespace = false);

std::string chartRichTextTxXml(const xlpp::ChartRichText& richText, bool prefixed, bool strict) {
    const auto c = prefixed ? "c:" : "";
    const auto drawingMain = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                    : "http://schemas.openxmlformats.org/drawingml/2006/main";
    std::string xml = "<" + std::string(c) + "tx><" + std::string(c) + "rich xmlns:a=\"" + drawingMain +
                      "\"><a:bodyPr/><a:lstStyle/><a:p>";
    for (const auto& run : richText.runs) {
        xml += "<a:r>";
        if (run.bold || run.italic || run.fontSizePoints > 0.0 || !run.typeface.empty() || run.color.present()) {
            xml += "<a:rPr";
            if (run.bold) xml += " b=\"1\"";
            if (run.italic) xml += " i=\"1\"";
            if (run.fontSizePoints > 0.0)
                xml += " sz=\"" + std::to_string(static_cast<long long>(std::llround(run.fontSizePoints * 100.0))) + "\"";
            xml += ">";
            if (run.color.present()) xml += chartSolidFillXml(run.color, false);
            if (!run.typeface.empty()) xml += "<a:latin typeface=\"" + xmlEscape(run.typeface) + "\"/>";
            xml += "</a:rPr>";
        }
        xml += "<a:t>" + xmlEscape(run.text) + "</a:t></a:r>";
    }
    xml += "</a:p></" + std::string(c) + "rich></" + std::string(c) + "tx>";
    return xml;
}

bool patchImportedChartTitleRichText(std::string& chartXmlText, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chartNode = originalChart;
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    if (plotPosition == std::string::npos) return false;
    std::string titleNode;
    for (const auto& candidate : drawingTags(chartNode, "c:title", "title")) {
        const auto position = chartNode.find(candidate);
        if (position < plotPosition) { titleNode = candidate; break; }
    }
    const bool prefixed = chartNode.find("<c:chart") != std::string::npos;
    const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
    const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
    if (titleNode.empty()) {
        const auto c = prefixed ? "c:" : "";
        const auto generated = "<" + std::string(c) + "title>" + txXml + "<" + std::string(c) +
                               "overlay val=\"0\"/></" + std::string(c) + "title>";
        chartNode.insert(plotPosition, generated);
    } else {
        auto patchedTitle = titleNode;
        const auto txNodes = drawingTags(patchedTitle, "c:tx", "tx");
        if (!txNodes.empty()) {
            const auto position = patchedTitle.find(txNodes.front());
            if (position == std::string::npos) return false;
            patchedTitle.replace(position, txNodes.front().size(), txXml);
        } else {
            const auto close = patchedTitle.rfind("</");
            if (close == std::string::npos) return false;
            patchedTitle.insert(close, txXml);
        }
        const auto position = chartNode.find(titleNode);
        if (position == std::string::npos) return false;
        chartNode.replace(position, titleNode.size(), patchedTitle);
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chartNode);
    return true;
}

bool patchImportedChartTitle(std::string& chartXmlText, const std::string& title) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chartNode = originalChart;
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    if (plotPosition == std::string::npos) return false;

    std::string titleNode;
    for (const auto& candidate : drawingTags(chartNode, "c:title", "title")) {
        const auto position = chartNode.find(candidate);
        if (position < plotPosition) { titleNode = candidate; break; }
    }
    if (title.empty()) {
        if (!titleNode.empty()) {
            const auto position = chartNode.find(titleNode);
            if (position == std::string::npos) return false;
            chartNode.erase(position, titleNode.size());
        }
    } else if (!titleNode.empty()) {
        auto patchedTitle = titleNode;
        if (!replaceSimpleElementText(patchedTitle, "a:t", "t", title) &&
            !replaceSimpleElementText(patchedTitle, "c:v", "v", title)) {
            const auto prefixed = chartNode.find("<c:chart") != std::string::npos;
            const auto strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            patchedTitle = generatedChartTitleXml(title, prefixed, strict);
        }
        const auto position = chartNode.find(titleNode);
        if (position == std::string::npos) return false;
        chartNode.replace(position, titleNode.size(), patchedTitle);
    } else {
        const auto prefixed = chartNode.find("<c:chart") != std::string::npos;
        const auto strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        chartNode.insert(plotPosition, generatedChartTitleXml(title, prefixed, strict));
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chartNode);
    return true;
}

bool patchAxisTitleNode(std::string& axis, const std::string& chartXmlText, const std::string& title) {
    const auto titleNodes = drawingTags(axis, "c:title", "title");
    if (title.empty()) {
        if (!titleNodes.empty()) {
            const auto position = axis.find(titleNodes.front());
            if (position == std::string::npos) return false;
            axis.erase(position, titleNodes.front().size());
        }
    } else if (!titleNodes.empty()) {
        auto patched = titleNodes.front();
        if (!replaceSimpleElementText(patched, "a:t", "t", title) &&
            !replaceSimpleElementText(patched, "c:v", "v", title)) {
            const bool prefixed = axis.find("<c:") != std::string::npos;
            const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            patched = generatedChartTitleXml(title, prefixed, strict);
        }
        const auto position = axis.find(titleNodes.front());
        if (position == std::string::npos) return false;
        axis.replace(position, titleNodes.front().size(), patched);
    } else {
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto generated = generatedChartTitleXml(title, prefixed, strict);
        std::size_t insertion = std::string::npos;
        for (const auto& candidate : std::array<std::pair<const char*, const char*>, 7>{
                 std::pair{"c:numFmt", "numFmt"}, std::pair{"c:majorTickMark", "majorTickMark"},
                 std::pair{"c:minorTickMark", "minorTickMark"}, std::pair{"c:tickLblPos", "tickLblPos"},
                 std::pair{"c:spPr", "spPr"}, std::pair{"c:txPr", "txPr"},
                 std::pair{"c:crossAx", "crossAx"}}) {
            const auto nodes = drawingTags(axis, candidate.first, candidate.second);
            if (nodes.empty()) continue;
            const auto position = axis.find(nodes.front());
            if (position != std::string::npos) insertion = std::min(insertion, position);
        }
        if (insertion == std::string::npos) {
            const auto close = axis.rfind("</");
            if (close == std::string::npos) return false;
            insertion = close;
        }
        axis.insert(insertion, generated);
    }
    return true;
}

bool patchImportedAxisTitle(std::string& chartXmlText,
                            const char* prefixedAxis,
                            const char* localAxis,
                            std::size_t axisIndex,
                            const std::string& title) {
    const auto axes = drawingTags(chartXmlText, prefixedAxis, localAxis);
    if (axisIndex >= axes.size()) return false;
    const auto originalAxis = axes[axisIndex];
    auto axis = originalAxis;
    if (!patchAxisTitleNode(axis, chartXmlText, title)) return false;
    const auto position = chartXmlText.find(originalAxis);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalAxis.size(), axis);
    return true;
}

bool patchImportedAxisTitleById(std::string& chartXmlText, std::uint64_t axisId, const std::string& title) {
    if (axisId == 0) return false;
    const auto tryAxis = [&](const char* prefixedAxis, const char* localAxis) -> bool {
        for (const auto& originalAxis : drawingTags(chartXmlText, prefixedAxis, localAxis)) {
            const auto ids = drawingTags(originalAxis, "c:axId", "axId");
            if (ids.empty()) continue;
            const auto value = xlpp::internal::attribute(ids.front(), "val");
            std::uint64_t parsed = 0;
            try { if (!value.empty()) parsed = std::stoull(value); } catch (...) { continue; }
            if (parsed != axisId) continue;
            auto axis = originalAxis;
            if (!patchAxisTitleNode(axis, chartXmlText, title)) return false;
            const auto position = chartXmlText.find(originalAxis);
            if (position == std::string::npos) return false;
            chartXmlText.replace(position, originalAxis.size(), axis);
            return true;
        }
        return false;
    };
    return tryAxis("c:catAx", "catAx") || tryAxis("c:valAx", "valAx") ||
           tryAxis("c:dateAx", "dateAx") || tryAxis("c:serAx", "serAx");
}

bool patchOpeningTagAttribute(std::string& node, const std::string& name, const std::string& value, bool removeWhenEmpty);
void removeDrawingChild(std::string& container, const char* prefixed, const char* local);
bool patchOrInsertValChild(std::string& container, const char* prefixed, const char* local, const std::string& value, bool insertWhenMissing);
bool patchImportedChartStyle(std::string& chartXmlText, const std::string& style) {
    if (style.empty()) return false;
    const auto existing = drawingTags(chartXmlText, "c:style", "style");
    if (!existing.empty()) {
        auto node = existing.front();
        if (!patchOpeningTagAttribute(node, "val", style, false)) return false;
        const auto pos = chartXmlText.find(existing.front());
        if (pos == std::string::npos) return false;
        chartXmlText.replace(pos, existing.front().size(), node);
        return true;
    }
    const auto charts = drawingTags(chartXmlText, "c:chart", "chart");
    if (charts.empty()) return false;
    const auto pos = chartXmlText.find(charts.front());
    if (pos == std::string::npos) return false;
    const bool prefixed = chartXmlText.find("<c:chartSpace") != std::string::npos;
    chartXmlText.insert(pos, std::string("<") + (prefixed ? "c:" : "") + "style val=\"" + xmlEscape(style) + "\"/>");
    return true;
}

bool patchChartLineFormatInSpPr(std::string& spPr, const xlpp::ChartLineFormat& format);
bool patchChartFillFormatInSpPr(std::string& spPr, const xlpp::ChartFillFormat& format);
bool ensureChartSpPr(std::string& owner, std::string& spPr, const std::string& beforeXml);
bool patchNestedLineFormat(std::string& owner, const xlpp::ChartLineFormat& format);
std::string axisDirectSpPr(const std::string& axisXml);

bool patchAxisNodeById(std::string& chartXmlText, std::uint64_t axisId,
                       const std::function<bool(std::string&)>& patcher) {
    if (axisId == 0) return false;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
             {"c:catAx", "catAx"}, {"c:valAx", "valAx"}, {"c:dateAx", "dateAx"}, {"c:serAx", "serAx"}}}) {
        for (const auto& originalAxis : drawingTags(chartXmlText, pair.first, pair.second)) {
            const auto ids = drawingTags(originalAxis, "c:axId", "axId");
            if (ids.empty()) continue;
            std::uint64_t parsed = 0;
            try { const auto value=xlpp::internal::attribute(ids.front(),"val"); if(!value.empty()) parsed=std::stoull(value); } catch (...) { continue; }
            if (parsed != axisId) continue;
            auto axis = originalAxis;
            if (!patcher(axis)) return false;
            const auto position = chartXmlText.find(originalAxis);
            if (position == std::string::npos) return false;
            chartXmlText.replace(position, originalAxis.size(), axis);
            return true;
        }
    }
    return false;
}

bool patchImportedAxisTitleRichTextById(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
        const auto titles = drawingTags(axis, "c:title", "title");
        if (!titles.empty()) {
            auto title = titles.front();
            const auto tx = drawingTags(title, "c:tx", "tx");
            if (!tx.empty()) { const auto pos=title.find(tx.front()); if(pos==std::string::npos) return false; title.replace(pos,tx.front().size(),txXml); }
            else { const auto close=title.rfind("</"); if(close==std::string::npos) return false; title.insert(close,txXml); }
            const auto pos=axis.find(titles.front()); if(pos==std::string::npos) return false; axis.replace(pos,titles.front().size(),title);
        } else {
            const auto generated="<"+std::string(c)+"title>"+txXml+"<"+std::string(c)+"overlay val=\"0\"/></"+std::string(c)+"title>";
            std::size_t insertion=axis.rfind("</");
            const auto cross=drawingTags(axis,"c:crossAx","crossAx"); if(!cross.empty()) insertion=axis.find(cross.front());
            if(insertion==std::string::npos) return false; axis.insert(insertion,generated);
        }
        return true;
    });
}

bool patchImportedAxisNumberFormat(std::string& chartXmlText, std::uint64_t axisId, const std::string& formatCode, bool sourceLinked) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto nodes=drawingTags(axis,"c:numFmt","numFmt");
        if(!nodes.empty()) { auto node=nodes.front(); if(!patchOpeningTagAttribute(node,"formatCode",formatCode,false) || !patchOpeningTagAttribute(node,"sourceLinked",sourceLinked?"1":"0",false)) return false; const auto pos=axis.find(nodes.front()); if(pos==std::string::npos) return false; axis.replace(pos,nodes.front().size(),node); return true; }
        const bool prefixed=axis.find("<c:")!=std::string::npos; const auto c=prefixed?"c:":"";
        std::size_t insertion=axis.rfind("</"); const auto major=drawingTags(axis,"c:majorTickMark","majorTickMark"); if(!major.empty()) insertion=axis.find(major.front()); const auto cross=drawingTags(axis,"c:crossAx","crossAx"); if(insertion==std::string::npos && !cross.empty()) insertion=axis.find(cross.front()); if(insertion==std::string::npos) return false;
        axis.insert(insertion,"<"+std::string(c)+"numFmt formatCode=\""+xmlEscape(formatCode)+"\" sourceLinked=\""+(sourceLinked?"1":"0")+"\"/>"); return true;
    });
}

bool patchImportedAxisTicks(std::string& chartXmlText, std::uint64_t axisId, const std::string& major, const std::string& minor, const std::string& labelPos) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        if(!major.empty() && !patchOrInsertValChild(axis,"c:majorTickMark","majorTickMark",major,true)) return false;
        if(!minor.empty() && !patchOrInsertValChild(axis,"c:minorTickMark","minorTickMark",minor,true)) return false;
        if(!labelPos.empty() && !patchOrInsertValChild(axis,"c:tickLblPos","tickLblPos",labelPos,true)) return false;
        return true;
    });
}

bool patchImportedAxisUnits(std::string& chartXmlText, std::uint64_t axisId, double majorUnit, double minorUnit) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        std::ostringstream major; major<<std::setprecision(15)<<majorUnit;
        if(!patchOrInsertValChild(axis,"c:majorUnit","majorUnit",major.str(),true)) return false;
        if(minorUnit>0.0) { std::ostringstream minor; minor<<std::setprecision(15)<<minorUnit; if(!patchOrInsertValChild(axis,"c:minorUnit","minorUnit",minor.str(),true)) return false; }
        return true;
    });
}

bool patchScalingValChild(std::string& scaling, const char* prefixed, const char* local, const std::string& value) {
    const auto nodes = drawingTags(scaling, prefixed, local);
    const bool usePrefix = scaling.find("<c:scaling") != std::string::npos;
    const auto c = usePrefix ? "c:" : "";
    const auto generated = "<" + std::string(c) + local + " val=\"" + xmlEscape(value) + "\"/>";
    if (!nodes.empty()) {
        const auto pos = scaling.find(nodes.front());
        if (pos == std::string::npos) return false;
        scaling.replace(pos, nodes.front().size(), generated);
        return true;
    }
    const auto ext = drawingTags(scaling, "c:extLst", "extLst");
    const auto insertion = !ext.empty() ? scaling.find(ext.front()) : scaling.rfind("</");
    if (insertion == std::string::npos) return false;
    scaling.insert(insertion, generated);
    return true;
}

bool patchImportedAxisScaling(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartAxisScaling& scalingValue) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto nodes = drawingTags(axis, "c:scaling", "scaling");
        std::string scaling;
        if (!nodes.empty()) scaling = nodes.front();
        else {
            const bool prefixed = axis.find("<c:") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            scaling = "<" + std::string(c) + "scaling></" + std::string(c) + "scaling>";
        }
        if (scalingValue.hasLogBase) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.logBase;
            if (!patchScalingValChild(scaling, "c:logBase", "logBase", value.str())) return false;
        } else removeDrawingChild(scaling, "c:logBase", "logBase");
        if (!patchScalingValChild(scaling, "c:orientation", "orientation", scalingValue.reverseOrder ? "maxMin" : "minMax")) return false;
        if (scalingValue.hasMaximum) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.maximum;
            if (!patchScalingValChild(scaling, "c:max", "max", value.str())) return false;
        } else removeDrawingChild(scaling, "c:max", "max");
        if (scalingValue.hasMinimum) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.minimum;
            if (!patchScalingValChild(scaling, "c:min", "min", value.str())) return false;
        } else removeDrawingChild(scaling, "c:min", "min");
        if (!nodes.empty()) {
            const auto pos = axis.find(nodes.front());
            if (pos == std::string::npos) return false;
            axis.replace(pos, nodes.front().size(), scaling);
        } else {
            const auto ids = drawingTags(axis, "c:axId", "axId");
            if (ids.empty()) return false;
            const auto pos = axis.find(ids.front());
            if (pos == std::string::npos) return false;
            axis.insert(pos + ids.front().size(), scaling);
        }
        return true;
    });
}

bool patchImportedAxisCrossesAt(std::string& chartXmlText, std::uint64_t axisId, double crossesAt, bool clear) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        removeDrawingChild(axis, "c:crossesAt", "crossesAt");
        if (clear) return true;
        removeDrawingChild(axis, "c:crosses", "crosses");
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::ostringstream value; value << std::setprecision(15) << crossesAt;
        const auto generated = "<" + std::string(c) + "crossesAt val=\"" + value.str() + "\"/>";
        std::size_t insertion = axis.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
                 {"c:crossBetween", "crossBetween"}, {"c:majorUnit", "majorUnit"}, {"c:minorUnit", "minorUnit"}, {"c:dispUnits", "dispUnits"}}}) {
            const auto following = drawingTags(axis, pair.first, pair.second);
            if (!following.empty()) { const auto pos=axis.find(following.front()); if(pos!=std::string::npos) insertion=std::min(insertion,pos); }
        }
        const auto ext = drawingTags(axis, "c:extLst", "extLst");
        if (!ext.empty()) { const auto pos=axis.find(ext.front()); if(pos!=std::string::npos) insertion=std::min(insertion,pos); }
        if (insertion == std::string::npos) return false;
        axis.insert(insertion, generated);
        return true;
    });
}

bool patchImportedAxisDisplayUnits(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartDisplayUnits* units) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const auto existing = drawingTags(axis, "c:dispUnits", "dispUnits");
        if (!units) {
            if (!existing.empty()) { const auto pos=axis.find(existing.front()); if(pos==std::string::npos) return false; axis.erase(pos,existing.front().size()); }
            return true;
        }
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string generated = "<" + std::string(c) + "dispUnits>";
        if (units->hasCustomUnit) {
            std::ostringstream value; value << std::setprecision(15) << units->customUnit;
            generated += "<" + std::string(c) + "custUnit val=\"" + value.str() + "\"/>";
        } else generated += "<" + std::string(c) + "builtInUnit val=\"" + xmlEscape(units->builtInUnit) + "\"/>";
        if (units->showLabel) {
            generated += "<" + std::string(c) + "dispUnitsLbl>";
            if (units->labelRichText.present && !units->labelRichText.runs.empty()) generated += chartRichTextTxXml(units->labelRichText, prefixed, strict);
            generated += "</" + std::string(c) + "dispUnitsLbl>";
        }
        generated += "</" + std::string(c) + "dispUnits>";
        if (!existing.empty()) {
            const auto pos = axis.find(existing.front()); if(pos==std::string::npos) return false; axis.replace(pos,existing.front().size(),generated); return true;
        }
        const auto ext = drawingTags(axis, "c:extLst", "extLst");
        const auto insertion = !ext.empty() ? axis.find(ext.front()) : axis.rfind("</");
        if (insertion == std::string::npos) return false;
        axis.insert(insertion, generated);
        return true;
    });
}

bool patchImportedAxisCrossing(std::string& chartXmlText, std::uint64_t axisId, const std::string& crosses, const std::string& crossBetween) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        if(!crosses.empty()) removeDrawingChild(axis,"c:crossesAt","crossesAt");
        if(!crosses.empty() && !patchOrInsertValChild(axis,"c:crosses","crosses",crosses,true)) return false;
        if(!crossBetween.empty() && !patchOrInsertValChild(axis,"c:crossBetween","crossBetween",crossBetween,true)) return false;
        return true;
    });
}

bool patchImportedAxisLineFormat(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartLineFormat& format) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto spPr=axisDirectSpPr(axis); if(spPr.empty() && !ensureChartSpPr(axis,spPr,{})) return false; auto patched=spPr; if(!patchChartLineFormatInSpPr(patched,format)) return false; const auto pos=axis.find(spPr); if(pos==std::string::npos) return false; axis.replace(pos,spPr.size(),patched); return true;
    });
}

bool patchImportedAxisGridlineFormat(std::string& chartXmlText, std::uint64_t axisId, bool major, const xlpp::ChartLineFormat& format) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const char* pref=major?"c:majorGridlines":"c:minorGridlines"; const char* local=major?"majorGridlines":"minorGridlines";
        auto grids=drawingTags(axis,pref,local); std::string grid;
        if(!grids.empty()) grid=grids.front(); else { const bool prefixed=axis.find("<c:")!=std::string::npos; const auto c=prefixed?"c:":""; grid="<"+std::string(c)+local+"></"+std::string(c)+local+">"; }
        if(!patchNestedLineFormat(grid,format)) return false;
        if(!grids.empty()) { const auto pos=axis.find(grids.front()); if(pos==std::string::npos) return false; axis.replace(pos,grids.front().size(),grid); }
        else { const auto title=drawingTags(axis,"c:title","title"); std::size_t pos=!title.empty()?axis.find(title.front()):axis.rfind("</"); if(pos==std::string::npos) return false; axis.insert(pos,grid); }
        return true;
    });
}

bool removeImportedAxisGridlines(std::string& chartXmlText, std::uint64_t axisId, bool major) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        removeDrawingChild(axis, major ? "c:majorGridlines" : "c:minorGridlines", major ? "majorGridlines" : "minorGridlines");
        return true;
    });
}

std::string chartSpaceDirectSpPr(const std::string& chartXmlText) {
    const auto candidates = drawingTags(chartXmlText, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    const auto charts = drawingTags(chartXmlText, "c:chart", "chart");
    for (const auto& candidate : candidates) {
        if (std::none_of(charts.begin(), charts.end(), [&](const auto& chart) { return chart.find(candidate) != std::string::npos; })) return candidate;
    }
    return {};
}

std::string plotAreaDirectSpPr(const std::string& plotArea) {
    const auto candidates = drawingTags(plotArea, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    std::vector<std::string> nested;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 20>{{
             {"c:barChart","barChart"},{"c:lineChart","lineChart"},{"c:pieChart","pieChart"},{"c:scatterChart","scatterChart"},
             {"c:doughnutChart","doughnutChart"},{"c:radarChart","radarChart"},{"c:areaChart","areaChart"},{"c:bubbleChart","bubbleChart"},{"c:stockChart","stockChart"},
             {"c:bar3DChart","bar3DChart"},{"c:line3DChart","line3DChart"},{"c:area3DChart","area3DChart"},{"c:pie3DChart","pie3DChart"},
             {"c:surfaceChart","surfaceChart"},{"c:surface3DChart","surface3DChart"},
             {"c:catAx","catAx"},{"c:valAx","valAx"},{"c:dateAx","dateAx"},{"c:serAx","serAx"},{"c:dTable","dTable"}}}) {
        const auto nodes = drawingTags(plotArea, pair.first, pair.second); nested.insert(nested.end(), nodes.begin(), nodes.end());
    }
    for (const auto& candidate : candidates)
        if (std::none_of(nested.begin(), nested.end(), [&](const auto& owner){ return owner.find(candidate) != std::string::npos; })) return candidate;
    return {};
}

bool patchImportedAreaFormat(std::string& chartXmlText, bool chartArea, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    std::string owner;
    std::string originalOwner;
    if (chartArea) owner = chartXmlText;
    else {
        const auto plots = drawingTags(chartXmlText, "c:plotArea", "plotArea");
        if (plots.empty()) return false;
        originalOwner = plots.front(); owner = originalOwner;
    }
    auto spPr = chartArea ? chartSpaceDirectSpPr(owner) : plotAreaDirectSpPr(owner);
    if (spPr.empty()) {
        const bool prefixed = owner.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        spPr = "<" + std::string(c) + "spPr></" + std::string(c) + "spPr>";
        std::size_t insertion = std::string::npos;
        if (chartArea) {
            const auto charts = drawingTags(owner, "c:chart", "chart");
            if (charts.empty()) return false;
            const auto chartPos = owner.find(charts.front());
            if (chartPos == std::string::npos) return false;
            insertion = chartPos + charts.front().size();
        } else {
            const auto ext = drawingTags(owner, "c:extLst", "extLst");
            insertion = !ext.empty() ? owner.find(ext.back()) : owner.rfind("</");
        }
        if (insertion == std::string::npos) return false;
        owner.insert(insertion, spPr);
    }
    auto patched = spPr;
    if (line && !patchChartLineFormatInSpPr(patched, *line)) return false;
    if (fill && !patchChartFillFormatInSpPr(patched, *fill)) return false;
    const auto spPos = owner.find(spPr);
    if (spPos == std::string::npos) return false;
    owner.replace(spPos, spPr.size(), patched);
    if (chartArea) { chartXmlText = std::move(owner); return true; }
    const auto position = chartXmlText.find(originalOwner);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalOwner.size(), owner);
    return true;
}

std::string chartManualLayoutXml(const xlpp::ChartManualLayout& layout, bool prefixed) {
    const auto c=prefixed?"c:":""; std::ostringstream xml; xml<<"<"<<c<<"layout><"<<c<<"manualLayout>";
    const auto addText=[&](const char* tag,const std::string& value){ if(!value.empty()) xml<<"<"<<c<<tag<<" val=\""<<xmlEscape(value)<<"\"/>"; };
    const auto addNum=[&](const char* tag,bool has,double value){ if(has) xml<<"<"<<c<<tag<<" val=\""<<std::setprecision(15)<<value<<"\"/>"; };
    addText("layoutTarget",layout.target); addText("xMode",layout.xMode); addText("yMode",layout.yMode); addText("wMode",layout.widthMode); addText("hMode",layout.heightMode);
    addNum("x",layout.hasX,layout.x); addNum("y",layout.hasY,layout.y); addNum("w",layout.hasWidth,layout.width); addNum("h",layout.hasHeight,layout.height);
    xml<<"</"<<c<<"manualLayout></"<<c<<"layout>"; return xml.str();
}

bool patchManualLayoutOwner(std::string& owner, const xlpp::ChartManualLayout& layout) {
    if(!layout.present) return false; const bool prefixed=owner.find("<c:")!=std::string::npos; const auto generated=chartManualLayoutXml(layout,prefixed); const auto layouts=drawingTags(owner,"c:layout","layout");
    if(!layouts.empty()) { const auto pos=owner.find(layouts.front()); if(pos==std::string::npos) return false; owner.replace(pos,layouts.front().size(),generated); }
    else { const auto openEnd=owner.find('>'); if(openEnd==std::string::npos) return false; owner.insert(openEnd+1,generated); }
    return true;
}

bool patchImportedPlotAreaLayout(std::string& chartXmlText, const xlpp::ChartManualLayout& layout) {
    const auto plots=drawingTags(chartXmlText,"c:plotArea","plotArea"); if(plots.empty()) return false; auto plot=plots.front(); if(!patchManualLayoutOwner(plot,layout)) return false; const auto pos=chartXmlText.find(plots.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,plots.front().size(),plot); return true;
}

bool patchImportedLegendLayout(std::string& chartXmlText, const xlpp::ChartManualLayout& layout) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); if(!patchManualLayoutOwner(legend,layout)) return false; const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedLegendOverlay(std::string& chartXmlText, bool overlay) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); if(!patchOrInsertValChild(legend,"c:overlay","overlay",overlay?"1":"0",true)) return false; const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedLegendFormat(std::string& chartXmlText, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); auto spPrNodes=drawingTags(legend,"c:spPr","spPr"); std::string spPr=spPrNodes.empty()?std::string{}:spPrNodes.front(); if(spPr.empty() && !ensureChartSpPr(legend,spPr,{})) return false; auto patched=spPr; if(line && !patchChartLineFormatInSpPr(patched,*line)) return false; if(fill && !patchChartFillFormatInSpPr(patched,*fill)) return false; const auto spos=legend.find(spPr); if(spos==std::string::npos) return false; legend.replace(spos,spPr.size(),patched); const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedChartLegend(std::string& chartXmlText, bool show, const std::string& legendPosition) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const auto legends = drawingTags(chart, "c:legend", "legend");
    if (!show) {
        if (!legends.empty()) {
            const auto position = chart.find(legends.front());
            if (position == std::string::npos) return false;
            chart.erase(position, legends.front().size());
        }
    } else if (!legends.empty()) {
        auto legend = legends.front();
        const auto positionNodes = drawingTags(legend, "c:legendPos", "legendPos");
        if (!positionNodes.empty()) {
            auto node = positionNodes.front();
            const auto val = xlpp::internal::attribute(node, "val");
            const auto attr = std::string("val=\"") + val + "\"";
            const auto attrPosition = node.find(attr);
            if (attrPosition == std::string::npos) return false;
            node.replace(attrPosition, attr.size(), "val=\"" + xmlEscape(legendPosition) + "\"");
            const auto nodePosition = legend.find(positionNodes.front());
            if (nodePosition == std::string::npos) return false;
            legend.replace(nodePosition, positionNodes.front().size(), node);
        } else {
            const bool prefixed = legend.find("<c:legend") != std::string::npos;
            const auto openEnd = legend.find('>');
            if (openEnd == std::string::npos) return false;
            const auto c = prefixed ? "c:" : "";
            legend.insert(openEnd + 1, "<" + std::string(c) + "legendPos val=\"" + xmlEscape(legendPosition) + "\"/>");
        }
        const auto position = chart.find(legends.front());
        if (position == std::string::npos) return false;
        chart.replace(position, legends.front().size(), legend);
    } else {
        const bool prefixed = chart.find("<c:chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const std::string legend = "<" + std::string(c) + "legend><" + std::string(c) +
            "legendPos val=\"" + xmlEscape(legendPosition) + "\"/><" + std::string(c) +
            "layout/><" + std::string(c) + "overlay val=\"0\"/></" + std::string(c) + "legend>";
        const auto plotAreas = drawingTags(chart, "c:plotArea", "plotArea");
        if (plotAreas.empty()) return false;
        const auto plotPosition = chart.find(plotAreas.front());
        if (plotPosition == std::string::npos) return false;
        chart.insert(plotPosition + plotAreas.front().size(), legend);
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chart);
    return true;
}

bool patchValAttribute(std::string& node, const std::string& value) {
    const auto oldValue = xlpp::internal::attribute(node, "val");
    if (!oldValue.empty()) {
        const auto token = std::string("val=\"") + oldValue + "\"";
        const auto position = node.find(token);
        if (position == std::string::npos) return false;
        node.replace(position, token.size(), "val=\"" + xmlEscape(value) + "\"");
        return true;
    }
    const auto close = node.find("/>");
    const auto openEnd = node.find('>');
    const auto insertion = close != std::string::npos ? close : openEnd;
    if (insertion == std::string::npos) return false;
    node.insert(insertion, " val=\"" + xmlEscape(value) + "\"");
    return true;
}

bool patchOrInsertValChild(std::string& container, const char* prefixed, const char* local,
                           const std::string& value, bool insertWhenMissing = true) {
    const auto nodes = drawingTags(container, prefixed, local);
    if (!nodes.empty()) {
        auto patched = nodes.front();
        if (!patchValAttribute(patched, value)) return false;
        const auto position = container.find(nodes.front());
        if (position == std::string::npos) return false;
        container.replace(position, nodes.front().size(), patched);
        return true;
    }
    if (!insertWhenMissing) return true;
    const bool prefixedContainer = container.find("<c:") != std::string::npos;
    const auto c = prefixedContainer ? "c:" : "";
    const auto close = container.rfind("</");
    if (close == std::string::npos) return false;
    container.insert(close, "<" + std::string(c) + local + " val=\"" + xmlEscape(value) + "\"/>");
    return true;
}

std::string dataLabelsAggregateMask(const std::string& dLbls) {
    auto mask = dLbls;
    std::size_t cursor = 0;
    for (const auto& point : drawingTags(dLbls, "c:dLbl", "dLbl")) {
        const auto position = dLbls.find(point, cursor);
        if (position == std::string::npos) continue;
        std::fill(mask.begin() + static_cast<std::ptrdiff_t>(position),
                  mask.begin() + static_cast<std::ptrdiff_t>(position + point.size()), ' ');
        cursor = position + point.size();
    }
    return mask;
}

bool patchOrInsertAggregateDataLabelVal(std::string& dLbls, const char* prefixed, const char* local,
                                        const std::string& value, bool insertWhenMissing = true) {
    const auto mask = dataLabelsAggregateMask(dLbls);
    const auto nodes = drawingTags(mask, prefixed, local);
    if (!nodes.empty()) {
        const auto position = mask.find(nodes.front());
        if (position == std::string::npos || position + nodes.front().size() > dLbls.size()) return false;
        auto patched = dLbls.substr(position, nodes.front().size());
        if (!patchValAttribute(patched, value)) return false;
        dLbls.replace(position, nodes.front().size(), patched);
        return true;
    }
    if (!insertWhenMissing) return true;
    const bool prefixedContainer = dLbls.find("<c:dLbls") != std::string::npos;
    const auto c = prefixedContainer ? "c:" : "";
    const auto close = dLbls.rfind("</");
    if (close == std::string::npos) return false;
    dLbls.insert(close, "<" + std::string(c) + local + " val=\"" + xmlEscape(value) + "\"/>");
    return true;
}

bool patchOrInsertAggregateDataLabelText(std::string& dLbls, const char* prefixed, const char* local,
                                         const std::string& value) {
    const auto mask = dataLabelsAggregateMask(dLbls);
    const auto nodes = drawingTags(mask, prefixed, local);
    const bool prefixedContainer = dLbls.find("<c:dLbls") != std::string::npos;
    const auto c = prefixedContainer ? "c:" : "";
    const auto generated = "<" + std::string(c) + local + ">" + xmlEscape(value) + "</" + std::string(c) + local + ">";
    if (!nodes.empty()) {
        const auto position = mask.find(nodes.front());
        if (position == std::string::npos || position + nodes.front().size() > dLbls.size()) return false;
        dLbls.replace(position, nodes.front().size(), generated);
        return true;
    }
    const auto close = dLbls.rfind("</");
    if (close == std::string::npos) return false;
    dLbls.insert(close, generated);
    return true;
}

void removeDrawingChild(std::string& container, const char* prefixed, const char* local) {
    for (;;) {
        const auto nodes = drawingTags(container, prefixed, local);
        if (nodes.empty()) return;
        const auto position = container.find(nodes.front());
        if (position == std::string::npos) return;
        container.erase(position, nodes.front().size());
    }
}

bool patchOpeningTagAttribute(std::string& node, const std::string& name, const std::string& value, bool removeWhenEmpty = false) {
    const auto openEnd = node.find('>');
    if (openEnd == std::string::npos) return false;
    const auto key = name + "=\"";
    auto position = node.find(key);
    if (position != std::string::npos && position < openEnd) {
        const auto valueStart = position + key.size();
        const auto valueEnd = node.find('"', valueStart);
        if (valueEnd == std::string::npos || valueEnd > openEnd) return false;
        if (removeWhenEmpty && value.empty()) {
            auto eraseStart = position;
            if (eraseStart > 0 && std::isspace(static_cast<unsigned char>(node[eraseStart - 1]))) --eraseStart;
            node.erase(eraseStart, valueEnd + 1 - eraseStart);
        } else node.replace(valueStart, valueEnd - valueStart, xmlEscape(value));
        return true;
    }
    if (removeWhenEmpty && value.empty()) return true;
    const auto insertion = node.find("/>") < openEnd ? node.find("/>") : openEnd;
    node.insert(insertion, " " + name + "=\"" + xmlEscape(value) + "\"");
    return true;
}

const char* chartColorTransformTag(xlpp::ChartColorTransform::Kind kind) {
    using Kind = xlpp::ChartColorTransform::Kind;
    switch (kind) {
    case Kind::Alpha: return "alpha";
    case Kind::AlphaMod: return "alphaMod";
    case Kind::AlphaOff: return "alphaOff";
    case Kind::Tint: return "tint";
    case Kind::Shade: return "shade";
    case Kind::LumMod: return "lumMod";
    case Kind::LumOff: return "lumOff";
    case Kind::SatMod: return "satMod";
    case Kind::SatOff: return "satOff";
    }
    return "alpha";
}

std::string chartColorElement(const xlpp::ChartColor& color) {
    using Kind = xlpp::ChartColor::Kind;
    if (!color.present()) return {};
    const char* tag = "srgbClr";
    switch (color.kind) {
    case Kind::SRgb: tag = "srgbClr"; break;
    case Kind::Scheme: tag = "schemeClr"; break;
    case Kind::System: tag = "sysClr"; break;
    case Kind::Preset: tag = "prstClr"; break;
    case Kind::Unknown: tag = "srgbClr"; break;
    case Kind::None: return {};
    }
    if (color.transforms.empty())
        return "<a:" + std::string(tag) + " val=\"" + xmlEscape(color.value) + "\"/>";
    std::string xml = "<a:" + std::string(tag) + " val=\"" + xmlEscape(color.value) + "\">";
    for (const auto& transform : color.transforms)
        xml += "<a:" + std::string(chartColorTransformTag(transform.kind)) + " val=\"" + std::to_string(transform.value) + "\"/>";
    xml += "</a:" + std::string(tag) + ">";
    return xml;
}

std::string chartSolidFillXml(const xlpp::ChartColor& color, bool declareNamespace) {
    if (!color.present()) return {};
    return "<a:solidFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") +
           ">" + chartColorElement(color) + "</a:solidFill>";
}

std::string chartGradientFillXml(const xlpp::ChartFillFormat& format, bool declareNamespace = false) {
    if (format.gradientStops.empty()) return {};
    std::string xml = "<a:gradFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") + "><a:gsLst>";
    for (const auto& stop : format.gradientStops)
        if (stop.color.present())
            xml += "<a:gs pos=\"" + std::to_string(std::clamp(stop.position, 0, 100000)) + "\">" +
                   chartColorElement(stop.color) + "</a:gs>";
    xml += "</a:gsLst>";
    if (std::isfinite(format.gradientAngleDegrees) && std::abs(format.gradientAngleDegrees) > 1e-12) {
        const auto angle = static_cast<long long>(std::llround(format.gradientAngleDegrees * 60000.0));
        xml += "<a:lin ang=\"" + std::to_string(angle) + "\" scaled=\"1\"/>";
    }
    xml += "</a:gradFill>";
    return xml;
}

std::string chartPatternFillXml(const xlpp::ChartFillFormat& format, bool declareNamespace = false) {
    if (format.pattern.empty()) return {};
    std::string xml = "<a:pattFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") +
                      " prst=\"" + xmlEscape(format.pattern) + "\">";
    if (format.foregroundColor.present()) xml += "<a:fgClr>" + chartColorElement(format.foregroundColor) + "</a:fgClr>";
    if (format.backgroundColor.present()) xml += "<a:bgClr>" + chartColorElement(format.backgroundColor) + "</a:bgClr>";
    xml += "</a:pattFill>";
    return xml;
}

bool replaceFirstChild(std::string& container, const char* prefixed, const char* local, const std::string& replacement) {
    const auto nodes = drawingTags(container, prefixed, local);
    if (nodes.empty()) return false;
    const auto position = container.find(nodes.front());
    if (position == std::string::npos) return false;
    container.replace(position, nodes.front().size(), replacement);
    return true;
}

std::string directSpPrFillNode(const std::string& spPr, const char* prefixed, const char* local) {
    const auto lines = drawingTags(spPr, "a:ln", "ln");
    for (const auto& candidate : drawingTags(spPr, prefixed, local)) {
        if (std::none_of(lines.begin(), lines.end(), [&](const auto& line) { return line.find(candidate) != std::string::npos; }))
            return candidate;
    }
    return {};
}

bool patchChartLineFormatInSpPr(std::string& spPr, const xlpp::ChartLineFormat& format) {
    auto lines = drawingTags(spPr, "a:ln", "ln");
    std::string line;
    if (!lines.empty()) line = lines.front();
    else line = "<a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"></a:ln>";

    if (format.widthPoints > 0.0) {
        const auto widthEmu = static_cast<long long>(std::llround(format.widthPoints * 12700.0));
        if (!patchOpeningTagAttribute(line, "w", std::to_string(widthEmu))) return false;
    }
    if (!patchOpeningTagAttribute(line, "cap", format.cap, true)) return false;
    if (!patchOpeningTagAttribute(line, "cmpd", format.compound, true)) return false;

    removeDrawingChild(line, "a:noFill", "noFill");
    removeDrawingChild(line, "a:solidFill", "solidFill");
    if (format.noFill) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, "<a:noFill/>");
    } else if (format.color.present()) {
        const auto dashes = drawingTags(line, "a:prstDash", "prstDash");
        const auto customDashes = drawingTags(line, "a:custDash", "custDash");
        std::size_t insertion = line.rfind("</");
        if (!dashes.empty()) insertion = line.find(dashes.front());
        else if (!customDashes.empty()) insertion = line.find(customDashes.front());
        if (insertion == std::string::npos) return false;
        line.insert(insertion, chartSolidFillXml(format.color));
    }

    removeDrawingChild(line, "a:prstDash", "prstDash");
    removeDrawingChild(line, "a:custDash", "custDash");
    if (!format.customDash.empty()) {
        std::string custom = "<a:custDash>";
        for (const auto& stop : format.customDash) {
            const auto d = static_cast<long long>(std::llround(std::max(0.0, stop.dash) * 1000.0));
            const auto sp = static_cast<long long>(std::llround(std::max(0.0, stop.space) * 1000.0));
            custom += "<a:ds d=\"" + std::to_string(d) + "\" sp=\"" + std::to_string(sp) + "\"/>";
        }
        custom += "</a:custDash>";
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, custom);
    } else if (!format.dash.empty()) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, "<a:prstDash val=\"" + xmlEscape(format.dash) + "\"/>");
    }

    removeDrawingChild(line, "a:round", "round");
    removeDrawingChild(line, "a:bevel", "bevel");
    removeDrawingChild(line, "a:miter", "miter");
    if (!format.join.empty()) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        if (format.join == "round") line.insert(close, "<a:round/>");
        else if (format.join == "bevel") line.insert(close, "<a:bevel/>");
        else if (format.join == "miter") line.insert(close, "<a:miter/>");
        else return false;
    }

    if (!lines.empty()) {
        const auto position = spPr.find(lines.front());
        if (position == std::string::npos) return false;
        spPr.replace(position, lines.front().size(), line);
    } else {
        const auto close = spPr.rfind("</");
        if (close == std::string::npos) return false;
        spPr.insert(close, line);
    }
    return true;
}

bool patchChartFillFormatInSpPr(std::string& spPr, const xlpp::ChartFillFormat& format) {
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
             {"a:noFill", "noFill"}, {"a:solidFill", "solidFill"}, {"a:gradFill", "gradFill"}, {"a:pattFill", "pattFill"}}}) {
        const auto existing = directSpPrFillNode(spPr, pair.first, pair.second);
        if (!existing.empty()) {
            const auto position = spPr.find(existing);
            if (position != std::string::npos) spPr.erase(position, existing.size());
        }
    }
    std::string generated;
    const auto kind = format.noFill ? xlpp::ChartFillFormat::Kind::NoFill : format.kind;
    if (kind == xlpp::ChartFillFormat::Kind::NoFill)
        generated = "<a:noFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>";
    else if (kind == xlpp::ChartFillFormat::Kind::Gradient)
        generated = chartGradientFillXml(format, true);
    else if (kind == xlpp::ChartFillFormat::Kind::Pattern)
        generated = chartPatternFillXml(format, true);
    else if (format.color.present())
        generated = chartSolidFillXml(format.color, true);
    if (!generated.empty()) {
        const auto lines = drawingTags(spPr, "a:ln", "ln");
        const auto insertion = !lines.empty() ? spPr.find(lines.front()) : spPr.rfind("</");
        if (insertion == std::string::npos) return false;
        spPr.insert(insertion, generated);
    }
    return true;
}

bool ensureChartSpPr(std::string& owner, std::string& spPr, const std::string& beforeXml = {}) {
    if (!spPr.empty()) return true;
    const bool prefixed = owner.find("<c:") != std::string::npos;
    const auto c = prefixed ? "c:" : "";
    spPr = "<" + std::string(c) + "spPr></" + std::string(c) + "spPr>";
    std::size_t insertion = std::string::npos;
    if (!beforeXml.empty()) insertion = owner.find(beforeXml);
    if (insertion == std::string::npos) insertion = owner.rfind("</");
    if (insertion == std::string::npos) return false;
    owner.insert(insertion, spPr);
    return true;
}

bool patchSeriesLineOrFill(std::string& chartXmlText, std::size_t seriesIndex,
                           const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    auto spPr = seriesDirectSpPr(series);
    if (spPr.empty()) {
        std::string before;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 3>{{{"c:marker", "marker"}, {"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { before = nodes.front(); break; }
        }
        if (!ensureChartSpPr(series, spPr, before)) return false;
    }
    auto patched = spPr;
    if (line && !patchChartLineFormatInSpPr(patched, *line)) return false;
    if (fill && !patchChartFillFormatInSpPr(patched, *fill)) return false;
    const auto spPos = series.find(spPr);
    if (spPos == std::string::npos) return false;
    series.replace(spPos, spPr.size(), patched);
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchSeriesMarkerFormat(std::string& chartXmlText, std::size_t seriesIndex, const xlpp::ChartMarkerFormat& format) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    auto markers = drawingTags(series, "c:marker", "marker");
    std::string marker;
    if (!markers.empty()) marker = markers.front();
    else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        marker = "<" + std::string(c) + "marker></" + std::string(c) + "marker>";
    }
    if (!format.symbol.empty() && !patchOrInsertValChild(marker, "c:symbol", "symbol", format.symbol)) return false;
    if (format.size > 0 && !patchOrInsertValChild(marker, "c:size", "size", std::to_string(format.size))) return false;
    auto spPrNodes = drawingTags(marker, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if ((format.line.present || format.fill.present) && spPr.empty()) {
        if (!ensureChartSpPr(marker, spPr)) return false;
    }
    if (!spPr.empty()) {
        auto patched = spPr;
        if (format.line.present && !patchChartLineFormatInSpPr(patched, format.line)) return false;
        if (format.fill.present && !patchChartFillFormatInSpPr(patched, format.fill)) return false;
        const auto position = marker.find(spPr);
        if (position == std::string::npos) return false;
        marker.replace(position, spPr.size(), patched);
    }
    if (!markers.empty()) {
        const auto position = series.find(markers.front());
        if (position == std::string::npos) return false;
        series.replace(position, markers.front().size(), marker);
    } else {
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 3>{{{"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}, {"c:errBars", "errBars"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { insertion = series.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = series.rfind("</");
        if (insertion == std::string::npos) return false;
        series.insert(insertion, marker);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchNestedLineFormat(std::string& owner, const xlpp::ChartLineFormat& format) {
    auto spPrNodes = drawingTags(owner, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if (spPr.empty() && !ensureChartSpPr(owner, spPr)) return false;
    auto patched = spPr;
    if (!patchChartLineFormatInSpPr(patched, format)) return false;
    const auto position = owner.find(spPr);
    if (position == std::string::npos) return false;
    owner.replace(position, spPr.size(), patched);
    return true;
}

std::string trendlineTypeValue(xlpp::ChartSeries::TrendlineType type) {
    using T = xlpp::ChartSeries::TrendlineType;
    switch (type) {
    case T::Linear: return "linear";
    case T::Exponential: return "exp";
    case T::Logarithmic: return "log";
    case T::Polynomial: return "poly";
    case T::Power: return "power";
    case T::MovingAverage: return "movingAvg";
    }
    return "linear";
}

std::string errorBarDirectionValue(xlpp::ChartSeries::ErrorBarDirection direction) {
    return direction == xlpp::ChartSeries::ErrorBarDirection::X ? "x" : "y";
}

std::string errorBarTypeValue(xlpp::ChartSeries::ErrorBarType type) {
    using T = xlpp::ChartSeries::ErrorBarType;
    switch (type) {
    case T::Both: return "both";
    case T::Plus: return "plus";
    case T::Minus: return "minus";
    }
    return "both";
}

std::string errorValueTypeValue(xlpp::ChartSeries::ErrorValueType type) {
    using T = xlpp::ChartSeries::ErrorValueType;
    switch (type) {
    case T::FixedValue: return "fixedVal";
    case T::Percentage: return "percentage";
    case T::StandardDeviation: return "stdDev";
    case T::StandardError: return "stdErr";
    case T::Custom: return "cust";
    }
    return "fixedVal";
}

std::string formatChartDouble(double value) {
    std::ostringstream stream;
    stream << std::setprecision(15) << value;
    return stream.str();
}

bool patchTrendlineNode(std::string& trendlineXml, const xlpp::ChartSeries::Trendline& trendline) {
    if (!patchOrInsertValChild(trendlineXml, "c:trendlineType", "trendlineType", trendlineTypeValue(trendline.type))) return false;
    if (trendline.type == xlpp::ChartSeries::TrendlineType::Polynomial) {
        if (!patchOrInsertValChild(trendlineXml, "c:order", "order", std::to_string(trendline.order))) return false;
    } else removeDrawingChild(trendlineXml, "c:order", "order");
    if (trendline.type == xlpp::ChartSeries::TrendlineType::MovingAverage) {
        if (!patchOrInsertValChild(trendlineXml, "c:period", "period", std::to_string(trendline.period))) return false;
    } else removeDrawingChild(trendlineXml, "c:period", "period");
    if (trendline.forward > 0.0) {
        if (!patchOrInsertValChild(trendlineXml, "c:forward", "forward", formatChartDouble(trendline.forward))) return false;
    } else removeDrawingChild(trendlineXml, "c:forward", "forward");
    if (trendline.backward > 0.0) {
        if (!patchOrInsertValChild(trendlineXml, "c:backward", "backward", formatChartDouble(trendline.backward))) return false;
    } else removeDrawingChild(trendlineXml, "c:backward", "backward");
    if (!patchOrInsertValChild(trendlineXml, "c:dispRSqr", "dispRSqr", trendline.displayRSquared ? "1" : "0")) return false;
    if (!patchOrInsertValChild(trendlineXml, "c:dispEq", "dispEq", trendline.displayEquation ? "1" : "0")) return false;
    if (trendline.lineFormat.present && !patchNestedLineFormat(trendlineXml, trendline.lineFormat)) return false;
    return true;
}

std::string makeTrendlineXml(const xlpp::ChartSeries::Trendline& trendline, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::string xml = "<" + std::string(c) + "trendline><" + std::string(c) + "trendlineType val=\"" + trendlineTypeValue(trendline.type) + "\"/>";
    if (trendline.type == xlpp::ChartSeries::TrendlineType::Polynomial)
        xml += "<" + std::string(c) + "order val=\"" + std::to_string(trendline.order) + "\"/>";
    if (trendline.type == xlpp::ChartSeries::TrendlineType::MovingAverage)
        xml += "<" + std::string(c) + "period val=\"" + std::to_string(trendline.period) + "\"/>";
    if (trendline.forward > 0.0) xml += "<" + std::string(c) + "forward val=\"" + formatChartDouble(trendline.forward) + "\"/>";
    if (trendline.backward > 0.0) xml += "<" + std::string(c) + "backward val=\"" + formatChartDouble(trendline.backward) + "\"/>";
    xml += "<" + std::string(c) + "dispRSqr val=\"" + (trendline.displayRSquared ? "1" : "0") + "\"/>";
    xml += "<" + std::string(c) + "dispEq val=\"" + (trendline.displayEquation ? "1" : "0") + "\"/></" + std::string(c) + "trendline>";
    if (trendline.lineFormat.present && !patchNestedLineFormat(xml, trendline.lineFormat)) return {};
    return xml;
}

bool patchImportedChartSeriesTrendline(std::string& chartXmlText, std::size_t seriesIndex,
                                       std::size_t trendlineIndex,
                                       const xlpp::ChartSeries::Trendline* trendline,
                                       bool add) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    const auto nodes = drawingTags(series, "c:trendline", "trendline");
    if (add) {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto generated = makeTrendlineXml(*trendline, prefixed);
        if (generated.empty()) return false;
        std::size_t insertion = std::string::npos;
        const auto errBars = drawingTags(series, "c:errBars", "errBars");
        if (!errBars.empty()) insertion = series.find(errBars.front());
        if (insertion == std::string::npos) {
            for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{{"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
                const auto refs = drawingTags(series, pair.first, pair.second);
                if (!refs.empty()) { insertion = series.find(refs.front()); if (insertion != std::string::npos) break; }
            }
        }
        if (insertion == std::string::npos) insertion = series.rfind("</");
        if (insertion == std::string::npos) return false;
        series.insert(insertion, generated);
    } else {
        if (trendlineIndex >= nodes.size()) return false;
        const auto position = series.find(nodes[trendlineIndex]);
        if (position == std::string::npos) return false;
        if (!trendline) series.erase(position, nodes[trendlineIndex].size());
        else {
            auto patched = nodes[trendlineIndex];
            if (!patchTrendlineNode(patched, *trendline)) return false;
            series.replace(position, nodes[trendlineIndex].size(), patched);
        }
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchImportedChartSeriesTrendlineLineFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                                   std::size_t trendlineIndex, const xlpp::ChartLineFormat& format) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    const auto trendlines = drawingTags(series, "c:trendline", "trendline");
    if (trendlineIndex >= trendlines.size()) return false;
    auto trendline = trendlines[trendlineIndex];
    if (!patchNestedLineFormat(trendline, format)) return false;
    const auto trendPosition = series.find(trendlines[trendlineIndex]);
    if (trendPosition == std::string::npos) return false;
    series.replace(trendPosition, trendlines[trendlineIndex].size(), trendline);
    const auto position = chartXmlText.find(originalSeries);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series);
    return true;
}

bool patchImportedChartSeriesErrorBarsLineFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                                  xlpp::ChartSeries::ErrorBarDirection direction,
                                                  const xlpp::ChartLineFormat& format) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    const auto nodes = drawingTags(series, "c:errBars", "errBars");
    const auto found = std::find_if(nodes.begin(), nodes.end(), [&](const auto& node) {
        const auto dirs = drawingTags(node, "c:errDir", "errDir");
        const auto actual = dirs.empty() ? std::string("y") : xlpp::internal::attribute(dirs.front(), "val");
        return actual == errorBarDirectionValue(direction);
    });
    if (found == nodes.end()) return false;
    auto errorBars = *found;
    if (!patchNestedLineFormat(errorBars, format)) return false;
    const auto barsPosition = series.find(*found);
    if (barsPosition == std::string::npos) return false;
    series.replace(barsPosition, found->size(), errorBars);
    const auto position = chartXmlText.find(originalSeries);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series);
    return true;
}

bool patchErrorBarReference(std::string& errorBarsXml, const char* prefixed, const char* local, const std::string& reference) {
    if (reference.empty()) return false;
    const bool prefixedContainer = errorBarsXml.find("<c:errBars") != std::string::npos;
    const auto c = prefixedContainer ? "c:" : "";
    const auto generated = "<" + std::string(c) + local + "><" + std::string(c) + "numRef><" + std::string(c) + "f>" +
                           xmlEscape(reference) + "</" + std::string(c) + "f></" + std::string(c) + "numRef></" + std::string(c) + local + ">";
    const auto nodes = drawingTags(errorBarsXml, prefixed, local);
    if (!nodes.empty()) {
        auto node = nodes.front();
        const auto refs = drawingTags(node, "c:numRef", "numRef");
        if (!refs.empty()) {
            auto numRef = refs.front();
            const auto formulas = drawingTags(numRef, "c:f", "f");
            const auto generatedFormula = "<" + std::string(c) + "f>" + xmlEscape(reference) + "</" + std::string(c) + "f>";
            if (!formulas.empty()) {
                const auto pos = numRef.find(formulas.front());
                if (pos == std::string::npos) return false;
                numRef.replace(pos, formulas.front().size(), generatedFormula);
            } else {
                const auto close = numRef.rfind("</");
                if (close == std::string::npos) return false;
                numRef.insert(close, generatedFormula);
            }
            const auto refPos = node.find(refs.front());
            if (refPos == std::string::npos) return false;
            node.replace(refPos, refs.front().size(), numRef);
            const auto position = errorBarsXml.find(nodes.front());
            if (position == std::string::npos) return false;
            errorBarsXml.replace(position, nodes.front().size(), node);
            return true;
        }
        const auto position = errorBarsXml.find(nodes.front());
        if (position == std::string::npos) return false;
        errorBarsXml.replace(position, nodes.front().size(), generated);
        return true;
    }
    const auto spPr = drawingTags(errorBarsXml, "c:spPr", "spPr");
    const auto insertion = !spPr.empty() ? errorBarsXml.find(spPr.front()) : errorBarsXml.rfind("</");
    if (insertion == std::string::npos) return false;
    errorBarsXml.insert(insertion, generated);
    return true;
}

bool patchErrorBarsNode(std::string& errorBarsXml, const xlpp::ChartSeries::ErrorBars& errorBars) {
    if (!patchOrInsertValChild(errorBarsXml, "c:errDir", "errDir", errorBarDirectionValue(errorBars.direction))) return false;
    if (!patchOrInsertValChild(errorBarsXml, "c:errBarType", "errBarType", errorBarTypeValue(errorBars.barType))) return false;
    if (!patchOrInsertValChild(errorBarsXml, "c:errValType", "errValType", errorValueTypeValue(errorBars.valueType))) return false;
    if (!patchOrInsertValChild(errorBarsXml, "c:noEndCap", "noEndCap", errorBars.noEndCap ? "1" : "0")) return false;
    if (errorBars.valueType != xlpp::ChartSeries::ErrorValueType::Custom) {
        removeDrawingChild(errorBarsXml, "c:plus", "plus");
        removeDrawingChild(errorBarsXml, "c:minus", "minus");
        if (!patchOrInsertValChild(errorBarsXml, "c:val", "val", formatChartDouble(errorBars.value))) return false;
    } else {
        removeDrawingChild(errorBarsXml, "c:val", "val");
        if (!patchErrorBarReference(errorBarsXml, "c:minus", "minus", errorBars.minusReference)) return false;
        if (!patchErrorBarReference(errorBarsXml, "c:plus", "plus", errorBars.plusReference)) return false;
    }
    if (errorBars.lineFormat.present && !patchNestedLineFormat(errorBarsXml, errorBars.lineFormat)) return false;
    return true;
}

std::string makeErrorBarsXml(const xlpp::ChartSeries::ErrorBars& errorBars, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::string xml = "<" + std::string(c) + "errBars>";
    xml += "<" + std::string(c) + "errDir val=\"" + errorBarDirectionValue(errorBars.direction) + "\"/>";
    xml += "<" + std::string(c) + "errBarType val=\"" + errorBarTypeValue(errorBars.barType) + "\"/>";
    xml += "<" + std::string(c) + "errValType val=\"" + errorValueTypeValue(errorBars.valueType) + "\"/>";
    xml += "<" + std::string(c) + "noEndCap val=\"" + (errorBars.noEndCap ? "1" : "0") + "\"/>";
    if (errorBars.valueType != xlpp::ChartSeries::ErrorValueType::Custom)
        xml += "<" + std::string(c) + "val val=\"" + formatChartDouble(errorBars.value) + "\"/>";
    else {
        xml += "<" + std::string(c) + "minus><" + std::string(c) + "numRef><" + std::string(c) + "f>" +
               xmlEscape(errorBars.minusReference) + "</" + std::string(c) + "f></" + std::string(c) + "numRef></" + std::string(c) + "minus>";
        xml += "<" + std::string(c) + "plus><" + std::string(c) + "numRef><" + std::string(c) + "f>" +
               xmlEscape(errorBars.plusReference) + "</" + std::string(c) + "f></" + std::string(c) + "numRef></" + std::string(c) + "plus>";
    }
    xml += "</" + std::string(c) + "errBars>";
    if (errorBars.lineFormat.present && !patchNestedLineFormat(xml, errorBars.lineFormat)) return {};
    return xml;
}

bool patchImportedChartSeriesErrorBars(std::string& chartXmlText, std::size_t seriesIndex,
                                       xlpp::ChartSeries::ErrorBarDirection direction,
                                       const xlpp::ChartSeries::ErrorBars* errorBars) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    auto matchesDirection = [&](const std::string& node) {
        const auto dirs = drawingTags(node, "c:errDir", "errDir");
        if (dirs.empty()) return direction == xlpp::ChartSeries::ErrorBarDirection::Y;
        return xlpp::internal::attribute(dirs.front(), "val") == errorBarDirectionValue(direction);
    };
    const auto nodes = drawingTags(series, "c:errBars", "errBars");
    const auto found = std::find_if(nodes.begin(), nodes.end(), matchesDirection);
    if (!errorBars) {
        if (found == nodes.end()) return false;
        const auto position = series.find(*found);
        if (position == std::string::npos) return false;
        series.erase(position, found->size());
    } else if (found != nodes.end()) {
        auto patched = *found;
        if (!patchErrorBarsNode(patched, *errorBars)) return false;
        const auto position = series.find(*found);
        if (position == std::string::npos) return false;
        series.replace(position, found->size(), patched);
    } else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto generated = makeErrorBarsXml(*errorBars, prefixed);
        if (generated.empty()) return false;
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{{"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto refs = drawingTags(series, pair.first, pair.second);
            if (!refs.empty()) { insertion = series.find(refs.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = series.rfind("</");
        if (insertion == std::string::npos) return false;
        series.insert(insertion, generated);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

std::vector<std::pair<std::size_t, std::string>> patchablePlotNodesInOrder(const std::string& chartXmlText) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return {};
    const auto& plotArea = plotAreas.front();
    std::vector<std::pair<std::size_t, std::string>> result;
    const auto collect = [&](const char* prefixed, const char* local) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, prefixed, local)) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            result.emplace_back(position, node);
            cursor = position + node.size();
        }
    };
    collect("c:barChart", "barChart"); collect("c:lineChart", "lineChart"); collect("c:pieChart", "pieChart");
    collect("c:scatterChart", "scatterChart"); collect("c:doughnutChart", "doughnutChart"); collect("c:radarChart", "radarChart");
    collect("c:areaChart", "areaChart"); collect("c:bubbleChart", "bubbleChart"); collect("c:stockChart", "stockChart");
    collect("c:ofPieChart", "ofPieChart");
    collect("c:bar3DChart", "bar3DChart"); collect("c:line3DChart", "line3DChart"); collect("c:area3DChart", "area3DChart");
    collect("c:pie3DChart", "pie3DChart"); collect("c:surfaceChart", "surfaceChart"); collect("c:surface3DChart", "surface3DChart");
    std::sort(result.begin(), result.end(), [](const auto& a, const auto& b) { return a.first < b.first; });
    return result;
}

bool patchImportedChartSeriesDataLabels(std::string& chartXmlText, std::size_t seriesIndex, const xlpp::Chart::DataLabels& labels) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    const auto existingLabels = drawingTags(series, "c:dLbls", "dLbls");
    std::string dLbls;
    if (!existingLabels.empty()) dLbls = existingLabels.front();
    else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!labels.position.empty() && !patchOrInsertAggregateDataLabelVal(dLbls, "c:dLblPos", "dLblPos", labels.position)) return false;
    const auto patchFlag = [&](const char* prefixed, const char* local, bool value) {
        const auto mask = dataLabelsAggregateMask(dLbls);
        const bool exists = !drawingTags(mask, prefixed, local).empty();
        return patchOrInsertAggregateDataLabelVal(dLbls, prefixed, local, value ? "1" : "0", value || exists);
    };
    if (!patchFlag("c:showLegendKey", "showLegendKey", labels.showLegendKey) ||
        !patchFlag("c:showVal", "showVal", labels.showValue) ||
        !patchFlag("c:showCatName", "showCatName", labels.showCategoryName) ||
        !patchFlag("c:showSerName", "showSerName", labels.showSeriesName) ||
        !patchFlag("c:showPercent", "showPercent", labels.showPercent) ||
        !patchFlag("c:showBubbleSize", "showBubbleSize", labels.showBubbleSize) ||
        !patchFlag("c:showLeaderLines", "showLeaderLines", labels.showLeaderLines)) return false;
    if (!labels.separator.empty() &&
        !patchOrInsertAggregateDataLabelText(dLbls, "c:separator", "separator", labels.separator)) return false;
    if (!existingLabels.empty()) {
        const auto position = series.find(existingLabels.front());
        if (position == std::string::npos) return false;
        series.replace(position, existingLabels.front().size(), dLbls);
    } else {
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 6>{{
                 {"c:trendline", "trendline"}, {"c:errBars", "errBars"}, {"c:cat", "cat"},
                 {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { insertion = series.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = series.rfind("</");
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(originalSeries);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series);
    return true;
}

bool patchImportedChartPlotDataLabels(std::string& chartXmlText, std::size_t plotIndex, const xlpp::Chart::DataLabels& labels) {
    const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plotNodes.size()) return false;
    const auto originalPlot = plotNodes[plotIndex].second;
    auto plot = originalPlot;
    auto existingLabels = drawingTags(plot, "c:dLbls", "dLbls");
    const auto seriesNodes = drawingTags(plot, "c:ser", "ser");
    existingLabels.erase(std::remove_if(existingLabels.begin(), existingLabels.end(), [&](const std::string& node) {
        return std::any_of(seriesNodes.begin(), seriesNodes.end(), [&](const std::string& series) {
            return series.find(node) != std::string::npos;
        });
    }), existingLabels.end());
    std::string dLbls;
    if (!existingLabels.empty()) dLbls = existingLabels.front();
    else {
        const bool prefixed = plot.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!labels.position.empty() && !patchOrInsertAggregateDataLabelVal(dLbls, "c:dLblPos", "dLblPos", labels.position)) return false;
    const auto patchFlag = [&](const char* prefixed, const char* local, bool value) {
        const auto mask = dataLabelsAggregateMask(dLbls);
        const bool exists = !drawingTags(mask, prefixed, local).empty();
        return patchOrInsertAggregateDataLabelVal(dLbls, prefixed, local, value ? "1" : "0", value || exists);
    };
    if (!patchFlag("c:showLegendKey", "showLegendKey", labels.showLegendKey) ||
        !patchFlag("c:showVal", "showVal", labels.showValue) ||
        !patchFlag("c:showCatName", "showCatName", labels.showCategoryName) ||
        !patchFlag("c:showSerName", "showSerName", labels.showSeriesName) ||
        !patchFlag("c:showPercent", "showPercent", labels.showPercent) ||
        !patchFlag("c:showBubbleSize", "showBubbleSize", labels.showBubbleSize) ||
        !patchFlag("c:showLeaderLines", "showLeaderLines", labels.showLeaderLines)) return false;
    if (!labels.separator.empty() &&
        !patchOrInsertAggregateDataLabelText(dLbls, "c:separator", "separator", labels.separator)) return false;
    if (!existingLabels.empty()) {
        const auto position = plot.find(existingLabels.front());
        if (position == std::string::npos) return false;
        plot.replace(position, existingLabels.front().size(), dLbls);
    } else {
        const auto axisIds = drawingTags(plot, "c:axId", "axId");
        std::size_t insertion = !axisIds.empty() ? plot.find(axisIds.front()) : plot.rfind("</");
        if (insertion == std::string::npos) return false;
        plot.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(originalPlot);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlot.size(), plot);
    return true;
}

bool patchShapeOwnerFormat(std::string& owner, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    auto spPrNodes = drawingTags(owner, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if (spPr.empty() && !ensureChartSpPr(owner, spPr)) return false;
    auto patched = spPr;
    if (line && line->present && !patchChartLineFormatInSpPr(patched, *line)) return false;
    if (fill && fill->present && !patchChartFillFormatInSpPr(patched, *fill)) return false;
    const auto position = owner.find(spPr);
    if (position == std::string::npos) return false;
    owner.replace(position, spPr.size(), patched);
    return true;
}

std::string chartView3DXml(const xlpp::ChartView3D& view, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::ostringstream xml;
    xml << "<" << c << "view3D>";
    if (view.hasRotationX) xml << "<" << c << "rotX val=\"" << view.rotationX << "\"/>";
    if (view.hasHeightPercent) xml << "<" << c << "hPercent val=\"" << view.heightPercent << "\"/>";
    if (view.hasRotationY) xml << "<" << c << "rotY val=\"" << view.rotationY << "\"/>";
    if (view.hasDepthPercent) xml << "<" << c << "depthPercent val=\"" << view.depthPercent << "\"/>";
    if (view.hasRightAngleAxes) xml << "<" << c << "rAngAx val=\"" << (view.rightAngleAxes ? "1" : "0") << "\"/>";
    if (view.hasPerspective) xml << "<" << c << "perspective val=\"" << view.perspective << "\"/>";
    xml << "</" << c << "view3D>";
    return xml.str();
}

std::string generatedChartWallXml(const char* localName, const xlpp::ChartWallFormat& format, bool prefixed) {
    if (!format.present) return {};
    const auto c = prefixed ? "c:" : "";
    std::string wall = "<" + std::string(c) + localName + ">";
    if (format.hasThickness)
        wall += "<" + std::string(c) + "thickness val=\"" + std::to_string(format.thickness) + "\"/>";
    wall += "</" + std::string(c) + localName + ">";
    if ((format.line.present || format.fill.present) &&
        !patchShapeOwnerFormat(wall, format.line.present ? &format.line : nullptr, format.fill.present ? &format.fill : nullptr))
        return {};
    return wall;
}

bool patchImportedChartView3D(std::string& chartXmlText, const xlpp::ChartView3D& view) {
    if (!view.present) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const bool prefixed = chart.find("<c:") != std::string::npos;
    const auto generated = chartView3DXml(view, prefixed);
    const auto existing = drawingTags(chart, "c:view3D", "view3D");
    if (!existing.empty()) {
        const auto pos = chart.find(existing.front());
        if (pos == std::string::npos) return false;
        chart.replace(pos, existing.front().size(), generated);
    } else {
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
                 {"c:floor","floor"},{"c:sideWall","sideWall"},{"c:backWall","backWall"},{"c:plotArea","plotArea"}}}) {
            const auto nodes = drawingTags(chart, pair.first, pair.second);
            if (!nodes.empty()) { insertion = chart.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = chart.rfind("</");
        if (insertion == std::string::npos) return false;
        chart.insert(insertion, generated);
    }
    const auto pos = chartXmlText.find(originalChart);
    if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, originalChart.size(), chart);
    return true;
}

bool patchImportedChartWallFormat(std::string& chartXmlText, const char* prefixedName, const char* localName,
                                  const xlpp::ChartWallFormat& format) {
    if (!format.present) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const bool prefixed = chart.find("<c:") != std::string::npos;
    const auto c = prefixed ? "c:" : "";
    const auto existing = drawingTags(chart, prefixedName, localName);
    std::string wall = existing.empty()
        ? "<" + std::string(c) + localName + "></" + std::string(c) + localName + ">"
        : existing.front();
    if (format.hasThickness && !patchOrInsertValChild(wall, "c:thickness", "thickness", std::to_string(format.thickness), true)) return false;
    if ((format.line.present || format.fill.present) &&
        !patchShapeOwnerFormat(wall, format.line.present ? &format.line : nullptr, format.fill.present ? &format.fill : nullptr)) return false;
    if (!existing.empty()) {
        const auto pos = chart.find(existing.front());
        if (pos == std::string::npos) return false;
        chart.replace(pos, existing.front().size(), wall);
    } else {
        std::vector<std::pair<const char*, const char*>> later;
        const std::string name(localName);
        if (name == "floor") later = {{"c:sideWall","sideWall"},{"c:backWall","backWall"},{"c:plotArea","plotArea"}};
        else if (name == "sideWall") later = {{"c:backWall","backWall"},{"c:plotArea","plotArea"}};
        else later = {{"c:plotArea","plotArea"}};
        std::size_t insertion = std::string::npos;
        for (const auto& pair : later) {
            const auto nodes = drawingTags(chart, pair.first, pair.second);
            if (!nodes.empty()) { insertion = chart.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = chart.rfind("</");
        if (insertion == std::string::npos) return false;
        chart.insert(insertion, wall);
    }
    const auto pos = chartXmlText.find(originalChart);
    if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, originalChart.size(), chart);
    return true;
}

std::string chartTextStyleTxPrXml(const xlpp::ChartTextStyle& style, bool prefixed, bool strict) {
    if (!style.present) return {};
    const auto c = prefixed ? "c:" : "";
    std::string rPr = "<a:defRPr";
    if (style.bold) rPr += " b=\"1\"";
    if (style.italic) rPr += " i=\"1\"";
    if (style.fontSizePoints > 0.0) {
        const auto size = static_cast<long long>(std::llround(style.fontSizePoints * 100.0));
        rPr += " sz=\"" + std::to_string(size) + "\"";
    }
    rPr += ">";
    if (style.color.present()) rPr += chartSolidFillXml(style.color, false);
    if (!style.typeface.empty()) rPr += "<a:latin typeface=\"" + xmlEscape(style.typeface) + "\"/>";
    rPr += "</a:defRPr>";
    const auto drawingNs = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                  : "http://schemas.openxmlformats.org/drawingml/2006/main";
    return "<" + std::string(c) + "txPr xmlns:a=\"" + std::string(drawingNs) + "\"><a:bodyPr/><a:lstStyle/><a:p><a:pPr>" + rPr +
           "</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></" + std::string(c) + "txPr>";
}

std::string generatedDataLabelsXml(const xlpp::ChartDataLabels& labels) {
    if (!labels.present && !labels.hasLeaderLines) return {};
    std::string xml = "<c:dLbls>";
    const auto flag=[&](const char* name,bool value){ xml += "<c:" + std::string(name) + " val=\"" + (value ? "1" : "0") + "\"/>"; };
    flag("showLegendKey", labels.showLegendKey); flag("showVal", labels.showValue); flag("showCatName", labels.showCategoryName);
    flag("showSerName", labels.showSeriesName); flag("showPercent", labels.showPercent); flag("showBubbleSize", labels.showBubbleSize);
    if (!labels.position.empty()) xml += "<c:dLblPos val=\"" + xmlEscape(labels.position) + "\"/>";
    if (!labels.separator.empty()) xml += "<c:separator>" + xmlEscape(labels.separator) + "</c:separator>";
    if (labels.hasLeaderLines) {
        flag("showLeaderLines", true);
        std::string leader = "<c:leaderLines></c:leaderLines>";
        if (labels.leaderLineFormat.present && !patchNestedLineFormat(leader, labels.leaderLineFormat)) return {};
        xml += leader;
    } else if (labels.showLeaderLines) flag("showLeaderLines", true);
    xml += "</c:dLbls>";
    return xml;
}

std::string generatedPlotAuxiliaryXml(const xlpp::Chart::Plot& plot, bool /*strict*/) {
    std::string xml;
    xml += generatedDataLabelsXml(plot.dataLabels);
    const auto lineObject=[&](const char* name, bool present, const xlpp::ChartLineFormat& format) {
        if (!present) return std::string{};
        std::string object = "<c:" + std::string(name) + "></c:" + std::string(name) + ">";
        if (format.present && !patchNestedLineFormat(object, format)) return std::string{};
        return object;
    };
    xml += lineObject("dropLines", plot.hasDropLines, plot.dropLinesFormat);
    xml += lineObject("hiLowLines", plot.hasHighLowLines, plot.highLowLinesFormat);
    if (plot.upDownBars.present) {
        std::string object = "<c:upDownBars><c:gapWidth val=\"" + std::to_string(plot.upDownBars.gapWidth) + "\"/>";
        const auto bar=[&](const char* name, const xlpp::ChartLineFormat& line, const xlpp::ChartFillFormat& fill) {
            std::string owner = "<c:" + std::string(name) + "></c:" + std::string(name) + ">";
            if ((line.present || fill.present) && !patchShapeOwnerFormat(owner, line.present ? &line : nullptr, fill.present ? &fill : nullptr)) return std::string{};
            return owner;
        };
        const auto up = bar("upBars", plot.upDownBars.upLine, plot.upDownBars.upFill);
        const auto down = bar("downBars", plot.upDownBars.downLine, plot.upDownBars.downFill);
        if (up.empty() || down.empty()) return {};
        object += up + down + "</c:upDownBars>";
        xml += object;
    }
    return xml;
}

std::string generatedDataTableXml(const xlpp::ChartDataTable& table, bool strict) {
    if (!table.present) return {};
    std::string xml = "<c:dTable>";
    const auto flag=[&](const char* name,bool value){ xml += "<c:" + std::string(name) + " val=\"" + (value ? "1" : "0") + "\"/>"; };
    flag("showHorzBorder", table.showHorizontalBorder); flag("showVertBorder", table.showVerticalBorder);
    flag("showOutline", table.showOutline); flag("showKeys", table.showLegendKeys);
    xml += "</c:dTable>";
    if ((table.line.present || table.fill.present) &&
        !patchShapeOwnerFormat(xml, table.line.present ? &table.line : nullptr, table.fill.present ? &table.fill : nullptr)) return {};
    if (table.textStyle.present) {
        const auto txPr = chartTextStyleTxPrXml(table.textStyle, true, strict);
        const auto insertion = xml.rfind("</c:dTable>");
        if (insertion == std::string::npos) return {};
        xml.insert(insertion, txPr);
    }
    return xml;
}

bool patchImportedChartDataTable(std::string& chartXmlText, const xlpp::ChartDataTable* table) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return false;
    const auto originalPlotArea = plotAreas.front();
    auto plotArea = originalPlotArea;
    auto tables = drawingTags(plotArea, "c:dTable", "dTable");
    if (!table) {
        if (!tables.empty()) {
            const auto position = plotArea.find(tables.front());
            if (position == std::string::npos) return false;
            plotArea.erase(position, tables.front().size());
        }
    } else {
        const bool prefixed = plotArea.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string generated = "<" + std::string(c) + "dTable>";
        const auto addFlag = [&](const char* name, bool value) {
            generated += "<" + std::string(c) + name + " val=\"" + (value ? "1" : "0") + "\"/>";
        };
        addFlag("showHorzBorder", table->showHorizontalBorder);
        addFlag("showVertBorder", table->showVerticalBorder);
        addFlag("showOutline", table->showOutline);
        addFlag("showKeys", table->showLegendKeys);
        generated += "</" + std::string(c) + "dTable>";
        if ((table->line.present || table->fill.present) &&
            !patchShapeOwnerFormat(generated, table->line.present ? &table->line : nullptr,
                                   table->fill.present ? &table->fill : nullptr)) return false;
        if (table->textStyle.present) {
            const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            const auto txPr = chartTextStyleTxPrXml(table->textStyle, prefixed, strict);
            const auto closing = generated.rfind(prefixed ? "</c:dTable>" : "</dTable>");
            if (closing == std::string::npos) return false;
            generated.insert(closing, txPr);
        }
        if (!tables.empty()) {
            const auto position = plotArea.find(tables.front());
            if (position == std::string::npos) return false;
            plotArea.replace(position, tables.front().size(), generated);
        } else {
            const auto directSpPr = plotAreaDirectSpPr(plotArea);
            std::size_t insertion = !directSpPr.empty() ? plotArea.find(directSpPr) : std::string::npos;
            if (insertion == std::string::npos) {
                const auto ext = drawingTags(plotArea, "c:extLst", "extLst");
                insertion = !ext.empty() ? plotArea.find(ext.front()) : plotArea.rfind("</");
            }
            if (insertion == std::string::npos) return false;
            plotArea.insert(insertion, generated);
        }
    }
    const auto position = chartXmlText.find(originalPlotArea);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlotArea.size(), plotArea);
    return true;
}

std::size_t plotAuxiliaryInsertion(const std::string& plot, const char* local) {
    std::vector<std::pair<const char*, const char*>> later;
    const std::string tag(local);
    if (tag == "dropLines") later = {{"c:hiLowLines","hiLowLines"},{"c:upDownBars","upDownBars"},{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    else if (tag == "hiLowLines") later = {{"c:upDownBars","upDownBars"},{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    else later = {{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    for (const auto& item : later) {
        const auto nodes = drawingTags(plot, item.first, item.second);
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position != std::string::npos) return position;
        }
    }
    return plot.rfind("</");
}

bool patchImportedChartPlotLineObject(std::string& chartXmlText, std::size_t plotIndex,
                                      const char* prefixed, const char* local,
                                      const xlpp::ChartLineFormat* format) {
    const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plotNodes.size()) return false;
    const auto originalPlot = plotNodes[plotIndex].second;
    auto plot = originalPlot;
    auto nodes = drawingTags(plot, prefixed, local);
    if (!format) {
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.erase(position, nodes.front().size());
        }
    } else {
        std::string object;
        if (!nodes.empty()) object = nodes.front();
        else {
            const bool hasPrefix = plot.find("<c:") != std::string::npos;
            const auto c = hasPrefix ? "c:" : "";
            object = "<" + std::string(c) + local + "></" + std::string(c) + local + ">";
        }
        if (format->present && !patchNestedLineFormat(object, *format)) return false;
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.replace(position, nodes.front().size(), object);
        } else {
            const auto insertion = plotAuxiliaryInsertion(plot, local);
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, object);
        }
    }
    const auto position = chartXmlText.find(originalPlot);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlot.size(), plot);
    return true;
}

bool patchImportedChartPlotUpDownBars(std::string& chartXmlText, std::size_t plotIndex,
                                      const xlpp::ChartUpDownBars* bars) {
    const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plotNodes.size()) return false;
    const auto originalPlot = plotNodes[plotIndex].second;
    auto plot = originalPlot;
    auto nodes = drawingTags(plot, "c:upDownBars", "upDownBars");
    if (!bars) {
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.erase(position, nodes.front().size());
        }
    } else {
        const bool prefixed = plot.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string object = "<" + std::string(c) + "upDownBars><" + std::string(c) + "gapWidth val=\"" +
                             std::to_string(bars->gapWidth) + "\"/>";
        auto makeBar = [&](const char* name, const xlpp::ChartLineFormat& line, const xlpp::ChartFillFormat& fill) {
            std::string bar = "<" + std::string(c) + name + "></" + std::string(c) + name + ">";
            if ((line.present || fill.present) && !patchShapeOwnerFormat(bar, line.present ? &line : nullptr, fill.present ? &fill : nullptr)) return std::string{};
            return bar;
        };
        const auto up = makeBar("upBars", bars->upLine, bars->upFill);
        const auto down = makeBar("downBars", bars->downLine, bars->downFill);
        if (up.empty() || down.empty()) return false;
        object += up + down + "</" + std::string(c) + "upDownBars>";
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.replace(position, nodes.front().size(), object);
        } else {
            const auto insertion = plotAuxiliaryInsertion(plot, "upDownBars");
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, object);
        }
    }
    const auto position = chartXmlText.find(originalPlot);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlot.size(), plot);
    return true;
}

bool patchLeaderLinesInDataLabels(std::string& dLbls, const xlpp::ChartLineFormat* format, bool remove) {
    auto leaderLines = drawingTags(dLbls, "c:leaderLines", "leaderLines");
    if (remove) {
        if (!leaderLines.empty()) {
            const auto position = dLbls.find(leaderLines.front());
            if (position == std::string::npos) return false;
            dLbls.erase(position, leaderLines.front().size());
        }
        const auto mask = dataLabelsAggregateMask(dLbls);
        const bool exists = !drawingTags(mask, "c:showLeaderLines", "showLeaderLines").empty();
        if (exists && !patchOrInsertAggregateDataLabelVal(dLbls, "c:showLeaderLines", "showLeaderLines", "0", true)) return false;
        return true;
    }
    if (!format || !format->present) return false;
    if (!patchOrInsertAggregateDataLabelVal(dLbls, "c:showLeaderLines", "showLeaderLines", "1", true)) return false;
    std::string leader;
    if (!leaderLines.empty()) leader = leaderLines.front();
    else {
        const bool prefixed = dLbls.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        leader = "<" + std::string(c) + "leaderLines></" + std::string(c) + "leaderLines>";
    }
    if (!patchNestedLineFormat(leader, *format)) return false;
    if (!leaderLines.empty()) {
        const auto position = dLbls.find(leaderLines.front());
        if (position == std::string::npos) return false;
        dLbls.replace(position, leaderLines.front().size(), leader);
    } else {
        const auto ext = drawingTags(dLbls, "c:extLst", "extLst");
        const auto insertion = !ext.empty() ? dLbls.find(ext.front()) : dLbls.rfind("</");
        if (insertion == std::string::npos) return false;
        dLbls.insert(insertion, leader);
    }
    return true;
}

bool patchImportedChartPlotSimpleValue(std::string& chartXmlText, std::size_t plotIndex,
                                       const char* prefixed, const char* local, const std::string& value,
                                       bool beforeSeries) {
    const auto plotNodes=patchablePlotNodesInOrder(chartXmlText); if(plotIndex>=plotNodes.size()) return false;
    const auto original=plotNodes[plotIndex].second; auto plot=original; auto nodes=drawingTags(plot,prefixed,local);
    const bool hasPrefix=plot.find("<c:")!=std::string::npos; const auto c=hasPrefix?"c:":"";
    const std::string generated="<"+std::string(c)+local+" val=\""+xmlEscape(value)+"\"/>";
    if(!nodes.empty()){ const auto pos=plot.find(nodes.front()); if(pos==std::string::npos) return false; plot.replace(pos,nodes.front().size(),generated); }
    else {
        std::size_t insertion=std::string::npos;
        if(beforeSeries){ const auto ser=drawingTags(plot,"c:ser","ser"); if(!ser.empty()) insertion=plot.find(ser.front()); }
        else { const auto ext=drawingTags(plot,"c:extLst","extLst"); insertion=!ext.empty()?plot.find(ext.front()):plot.rfind("</"); }
        if(insertion==std::string::npos) return false; plot.insert(insertion,generated);
    }
    const auto pos=chartXmlText.find(original); if(pos==std::string::npos) return false; chartXmlText.replace(pos,original.size(),plot); return true;
}

bool patchImportedChartProjectedPie(std::string& chartXmlText, std::size_t plotIndex, const xlpp::ChartProjectedPieOptions& options) {
    const auto plotNodes=patchablePlotNodesInOrder(chartXmlText); if(plotIndex>=plotNodes.size()) return false;
    const auto original=plotNodes[plotIndex].second; auto plot=original;
    if(plot.find("ofPieChart")==std::string::npos) return false;
    const bool hasPrefix=plot.find("<c:")!=std::string::npos; const auto c=hasPrefix?"c:":"";
    for(const auto& tag:std::array<std::pair<const char*,const char*>,7>{{
        {"c:ofPieType","ofPieType"},{"c:gapWidth","gapWidth"},{"c:splitType","splitType"},{"c:splitPos","splitPos"},
        {"c:custSplit","custSplit"},{"c:secondPieSize","secondPieSize"},{"c:serLines","serLines"}}}){
        const auto nodes=drawingTags(plot,tag.first,tag.second);
        if(!nodes.empty()){ const auto pos=plot.find(nodes.front()); if(pos==std::string::npos) return false; plot.erase(pos,nodes.front().size()); }
    }
    std::ostringstream block;
    block<<"<"<<c<<"ofPieType val=\""<<xmlEscape(options.ofPieType)<<"\"/>";
    block<<"<"<<c<<"gapWidth val=\""<<options.gapWidth<<"\"/><"<<c<<"splitType val=\""<<xmlEscape(options.splitType)<<"\"/>";
    if(options.hasSplitPosition) block<<"<"<<c<<"splitPos val=\""<<options.splitPosition<<"\"/>";
    if(!options.customSplitPoints.empty()){ block<<"<"<<c<<"custSplit>"; for(const auto point:options.customSplitPoints) block<<"<"<<c<<"secondPiePt val=\""<<point<<"\"/>"; block<<"</"<<c<<"custSplit>"; }
    block<<"<"<<c<<"secondPieSize val=\""<<options.secondPlotSize<<"\"/>";
    if(options.hasSeriesLines){ std::string lines="<"+std::string(c)+"serLines></"+std::string(c)+"serLines>"; if(options.seriesLinesFormat.present&&!patchNestedLineFormat(lines,options.seriesLinesFormat)) return false; block<<lines; }
    const auto ext=drawingTags(plot,"c:extLst","extLst"); const auto insertion=!ext.empty()?plot.find(ext.front()):plot.rfind("</"); if(insertion==std::string::npos) return false; plot.insert(insertion,block.str());
    const auto pos=chartXmlText.find(original); if(pos==std::string::npos) return false; chartXmlText.replace(pos,original.size(),plot); return true;
}

bool patchImportedChartLeaderLines(std::string& chartXmlText, bool plotLevel, std::size_t ownerIndex,
                                   const xlpp::ChartLineFormat* format, bool remove) {
    if (plotLevel) {
        const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
        if (ownerIndex >= plotNodes.size()) return false;
        const auto originalPlot = plotNodes[ownerIndex].second;
        auto plot = originalPlot;
        auto labels = drawingTags(plot, "c:dLbls", "dLbls");
        const auto seriesNodes = drawingTags(plot, "c:ser", "ser");
        labels.erase(std::remove_if(labels.begin(), labels.end(), [&](const std::string& node) {
            return std::any_of(seriesNodes.begin(), seriesNodes.end(), [&](const std::string& series){ return series.find(node) != std::string::npos; });
        }), labels.end());
        if (labels.empty()) {
            if (remove) return true;
            const bool prefixed = plot.find("<c:") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            std::string dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
            if (!patchLeaderLinesInDataLabels(dLbls, format, false)) return false;
            const auto axisIds = drawingTags(plot, "c:axId", "axId");
            const auto insertion = !axisIds.empty() ? plot.find(axisIds.front()) : plot.rfind("</");
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, dLbls);
        } else {
            auto dLbls = labels.front();
            if (!patchLeaderLinesInDataLabels(dLbls, format, remove)) return false;
            const auto position = plot.find(labels.front()); if (position == std::string::npos) return false;
            plot.replace(position, labels.front().size(), dLbls);
        }
        const auto position = chartXmlText.find(originalPlot); if (position == std::string::npos) return false;
        chartXmlText.replace(position, originalPlot.size(), plot); return true;
    }
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (ownerIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[ownerIndex];
    auto series = originalSeries;
    auto labels = drawingTags(series, "c:dLbls", "dLbls");
    if (labels.empty()) {
        if (remove) return true;
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
        if (!patchLeaderLinesInDataLabels(dLbls, format, false)) return false;
        std::size_t insertion = series.rfind("</");
        for (const auto& item : std::array<std::pair<const char*, const char*>, 6>{{{"c:trendline","trendline"},{"c:errBars","errBars"},{"c:cat","cat"},{"c:xVal","xVal"},{"c:val","val"},{"c:yVal","yVal"}}}) {
            const auto nodes = drawingTags(series, item.first, item.second);
            if (!nodes.empty()) { const auto pos=series.find(nodes.front()); if(pos!=std::string::npos) { insertion=pos; break; } }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    } else {
        auto dLbls = labels.front();
        if (!patchLeaderLinesInDataLabels(dLbls, format, remove)) return false;
        const auto position = series.find(labels.front()); if(position==std::string::npos) return false;
        series.replace(position, labels.front().size(), dLbls);
    }
    const auto position = chartXmlText.find(originalSeries); if(position==std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series); return true;
}

std::string dataLabelPointXml(const xlpp::ChartDataLabelPoint& label, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::string xml = "<" + std::string(c) + "dLbl><" + std::string(c) + "idx val=\"" + std::to_string(label.index) + "\"/>";
    if (label.deleted) xml += "<" + std::string(c) + "delete val=\"1\"/>";
    if (!label.position.empty()) xml += "<" + std::string(c) + "dLblPos val=\"" + xmlEscape(label.position) + "\"/>";
    if (label.showLegendKey) xml += "<" + std::string(c) + "showLegendKey val=\"1\"/>";
    if (label.showValue) xml += "<" + std::string(c) + "showVal val=\"1\"/>";
    if (label.showCategoryName) xml += "<" + std::string(c) + "showCatName val=\"1\"/>";
    if (label.showSeriesName) xml += "<" + std::string(c) + "showSerName val=\"1\"/>";
    if (label.showPercent) xml += "<" + std::string(c) + "showPercent val=\"1\"/>";
    if (label.showBubbleSize) xml += "<" + std::string(c) + "showBubbleSize val=\"1\"/>";
    if (label.showLeaderLines) xml += "<" + std::string(c) + "showLeaderLines val=\"1\"/>";
    if (!label.separator.empty()) xml += "<" + std::string(c) + "separator>" + xmlEscape(label.separator) + "</" + std::string(c) + "separator>";
    xml += "</" + std::string(c) + "dLbl>";
    return xml;
}

bool patchDataLabelPointNode(std::string& dLbls, const xlpp::ChartDataLabelPoint& label, bool remove) {
    auto points = drawingTags(dLbls, "c:dLbl", "dLbl");
    const auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto idx = drawingTags(point, "c:idx", "idx");
        if (idx.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(idx.front(), "val")) == label.index; } catch (...) { return false; }
    });
    if (remove) {
        if (found == points.end()) return false;
        const auto position = dLbls.find(*found);
        if (position == std::string::npos) return false;
        dLbls.erase(position, found->size());
        return true;
    }
    if (found == points.end()) {
        const bool prefixed = dLbls.find("<c:dLbls") != std::string::npos;
        const auto generated = dataLabelPointXml(label, prefixed);
        std::size_t insertion = dLbls.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 8>{{
                 {"c:delete", "delete"}, {"c:dLblPos", "dLblPos"}, {"c:showLegendKey", "showLegendKey"}, {"c:showVal", "showVal"},
                 {"c:showCatName", "showCatName"}, {"c:showSerName", "showSerName"}, {"c:showPercent", "showPercent"}, {"c:showBubbleSize", "showBubbleSize"}}}) {
            const auto nodes = drawingTags(dLbls, pair.first, pair.second);
            for (const auto& node : nodes) {
                if (std::any_of(points.begin(), points.end(), [&](const auto& point) { return point.find(node) != std::string::npos; })) continue;
                const auto pos = dLbls.find(node);
                if (pos != std::string::npos) { insertion = std::min(insertion, pos); break; }
            }
        }
        if (insertion == std::string::npos) return false;
        dLbls.insert(insertion, generated);
        return true;
    }

    auto patched = *found;
    const auto patchFlag = [&](const char* prefixed, const char* local, bool value) {
        const bool exists = !drawingTags(patched, prefixed, local).empty();
        return patchOrInsertValChild(patched, prefixed, local, value ? "1" : "0", value || exists);
    };
    if (!patchFlag("c:delete", "delete", label.deleted) ||
        !patchFlag("c:showLegendKey", "showLegendKey", label.showLegendKey) ||
        !patchFlag("c:showVal", "showVal", label.showValue) ||
        !patchFlag("c:showCatName", "showCatName", label.showCategoryName) ||
        !patchFlag("c:showSerName", "showSerName", label.showSeriesName) ||
        !patchFlag("c:showPercent", "showPercent", label.showPercent) ||
        !patchFlag("c:showBubbleSize", "showBubbleSize", label.showBubbleSize) ||
        !patchFlag("c:showLeaderLines", "showLeaderLines", label.showLeaderLines)) return false;
    if (!label.position.empty() && !patchOrInsertValChild(patched, "c:dLblPos", "dLblPos", label.position)) return false;
    if (!label.separator.empty()) {
        const auto separators = drawingTags(patched, "c:separator", "separator");
        const bool prefixed = patched.find("<c:dLbl") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const auto generated = "<" + std::string(c) + "separator>" + xmlEscape(label.separator) + "</" + std::string(c) + "separator>";
        if (!separators.empty()) {
            const auto pos = patched.find(separators.front()); if (pos == std::string::npos) return false;
            patched.replace(pos, separators.front().size(), generated);
        } else {
            const auto close = patched.rfind("</"); if (close == std::string::npos) return false;
            patched.insert(close, generated);
        }
    }
    const auto position = dLbls.find(*found);
    if (position == std::string::npos) return false;
    dLbls.replace(position, found->size(), patched);
    return true;
}

std::string plotDirectDataLabels(const std::string& plot) {
    auto labels = drawingTags(plot, "c:dLbls", "dLbls");
    const auto series = drawingTags(plot, "c:ser", "ser");
    labels.erase(std::remove_if(labels.begin(), labels.end(), [&](const auto& node) {
        return std::any_of(series.begin(), series.end(), [&](const auto& seriesNode) { return seriesNode.find(node) != std::string::npos; });
    }), labels.end());
    return labels.empty() ? std::string{} : labels.front();
}

bool patchImportedChartSeriesDataLabelPoint(std::string& chartXmlText, std::size_t seriesIndex,
                                            const xlpp::ChartDataLabelPoint& label, bool remove) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    const auto labelNodes = drawingTags(series, "c:dLbls", "dLbls");
    std::string dLbls = labelNodes.empty() ? std::string{} : labelNodes.front();
    if (dLbls.empty()) {
        if (remove) return false;
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!patchDataLabelPointNode(dLbls, label, remove)) return false;
    if (!labelNodes.empty()) {
        const auto position = series.find(labelNodes.front());
        if (position == std::string::npos) return false;
        series.replace(position, labelNodes.front().size(), dLbls);
    } else {
        std::size_t insertion = series.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 6>{{
                 {"c:trendline", "trendline"}, {"c:errBars", "errBars"}, {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { const auto pos = series.find(nodes.front()); if (pos != std::string::npos) insertion = std::min(insertion, pos); }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchImportedChartSeriesDataLabelPointRichText(std::string& chartXmlText, std::size_t seriesIndex,
                                                    std::size_t pointIndex, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    auto labelNodes = drawingTags(series, "c:dLbls", "dLbls");
    std::string dLbls;
    if (labelNodes.empty()) {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
        xlpp::ChartDataLabelPoint point;
        point.index = pointIndex;
        if (!patchDataLabelPointNode(dLbls, point, false)) return false;
    } else dLbls = labelNodes.front();

    auto points = drawingTags(dLbls, "c:dLbl", "dLbl");
    auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto indices = drawingTags(point, "c:idx", "idx");
        if (indices.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == pointIndex; } catch (...) { return false; }
    });
    if (found == points.end()) {
        xlpp::ChartDataLabelPoint point;
        point.index = pointIndex;
        if (!patchDataLabelPointNode(dLbls, point, false)) return false;
        points = drawingTags(dLbls, "c:dLbl", "dLbl");
        found = std::find_if(points.begin(), points.end(), [&](const auto& pointXml) {
            const auto indices = drawingTags(pointXml, "c:idx", "idx");
            if (indices.empty()) return false;
            try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == pointIndex; } catch (...) { return false; }
        });
        if (found == points.end()) return false;
    }
    auto pointXml = *found;
    const bool prefixed = pointXml.find("<c:dLbl") != std::string::npos;
    const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
    const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
    const auto txNodes = drawingTags(pointXml, "c:tx", "tx");
    if (!txNodes.empty()) {
        const auto pos = pointXml.find(txNodes.front());
        if (pos == std::string::npos) return false;
        pointXml.replace(pos, txNodes.front().size(), txXml);
    } else {
        std::size_t insertion = pointXml.rfind("</");
        const auto positions = drawingTags(pointXml, "c:dLblPos", "dLblPos");
        if (!positions.empty()) insertion = pointXml.find(positions.front());
        if (insertion == std::string::npos) return false;
        pointXml.insert(insertion, txXml);
    }
    const auto pointPos = dLbls.find(*found);
    if (pointPos == std::string::npos) return false;
    dLbls.replace(pointPos, found->size(), pointXml);

    if (!labelNodes.empty()) {
        const auto pos = series.find(labelNodes.front());
        if (pos == std::string::npos) return false;
        series.replace(pos, labelNodes.front().size(), dLbls);
    } else {
        std::size_t insertion = series.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 6>{{
                 {"c:trendline", "trendline"}, {"c:errBars", "errBars"}, {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) {
                const auto pos = series.find(nodes.front());
                if (pos != std::string::npos) insertion = std::min(insertion, pos);
            }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    }
    const auto seriesPos = chartXmlText.find(originalSeries);
    if (seriesPos == std::string::npos) return false;
    chartXmlText.replace(seriesPos, originalSeries.size(), series);
    return true;
}

bool patchMarkerFormatInOwner(std::string& owner, const xlpp::ChartMarkerFormat& format) {
    auto markers = drawingTags(owner, "c:marker", "marker");
    std::string marker;
    if (!markers.empty()) marker = markers.front();
    else {
        const bool prefixed = owner.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        marker = "<" + std::string(c) + "marker></" + std::string(c) + "marker>";
    }
    if (!format.symbol.empty() && !patchOrInsertValChild(marker, "c:symbol", "symbol", format.symbol)) return false;
    if (format.size > 0 && !patchOrInsertValChild(marker, "c:size", "size", std::to_string(format.size))) return false;
    auto spPrNodes = drawingTags(marker, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if ((format.line.present || format.fill.present) && spPr.empty()) {
        if (!ensureChartSpPr(marker, spPr)) return false;
    }
    if (!spPr.empty()) {
        auto patched = spPr;
        if (format.line.present && !patchChartLineFormatInSpPr(patched, format.line)) return false;
        if (format.fill.present && !patchChartFillFormatInSpPr(patched, format.fill)) return false;
        const auto position = marker.find(spPr);
        if (position == std::string::npos) return false;
        marker.replace(position, spPr.size(), patched);
    }
    if (!markers.empty()) {
        const auto position = owner.find(markers.front());
        if (position == std::string::npos) return false;
        owner.replace(position, markers.front().size(), marker);
    } else {
        const auto close = owner.rfind("</");
        if (close == std::string::npos) return false;
        owner.insert(close, marker);
    }
    return true;
}

bool patchImportedChartSeriesDataPointFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                             const xlpp::ChartDataPointFormat& format, bool remove) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    auto points = drawingTags(series, "c:dPt", "dPt");
    auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto indices = drawingTags(point, "c:idx", "idx");
        if (indices.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == format.index; } catch (...) { return false; }
    });
    if (remove) {
        if (found == points.end()) return false;
        const auto position = series.find(*found);
        if (position == std::string::npos) return false;
        series.erase(position, found->size());
    } else {
        std::string pointXml;
        if (found != points.end()) pointXml = *found;
        else {
            const bool prefixed = series.find("<c:ser") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            pointXml = "<" + std::string(c) + "dPt><" + std::string(c) + "idx val=\"" +
                       std::to_string(format.index) + "\"/></" + std::string(c) + "dPt>";
        }
        auto spPrNodes = drawingTags(pointXml, "c:spPr", "spPr");
        std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
        if ((format.line.present || format.fill.present) && spPr.empty()) {
            std::string before;
            const auto markerNodes = drawingTags(pointXml, "c:marker", "marker");
            if (!markerNodes.empty()) before = markerNodes.front();
            if (!ensureChartSpPr(pointXml, spPr, before)) return false;
        }
        if (!spPr.empty()) {
            auto patched = spPr;
            if (format.line.present && !patchChartLineFormatInSpPr(patched, format.line)) return false;
            if (format.fill.present && !patchChartFillFormatInSpPr(patched, format.fill)) return false;
            const auto spPos = pointXml.find(spPr);
            if (spPos == std::string::npos) return false;
            pointXml.replace(spPos, spPr.size(), patched);
        }
        if (format.marker.present && !patchMarkerFormatInOwner(pointXml, format.marker)) return false;

        if (found != points.end()) {
            const auto position = series.find(*found);
            if (position == std::string::npos) return false;
            series.replace(position, found->size(), pointXml);
        } else {
            std::size_t insertion = series.rfind("</");
            for (const auto& pair : std::array<std::pair<const char*, const char*>, 7>{{
                     {"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}, {"c:errBars", "errBars"},
                     {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
                const auto nodes = drawingTags(series, pair.first, pair.second);
                if (!nodes.empty()) {
                    const auto pos = series.find(nodes.front());
                    if (pos != std::string::npos) insertion = std::min(insertion, pos);
                }
            }
            if (insertion == std::string::npos) return false;
            series.insert(insertion, pointXml);
        }
    }
    const auto seriesPosition = chartXmlText.find(originalSeries);
    if (seriesPosition == std::string::npos) return false;
    chartXmlText.replace(seriesPosition, originalSeries.size(), series);
    return true;
}

bool patchImportedChartPlotDataLabelPoint(std::string& chartXmlText, std::size_t plotIndex,
                                          const xlpp::ChartDataLabelPoint& label, bool remove) {
    const auto plots = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plots.size()) return false;
    const auto original = plots[plotIndex].second;
    auto plot = original;
    const auto existing = plotDirectDataLabels(plot);
    std::string dLbls = existing;
    if (dLbls.empty()) {
        if (remove) return false;
        const bool prefixed = plot.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!patchDataLabelPointNode(dLbls, label, remove)) return false;
    if (!existing.empty()) {
        const auto pos = plot.find(existing); if (pos == std::string::npos) return false;
        plot.replace(pos, existing.size(), dLbls);
    } else {
        const auto axisIds = drawingTags(plot, "c:axId", "axId");
        const auto insertion = !axisIds.empty() ? plot.find(axisIds.front()) : plot.rfind("</");
        if (insertion == std::string::npos) return false;
        plot.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), plot);
    return true;
}

bool patchImportedChartSeriesTitle(std::string& chartXmlText,
                                   std::size_t seriesIndex,
                                   const std::string& title) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    const auto txNodes = drawingTags(series, "c:tx", "tx");
    if (title.empty()) {
        if (!txNodes.empty()) {
            const auto position = series.find(txNodes.front());
            if (position == std::string::npos) return false;
            series.erase(position, txNodes.front().size());
        }
    } else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const std::string generated = "<" + std::string(c) + "tx><" + std::string(c) + "v>" +
            xmlEscape(title) + "</" + std::string(c) + "v></" + std::string(c) + "tx>";
        if (!txNodes.empty()) {
            const auto position = series.find(txNodes.front());
            if (position == std::string::npos) return false;
            series.replace(position, txNodes.front().size(), generated);
        } else {
            const auto idxNodes = drawingTags(series, "c:idx", "idx");
            const auto orderNodes = drawingTags(series, "c:order", "order");
            std::size_t insertion = std::string::npos;
            if (!orderNodes.empty()) {
                const auto pos = series.find(orderNodes.front());
                if (pos != std::string::npos) insertion = pos + orderNodes.front().size();
            } else if (!idxNodes.empty()) {
                const auto pos = series.find(idxNodes.front());
                if (pos != std::string::npos) insertion = pos + idxNodes.front().size();
            }
            if (insertion == std::string::npos) {
                const auto open = series.find('>');
                if (open == std::string::npos) return false;
                insertion = open + 1;
            }
            series.insert(insertion, generated);
        }
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchCacheInReferenceOwner(std::string& owner, const xlpp::ChartSeriesCache* cache) {
    auto refs = drawingTags(owner, "c:numRef", "numRef");
    bool numericRef = true;
    if (refs.empty()) { refs = drawingTags(owner, "c:strRef", "strRef"); numericRef = false; }
    if (refs.empty()) return cache == nullptr;
    const auto originalRef = refs.front();
    auto ref = originalRef;
    if (!cache) { eraseChartCacheBlocks(ref); }
    else {
        if (!cache->present || cache->numeric != numericRef) return false;
        const auto existingNum = drawingTags(ref, "c:numCache", "numCache");
        const auto existingStr = drawingTags(ref, "c:strCache", "strCache");
        const auto existing = !existingNum.empty() ? existingNum.front() : (!existingStr.empty() ? existingStr.front() : std::string{});
        const bool prefixed = ref.find("<c:") != std::string::npos;
        const auto generated = chartSeriesCacheXml(*cache, prefixed);
        if (!existing.empty()) {
            const auto pos = ref.find(existing); if (pos == std::string::npos) return false; ref.replace(pos, existing.size(), generated);
        } else {
            const auto close = ref.rfind("</"); if (close == std::string::npos) return false; ref.insert(close, generated);
        }
    }
    const auto pos = owner.find(originalRef);
    if (pos == std::string::npos) return false;
    owner.replace(pos, originalRef.size(), ref);
    return true;
}

bool patchImportedChartSeriesCache(std::string& chartXmlText, std::size_t seriesIndex, int kind, const xlpp::ChartSeriesCache* cache) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    if (kind == 3) eraseChartCacheBlocks(series);
    else {
        const char* prefixed = kind == 0 ? "c:cat" : (kind == 1 ? "c:val" : "c:tx");
        const char* local = kind == 0 ? "cat" : (kind == 1 ? "val" : "tx");
        auto owners = drawingTags(series, prefixed, local);
        if (owners.empty() && kind == 0) owners = drawingTags(series, "c:xVal", "xVal");
        if (owners.empty() && kind == 1) owners = drawingTags(series, "c:yVal", "yVal");
        if (owners.empty()) return false;
        auto owner = owners.front();
        if (!patchCacheInReferenceOwner(owner, cache)) return false;
        const auto ownerPos = series.find(owners.front()); if (ownerPos == std::string::npos) return false;
        series.replace(ownerPos, owners.front().size(), owner);
    }
    const auto pos = chartXmlText.find(original); if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, original.size(), series);
    return true;
}

bool patchImportedChartSeriesReferences(std::string& chartXmlText,
                                        std::size_t seriesIndex,
                                        const std::string& categoriesReference,
                                        const std::string& valuesReference) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto patched = original;
    bool categoriesOk = patchSeriesReferenceContainer(patched, "c:cat", "cat", categoriesReference);
    if (!categoriesOk) categoriesOk = patchSeriesReferenceContainer(patched, "c:xVal", "xVal", categoriesReference);
    bool valuesOk = patchSeriesReferenceContainer(patched, "c:val", "val", valuesReference);
    if (!valuesOk) valuesOk = patchSeriesReferenceContainer(patched, "c:yVal", "yVal", valuesReference);
    if (!categoriesOk || !valuesOk) return false;
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), patched);
    return true;
}

void suppressExclusivePartClosure(const std::string& rootPart,
                                  const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                  std::set<std::string>& suppressedPreservedParts) {
    std::unordered_set<std::string> closure;
    std::vector<std::string> stack{rootPart};
    while (!stack.empty()) {
        auto part = std::move(stack.back());
        stack.pop_back();
        if (part.empty() || !closure.insert(part).second) continue;
        for (const auto& relationship : allRelationships) {
            if (relationship.sourcePart != part || relationship.targetMode == "External") continue;
            const auto target = resolvePackagePart(part, relationship.target);
            if (!target.empty() && !closure.count(target)) stack.push_back(target);
        }
    }

    std::unordered_set<std::string> protectedParts;
    for (const auto& candidate : closure) {
        if (candidate == rootPart) continue;
        const bool externallyReferenced = std::any_of(allRelationships.begin(), allRelationships.end(), [&](const auto& relationship) {
            if (relationship.targetMode == "External" || closure.count(relationship.sourcePart)) return false;
            return resolvePackagePart(relationship.sourcePart, relationship.target) == candidate;
        });
        if (externallyReferenced) protectedParts.insert(candidate);
    }
    std::vector<std::string> protectStack(protectedParts.begin(), protectedParts.end());
    while (!protectStack.empty()) {
        auto part = std::move(protectStack.back());
        protectStack.pop_back();
        for (const auto& relationship : allRelationships) {
            if (relationship.sourcePart != part || relationship.targetMode == "External") continue;
            const auto target = resolvePackagePart(part, relationship.target);
            if (closure.count(target) && protectedParts.insert(target).second) protectStack.push_back(target);
        }
    }
    for (const auto& part : closure) {
        if (protectedParts.count(part)) continue;
        suppressedPreservedParts.insert(part);
        suppressedPreservedParts.insert(xlpp::internal::RelationshipGraph::relationshipsPartForSource(part));
    }
}

bool applyChartChangesToPreservedDrawing(xlpp::internal::ZipArchive& z,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextChartId,
                                         std::set<std::string>& suppressedPreservedParts) {
    const auto& edits = xlpp::internal::WorkbookDrawingAccess::chartEdits(sheet);
    if (sheet.appendedChartCount() == 0 && edits.empty()) return true;
    if (sourceSheetPart.empty() || sourceSheetXml.empty()) return false;

    const auto sheetRelationships = relationshipsForSource(allRelationships, sourceSheetPart);
    std::vector<std::string> ownedDrawingParts;
    std::unordered_set<std::string> seenDrawingParts;
    for (const auto& drawingNode : xlpp::internal::tags(sourceSheetXml, "drawing")) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = std::find_if(sheetRelationships.begin(), sheetRelationships.end(), [&](const auto& candidate) {
            return candidate.id == relationshipId && relationshipKind(candidate) == "drawing"
                && candidate.targetMode != "External";
        });
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sourceSheetPart, relationship->target);
        if (!drawingPart.empty() && seenDrawingParts.insert(drawingPart).second) ownedDrawingParts.push_back(drawingPart);
    }
    if (ownedDrawingParts.empty()) return false;
    for (const auto& edit : edits)
        if (!seenDrawingParts.count(edit.sourceDrawingPart)) return false;

    std::string appendDrawingPart = ownedDrawingParts.front();
    const auto& charts = static_cast<const xlpp::Worksheet&>(sheet).charts();
    for (std::size_t index = 0; index < std::min(sheet.loadedChartCount(), charts.size()); ++index) {
        if (seenDrawingParts.count(charts[index].sourceDrawingPart())) {
            appendDrawingPart = charts[index].sourceDrawingPart();
            break;
        }
    }
    if (sheet.loadedChartCount() == 0) {
        const auto& images = static_cast<const xlpp::Worksheet&>(sheet).images();
        for (std::size_t index = 0; index < std::min(sheet.loadedImageCount(), images.size()); ++index) {
            if (seenDrawingParts.count(images[index].sourceDrawingPart())) {
                appendDrawingPart = images[index].sourceDrawingPart();
                break;
            }
        }
    }

    for (const auto& drawingPart : ownedDrawingParts) {
        const bool appendHere = sheet.appendedChartCount() > 0 && drawingPart == appendDrawingPart;
        const bool hasEditsHere = std::any_of(edits.begin(), edits.end(), [&](const auto& edit) {
            return edit.sourceDrawingPart == drawingPart;
        });
        if (!appendHere && !hasEditsHere) continue;

        std::string drawingXmlText;
        if (z.contains(drawingPart)) drawingXmlText = z.get(drawingPart);
        else {
            const auto* raw = findPreservedPart(preservedParts, drawingPart);
            if (!raw) return false;
            drawingXmlText = raw->data;
        }
        const bool sourceStrict = drawingXmlText.find("http://purl.oclc.org/ooxml/drawingml/") != std::string::npos;
        const auto drawingRelsPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(drawingPart);
        auto drawingRelationships = z.contains(drawingRelsPart)
            ? xlpp::internal::RelationshipGraph::parseRelationshipsXml(drawingPart, z.get(drawingRelsPart))
            : relationshipsForSource(allRelationships, drawingPart);
        bool relationshipsChanged = false;

        std::unordered_map<std::string, std::string> chartWorkingCopies;
        for (const auto& edit : edits) {
            if (edit.sourceDrawingPart != drawingPart) continue;
            auto relationship = std::find_if(drawingRelationships.begin(), drawingRelationships.end(), [&](const auto& candidate) {
                return candidate.id == edit.sourceRelationshipId;
            });
            if (relationship == drawingRelationships.end()) return false;

            if (edit.removed) {
                if (!removeImportedChartAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId)) return false;
                if (!drawingReferencesChartRelationship(drawingXmlText, edit.sourceRelationshipId)) {
                    const bool shared = std::any_of(allRelationships.begin(), allRelationships.end(), [&](const auto& candidate) {
                        if (candidate.sourcePart == drawingPart && candidate.id == edit.sourceRelationshipId) return false;
                        return candidate.targetMode != "External" &&
                            resolvePackagePart(candidate.sourcePart, candidate.target) == edit.sourceChartPart;
                    });
                    drawingRelationships.erase(relationship);
                    relationshipsChanged = true;
                    if (!shared) suppressExclusivePartClosure(edit.sourceChartPart, allRelationships, suppressedPreservedParts);
                }
                continue;
            }

            if ((edit.moved || edit.resized) &&
                !patchImportedChartAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId,
                                          edit.anchor, edit.moved, edit.resized)) return false;

            const bool requiresChartXml = edit.titleChanged || edit.styleChanged || edit.titleRichTextChanged || edit.xAxisTitleChanged || edit.yAxisTitleChanged ||
                !edit.axisTitleEdits.empty() || !edit.axisRichTitleEdits.empty() || !edit.axisFormatEdits.empty() || !edit.areaFormatEdits.empty() || !edit.layoutEdits.empty() ||
                !edit.dataTableEdits.empty() || !edit.view3DEdits.empty() || !edit.wallFormatEdits.empty() || !edit.plotAuxiliaryEdits.empty() || !edit.plotTypeSpecificEdits.empty() || !edit.leaderLineEdits.empty() ||
                edit.legendChanged || edit.legendOverlayChanged || edit.legendLineFormatChanged || edit.legendFillFormatChanged ||
                !edit.seriesTitleEdits.empty() || !edit.seriesReferenceEdits.empty() || !edit.seriesCacheEdits.empty() ||
                !edit.plotDataLabelsEdits.empty() || !edit.seriesDataLabelsEdits.empty() || !edit.pointDataLabelEdits.empty() ||
                !edit.pointDataLabelRichTextEdits.empty() || !edit.dataPointFormatEdits.empty() ||
                !edit.seriesFormatEdits.empty() || !edit.trendlineFormatEdits.empty() || !edit.errorBarsFormatEdits.empty() ||
                !edit.trendlineEdits.empty() || !edit.errorBarsEdits.empty();
            if (!requiresChartXml) continue;
            auto chartIt = chartWorkingCopies.find(edit.sourceChartPart);
            if (chartIt == chartWorkingCopies.end()) {
                std::string chartXmlText;
                if (z.contains(edit.sourceChartPart)) chartXmlText = z.get(edit.sourceChartPart);
                else {
                    const auto* raw = findPreservedPart(preservedParts, edit.sourceChartPart);
                    if (!raw) return false;
                    chartXmlText = raw->data;
                }
                chartIt = chartWorkingCopies.emplace(edit.sourceChartPart, std::move(chartXmlText)).first;
            }
            if (edit.titleChanged && !patchImportedChartTitle(chartIt->second, edit.title)) return false;
            if (edit.styleChanged && !patchImportedChartStyle(chartIt->second, edit.style)) return false;
            if (edit.titleRichTextChanged && !patchImportedChartTitleRichText(chartIt->second, edit.titleRichText)) return false;
            const bool xyValueAxes = edit.chartType == xlpp::Chart::Type::Scatter || edit.chartType == xlpp::Chart::Type::Bubble;
            if (edit.xAxisTitleChanged) {
                if (edit.primaryXAxisId != 0) {
                    if (!patchImportedAxisTitleById(chartIt->second, edit.primaryXAxisId, edit.xAxisTitle)) return false;
                } else if (!patchImportedAxisTitle(chartIt->second,
                        xyValueAxes ? "c:valAx" : "c:catAx", xyValueAxes ? "valAx" : "catAx", 0, edit.xAxisTitle)) return false;
            }
            if (edit.yAxisTitleChanged) {
                if (edit.primaryYAxisId != 0) {
                    if (!patchImportedAxisTitleById(chartIt->second, edit.primaryYAxisId, edit.yAxisTitle)) return false;
                } else if (!patchImportedAxisTitle(chartIt->second, "c:valAx", "valAx",
                        xyValueAxes ? 1 : 0, edit.yAxisTitle)) return false;
            }
            for (const auto& axisEdit : edit.axisTitleEdits)
                if (!patchImportedAxisTitleById(chartIt->second, axisEdit.axisId, axisEdit.title)) return false;
            for (const auto& richEdit : edit.axisRichTitleEdits)
                if (!patchImportedAxisTitleRichTextById(chartIt->second, richEdit.axisId, richEdit.richText)) return false;
            for (const auto& axisEdit : edit.axisFormatEdits) {
                using Kind = std::decay_t<decltype(axisEdit.kind)>;
                if (axisEdit.kind == Kind::NumberFormat) { if (!patchImportedAxisNumberFormat(chartIt->second, axisEdit.axisId, axisEdit.value1, axisEdit.flag)) return false; }
                else if (axisEdit.kind == Kind::Ticks) { if (!patchImportedAxisTicks(chartIt->second, axisEdit.axisId, axisEdit.value1, axisEdit.value2, axisEdit.value3)) return false; }
                else if (axisEdit.kind == Kind::Units) { if (!patchImportedAxisUnits(chartIt->second, axisEdit.axisId, axisEdit.number1, axisEdit.number2)) return false; }
                else if (axisEdit.kind == Kind::Scaling) { if (!patchImportedAxisScaling(chartIt->second, axisEdit.axisId, axisEdit.scaling)) return false; }
                else if (axisEdit.kind == Kind::Crossing) { if (!patchImportedAxisCrossing(chartIt->second, axisEdit.axisId, axisEdit.value1, axisEdit.value2)) return false; }
                else if (axisEdit.kind == Kind::CrossesAt) { if (!patchImportedAxisCrossesAt(chartIt->second, axisEdit.axisId, axisEdit.number1, false)) return false; }
                else if (axisEdit.kind == Kind::ClearCrossesAt) { if (!patchImportedAxisCrossesAt(chartIt->second, axisEdit.axisId, 0.0, true)) return false; }
                else if (axisEdit.kind == Kind::DisplayUnits) { if (!patchImportedAxisDisplayUnits(chartIt->second, axisEdit.axisId, &axisEdit.displayUnits)) return false; }
                else if (axisEdit.kind == Kind::ClearDisplayUnits) { if (!patchImportedAxisDisplayUnits(chartIt->second, axisEdit.axisId, nullptr)) return false; }
                else if (axisEdit.kind == Kind::Line) { if (!patchImportedAxisLineFormat(chartIt->second, axisEdit.axisId, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::MajorGridline) { if (!patchImportedAxisGridlineFormat(chartIt->second, axisEdit.axisId, true, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::MinorGridline) { if (!patchImportedAxisGridlineFormat(chartIt->second, axisEdit.axisId, false, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::RemoveMajorGridline) { if (!removeImportedAxisGridlines(chartIt->second, axisEdit.axisId, true)) return false; }
                else if (!removeImportedAxisGridlines(chartIt->second, axisEdit.axisId, false)) return false;
            }
            for (const auto& areaEdit : edit.areaFormatEdits) {
                using Owner = std::decay_t<decltype(areaEdit.owner)>;
                using Kind = std::decay_t<decltype(areaEdit.kind)>;
                const bool chartArea = areaEdit.owner == Owner::ChartArea;
                if (areaEdit.kind == Kind::Line) { if (!patchImportedAreaFormat(chartIt->second, chartArea, &areaEdit.line, nullptr)) return false; }
                else if (!patchImportedAreaFormat(chartIt->second, chartArea, nullptr, &areaEdit.fill)) return false;
            }
            for (const auto& layoutEdit : edit.layoutEdits) {
                using Owner = std::decay_t<decltype(layoutEdit.owner)>;
                if (layoutEdit.owner == Owner::PlotArea) { if (!patchImportedPlotAreaLayout(chartIt->second, layoutEdit.layout)) return false; }
                else if (!patchImportedLegendLayout(chartIt->second, layoutEdit.layout)) return false;
            }
            for (const auto& tableEdit : edit.dataTableEdits)
                if (!patchImportedChartDataTable(chartIt->second, tableEdit.remove ? nullptr : &tableEdit.table)) return false;
            for (const auto& viewEdit : edit.view3DEdits)
                if (!patchImportedChartView3D(chartIt->second, viewEdit.view)) return false;
            for (const auto& wallEdit : edit.wallFormatEdits) {
                using Owner = std::decay_t<decltype(wallEdit.owner)>;
                if (wallEdit.owner == Owner::Floor) { if (!patchImportedChartWallFormat(chartIt->second, "c:floor", "floor", wallEdit.format)) return false; }
                else if (wallEdit.owner == Owner::SideWall) { if (!patchImportedChartWallFormat(chartIt->second, "c:sideWall", "sideWall", wallEdit.format)) return false; }
                else if (!patchImportedChartWallFormat(chartIt->second, "c:backWall", "backWall", wallEdit.format)) return false;
            }
            for (const auto& auxEdit : edit.plotAuxiliaryEdits) {
                using Kind = std::decay_t<decltype(auxEdit.kind)>;
                if (auxEdit.kind == Kind::DropLines) {
                    if (!patchImportedChartPlotLineObject(chartIt->second, auxEdit.plotIndex, "c:dropLines", "dropLines", auxEdit.remove ? nullptr : &auxEdit.line)) return false;
                } else if (auxEdit.kind == Kind::HighLowLines) {
                    if (!patchImportedChartPlotLineObject(chartIt->second, auxEdit.plotIndex, "c:hiLowLines", "hiLowLines", auxEdit.remove ? nullptr : &auxEdit.line)) return false;
                } else if (!patchImportedChartPlotUpDownBars(chartIt->second, auxEdit.plotIndex, auxEdit.remove ? nullptr : &auxEdit.upDownBars)) return false;
            }
            for (const auto& typeEdit : edit.plotTypeSpecificEdits) {
                using Kind = std::decay_t<decltype(typeEdit.kind)>;
                if (typeEdit.kind == Kind::FirstSliceAngle) { if (!patchImportedChartPlotSimpleValue(chartIt->second,typeEdit.plotIndex,"c:firstSliceAng","firstSliceAng",std::to_string(typeEdit.integerValue),false)) return false; }
                else if (typeEdit.kind == Kind::DoughnutHoleSize) { if (!patchImportedChartPlotSimpleValue(chartIt->second,typeEdit.plotIndex,"c:holeSize","holeSize",std::to_string(typeEdit.integerValue),false)) return false; }
                else if (typeEdit.kind == Kind::RadarStyle) { if (!patchImportedChartPlotSimpleValue(chartIt->second,typeEdit.plotIndex,"c:radarStyle","radarStyle",typeEdit.textValue,true)) return false; }
                else if (!patchImportedChartProjectedPie(chartIt->second,typeEdit.plotIndex,typeEdit.projectedPie)) return false;
            }
            for (const auto& leaderEdit : edit.leaderLineEdits)
                if (!patchImportedChartLeaderLines(chartIt->second, leaderEdit.plotLevel, leaderEdit.ownerIndex, leaderEdit.remove ? nullptr : &leaderEdit.line, leaderEdit.remove)) return false;
            if (edit.legendChanged && !patchImportedChartLegend(chartIt->second, edit.showLegend, edit.legendPosition)) return false;
            if (edit.legendOverlayChanged && !patchImportedLegendOverlay(chartIt->second, edit.legendOverlay)) return false;
            if (edit.legendLineFormatChanged && !patchImportedLegendFormat(chartIt->second, &edit.legendLineFormat, nullptr)) return false;
            if (edit.legendFillFormatChanged && !patchImportedLegendFormat(chartIt->second, nullptr, &edit.legendFillFormat)) return false;
            for (const auto& seriesEdit : edit.seriesTitleEdits)
                if (!patchImportedChartSeriesTitle(chartIt->second, seriesEdit.seriesIndex, seriesEdit.title)) return false;
            for (const auto& seriesEdit : edit.seriesReferenceEdits) {
                if (!patchImportedChartSeriesReferences(chartIt->second, seriesEdit.seriesIndex,
                                                        seriesEdit.categoriesReference,
                                                        seriesEdit.valuesReference)) return false;
            }
            for (const auto& cacheEdit : edit.seriesCacheEdits) {
                using Kind = std::decay_t<decltype(cacheEdit.kind)>;
                int kind = cacheEdit.kind == Kind::Categories ? 0 : (cacheEdit.kind == Kind::Values ? 1 : (cacheEdit.kind == Kind::Title ? 2 : 3));
                if (!patchImportedChartSeriesCache(chartIt->second, cacheEdit.seriesIndex, kind,
                                                   cacheEdit.kind == Kind::ClearAll ? nullptr : &cacheEdit.cache)) return false;
            }
            for (const auto& labelsEdit : edit.plotDataLabelsEdits)
                if (!patchImportedChartPlotDataLabels(chartIt->second, labelsEdit.plotIndex, labelsEdit.labels)) return false;
            for (const auto& labelsEdit : edit.seriesDataLabelsEdits)
                if (!patchImportedChartSeriesDataLabels(chartIt->second, labelsEdit.seriesIndex, labelsEdit.labels)) return false;
            for (const auto& pointEdit : edit.pointDataLabelEdits) {
                if (pointEdit.plotLevel) {
                    if (!patchImportedChartPlotDataLabelPoint(chartIt->second, pointEdit.ownerIndex, pointEdit.label, pointEdit.remove)) return false;
                } else if (!patchImportedChartSeriesDataLabelPoint(chartIt->second, pointEdit.ownerIndex, pointEdit.label, pointEdit.remove)) return false;
            }
            for (const auto& richEdit : edit.pointDataLabelRichTextEdits)
                if (!patchImportedChartSeriesDataLabelPointRichText(chartIt->second, richEdit.seriesIndex, richEdit.pointIndex, richEdit.richText)) return false;
            for (const auto& pointFormatEdit : edit.dataPointFormatEdits)
                if (!patchImportedChartSeriesDataPointFormat(chartIt->second, pointFormatEdit.seriesIndex, pointFormatEdit.format, pointFormatEdit.remove)) return false;
            for (const auto& formatEdit : edit.seriesFormatEdits) {
                using Kind = std::decay_t<decltype(formatEdit.kind)>;
                if (formatEdit.kind == Kind::Line) {
                    if (!patchSeriesLineOrFill(chartIt->second, formatEdit.seriesIndex, &formatEdit.line, nullptr)) return false;
                } else if (formatEdit.kind == Kind::Fill) {
                    if (!patchSeriesLineOrFill(chartIt->second, formatEdit.seriesIndex, nullptr, &formatEdit.fill)) return false;
                } else if (!patchSeriesMarkerFormat(chartIt->second, formatEdit.seriesIndex, formatEdit.marker)) return false;
            }
            for (const auto& trendlineEdit : edit.trendlineEdits) {
                using Action = std::decay_t<decltype(trendlineEdit.action)>;
                if (trendlineEdit.action == Action::Add) {
                    if (!patchImportedChartSeriesTrendline(chartIt->second, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                           &trendlineEdit.trendline, true)) return false;
                } else if (trendlineEdit.action == Action::Remove) {
                    if (!patchImportedChartSeriesTrendline(chartIt->second, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                           nullptr, false)) return false;
                } else if (!patchImportedChartSeriesTrendline(chartIt->second, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                               &trendlineEdit.trendline, false)) return false;
            }
            for (const auto& barsEdit : edit.errorBarsEdits)
                if (!patchImportedChartSeriesErrorBars(chartIt->second, barsEdit.seriesIndex, barsEdit.direction,
                                                       barsEdit.remove ? nullptr : &barsEdit.errorBars)) return false;
            for (const auto& formatEdit : edit.trendlineFormatEdits)
                if (!patchImportedChartSeriesTrendlineLineFormat(chartIt->second, formatEdit.seriesIndex, formatEdit.trendlineIndex, formatEdit.line)) return false;
            for (const auto& formatEdit : edit.errorBarsFormatEdits)
                if (!patchImportedChartSeriesErrorBarsLineFormat(chartIt->second, formatEdit.seriesIndex, formatEdit.direction, formatEdit.line)) return false;
        }

        if (appendHere) {
            std::set<std::string> usedRelationshipIds;
            for (const auto& relationship : drawingRelationships) usedRelationshipIds.insert(relationship.id);
            std::size_t objectId = maximumDrawingObjectId(drawingXmlText) + 1;
            std::string appendedAnchors;
            for (std::size_t index = sheet.loadedChartCount(); index < charts.size(); ++index) {
                const auto& chart = charts[index];
                const auto relationshipId = allocateRelationshipId(usedRelationshipIds);
                usedRelationshipIds.insert(relationshipId);
                const auto chartId = nextChartId++;
                xlpp::PreservedRelationship relationship;
                relationship.sourcePart = drawingPart;
                relationship.id = relationshipId;
                relationship.type = nsRelsDoc(sourceStrict) + "/chart";
                relationship.target = "../charts/chart" + std::to_string(chartId) + ".xml";
                drawingRelationships.push_back(std::move(relationship));
                relationshipsChanged = true;
                appendedAnchors += appendedChartAnchorXml(chart, relationshipId, objectId++, index, sourceStrict);
                z.add("xl/charts/chart" + std::to_string(chartId) + ".xml", chartXml(chart, sourceStrict));
            }
            if (!appendedAnchors.empty()) {
                const auto closing = drawingXmlText.rfind("</xdr:wsDr>") != std::string::npos
                    ? std::string("</xdr:wsDr>") : std::string("</wsDr>");
                const auto closingPosition = drawingXmlText.rfind(closing);
                if (closingPosition == std::string::npos) return false;
                drawingXmlText.insert(closingPosition, appendedAnchors);
            }
        }

        z.add(drawingPart, std::move(drawingXmlText));
        if (relationshipsChanged)
            z.add(drawingRelsPart, xlpp::internal::RelationshipGraph::serializeRelationships(drawingRelationships, sourceStrict));
        for (auto& [part, data] : chartWorkingCopies) z.add(part, std::move(data));
    }
    return true;
}

bool relationshipTargetsPart(const xlpp::PreservedRelationship& relationship, const std::string& part) {
    return relationship.targetMode != "External"
        && resolvePackagePart(relationship.sourcePart, relationship.target) == part;
}

bool hasOtherRelationshipToPart(const std::vector<xlpp::PreservedRelationship>& relationships,
                                const std::string& sourcePart,
                                const std::string& relationshipId,
                                const std::string& part) {
    return std::any_of(relationships.begin(), relationships.end(), [&](const auto& relationship) {
        if (relationship.sourcePart == sourcePart && relationship.id == relationshipId) return false;
        return relationshipTargetsPart(relationship, part);
    });
}

bool applyImageChangesToPreservedDrawing(xlpp::internal::ZipArchive& z,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextMediaId,
                                         std::set<std::string>& suppressedPreservedParts) {
    const auto& edits = xlpp::internal::WorkbookDrawingAccess::imageEdits(sheet);
    if (sheet.appendedImageCount() == 0 && edits.empty()) return true;
    if (sourceSheetPart.empty()) return false;

    if (sourceSheetXml.empty()) return false;
    const auto sheetRelationships = relationshipsForSource(allRelationships, sourceSheetPart);

    // A normal worksheet owns one DrawingML part, but preserving producer data
    // means we must not assume that.  Resolve every explicit <drawing r:id>
    // owner node and patch only the drawing part that owns each imported image.
    std::vector<std::string> ownedDrawingParts;
    std::unordered_set<std::string> seenDrawingParts;
    for (const auto& drawingNode : xlpp::internal::tags(sourceSheetXml, "drawing")) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = std::find_if(sheetRelationships.begin(), sheetRelationships.end(), [&](const auto& candidate) {
            return candidate.id == relationshipId && relationshipKind(candidate) == "drawing"
                && candidate.targetMode != "External";
        });
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sourceSheetPart, relationship->target);
        if (!drawingPart.empty() && seenDrawingParts.insert(drawingPart).second)
            ownedDrawingParts.push_back(drawingPart);
    }
    if (ownedDrawingParts.empty()) return false;

    for (const auto& edit : edits)
        if (!seenDrawingParts.count(edit.sourceDrawingPart)) return false;

    // New images are appended to the first preserved drawing by default.  If
    // the sheet already has imported images, prefer the drawing that owns the
    // first one so append behavior stays stable for multi-drawing producers.
    std::string appendDrawingPart = ownedDrawingParts.front();
    for (std::size_t index = 0; index < std::min(sheet.loadedImageCount(), sheet.images().size()); ++index) {
        const auto& candidate = sheet.images()[index];
        if (seenDrawingParts.count(candidate.sourceDrawingPart())) {
            appendDrawingPart = candidate.sourceDrawingPart();
            break;
        }
    }

    for (const auto& drawingPart : ownedDrawingParts) {
        const bool appendHere = sheet.appendedImageCount() > 0 && drawingPart == appendDrawingPart;
        const bool hasEditsHere = std::any_of(edits.begin(), edits.end(), [&](const auto& edit) {
            return edit.sourceDrawingPart == drawingPart;
        });
        if (!appendHere && !hasEditsHere) continue; // keep untouched drawing bytes exactly as loaded

        const auto* rawDrawingPart = findPreservedPart(preservedParts, drawingPart);
        if (!rawDrawingPart) return false;
        auto drawingXmlText = rawDrawingPart->data;
        const bool sourceStrict = drawingXmlText.find("http://purl.oclc.org/ooxml/drawingml/") != std::string::npos;
        auto drawingRelationships = relationshipsForSource(allRelationships, drawingPart);

        for (const auto& edit : edits) {
            if (edit.sourceDrawingPart != drawingPart) continue;
            if (!patchImportedImageAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId, edit.anchor,
                                          edit.moved, edit.resized, edit.removed)) return false;
            auto relationship = std::find_if(drawingRelationships.begin(), drawingRelationships.end(), [&](const auto& candidate) {
                return candidate.id == edit.sourceRelationshipId;
            });
            if (relationship == drawingRelationships.end()) return false;

            if (edit.removed) {
                if (!drawingReferencesRelationship(drawingXmlText, edit.sourceRelationshipId)) {
                    const bool shared = hasOtherRelationshipToPart(allRelationships, drawingPart,
                                                                   edit.sourceRelationshipId, edit.sourceMediaPart);
                    drawingRelationships.erase(relationship);
                    if (!shared) suppressedPreservedParts.insert(edit.sourceMediaPart);
                }
                continue;
            }
            if (edit.replaced) {
                const bool shared = hasOtherRelationshipToPart(allRelationships, drawingPart,
                                                               edit.sourceRelationshipId, edit.sourceMediaPart);
                const auto oldExtension = partExtension(edit.sourceMediaPart);
                if (!shared && oldExtension == edit.replacementExtension) {
                    const std::string bytes(reinterpret_cast<const char*>(edit.replacementBytes.data()), edit.replacementBytes.size());
                    z.add(edit.sourceMediaPart, bytes, false);
                } else {
                    const auto mediaId = nextMediaId++;
                    const auto newMediaPart = "xl/media/image" + std::to_string(mediaId) + "." + edit.replacementExtension;
                    relationship->target = "../media/image" + std::to_string(mediaId) + "." + edit.replacementExtension;
                    const std::string bytes(reinterpret_cast<const char*>(edit.replacementBytes.data()), edit.replacementBytes.size());
                    z.add(newMediaPart, bytes, false);
                    if (!shared) suppressedPreservedParts.insert(edit.sourceMediaPart);
                }
            }
        }

        if (appendHere) {
            std::set<std::string> usedRelationshipIds;
            for (const auto& relationship : drawingRelationships) usedRelationshipIds.insert(relationship.id);
            std::size_t objectId = maximumDrawingObjectId(drawingXmlText) + 1;
            std::string appendedAnchors;
            const auto& images = sheet.images();
            for (std::size_t index = sheet.loadedImageCount(); index < images.size(); ++index) {
                const auto& image = images[index];
                const auto relationshipId = allocateRelationshipId(usedRelationshipIds);
                usedRelationshipIds.insert(relationshipId);
                const auto mediaId = nextMediaId++;
                xlpp::PreservedRelationship relationship;
                relationship.sourcePart = drawingPart;
                relationship.id = relationshipId;
                relationship.type = nsRelsDoc(sourceStrict) + "/image";
                relationship.target = "../media/image" + std::to_string(mediaId) + "." + image.extension();
                drawingRelationships.push_back(std::move(relationship));
                appendedAnchors += appendedImageAnchorXml(image, relationshipId, objectId++, sourceStrict);
                const std::string bytes(reinterpret_cast<const char*>(image.bytes().data()), image.bytes().size());
                z.add("xl/media/image" + std::to_string(mediaId) + "." + image.extension(), bytes, false);
            }
            if (!appendedAnchors.empty()) {
                const auto closing = drawingXmlText.rfind("</xdr:wsDr>") != std::string::npos ? std::string("</xdr:wsDr>") : std::string("</wsDr>");
                const auto closingPosition = drawingXmlText.rfind(closing);
                if (closingPosition == std::string::npos) return false;
                drawingXmlText.insert(closingPosition, appendedAnchors);
            }
        }

        z.add(drawingPart, std::move(drawingXmlText));
        const auto slash = drawingPart.find_last_of('/');
        const auto drawingFile = drawingPart.substr(slash + 1);
        const auto drawingRelsPart = drawingPart.substr(0, slash + 1) + "_rels/" + drawingFile + ".rels";
        z.add(drawingRelsPart, xlpp::internal::RelationshipGraph::serializeRelationships(drawingRelationships, sourceStrict));
    }
    return true;
}

std::string mergePivotTablePartsBlocks(const std::vector<std::string>& originalBlocks,
                                      const std::vector<std::string>& generatedBlocks,
                                      bool preserveOriginal) {
    if (!preserveOriginal || originalBlocks.empty()) return joinBlocks(generatedBlocks);
    if (generatedBlocks.empty()) return joinBlocks(originalBlocks);

    std::vector<std::string> nodes;
    for (const auto& block : originalBlocks) {
        const auto items = extractTagBlocks(block, "pivotTablePart");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    for (const auto& block : generatedBlocks) {
        const auto items = extractTagBlocks(block, "pivotTablePart");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    if (nodes.empty()) return {};
    std::ostringstream xml;
    xml << "<pivotTableParts count=\"" << nodes.size() << "\">";
    for (const auto& node : nodes) xml << node;
    xml << "</pivotTableParts>";
    return xml.str();
}

std::string rebuildWorksheetTail(std::string generated,
                                 const std::string& original,
                                 bool preserveDrawing,
                                 bool preservePivot,
                                 bool preserveTables,
                                 bool preserveComments) {
    static const std::array<const char*, 10> tailTags{
        "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects",
        "controls", "webPublishItems", "tableParts", "pivotTableParts", "extLst"
    };
    std::unordered_map<std::string, std::vector<std::string>> generatedBlocks;
    std::unordered_map<std::string, std::vector<std::string>> originalBlocks;
    for (const auto* tag : tailTags) {
        generatedBlocks[tag] = extractTagBlocks(generated, tag);
        originalBlocks[tag] = extractTagBlocks(original, tag);
        eraseTagBlocks(generated, tag);
    }

    std::string tail;
    tail += joinBlocks(preserveDrawing && !originalBlocks["drawing"].empty()
        ? originalBlocks["drawing"] : generatedBlocks["drawing"]);
    tail += joinBlocks(preserveComments && !originalBlocks["legacyDrawing"].empty()
        ? originalBlocks["legacyDrawing"] : generatedBlocks["legacyDrawing"]);
    for (const auto* tag : {"legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems"})
        tail += joinBlocks(!originalBlocks[tag].empty() ? originalBlocks[tag] : generatedBlocks[tag]);
    tail += joinBlocks(preserveTables && !originalBlocks["tableParts"].empty()
        ? originalBlocks["tableParts"] : generatedBlocks["tableParts"]);
    tail += mergePivotTablePartsBlocks(originalBlocks["pivotTableParts"], generatedBlocks["pivotTableParts"], preservePivot);
    tail += joinBlocks(!originalBlocks["extLst"].empty() ? originalBlocks["extLst"] : generatedBlocks["extLst"]);
    insertBefore(generated, "</worksheet>", tail);
    return generated;
}

std::string preserveWorkbookNodes(std::string generated,
                                  const std::string& original,
                                  bool preservePivotCaches) {
    if (original.empty()) return generated;
    const auto preserveBeforeWorkbookPr = [&](const char* tag) {
        const auto blocks = extractTagBlocks(original, tag);
        if (!blocks.empty() && extractTagBlocks(generated, tag).empty())
            insertBefore(generated, "<workbookPr", joinBlocks(blocks));
    };
    preserveBeforeWorkbookPr("fileVersion");
    preserveBeforeWorkbookPr("fileSharing");

    if (extractTagBlocks(generated, "bookViews").empty())
        insertBefore(generated, "<sheets>", joinBlocks(extractTagBlocks(original, "bookViews")));

    const auto sheetsEnd = generated.find("</sheets>");
    if (sheetsEnd != std::string::npos) {
        std::string afterSheets;
        for (const auto* tag : {"functionGroups", "externalReferences"})
            if (extractTagBlocks(generated, tag).empty()) afterSheets += joinBlocks(extractTagBlocks(original, tag));
        generated.insert(sheetsEnd + std::string("</sheets>").size(), afterSheets);
    }

    std::string finalNodes;
    for (const auto* tag : {"oleSize", "customWorkbookViews"})
        if (extractTagBlocks(generated, tag).empty()) finalNodes += joinBlocks(extractTagBlocks(original, tag));
    if (preservePivotCaches && extractTagBlocks(generated, "pivotCaches").empty())
        finalNodes += joinBlocks(extractTagBlocks(original, "pivotCaches"));
    for (const auto* tag : {"webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"})
        if (extractTagBlocks(generated, tag).empty()) finalNodes += joinBlocks(extractTagBlocks(original, tag));
    insertBefore(generated, "</workbook>", finalNodes);
    return generated;
}

std::string mergeWorkbookPivotCaches(const std::string& originalWorkbookXml,
                                     const std::string& generatedPivotCachesXml) {
    const auto originalContainers = extractTagBlocks(originalWorkbookXml, "pivotCaches");
    if (originalContainers.empty()) return generatedPivotCachesXml;
    if (generatedPivotCachesXml.empty()) return joinBlocks(originalContainers);

    std::vector<std::string> nodes;
    for (const auto& block : originalContainers) {
        const auto items = extractTagBlocks(block, "pivotCache");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    const auto generatedContainers = extractTagBlocks(generatedPivotCachesXml, "pivotCaches");
    for (const auto& block : generatedContainers) {
        const auto items = extractTagBlocks(block, "pivotCache");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    if (nodes.empty()) return {};
    std::ostringstream xml;
    xml << "<pivotCaches>";
    for (const auto& node : nodes) xml << node;
    xml << "</pivotCaches>";
    return xml.str();
}

std::size_t nextAvailablePivotCacheId(const std::string& workbookXml) {
    std::size_t maximum = 0;
    for (const auto& node : extractTagBlocks(workbookXml, "pivotCache")) {
        const auto value = xlpp::internal::attribute(node, "cacheId");
        if (value.empty() || !std::all_of(value.begin(), value.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        maximum = std::max(maximum, static_cast<std::size_t>(std::stoull(value)));
    }
    return maximum + 1;
}

std::size_t nextAvailablePartId(const std::vector<xlpp::PreservedPart>& parts,
                                const std::string& prefix,
                                const std::string& suffix) {
    std::size_t maximum = 0;
    for (const auto& part : parts) {
        if (part.name.rfind(prefix, 0) != 0 || part.name.size() <= prefix.size() + suffix.size()) continue;
        if (part.name.compare(part.name.size() - suffix.size(), suffix.size(), suffix) != 0) continue;
        const auto number = part.name.substr(prefix.size(), part.name.size() - prefix.size() - suffix.size());
        if (number.empty() || !std::all_of(number.begin(), number.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        maximum = std::max(maximum, static_cast<std::size_t>(std::stoull(number)));
    }
    return maximum + 1;
}


std::size_t nextAvailableMediaId(const std::vector<xlpp::PreservedPart>& parts) {
    std::size_t maximum = 0;
    const std::string prefix = "xl/media/image";
    for (const auto& part : parts) {
        if (part.name.rfind(prefix, 0) != 0) continue;
        const auto dot = part.name.find('.', prefix.size());
        if (dot == std::string::npos) continue;
        const auto number = part.name.substr(prefix.size(), dot - prefix.size());
        if (number.empty() || !std::all_of(number.begin(), number.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        maximum = std::max(maximum, static_cast<std::size_t>(std::stoull(number)));
    }
    return maximum + 1;
}


std::vector<std::string> drawingTags(const std::string& xml, const char* prefixed, const char* local) {
    auto result = xlpp::internal::tags(xml, prefixed);
    if (std::string_view(prefixed) != std::string_view(local)) {
        auto unprefixed = xlpp::internal::tags(xml, local);
        result.insert(result.end(), std::make_move_iterator(unprefixed.begin()), std::make_move_iterator(unprefixed.end()));
    }
    // A preserved drawing can mix producer-native default-namespace elements
    // with xdr:-prefixed nodes appended by XL++. Preserve document order among
    // equivalent spellings when both forms occur in the same anchor family.
    std::stable_sort(result.begin(), result.end(), [&](const auto& lhs, const auto& rhs) {
        return xml.find(lhs) < xml.find(rhs);
    });
    return result;
}

std::string drawingTagText(const std::string& xml, const char* prefixed, const char* local) {
    auto value = xlpp::internal::tagText(xml, prefixed);
    if (value.empty()) value = xlpp::internal::tagText(xml, local);
    return value;
}

long long drawingInteger(const std::string& xml, const char* prefixed, const char* local, long long fallback = 0) {
    const auto value = drawingTagText(xml, prefixed, local);
    if (value.empty()) return fallback;
    try { return std::stoll(value); } catch (...) { return fallback; }
}

std::string partExtension(const std::string& part) {
    const auto slash = part.find_last_of('/');
    const auto dot = part.find_last_of('.');
    if (dot == std::string::npos || (slash != std::string::npos && dot < slash)) return {};
    auto extension = part.substr(dot + 1);
    std::transform(extension.begin(), extension.end(), extension.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    if (extension == "jpeg") extension = "jpg";
    return extension;
}

xlpp::DrawingMarker parseDrawingMarker(const std::string& markerXml) {
    xlpp::DrawingMarker marker;
    marker.column = static_cast<std::size_t>(std::max<long long>(0, drawingInteger(markerXml, "xdr:col", "col"))) + 1;
    marker.row = static_cast<std::size_t>(std::max<long long>(0, drawingInteger(markerXml, "xdr:row", "row"))) + 1;
    marker.columnOffsetEmu = drawingInteger(markerXml, "xdr:colOff", "colOff");
    marker.rowOffsetEmu = drawingInteger(markerXml, "xdr:rowOff", "rowOff");
    return marker;
}

void loadImages(xlpp::Worksheet& ws, const std::string& sheetXml, const xlpp::internal::ZipArchive& z,
                const std::string& sheetPart) {
    const auto drawings = xlpp::internal::tags(sheetXml, "drawing");
    if (drawings.empty()) return;
    const auto sheetSlash = sheetPart.find_last_of('/');
    const auto sheetFile = sheetPart.substr(sheetSlash + 1);
    const auto sheetRelsPart = sheetPart.substr(0, sheetSlash + 1) + "_rels/" + sheetFile + ".rels";
    if (!z.contains(sheetRelsPart)) return;

    std::unordered_map<std::string, std::string> sheetRelationships;
    for (const auto& rel : xlpp::internal::tags(z.get(sheetRelsPart), "Relationship"))
        if (xlpp::internal::attribute(rel, "Type").find("/drawing") != std::string::npos)
            sheetRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

    for (const auto& drawingNode : drawings) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = sheetRelationships.find(relationshipId);
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sheetPart, relationship->second);
        if (!z.contains(drawingPart)) continue;
        const auto drawingXmlText = z.get(drawingPart);
        const auto drawingSlash = drawingPart.find_last_of('/');
        const auto drawingFile = drawingPart.substr(drawingSlash + 1);
        const auto drawingRelsPart = drawingPart.substr(0, drawingSlash + 1) + "_rels/" + drawingFile + ".rels";
        if (!z.contains(drawingRelsPart)) continue;

        std::unordered_map<std::string, std::string> imageRelationships;
        for (const auto& rel : xlpp::internal::tags(z.get(drawingRelsPart), "Relationship"))
            if (xlpp::internal::attribute(rel, "Type").find("/image") != std::string::npos)
                imageRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

        const auto loadAnchors = [&](const char* prefixedName, const char* localName, xlpp::DrawingAnchorType type) {
            for (const auto& anchorNode : drawingTags(drawingXmlText, prefixedName, localName)) {
                const auto pictureNodes = drawingTags(anchorNode, "xdr:pic", "pic");
                if (pictureNodes.empty()) continue;
                const auto& pictureNode = pictureNodes.front();
                const auto blips = drawingTags(pictureNode, "a:blip", "blip");
                if (blips.empty()) continue;
                auto imageRelId = xlpp::internal::attribute(blips.front(), "r:embed");
                if (imageRelId.empty()) imageRelId = xlpp::internal::attribute(blips.front(), "r:link");
                const auto imageRelationship = imageRelationships.find(imageRelId);
                if (imageRelationship == imageRelationships.end()) continue;
                const auto mediaPart = resolvePackagePart(drawingPart, imageRelationship->second);
                if (!z.contains(mediaPart)) continue;
                const auto extension = partExtension(mediaPart);
                if (extension.empty()) continue;

                xlpp::DrawingAnchorInfo anchorInfo;
                anchorInfo.type = type;
                anchorInfo.editAs = xlpp::internal::attribute(anchorNode, "editAs");
                const auto fromNodes = drawingTags(anchorNode, "xdr:from", "from");
                if (!fromNodes.empty()) anchorInfo.from = parseDrawingMarker(fromNodes.front());
                const auto toNodes = drawingTags(anchorNode, "xdr:to", "to");
                if (!toNodes.empty()) anchorInfo.to = parseDrawingMarker(toNodes.front());
                const auto posNodes = drawingTags(anchorNode, "xdr:pos", "pos");
                if (!posNodes.empty()) {
                    anchorInfo.xEmu = drawingInteger(posNodes.front(), "xdr:x", "x");
                    anchorInfo.yEmu = drawingInteger(posNodes.front(), "xdr:y", "y");
                    const auto x = xlpp::internal::attribute(posNodes.front(), "x");
                    const auto y = xlpp::internal::attribute(posNodes.front(), "y");
                    if (!x.empty()) anchorInfo.xEmu = std::stoll(x);
                    if (!y.empty()) anchorInfo.yEmu = std::stoll(y);
                }
                const auto extNodes = drawingTags(anchorNode, "xdr:ext", "ext");
                if (!extNodes.empty()) {
                    const auto cx = xlpp::internal::attribute(extNodes.front(), "cx");
                    const auto cy = xlpp::internal::attribute(extNodes.front(), "cy");
                    if (!cx.empty()) anchorInfo.widthEmu = std::stoll(cx);
                    if (!cy.empty()) anchorInfo.heightEmu = std::stoll(cy);
                }
                // twoCellAnchor stores image extents in a:xfrm rather than an
                // anchor-level xdr:ext. Capture those values for inspection.
                if ((anchorInfo.widthEmu <= 0 || anchorInfo.heightEmu <= 0) && type == xlpp::DrawingAnchorType::TwoCell) {
                    const auto transformExt = drawingTags(pictureNode, "a:ext", "ext");
                    if (!transformExt.empty()) {
                        const auto cx = xlpp::internal::attribute(transformExt.front(), "cx");
                        const auto cy = xlpp::internal::attribute(transformExt.front(), "cy");
                        if (!cx.empty()) anchorInfo.widthEmu = std::stoll(cx);
                        if (!cy.empty()) anchorInfo.heightEmu = std::stoll(cy);
                    }
                }

                const auto nonVisual = drawingTags(pictureNode, "xdr:cNvPr", "cNvPr");
                std::string objectId;
                std::string objectName = "Image";
                if (!nonVisual.empty()) {
                    objectId = xlpp::internal::attribute(nonVisual.front(), "id");
                    const auto parsedName = xlpp::internal::attribute(nonVisual.front(), "name");
                    if (!parsedName.empty()) objectName = parsedName;
                }
                const auto anchorAddress = xlpp::CellReference{anchorInfo.from.row, anchorInfo.from.column}.address();
                const auto bytesText = z.get(mediaPart);
                xlpp::Image image(anchorAddress,
                    std::vector<unsigned char>(bytesText.begin(), bytesText.end()), extension);
                image.setName(objectName);
                image.setAnchorInfo(anchorInfo);
                image.setStableId(drawingPart + "#" + (objectId.empty() ? imageRelId : objectId));
                image.setSourceDrawingPart(drawingPart);
                image.setSourceMediaPart(mediaPart);
                image.setSourceRelationshipId(imageRelId);
                image.setImported(true);
                if (anchorInfo.widthEmu > 0) image.setWidthPixels(static_cast<double>(anchorInfo.widthEmu) / 9525.0);
                if (anchorInfo.heightEmu > 0) image.setHeightPixels(static_cast<double>(anchorInfo.heightEmu) / 9525.0);
                ws.addLoadedImage(std::move(image));
            }
        };

        loadAnchors("xdr:oneCellAnchor", "oneCellAnchor", xlpp::DrawingAnchorType::OneCell);
        loadAnchors("xdr:twoCellAnchor", "twoCellAnchor", xlpp::DrawingAnchorType::TwoCell);
        loadAnchors("xdr:absoluteAnchor", "absoluteAnchor", xlpp::DrawingAnchorType::Absolute);
    }
}


struct ParsedChartPlotNode {
    std::size_t position{0};
    xlpp::Chart::Type type{xlpp::Chart::Type::Bar};
    std::string xml;
};

std::vector<ParsedChartPlotNode> chartPlotNodesInOrder(const std::string& chartXmlText) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return {};
    const auto& plotArea = plotAreas.front();
    std::vector<ParsedChartPlotNode> result;
    const auto collect = [&](const char* prefixed, const char* local, xlpp::Chart::Type type) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, prefixed, local)) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            result.push_back({position, type, node});
            cursor = position + node.size();
        }
    };
    collect("c:barChart", "barChart", xlpp::Chart::Type::Bar);
    collect("c:lineChart", "lineChart", xlpp::Chart::Type::Line);
    collect("c:pieChart", "pieChart", xlpp::Chart::Type::Pie);
    collect("c:scatterChart", "scatterChart", xlpp::Chart::Type::Scatter);
    collect("c:doughnutChart", "doughnutChart", xlpp::Chart::Type::Doughnut);
    collect("c:radarChart", "radarChart", xlpp::Chart::Type::Radar);
    collect("c:areaChart", "areaChart", xlpp::Chart::Type::Area);
    collect("c:bubbleChart", "bubbleChart", xlpp::Chart::Type::Bubble);
    collect("c:stockChart", "stockChart", xlpp::Chart::Type::Stock);
    {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, "c:ofPieChart", "ofPieChart")) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            const auto kinds = drawingTags(node, "c:ofPieType", "ofPieType");
            const auto value = kinds.empty() ? std::string{"pie"} : xlpp::internal::attribute(kinds.front(), "val");
            result.push_back({position, value == "bar" ? xlpp::Chart::Type::BarOfPie : xlpp::Chart::Type::PieOfPie, node});
            cursor = position + node.size();
        }
    }
    collect("c:bar3DChart", "bar3DChart", xlpp::Chart::Type::Bar3D);
    collect("c:line3DChart", "line3DChart", xlpp::Chart::Type::Line3D);
    collect("c:area3DChart", "area3DChart", xlpp::Chart::Type::Area3D);
    collect("c:pie3DChart", "pie3DChart", xlpp::Chart::Type::Pie3D);
    collect("c:surfaceChart", "surfaceChart", xlpp::Chart::Type::Surface);
    collect("c:surface3DChart", "surface3DChart", xlpp::Chart::Type::Surface3D);
    std::sort(result.begin(), result.end(), [](const auto& a, const auto& b) { return a.position < b.position; });
    return result;
}

xlpp::Chart::Grouping parseChartGroupingNode(const std::string& chartNode) {
    const auto groupingNodes = drawingTags(chartNode, "c:grouping", "grouping");
    if (groupingNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto value = xlpp::internal::attribute(groupingNodes.front(), "val");
    if (value == "stacked") return xlpp::Chart::Grouping::Stacked;
    if (value == "percentStacked") return xlpp::Chart::Grouping::PercentStacked;
    if (value == "clustered") return xlpp::Chart::Grouping::Clustered;
    return xlpp::Chart::Grouping::Standard;
}

std::vector<std::uint64_t> chartAxisIds(const std::string& chartNode) {
    std::vector<std::uint64_t> result;
    for (const auto& node : drawingTags(chartNode, "c:axId", "axId")) {
        const auto value = xlpp::internal::attribute(node, "val");
        if (value.empty()) continue;
        try { result.push_back(std::stoull(value)); } catch (...) {}
    }
    return result;
}

bool chartBoolValue(const std::string& container, const char* prefixed, const char* local, bool fallback = false) {
    const auto nodes = drawingTags(container, prefixed, local);
    if (nodes.empty()) return fallback;
    const auto value = xlpp::internal::attribute(nodes.front(), "val");
    return value == "1" || value == "true" || value == "True";
}

xlpp::ChartColor parseChartColor(const std::string& container) {
    using Kind = xlpp::ChartColor::Kind;
    using TransformKind = xlpp::ChartColorTransform::Kind;
    for (const auto& entry : std::array<std::tuple<const char*, const char*, Kind>, 4>{
             std::tuple{"a:srgbClr", "srgbClr", Kind::SRgb},
             std::tuple{"a:schemeClr", "schemeClr", Kind::Scheme},
             std::tuple{"a:sysClr", "sysClr", Kind::System},
             std::tuple{"a:prstClr", "prstClr", Kind::Preset}}) {
        const auto nodes = drawingTags(container, std::get<0>(entry), std::get<1>(entry));
        if (nodes.empty()) continue;
        const auto value = xlpp::internal::attribute(nodes.front(), "val");
        if (value.empty()) continue;
        xlpp::ChartColor color;
        color.kind = std::get<2>(entry);
        color.value = value;
        const auto& colorXml = nodes.front();
        for (const auto& transform : std::array<std::tuple<const char*, const char*, TransformKind>, 9>{
                 std::tuple{"a:alpha", "alpha", TransformKind::Alpha},
                 std::tuple{"a:alphaMod", "alphaMod", TransformKind::AlphaMod},
                 std::tuple{"a:alphaOff", "alphaOff", TransformKind::AlphaOff},
                 std::tuple{"a:tint", "tint", TransformKind::Tint},
                 std::tuple{"a:shade", "shade", TransformKind::Shade},
                 std::tuple{"a:lumMod", "lumMod", TransformKind::LumMod},
                 std::tuple{"a:lumOff", "lumOff", TransformKind::LumOff},
                 std::tuple{"a:satMod", "satMod", TransformKind::SatMod},
                 std::tuple{"a:satOff", "satOff", TransformKind::SatOff}}) {
            for (const auto& node : drawingTags(colorXml, std::get<0>(transform), std::get<1>(transform))) {
                try {
                    color.transforms.push_back({std::get<2>(transform), std::stoi(xlpp::internal::attribute(node, "val"))});
                } catch (...) {}
            }
        }
        return color;
    }
    return {};
}


xlpp::ChartSeriesCache parseChartSeriesCache(const std::string& container) {
    xlpp::ChartSeriesCache result;
    auto caches = drawingTags(container, "c:numCache", "numCache");
    if (!caches.empty()) result.numeric = true;
    else { caches = drawingTags(container, "c:strCache", "strCache"); result.numeric = false; }
    if (caches.empty()) return result;
    result.present = true;
    const auto& cache = caches.front();
    result.formatCode = drawingTagText(cache, "c:formatCode", "formatCode");
    const auto counts = drawingTags(cache, "c:ptCount", "ptCount");
    if (!counts.empty()) { try { result.pointCount = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(counts.front(), "val"))); } catch (...) {} }
    for (const auto& point : drawingTags(cache, "c:pt", "pt")) {
        const auto idx = xlpp::internal::attribute(point, "idx");
        if (idx.empty()) continue;
        try { result.points.push_back({static_cast<std::size_t>(std::stoull(idx)), drawingTagText(point, "c:v", "v")}); } catch (...) {}
    }
    return result;
}

xlpp::ChartThemePalette parseChartThemePalette(const xlpp::internal::ZipArchive& z) {
    xlpp::ChartThemePalette palette;
    if (!z.contains("xl/theme/theme1.xml")) return palette;
    const auto xml = z.get("xl/theme/theme1.xml");
    for (const auto& name : std::array<const char*, 12>{"dk1","lt1","dk2","lt2","accent1","accent2","accent3","accent4","accent5","accent6","hlink","folHlink"}) {
        const auto prefixed = std::string("a:") + name;
        const auto nodes = drawingTags(xml, prefixed.c_str(), name);
        if (nodes.empty()) continue;
        std::string value;
        const auto srgb = drawingTags(nodes.front(), "a:srgbClr", "srgbClr");
        if (!srgb.empty()) value = xlpp::internal::attribute(srgb.front(), "val");
        if (value.empty()) {
            const auto sys = drawingTags(nodes.front(), "a:sysClr", "sysClr");
            if (!sys.empty()) { value = xlpp::internal::attribute(sys.front(), "lastClr"); if (value.empty()) value = xlpp::internal::attribute(sys.front(), "val"); }
        }
        if (!value.empty()) palette.colors.push_back({name, value});
    }
    const auto fontSchemes = drawingTags(xml, "a:fontScheme", "fontScheme");
    if (!fontSchemes.empty()) {
        palette.fontScheme.present = true;
        palette.fontScheme.name = xlpp::internal::attribute(fontSchemes.front(), "name");
        const auto major = drawingTags(fontSchemes.front(), "a:majorFont", "majorFont");
        if (!major.empty()) {
            const auto latin = drawingTags(major.front(), "a:latin", "latin");
            if (!latin.empty()) palette.fontScheme.majorLatinTypeface = xlpp::internal::attribute(latin.front(), "typeface");
        }
        const auto minor = drawingTags(fontSchemes.front(), "a:minorFont", "minorFont");
        if (!minor.empty()) {
            const auto latin = drawingTags(minor.front(), "a:latin", "latin");
            if (!latin.empty()) palette.fontScheme.minorLatinTypeface = xlpp::internal::attribute(latin.front(), "typeface");
        }
    }
    const auto fmtSchemes = drawingTags(xml, "a:fmtScheme", "fmtScheme");
    if (!fmtSchemes.empty()) {
        auto& effects = palette.effectScheme; effects.present = true; effects.name = xlpp::internal::attribute(fmtSchemes.front(), "name");
        const auto countFills = [&](const std::string& list) {
            std::size_t count = 0;
            for (const auto* tag : {"solidFill", "gradFill", "pattFill", "noFill"}) count += drawingTags(list, (std::string("a:") + tag).c_str(), tag).size();
            return count;
        };
        const auto fills = drawingTags(fmtSchemes.front(), "a:fillStyleLst", "fillStyleLst"); if (!fills.empty()) effects.fillStyleCount = countFills(fills.front());
        const auto lines = drawingTags(fmtSchemes.front(), "a:lnStyleLst", "lnStyleLst"); if (!lines.empty()) effects.lineStyleCount = drawingTags(lines.front(), "a:ln", "ln").size();
        const auto effectStyles = drawingTags(fmtSchemes.front(), "a:effectStyleLst", "effectStyleLst"); if (!effectStyles.empty()) effects.effectStyleCount = drawingTags(effectStyles.front(), "a:effectStyle", "effectStyle").size();
        const auto bgFills = drawingTags(fmtSchemes.front(), "a:bgFillStyleLst", "bgFillStyleLst"); if (!bgFills.empty()) effects.backgroundFillStyleCount = countFills(bgFills.front());
    }
    palette.present = !palette.colors.empty() || palette.fontScheme.present || palette.effectScheme.present;
    return palette;
}

xlpp::ChartStyleResources parseChartStyleResources(const xlpp::internal::ZipArchive& z, const std::string& chartPart) {
    xlpp::ChartStyleResources resources;
    const auto slash = chartPart.find_last_of('/');
    if (slash == std::string::npos) return resources;
    const auto file = chartPart.substr(slash + 1);
    const auto relsPart = chartPart.substr(0, slash + 1) + "_rels/" + file + ".rels";
    if (!z.contains(relsPart)) return resources;
    for (const auto& rel : xlpp::internal::tags(z.get(relsPart), "Relationship")) {
        const auto type = xlpp::internal::attribute(rel, "Type");
        const auto target = xlpp::internal::attribute(rel, "Target");
        if (target.empty()) continue;
        if (type.find("/chartStyle") != std::string::npos) { resources.chartStylePresent = true; resources.chartStylePart = resolvePackagePart(chartPart, target); }
        else if (type.find("/chartColorStyle") != std::string::npos) { resources.colorStylePresent = true; resources.colorStylePart = resolvePackagePart(chartPart, target); }
    }
    return resources;
}

xlpp::ChartLineFormat parseChartLineFormat(const std::string& container) {
    xlpp::ChartLineFormat format;
    const auto lines = drawingTags(container, "a:ln", "ln");
    if (lines.empty()) return format;
    const auto& line = lines.front();
    format.present = true;
    const auto width = xlpp::internal::attribute(line, "w");
    if (!width.empty()) {
        try { format.widthPoints = std::stod(width) / 12700.0; } catch (...) {}
    }
    format.cap = xlpp::internal::attribute(line, "cap");
    format.compound = xlpp::internal::attribute(line, "cmpd");
    format.noFill = !drawingTags(line, "a:noFill", "noFill").empty();
    const auto fills = drawingTags(line, "a:solidFill", "solidFill");
    if (!fills.empty()) format.color = parseChartColor(fills.front());
    const auto dashes = drawingTags(line, "a:prstDash", "prstDash");
    if (!dashes.empty()) format.dash = xlpp::internal::attribute(dashes.front(), "val");
    const auto custom = drawingTags(line, "a:custDash", "custDash");
    if (!custom.empty()) {
        for (const auto& ds : drawingTags(custom.front(), "a:ds", "ds")) {
            try {
                format.customDash.push_back({
                    std::stod(xlpp::internal::attribute(ds, "d")) / 1000.0,
                    std::stod(xlpp::internal::attribute(ds, "sp")) / 1000.0
                });
            } catch (...) {}
        }
    }
    if (!drawingTags(line, "a:round", "round").empty()) format.join = "round";
    else if (!drawingTags(line, "a:bevel", "bevel").empty()) format.join = "bevel";
    else if (!drawingTags(line, "a:miter", "miter").empty()) format.join = "miter";
    return format;
}

xlpp::ChartFillFormat parseChartFillFormat(const std::string& container) {
    xlpp::ChartFillFormat format;
    auto scope = container;
    for (const auto& line : drawingTags(scope, "a:ln", "ln")) {
        const auto position = scope.find(line);
        if (position != std::string::npos) scope.erase(position, line.size());
    }
    if (!drawingTags(scope, "a:noFill", "noFill").empty()) {
        format.present = true;
        format.noFill = true;
        format.kind = xlpp::ChartFillFormat::Kind::NoFill;
        return format;
    }
    const auto fills = drawingTags(scope, "a:solidFill", "solidFill");
    if (!fills.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Solid;
        format.color = parseChartColor(fills.front());
        return format;
    }
    const auto gradients = drawingTags(scope, "a:gradFill", "gradFill");
    if (!gradients.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Gradient;
        for (const auto& gs : drawingTags(gradients.front(), "a:gs", "gs")) {
            xlpp::ChartGradientStop stop;
            try { stop.position = std::stoi(xlpp::internal::attribute(gs, "pos")); } catch (...) {}
            stop.color = parseChartColor(gs);
            if (stop.color.present()) format.gradientStops.push_back(std::move(stop));
        }
        const auto linear = drawingTags(gradients.front(), "a:lin", "lin");
        if (!linear.empty()) {
            try { format.gradientAngleDegrees = std::stod(xlpp::internal::attribute(linear.front(), "ang")) / 60000.0; } catch (...) {}
        }
        return format;
    }
    const auto patterns = drawingTags(scope, "a:pattFill", "pattFill");
    if (!patterns.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Pattern;
        format.pattern = xlpp::internal::attribute(patterns.front(), "prst");
        const auto foreground = drawingTags(patterns.front(), "a:fgClr", "fgClr");
        if (!foreground.empty()) format.foregroundColor = parseChartColor(foreground.front());
        const auto background = drawingTags(patterns.front(), "a:bgClr", "bgClr");
        if (!background.empty()) format.backgroundColor = parseChartColor(background.front());
    }
    return format;
}

xlpp::ChartMarkerFormat parseChartMarkerFormat(const std::string& seriesXml) {
    xlpp::ChartMarkerFormat format;
    const auto markers = drawingTags(seriesXml, "c:marker", "marker");
    if (markers.empty()) return format;
    const auto& marker = markers.front();
    format.present = true;
    const auto symbols = drawingTags(marker, "c:symbol", "symbol");
    if (!symbols.empty()) format.symbol = xlpp::internal::attribute(symbols.front(), "val");
    const auto sizes = drawingTags(marker, "c:size", "size");
    if (!sizes.empty()) { try { format.size = std::stoi(xlpp::internal::attribute(sizes.front(), "val")); } catch (...) {} }
    const auto spPr = drawingTags(marker, "c:spPr", "spPr");
    if (!spPr.empty()) {
        format.line = parseChartLineFormat(spPr.front());
        format.fill = parseChartFillFormat(spPr.front());
    }
    return format;
}

std::string seriesDirectSpPr(const std::string& seriesXml) {
    const auto candidates = drawingTags(seriesXml, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    std::vector<std::string> nested;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 5>{
             std::pair{"c:marker", "marker"}, {"c:dPt", "dPt"}, {"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}, {"c:errBars", "errBars"}}) {
        const auto nodes = drawingTags(seriesXml, pair.first, pair.second);
        nested.insert(nested.end(), nodes.begin(), nodes.end());
    }
    for (const auto& candidate : candidates) {
        if (std::none_of(nested.begin(), nested.end(), [&](const auto& node) { return node.find(candidate) != std::string::npos; }))
            return candidate;
    }
    return {};
}

xlpp::ChartRichText parseChartRichText(const std::string& owner) {
    xlpp::ChartRichText result;
    const auto richNodes = drawingTags(owner, "c:rich", "rich");
    if (richNodes.empty()) return result;
    result.present = true;
    const auto& rich = richNodes.front();
    for (const auto& runXml : drawingTags(rich, "a:r", "r")) {
        xlpp::ChartTextRun run;
        run.text = drawingTagText(runXml, "a:t", "t");
        const auto properties = drawingTags(runXml, "a:rPr", "rPr");
        if (!properties.empty()) {
            const auto& rPr = properties.front();
            const auto bold = xlpp::internal::attribute(rPr, "b");
            const auto italic = xlpp::internal::attribute(rPr, "i");
            run.bold = bold == "1" || bold == "true";
            run.italic = italic == "1" || italic == "true";
            const auto size = xlpp::internal::attribute(rPr, "sz");
            if (!size.empty()) { try { run.fontSizePoints = std::stod(size) / 100.0; } catch (...) {} }
            const auto latin = drawingTags(rPr, "a:latin", "latin");
            if (!latin.empty()) run.typeface = xlpp::internal::attribute(latin.front(), "typeface");
            run.color = parseChartColor(rPr);
        }
        result.runs.push_back(std::move(run));
    }
    if (result.runs.empty()) {
        const auto text = drawingTagText(rich, "a:t", "t");
        if (!text.empty()) result.runs.push_back({text});
    }
    return result;
}

xlpp::ChartTextStyle parseChartTextStyle(const std::string& owner) {
    xlpp::ChartTextStyle style;
    const auto txPr = drawingTags(owner, "c:txPr", "txPr");
    if (txPr.empty()) return style;
    std::string properties;
    const auto defaults = drawingTags(txPr.front(), "a:defRPr", "defRPr");
    if (!defaults.empty()) properties = defaults.front();
    else {
        const auto ends = drawingTags(txPr.front(), "a:endParaRPr", "endParaRPr");
        if (!ends.empty()) properties = ends.front();
        else {
            const auto runs = drawingTags(txPr.front(), "a:rPr", "rPr");
            if (!runs.empty()) properties = runs.front();
        }
    }
    if (properties.empty()) { style.present = true; return style; }
    style.present = true;
    const auto bold = xlpp::internal::attribute(properties, "b");
    const auto italic = xlpp::internal::attribute(properties, "i");
    style.bold = bold == "1" || bold == "true";
    style.italic = italic == "1" || italic == "true";
    const auto size = xlpp::internal::attribute(properties, "sz");
    if (!size.empty()) { try { style.fontSizePoints = std::stod(size) / 100.0; } catch (...) {} }
    const auto latin = drawingTags(properties, "a:latin", "latin");
    if (!latin.empty()) style.typeface = xlpp::internal::attribute(latin.front(), "typeface");
    style.color = parseChartColor(properties);
    return style;
}

std::vector<xlpp::ChartDataPointFormat> parseChartDataPoints(const std::string& seriesXml) {
    std::vector<xlpp::ChartDataPointFormat> result;
    for (const auto& node : drawingTags(seriesXml, "c:dPt", "dPt")) {
        xlpp::ChartDataPointFormat point;
        const auto indices = drawingTags(node, "c:idx", "idx");
        if (!indices.empty()) { try { point.index = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(indices.front(), "val"))); } catch (...) {} }
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) {
            point.line = parseChartLineFormat(spPr.front());
            point.fill = parseChartFillFormat(spPr.front());
        }
        point.marker = parseChartMarkerFormat(node);
        result.push_back(std::move(point));
    }
    return result;
}

xlpp::ChartDataLabelPoint parseChartDataLabelPoint(const std::string& xml) {
    xlpp::ChartDataLabelPoint point;
    const auto idx = drawingTags(xml, "c:idx", "idx");
    if (!idx.empty()) { try { point.index = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(idx.front(), "val"))); } catch (...) {} }
    point.deleted = chartBoolValue(xml, "c:delete", "delete");
    point.showLegendKey = chartBoolValue(xml, "c:showLegendKey", "showLegendKey");
    point.showValue = chartBoolValue(xml, "c:showVal", "showVal");
    point.showCategoryName = chartBoolValue(xml, "c:showCatName", "showCatName");
    point.showSeriesName = chartBoolValue(xml, "c:showSerName", "showSerName");
    point.showPercent = chartBoolValue(xml, "c:showPercent", "showPercent");
    point.showBubbleSize = chartBoolValue(xml, "c:showBubbleSize", "showBubbleSize");
    point.showLeaderLines = chartBoolValue(xml, "c:showLeaderLines", "showLeaderLines");
    const auto positions = drawingTags(xml, "c:dLblPos", "dLblPos");
    if (!positions.empty()) point.position = xlpp::internal::attribute(positions.front(), "val");
    point.separator = drawingTagText(xml, "c:separator", "separator");
    point.richText = parseChartRichText(xml);
    return point;
}

xlpp::Chart::DataLabels parseChartDataLabels(const std::string& plotXml, bool directPlotChild = false) {
    xlpp::Chart::DataLabels labels;
    auto nodes = drawingTags(plotXml, "c:dLbls", "dLbls");
    if (directPlotChild && !nodes.empty()) {
        const auto seriesNodes = drawingTags(plotXml, "c:ser", "ser");
        nodes.erase(std::remove_if(nodes.begin(), nodes.end(), [&](const std::string& node) {
            return std::any_of(seriesNodes.begin(), seriesNodes.end(), [&](const std::string& series) {
                return series.find(node) != std::string::npos;
            });
        }), nodes.end());
    }
    if (nodes.empty()) return labels;
    labels.present = true;
    const auto& xml = nodes.front();
    auto aggregateXml = xml;
    for (const auto& pointNode : drawingTags(xml, "c:dLbl", "dLbl")) {
        labels.points.push_back(parseChartDataLabelPoint(pointNode));
        const auto position = aggregateXml.find(pointNode);
        if (position != std::string::npos) aggregateXml.erase(position, pointNode.size());
    }
    labels.showLegendKey = chartBoolValue(aggregateXml, "c:showLegendKey", "showLegendKey");
    labels.showValue = chartBoolValue(aggregateXml, "c:showVal", "showVal");
    labels.showCategoryName = chartBoolValue(aggregateXml, "c:showCatName", "showCatName");
    labels.showSeriesName = chartBoolValue(aggregateXml, "c:showSerName", "showSerName");
    labels.showPercent = chartBoolValue(aggregateXml, "c:showPercent", "showPercent");
    labels.showBubbleSize = chartBoolValue(aggregateXml, "c:showBubbleSize", "showBubbleSize");
    labels.showLeaderLines = chartBoolValue(aggregateXml, "c:showLeaderLines", "showLeaderLines");
    const auto leaderLines = drawingTags(aggregateXml, "c:leaderLines", "leaderLines");
    if (!leaderLines.empty()) {
        labels.hasLeaderLines = true;
        const auto spPr = drawingTags(leaderLines.front(), "c:spPr", "spPr");
        if (!spPr.empty()) labels.leaderLineFormat = parseChartLineFormat(spPr.front());
    }
    const auto positions = drawingTags(aggregateXml, "c:dLblPos", "dLblPos");
    if (!positions.empty()) labels.position = xlpp::internal::attribute(positions.front(), "val");
    labels.separator = drawingTagText(aggregateXml, "c:separator", "separator");
    return labels;
}

xlpp::ChartSeries::TrendlineType parseTrendlineTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::TrendlineType;
    if (value == "exp") return T::Exponential;
    if (value == "log") return T::Logarithmic;
    if (value == "poly") return T::Polynomial;
    if (value == "power") return T::Power;
    if (value == "movingAvg") return T::MovingAverage;
    return T::Linear;
}

std::vector<xlpp::ChartSeries::Trendline> parseChartTrendlines(const std::string& seriesXml) {
    std::vector<xlpp::ChartSeries::Trendline> result;
    for (const auto& node : drawingTags(seriesXml, "c:trendline", "trendline")) {
        xlpp::ChartSeries::Trendline trendline;
        const auto types = drawingTags(node, "c:trendlineType", "trendlineType");
        if (!types.empty()) trendline.type = parseTrendlineTypeValue(xlpp::internal::attribute(types.front(), "val"));
        const auto order = drawingTags(node, "c:order", "order");
        if (!order.empty()) { try { trendline.order = std::stoi(xlpp::internal::attribute(order.front(), "val")); } catch (...) {} }
        const auto period = drawingTags(node, "c:period", "period");
        if (!period.empty()) { try { trendline.period = std::stoi(xlpp::internal::attribute(period.front(), "val")); } catch (...) {} }
        const auto forward = drawingTags(node, "c:forward", "forward");
        if (!forward.empty()) { try { trendline.forward = std::stod(xlpp::internal::attribute(forward.front(), "val")); } catch (...) {} }
        const auto backward = drawingTags(node, "c:backward", "backward");
        if (!backward.empty()) { try { trendline.backward = std::stod(xlpp::internal::attribute(backward.front(), "val")); } catch (...) {} }
        trendline.displayEquation = chartBoolValue(node, "c:dispEq", "dispEq");
        trendline.displayRSquared = chartBoolValue(node, "c:dispRSqr", "dispRSqr");
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) trendline.lineFormat = parseChartLineFormat(spPr.front());
        result.push_back(std::move(trendline));
    }
    return result;
}

xlpp::ChartSeries::ErrorBarDirection parseErrorBarDirectionValue(const std::string& value) {
    return value == "x" ? xlpp::ChartSeries::ErrorBarDirection::X : xlpp::ChartSeries::ErrorBarDirection::Y;
}

xlpp::ChartSeries::ErrorBarType parseErrorBarTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::ErrorBarType;
    if (value == "plus") return T::Plus;
    if (value == "minus") return T::Minus;
    return T::Both;
}

xlpp::ChartSeries::ErrorValueType parseErrorValueTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::ErrorValueType;
    if (value == "percentage") return T::Percentage;
    if (value == "stdDev") return T::StandardDeviation;
    if (value == "stdErr") return T::StandardError;
    if (value == "cust") return T::Custom;
    return T::FixedValue;
}

std::vector<xlpp::ChartSeries::ErrorBars> parseChartErrorBars(const std::string& seriesXml) {
    std::vector<xlpp::ChartSeries::ErrorBars> result;
    for (const auto& node : drawingTags(seriesXml, "c:errBars", "errBars")) {
        xlpp::ChartSeries::ErrorBars bars;
        const auto dirs = drawingTags(node, "c:errDir", "errDir");
        if (!dirs.empty()) bars.direction = parseErrorBarDirectionValue(xlpp::internal::attribute(dirs.front(), "val"));
        const auto types = drawingTags(node, "c:errBarType", "errBarType");
        if (!types.empty()) bars.barType = parseErrorBarTypeValue(xlpp::internal::attribute(types.front(), "val"));
        const auto valueTypes = drawingTags(node, "c:errValType", "errValType");
        if (!valueTypes.empty()) bars.valueType = parseErrorValueTypeValue(xlpp::internal::attribute(valueTypes.front(), "val"));
        bars.noEndCap = chartBoolValue(node, "c:noEndCap", "noEndCap");
        const auto values = drawingTags(node, "c:val", "val");
        if (!values.empty()) { try { bars.value = std::stod(xlpp::internal::attribute(values.front(), "val")); } catch (...) {} }
        const auto plus = drawingTags(node, "c:plus", "plus");
        if (!plus.empty()) bars.plusReference = drawingTagText(plus.front(), "c:f", "f");
        const auto minus = drawingTags(node, "c:minus", "minus");
        if (!minus.empty()) bars.minusReference = drawingTagText(minus.front(), "c:f", "f");
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) bars.lineFormat = parseChartLineFormat(spPr.front());
        result.push_back(std::move(bars));
    }
    return result;
}

xlpp::ChartDataTable parseChartDataTable(const std::string& plotArea) {
    xlpp::ChartDataTable table;
    const auto nodes = drawingTags(plotArea, "c:dTable", "dTable");
    if (nodes.empty()) return table;
    table.present = true;
    const auto& xml = nodes.front();
    table.showHorizontalBorder = chartBoolValue(xml, "c:showHorzBorder", "showHorzBorder");
    table.showVerticalBorder = chartBoolValue(xml, "c:showVertBorder", "showVertBorder");
    table.showOutline = chartBoolValue(xml, "c:showOutline", "showOutline");
    table.showLegendKeys = chartBoolValue(xml, "c:showKeys", "showKeys");
    const auto spPr = drawingTags(xml, "c:spPr", "spPr");
    if (!spPr.empty()) {
        table.line = parseChartLineFormat(spPr.front());
        table.fill = parseChartFillFormat(spPr.front());
    }
    table.textStyle = parseChartTextStyle(xml);
    return table;
}

xlpp::ChartLineFormat parsePlotLineObject(const std::string& plotXml, const char* prefixed, const char* local, bool& present) {
    present = false;
    const auto nodes = drawingTags(plotXml, prefixed, local);
    if (nodes.empty()) return {};
    present = true;
    const auto spPr = drawingTags(nodes.front(), "c:spPr", "spPr");
    return spPr.empty() ? xlpp::ChartLineFormat{} : parseChartLineFormat(spPr.front());
}

xlpp::ChartUpDownBars parseChartUpDownBars(const std::string& plotXml) {
    xlpp::ChartUpDownBars bars;
    const auto nodes = drawingTags(plotXml, "c:upDownBars", "upDownBars");
    if (nodes.empty()) return bars;
    bars.present = true;
    const auto& xml = nodes.front();
    const auto gaps = drawingTags(xml, "c:gapWidth", "gapWidth");
    if (!gaps.empty()) { try { bars.gapWidth = std::stoi(xlpp::internal::attribute(gaps.front(), "val")); } catch (...) {} }
    const auto parseBar = [&](const char* prefixed, const char* local, xlpp::ChartFillFormat& fill, xlpp::ChartLineFormat& line) {
        const auto barNodes = drawingTags(xml, prefixed, local);
        if (barNodes.empty()) return;
        const auto spPr = drawingTags(barNodes.front(), "c:spPr", "spPr");
        if (!spPr.empty()) { fill = parseChartFillFormat(spPr.front()); line = parseChartLineFormat(spPr.front()); }
    };
    parseBar("c:upBars", "upBars", bars.upFill, bars.upLine);
    parseBar("c:downBars", "downBars", bars.downFill, bars.downLine);
    return bars;
}

std::vector<xlpp::Chart::Plot> parseChartPlots(const std::string& chartXmlText) {
    std::vector<xlpp::Chart::Plot> plots;
    std::size_t firstSeries = 0;
    for (const auto& parsed : chartPlotNodesInOrder(chartXmlText)) {
        xlpp::Chart::Plot plot;
        plot.type = parsed.type;
        plot.grouping = parseChartGroupingNode(parsed.xml);
        plot.axisIds = chartAxisIds(parsed.xml);
        plot.firstSeries = firstSeries;
        plot.seriesCount = drawingTags(parsed.xml, "c:ser", "ser").size();
        plot.dataLabels = parseChartDataLabels(parsed.xml, true);
        plot.dropLinesFormat = parsePlotLineObject(parsed.xml, "c:dropLines", "dropLines", plot.hasDropLines);
        plot.highLowLinesFormat = parsePlotLineObject(parsed.xml, "c:hiLowLines", "hiLowLines", plot.hasHighLowLines);
        plot.upDownBars = parseChartUpDownBars(parsed.xml);
        const auto gapDepth = drawingTags(parsed.xml, "c:gapDepth", "gapDepth");
        if (!gapDepth.empty()) { try { plot.gapDepth=std::stoi(xlpp::internal::attribute(gapDepth.front(),"val")); plot.hasGapDepth=true; } catch (...) {} }
        const auto wireframe = drawingTags(parsed.xml, "c:wireframe", "wireframe");
        if (!wireframe.empty()) { plot.hasWireframe=true; const auto value=xlpp::internal::attribute(wireframe.front(),"val"); plot.wireframe=value=="1"||value=="true"||value=="True"; }
        const auto shape = drawingTags(parsed.xml, "c:shape", "shape");
        if (!shape.empty()) plot.shape=xlpp::internal::attribute(shape.front(),"val");
        const auto firstSlice = drawingTags(parsed.xml, "c:firstSliceAng", "firstSliceAng");
        if (!firstSlice.empty()) { try { plot.firstSliceAngle=std::stoi(xlpp::internal::attribute(firstSlice.front(),"val")); plot.hasFirstSliceAngle=true; } catch (...) {} }
        const auto holeSize = drawingTags(parsed.xml, "c:holeSize", "holeSize");
        if (!holeSize.empty()) { try { plot.holeSize=std::stoi(xlpp::internal::attribute(holeSize.front(),"val")); plot.hasHoleSize=true; } catch (...) {} }
        const auto radarStyle = drawingTags(parsed.xml, "c:radarStyle", "radarStyle");
        if (!radarStyle.empty()) plot.radarStyle=xlpp::internal::attribute(radarStyle.front(),"val");
        if (parsed.type == xlpp::Chart::Type::PieOfPie || parsed.type == xlpp::Chart::Type::BarOfPie) {
            auto& options=plot.projectedPie; options.present=true; options.ofPieType=parsed.type==xlpp::Chart::Type::BarOfPie?"bar":"pie";
            const auto gapWidth=drawingTags(parsed.xml,"c:gapWidth","gapWidth"); if(!gapWidth.empty()) { try { options.gapWidth=std::stoi(xlpp::internal::attribute(gapWidth.front(),"val")); } catch (...) {} }
            const auto splitType=drawingTags(parsed.xml,"c:splitType","splitType"); if(!splitType.empty()) options.splitType=xlpp::internal::attribute(splitType.front(),"val");
            const auto splitPos=drawingTags(parsed.xml,"c:splitPos","splitPos"); if(!splitPos.empty()) { try { options.splitPosition=std::stod(xlpp::internal::attribute(splitPos.front(),"val")); options.hasSplitPosition=true; } catch (...) {} }
            const auto secondSize=drawingTags(parsed.xml,"c:secondPieSize","secondPieSize"); if(!secondSize.empty()) { try { options.secondPlotSize=std::stoi(xlpp::internal::attribute(secondSize.front(),"val")); } catch (...) {} }
            const auto custom=drawingTags(parsed.xml,"c:custSplit","custSplit"); if(!custom.empty()) for(const auto& pt:drawingTags(custom.front(),"c:secondPiePt","secondPiePt")) { try { options.customSplitPoints.push_back(std::stoi(xlpp::internal::attribute(pt,"val"))); } catch (...) {} }
            options.seriesLinesFormat=parsePlotLineObject(parsed.xml,"c:serLines","serLines",options.hasSeriesLines);
        }
        firstSeries += plot.seriesCount;
        plots.push_back(std::move(plot));
    }
    return plots;
}

xlpp::ChartManualLayout parseChartManualLayout(const std::string& owner) {
    xlpp::ChartManualLayout layout;
    const auto layouts = drawingTags(owner, "c:layout", "layout");
    if (layouts.empty()) return layout;
    const auto manuals = drawingTags(layouts.front(), "c:manualLayout", "manualLayout");
    if (manuals.empty()) return layout;
    layout.present = true;
    const auto& xml = manuals.front();
    const auto readString=[&](const char* p,const char* l){ const auto n=drawingTags(xml,p,l); return n.empty()?std::string{}:xlpp::internal::attribute(n.front(),"val"); };
    const auto readDouble=[&](const char* p,const char* l,bool& has,double& value){ const auto n=drawingTags(xml,p,l); if(n.empty()) return; const auto v=xlpp::internal::attribute(n.front(),"val"); if(v.empty()) return; try{ value=std::stod(v); has=true; }catch(...){} };
    layout.target=readString("c:layoutTarget","layoutTarget");
    layout.xMode=readString("c:xMode","xMode"); layout.yMode=readString("c:yMode","yMode");
    layout.widthMode=readString("c:wMode","wMode"); layout.heightMode=readString("c:hMode","hMode");
    readDouble("c:x","x",layout.hasX,layout.x); readDouble("c:y","y",layout.hasY,layout.y);
    readDouble("c:w","w",layout.hasWidth,layout.width); readDouble("c:h","h",layout.hasHeight,layout.height);
    return layout;
}

xlpp::ChartView3D parseChartView3D(const std::string& chartXmlText) {
    xlpp::ChartView3D view;
    const auto nodes = drawingTags(chartXmlText, "c:view3D", "view3D");
    if (nodes.empty()) return view;
    view.present = true;
    const auto& xml = nodes.front();
    const auto readInt = [&](const char* prefixed, const char* local, bool& has, int& out) {
        const auto values = drawingTags(xml, prefixed, local);
        if (values.empty()) return;
        const auto value = xlpp::internal::attribute(values.front(), "val");
        if (value.empty()) return;
        try { out = std::stoi(value); has = true; } catch (...) {}
    };
    readInt("c:rotX", "rotX", view.hasRotationX, view.rotationX);
    readInt("c:rotY", "rotY", view.hasRotationY, view.rotationY);
    readInt("c:hPercent", "hPercent", view.hasHeightPercent, view.heightPercent);
    readInt("c:depthPercent", "depthPercent", view.hasDepthPercent, view.depthPercent);
    readInt("c:perspective", "perspective", view.hasPerspective, view.perspective);
    const auto rAng = drawingTags(xml, "c:rAngAx", "rAngAx");
    if (!rAng.empty()) {
        view.hasRightAngleAxes = true;
        const auto value = xlpp::internal::attribute(rAng.front(), "val");
        view.rightAngleAxes = value.empty() || value == "1" || value == "true" || value == "True";
    }
    return view;
}

xlpp::ChartWallFormat parseChartWallFormat(const std::string& chartXmlText, const char* prefixed, const char* local) {
    xlpp::ChartWallFormat wall;
    const auto nodes = drawingTags(chartXmlText, prefixed, local);
    if (nodes.empty()) return wall;
    wall.present = true;
    const auto& xml = nodes.front();
    const auto thickness = drawingTags(xml, "c:thickness", "thickness");
    if (!thickness.empty()) {
        try {
            const auto value = xlpp::internal::attribute(thickness.front(), "val");
            if (!value.empty()) { wall.thickness = std::stoi(value); wall.hasThickness = true; }
        } catch (...) {}
    }
    const auto spPr = drawingTags(xml, "c:spPr", "spPr");
    if (!spPr.empty()) {
        wall.line = parseChartLineFormat(spPr.front());
        wall.fill = parseChartFillFormat(spPr.front());
    }
    return wall;
}

std::string axisDirectSpPr(const std::string& axisXml) {
    const auto candidates=drawingTags(axisXml,"c:spPr","spPr");
    if(candidates.empty()) return {};
    std::vector<std::string> nested;
    for(const auto& pair: std::array<std::pair<const char*,const char*>,5>{{{"c:title","title"},{"c:majorGridlines","majorGridlines"},{"c:minorGridlines","minorGridlines"},{"c:txPr","txPr"},{"c:extLst","extLst"}}}) {
        const auto nodes=drawingTags(axisXml,pair.first,pair.second); nested.insert(nested.end(),nodes.begin(),nodes.end());
    }
    for(const auto& candidate:candidates)
        if(std::none_of(nested.begin(),nested.end(),[&](const auto& node){ return node.find(candidate)!=std::string::npos; })) return candidate;
    return {};
}

std::vector<xlpp::Chart::Axis> parseChartAxes(const std::string& chartXmlText,
                                               const std::vector<xlpp::Chart::Plot>& plots) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return {};
    const auto& plotArea = plotAreas.front();
    struct AxisNode { std::size_t position; xlpp::Chart::AxisKind kind; std::string xml; };
    std::vector<AxisNode> nodes;
    const auto collect = [&](const char* prefixed, const char* local, xlpp::Chart::AxisKind kind) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, prefixed, local)) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            nodes.push_back({position, kind, node});
            cursor = position + node.size();
        }
    };
    collect("c:catAx", "catAx", xlpp::Chart::AxisKind::Category);
    collect("c:valAx", "valAx", xlpp::Chart::AxisKind::Value);
    collect("c:dateAx", "dateAx", xlpp::Chart::AxisKind::Date);
    collect("c:serAx", "serAx", xlpp::Chart::AxisKind::Series);
    std::sort(nodes.begin(), nodes.end(), [](const auto& a, const auto& b) { return a.position < b.position; });

    std::set<std::uint64_t> primaryAxisIds;
    if (!plots.empty()) primaryAxisIds.insert(plots.front().axisIds.begin(), plots.front().axisIds.end());
    std::set<std::uint64_t> laterPlotAxisIds;
    for (std::size_t i = 1; i < plots.size(); ++i)
        laterPlotAxisIds.insert(plots[i].axisIds.begin(), plots[i].axisIds.end());

    std::vector<xlpp::Chart::Axis> axes;
    for (const auto& parsed : nodes) {
        xlpp::Chart::Axis axis;
        axis.kind = parsed.kind;
        const auto ids = drawingTags(parsed.xml, "c:axId", "axId");
        if (!ids.empty()) {
            const auto value = xlpp::internal::attribute(ids.front(), "val");
            try { if (!value.empty()) axis.id = std::stoull(value); } catch (...) {}
        }
        const auto crosses = drawingTags(parsed.xml, "c:crossAx", "crossAx");
        if (!crosses.empty()) {
            const auto value = xlpp::internal::attribute(crosses.front(), "val");
            try { if (!value.empty()) axis.crossAxisId = std::stoull(value); } catch (...) {}
        }
        const auto positions = drawingTags(parsed.xml, "c:axPos", "axPos");
        if (!positions.empty()) axis.position = xlpp::internal::attribute(positions.front(), "val");
        const auto scalings = drawingTags(parsed.xml, "c:scaling", "scaling");
        if (!scalings.empty()) {
            const auto& scaling = scalings.front();
            const auto minimum = drawingTags(scaling, "c:min", "min");
            if (!minimum.empty()) { try { axis.scaling.minimum=std::stod(xlpp::internal::attribute(minimum.front(),"val")); axis.scaling.hasMinimum=true; } catch (...) {} }
            const auto maximum = drawingTags(scaling, "c:max", "max");
            if (!maximum.empty()) { try { axis.scaling.maximum=std::stod(xlpp::internal::attribute(maximum.front(),"val")); axis.scaling.hasMaximum=true; } catch (...) {} }
            const auto logBase = drawingTags(scaling, "c:logBase", "logBase");
            if (!logBase.empty()) { try { axis.scaling.logBase=std::stod(xlpp::internal::attribute(logBase.front(),"val")); axis.scaling.hasLogBase=true; } catch (...) {} }
            const auto orientation = drawingTags(scaling, "c:orientation", "orientation");
            if (!orientation.empty()) axis.scaling.reverseOrder = xlpp::internal::attribute(orientation.front(),"val") == "maxMin";
        }
        const auto titles = drawingTags(parsed.xml, "c:title", "title");
        if (!titles.empty()) {
            axis.title = drawingTagText(titles.front(), "a:t", "t");
            if (axis.title.empty()) axis.title = drawingTagText(titles.front(), "c:v", "v");
            if (axis.title.empty()) axis.title = drawingTagText(titles.front(), "c:f", "f");
            axis.titleRichText = parseChartRichText(titles.front());
            if (axis.title.empty() && axis.titleRichText.present) axis.title = axis.titleRichText.plainText();
        }
        const auto numFmt = drawingTags(parsed.xml, "c:numFmt", "numFmt");
        if (!numFmt.empty()) {
            axis.numberFormat = xlpp::internal::attribute(numFmt.front(), "formatCode");
            const auto linked = xlpp::internal::attribute(numFmt.front(), "sourceLinked");
            axis.numberFormatSourceLinked = linked.empty() || linked == "1" || linked == "true";
        }
        const auto majorTick = drawingTags(parsed.xml, "c:majorTickMark", "majorTickMark");
        if (!majorTick.empty()) axis.majorTickMark = xlpp::internal::attribute(majorTick.front(), "val");
        const auto minorTick = drawingTags(parsed.xml, "c:minorTickMark", "minorTickMark");
        if (!minorTick.empty()) axis.minorTickMark = xlpp::internal::attribute(minorTick.front(), "val");
        const auto tickPos = drawingTags(parsed.xml, "c:tickLblPos", "tickLblPos");
        if (!tickPos.empty()) axis.tickLabelPosition = xlpp::internal::attribute(tickPos.front(), "val");
        const auto majorUnit = drawingTags(parsed.xml, "c:majorUnit", "majorUnit");
        if (!majorUnit.empty()) { try { axis.majorUnit = std::stod(xlpp::internal::attribute(majorUnit.front(), "val")); axis.hasMajorUnit = true; } catch (...) {} }
        const auto minorUnit = drawingTags(parsed.xml, "c:minorUnit", "minorUnit");
        if (!minorUnit.empty()) { try { axis.minorUnit = std::stod(xlpp::internal::attribute(minorUnit.front(), "val")); axis.hasMinorUnit = true; } catch (...) {} }
        const auto crossesNode = drawingTags(parsed.xml, "c:crosses", "crosses");
        if (!crossesNode.empty()) axis.crosses = xlpp::internal::attribute(crossesNode.front(), "val");
        const auto crossesAtNode = drawingTags(parsed.xml, "c:crossesAt", "crossesAt");
        if (!crossesAtNode.empty()) { try { axis.crossesAt=std::stod(xlpp::internal::attribute(crossesAtNode.front(),"val")); axis.hasCrossesAt=true; } catch (...) {} }
        const auto crossBetweenNode = drawingTags(parsed.xml, "c:crossBetween", "crossBetween");
        if (!crossBetweenNode.empty()) axis.crossBetween = xlpp::internal::attribute(crossBetweenNode.front(), "val");
        const auto displayUnits = drawingTags(parsed.xml, "c:dispUnits", "dispUnits");
        if (!displayUnits.empty()) {
            axis.displayUnits.present = true;
            const auto builtIn = drawingTags(displayUnits.front(), "c:builtInUnit", "builtInUnit");
            if (!builtIn.empty()) axis.displayUnits.builtInUnit = xlpp::internal::attribute(builtIn.front(), "val");
            const auto custom = drawingTags(displayUnits.front(), "c:custUnit", "custUnit");
            if (!custom.empty()) { try { axis.displayUnits.customUnit=std::stod(xlpp::internal::attribute(custom.front(),"val")); axis.displayUnits.hasCustomUnit=true; } catch (...) {} }
            const auto labels = drawingTags(displayUnits.front(), "c:dispUnitsLbl", "dispUnitsLbl");
            if (!labels.empty()) { axis.displayUnits.showLabel=true; axis.displayUnits.labelRichText=parseChartRichText(labels.front()); }
        }
        const auto axisSpPr = axisDirectSpPr(parsed.xml);
        if (!axisSpPr.empty()) axis.lineFormat = parseChartLineFormat(axisSpPr);
        const auto majorGrid = drawingTags(parsed.xml, "c:majorGridlines", "majorGridlines");
        axis.hasMajorGridlines = !majorGrid.empty();
        if (!majorGrid.empty()) { const auto spPr=drawingTags(majorGrid.front(),"c:spPr","spPr"); if(!spPr.empty()) axis.majorGridlineFormat=parseChartLineFormat(spPr.front()); }
        const auto minorGrid = drawingTags(parsed.xml, "c:minorGridlines", "minorGridlines");
        axis.hasMinorGridlines = !minorGrid.empty();
        if (!minorGrid.empty()) { const auto spPr=drawingTags(minorGrid.front(),"c:spPr","spPr"); if(!spPr.empty()) axis.minorGridlineFormat=parseChartLineFormat(spPr.front()); }
        axis.secondary = axis.id != 0 && primaryAxisIds.find(axis.id) == primaryAxisIds.end() &&
            laterPlotAxisIds.find(axis.id) != laterPlotAxisIds.end();
        axes.push_back(std::move(axis));
    }
    return axes;
}

std::string axisTitleById(const std::vector<xlpp::Chart::Axis>& axes, std::uint64_t axisId) {
    const auto it = std::find_if(axes.begin(), axes.end(), [&](const auto& axis) { return axis.id == axisId; });
    return it == axes.end() ? std::string{} : it->title;
}

xlpp::Chart::Type parseChartType(const std::string& chartXmlText) {
    using Type = xlpp::Chart::Type;
    if (!drawingTags(chartXmlText, "c:lineChart", "lineChart").empty()) return Type::Line;
    if (!drawingTags(chartXmlText, "c:pieChart", "pieChart").empty()) return Type::Pie;
    if (!drawingTags(chartXmlText, "c:scatterChart", "scatterChart").empty()) return Type::Scatter;
    if (!drawingTags(chartXmlText, "c:doughnutChart", "doughnutChart").empty()) return Type::Doughnut;
    if (!drawingTags(chartXmlText, "c:radarChart", "radarChart").empty()) return Type::Radar;
    if (!drawingTags(chartXmlText, "c:areaChart", "areaChart").empty()) return Type::Area;
    if (!drawingTags(chartXmlText, "c:bubbleChart", "bubbleChart").empty()) return Type::Bubble;
    if (!drawingTags(chartXmlText, "c:stockChart", "stockChart").empty()) return Type::Stock;
    const auto projectedPie = drawingTags(chartXmlText, "c:ofPieChart", "ofPieChart");
    if (!projectedPie.empty()) {
        const auto kinds=drawingTags(projectedPie.front(),"c:ofPieType","ofPieType");
        return !kinds.empty() && xlpp::internal::attribute(kinds.front(),"val")=="bar" ? Type::BarOfPie : Type::PieOfPie;
    }
    if (!drawingTags(chartXmlText, "c:bar3DChart", "bar3DChart").empty()) return Type::Bar3D;
    if (!drawingTags(chartXmlText, "c:line3DChart", "line3DChart").empty()) return Type::Line3D;
    if (!drawingTags(chartXmlText, "c:area3DChart", "area3DChart").empty()) return Type::Area3D;
    if (!drawingTags(chartXmlText, "c:pie3DChart", "pie3DChart").empty()) return Type::Pie3D;
    if (!drawingTags(chartXmlText, "c:surfaceChart", "surfaceChart").empty()) return Type::Surface;
    if (!drawingTags(chartXmlText, "c:surface3DChart", "surface3DChart").empty()) return Type::Surface3D;
    return Type::Bar;
}

xlpp::Chart::Grouping parseChartGrouping(const std::string& chartXmlText, xlpp::Chart::Type type) {
    const char* prefixed = "c:barChart";
    const char* local = "barChart";
    switch (type) {
        case xlpp::Chart::Type::Line: prefixed = "c:lineChart"; local = "lineChart"; break;
        case xlpp::Chart::Type::Area: prefixed = "c:areaChart"; local = "areaChart"; break;
        case xlpp::Chart::Type::Bar3D: prefixed = "c:bar3DChart"; local = "bar3DChart"; break;
        case xlpp::Chart::Type::Line3D: prefixed = "c:line3DChart"; local = "line3DChart"; break;
        case xlpp::Chart::Type::Area3D: prefixed = "c:area3DChart"; local = "area3DChart"; break;
        default: break;
    }
    const auto chartTypeNodes = drawingTags(chartXmlText, prefixed, local);
    if (chartTypeNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto groupingNodes = drawingTags(chartTypeNodes.front(), "c:grouping", "grouping");
    if (groupingNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto value = xlpp::internal::attribute(groupingNodes.front(), "val");
    if (value == "stacked") return xlpp::Chart::Grouping::Stacked;
    if (value == "percentStacked") return xlpp::Chart::Grouping::PercentStacked;
    if (value == "clustered") return xlpp::Chart::Grouping::Clustered;
    return xlpp::Chart::Grouping::Standard;
}

std::string chartTitleText(const std::string& chartXmlText) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return {};
    const auto& chartNode = chartNodes.front();
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    for (const auto& titleNode : drawingTags(chartNode, "c:title", "title")) {
        const auto titlePosition = chartNode.find(titleNode);
        if (plotPosition != std::string::npos && titlePosition > plotPosition) continue;
        auto value = drawingTagText(titleNode, "a:t", "t");
        if (value.empty()) value = drawingTagText(titleNode, "c:v", "v");
        if (value.empty()) value = drawingTagText(titleNode, "c:f", "f");
        return value;
    }
    return {};
}

std::string axisTitleText(const std::string& chartXmlText, const char* prefixedAxis, const char* localAxis, std::size_t axisIndex = 0) {
    const auto axes = drawingTags(chartXmlText, prefixedAxis, localAxis);
    if (axisIndex >= axes.size()) return {};
    const auto titles = drawingTags(axes[axisIndex], "c:title", "title");
    if (titles.empty()) return {};
    auto value = drawingTagText(titles.front(), "a:t", "t");
    if (value.empty()) value = drawingTagText(titles.front(), "c:v", "v");
    if (value.empty()) value = drawingTagText(titles.front(), "c:f", "f");
    return value;
}

xlpp::DrawingAnchorInfo parseChartAnchorInfo(const std::string& anchorNode,
                                             xlpp::DrawingAnchorType type,
                                             const std::string& graphicFrame) {
    xlpp::DrawingAnchorInfo info;
    info.type = type;
    info.editAs = xlpp::internal::attribute(anchorNode, "editAs");
    const auto fromNodes = drawingTags(anchorNode, "xdr:from", "from");
    if (!fromNodes.empty()) info.from = parseDrawingMarker(fromNodes.front());
    const auto toNodes = drawingTags(anchorNode, "xdr:to", "to");
    if (!toNodes.empty()) info.to = parseDrawingMarker(toNodes.front());
    const auto posNodes = drawingTags(anchorNode, "xdr:pos", "pos");
    if (!posNodes.empty()) {
        const auto x = xlpp::internal::attribute(posNodes.front(), "x");
        const auto y = xlpp::internal::attribute(posNodes.front(), "y");
        if (!x.empty()) info.xEmu = std::stoll(x);
        if (!y.empty()) info.yEmu = std::stoll(y);
    }
    const auto anchorExt = drawingTags(anchorNode, "xdr:ext", "ext");
    if (!anchorExt.empty()) {
        const auto cx = xlpp::internal::attribute(anchorExt.front(), "cx");
        const auto cy = xlpp::internal::attribute(anchorExt.front(), "cy");
        if (!cx.empty()) info.widthEmu = std::stoll(cx);
        if (!cy.empty()) info.heightEmu = std::stoll(cy);
    }
    if ((info.widthEmu <= 0 || info.heightEmu <= 0) && type == xlpp::DrawingAnchorType::TwoCell) {
        const auto frameExt = drawingTags(graphicFrame, "a:ext", "ext");
        if (!frameExt.empty()) {
            const auto cx = xlpp::internal::attribute(frameExt.front(), "cx");
            const auto cy = xlpp::internal::attribute(frameExt.front(), "cy");
            if (!cx.empty()) info.widthEmu = std::stoll(cx);
            if (!cy.empty()) info.heightEmu = std::stoll(cy);
        }
    }
    return info;
}

void loadCharts(xlpp::Worksheet& ws, const std::string& sheetXml, const xlpp::internal::ZipArchive& z,
                const std::string& sheetPart) {
    const auto drawings = xlpp::internal::tags(sheetXml, "drawing");
    if (drawings.empty()) return;
    const auto sheetSlash = sheetPart.find_last_of('/');
    const auto sheetFile = sheetPart.substr(sheetSlash + 1);
    const auto sheetRelsPart = sheetPart.substr(0, sheetSlash + 1) + "_rels/" + sheetFile + ".rels";
    if (!z.contains(sheetRelsPart)) return;
    std::unordered_map<std::string, std::string> sheetRelationships;
    for (const auto& rel : xlpp::internal::tags(z.get(sheetRelsPart), "Relationship"))
        if (xlpp::internal::attribute(rel, "Type").find("/drawing") != std::string::npos)
            sheetRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

    for (const auto& drawingNode : drawings) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = sheetRelationships.find(relationshipId);
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sheetPart, relationship->second);
        if (!z.contains(drawingPart)) continue;
        const auto drawingXmlText = z.get(drawingPart);
        const auto drawingSlash = drawingPart.find_last_of('/');
        const auto drawingFile = drawingPart.substr(drawingSlash + 1);
        const auto drawingRelsPart = drawingPart.substr(0, drawingSlash + 1) + "_rels/" + drawingFile + ".rels";
        if (!z.contains(drawingRelsPart)) continue;
        std::unordered_map<std::string, std::string> drawingRelationships;
        for (const auto& rel : xlpp::internal::tags(z.get(drawingRelsPart), "Relationship"))
            if (xlpp::internal::attribute(rel, "Type").find("/chart") != std::string::npos)
                drawingRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

        const auto loadAnchors = [&](const char* prefixedName, const char* localName, xlpp::DrawingAnchorType anchorType) {
            for (const auto& anchorNode : drawingTags(drawingXmlText, prefixedName, localName)) {
                const auto frames = drawingTags(anchorNode, "xdr:graphicFrame", "graphicFrame");
                if (frames.empty()) continue;
                const auto& frame = frames.front();
                const auto chartRefs = drawingTags(frame, "c:chart", "chart");
                if (chartRefs.empty()) continue;
                const auto chartRelationshipId = xlpp::internal::attribute(chartRefs.front(), "r:id");
                const auto chartRelationship = drawingRelationships.find(chartRelationshipId);
                if (chartRelationship == drawingRelationships.end()) continue;
                const auto chartPart = resolvePackagePart(drawingPart, chartRelationship->second);
                if (!z.contains(chartPart)) continue;
                const auto chartText = z.get(chartPart);

                auto plots = parseChartPlots(chartText);
                const auto primaryType = plots.empty() ? parseChartType(chartText) : plots.front().type;
                xlpp::Chart chart(primaryType);
                chart.setGrouping(plots.empty() ? parseChartGrouping(chartText, chart.type()) : plots.front().grouping);
                chart.setThemePalette(parseChartThemePalette(z));
                chart.setStyleResources(parseChartStyleResources(z, chartPart));
                chart.setTitle(chartTitleText(chartText));
                const auto chartTitleNodes = drawingTags(chartText, "c:title", "title");
                if (!chartTitleNodes.empty()) chart.setTitleRichText(parseChartRichText(chartTitleNodes.front()));
                chart.setView3D(parseChartView3D(chartText));
                chart.setFloorFormat(parseChartWallFormat(chartText, "c:floor", "floor"));
                chart.setSideWallFormat(parseChartWallFormat(chartText, "c:sideWall", "sideWall"));
                chart.setBackWallFormat(parseChartWallFormat(chartText, "c:backWall", "backWall"));
                auto axes = parseChartAxes(chartText, plots);
                std::uint64_t primaryXAxisId = 0;
                std::uint64_t primaryYAxisId = 0;
                if (!plots.empty() && plots.front().axisIds.size() >= 2) {
                    primaryXAxisId = plots.front().axisIds[0];
                    primaryYAxisId = plots.front().axisIds[1];
                }
                for (auto& plot : plots) {
                    plot.usesSecondaryAxes = std::any_of(plot.axisIds.begin(), plot.axisIds.end(), [&](std::uint64_t axisId) {
                        const auto it = std::find_if(axes.begin(), axes.end(), [&](const auto& axis) { return axis.id == axisId; });
                        return it != axes.end() && it->secondary;
                    });
                }
                chart.setPrimaryAxisIds(primaryXAxisId, primaryYAxisId);
                chart.setXAxisTitle(primaryXAxisId != 0 ? axisTitleById(axes, primaryXAxisId) : std::string{});
                chart.setYAxisTitle(primaryYAxisId != 0 ? axisTitleById(axes, primaryYAxisId) : std::string{});
                if (primaryXAxisId == 0 || primaryYAxisId == 0) {
                    const bool xyValueAxes = chart.type() == xlpp::Chart::Type::Scatter || chart.type() == xlpp::Chart::Type::Bubble;
                    if (primaryXAxisId == 0) chart.setXAxisTitle(axisTitleText(chartText, xyValueAxes ? "c:valAx" : "c:catAx",
                                                                               xyValueAxes ? "valAx" : "catAx", 0));
                    if (primaryYAxisId == 0) chart.setYAxisTitle(axisTitleText(chartText, "c:valAx", "valAx", xyValueAxes ? 1 : 0));
                }
                chart.setAxes(axes);
                chart.setPlots(plots);
                const auto plotAreas = drawingTags(chartText, "c:plotArea", "plotArea");
                if (!plotAreas.empty()) {
                    chart.setPlotAreaLayout(parseChartManualLayout(plotAreas.front()));
                    chart.setDataTable(parseChartDataTable(plotAreas.front()));
                    const auto plotSpPr = plotAreaDirectSpPr(plotAreas.front());
                    if (!plotSpPr.empty()) { chart.setPlotAreaLineFormat(parseChartLineFormat(plotSpPr)); chart.setPlotAreaFillFormat(parseChartFillFormat(plotSpPr)); }
                }
                const auto chartAreaSpPr = chartSpaceDirectSpPr(chartText);
                if (!chartAreaSpPr.empty()) { chart.setChartAreaLineFormat(parseChartLineFormat(chartAreaSpPr)); chart.setChartAreaFillFormat(parseChartFillFormat(chartAreaSpPr)); }
                const auto styleNodes = drawingTags(chartText, "c:style", "style");
                if (!styleNodes.empty()) chart.setStyle(xlpp::internal::attribute(styleNodes.front(), "val"));
                const auto legendNodes = drawingTags(chartText, "c:legend", "legend");
                chart.setShowLegend(!legendNodes.empty());
                if (!legendNodes.empty()) {
                    const auto& legendXml = legendNodes.front();
                    const auto positions = drawingTags(legendXml, "c:legendPos", "legendPos");
                    if (!positions.empty()) chart.setLegendPosition(xlpp::internal::attribute(positions.front(), "val"));
                    xlpp::ChartLegendFormat legendFormat; legendFormat.present = true;
                    legendFormat.overlay = chartBoolValue(legendXml, "c:overlay", "overlay");
                    legendFormat.layout = parseChartManualLayout(legendXml);
                    const auto spPr = drawingTags(legendXml, "c:spPr", "spPr");
                    if (!spPr.empty()) { legendFormat.line = parseChartLineFormat(spPr.front()); legendFormat.fill = parseChartFillFormat(spPr.front()); }
                    chart.setLegendFormat(std::move(legendFormat));
                }
                for (const auto& seriesNode : drawingTags(chartText, "c:ser", "ser")) {
                    xlpp::ChartSeries series;
                    const auto txNodes = drawingTags(seriesNode, "c:tx", "tx");
                    if (!txNodes.empty()) {
                        auto title = drawingTagText(txNodes.front(), "c:v", "v");
                        const auto titleReference = drawingTagText(txNodes.front(), "c:f", "f");
                        auto titleCache = parseChartSeriesCache(txNodes.front());
                        if (title.empty() && titleCache.present && !titleCache.points.empty()) title = titleCache.points.front().value;
                        if (title.empty()) title = titleReference;
                        series.setTitle(title);
                        series.setTitleReference(titleReference);
                        series.setTitleCache(std::move(titleCache));
                    }
                    const auto valNodes = drawingTags(seriesNode, "c:val", "val");
                    if (!valNodes.empty()) { series.setValuesReference(drawingTagText(valNodes.front(), "c:f", "f")); series.setValuesCache(parseChartSeriesCache(valNodes.front())); }
                    else {
                        const auto yValNodes = drawingTags(seriesNode, "c:yVal", "yVal");
                        if (!yValNodes.empty()) { series.setValuesReference(drawingTagText(yValNodes.front(), "c:f", "f")); series.setValuesCache(parseChartSeriesCache(yValNodes.front())); }
                    }
                    const auto catNodes = drawingTags(seriesNode, "c:cat", "cat");
                    if (!catNodes.empty()) { series.setCategoriesReference(drawingTagText(catNodes.front(), "c:f", "f")); series.setCategoriesCache(parseChartSeriesCache(catNodes.front())); }
                    else {
                        const auto xValNodes = drawingTags(seriesNode, "c:xVal", "xVal");
                        if (!xValNodes.empty()) { series.setCategoriesReference(drawingTagText(xValNodes.front(), "c:f", "f")); series.setCategoriesCache(parseChartSeriesCache(xValNodes.front())); }
                    }
                    series.setDataLabels(parseChartDataLabels(seriesNode));
                    series.setTrendlines(parseChartTrendlines(seriesNode));
                    series.setErrorBars(parseChartErrorBars(seriesNode));
                    const auto seriesSpPr = seriesDirectSpPr(seriesNode);
                    if (!seriesSpPr.empty()) {
                        series.setLineFormat(parseChartLineFormat(seriesSpPr));
                        series.setFillFormat(parseChartFillFormat(seriesSpPr));
                    }
                    series.setMarkerFormat(parseChartMarkerFormat(seriesNode));
                    series.setDataPoints(parseChartDataPoints(seriesNode));
                    chart.addSeries(std::move(series));
                }

                const auto anchorInfo = parseChartAnchorInfo(anchorNode, anchorType, frame);
                chart.setAnchorInfo(anchorInfo);
                if (anchorInfo.widthEmu > 0) chart.setWidth(std::max(1, static_cast<int>(std::llround(static_cast<double>(anchorInfo.widthEmu) / 9525.0))));
                if (anchorInfo.heightEmu > 0) chart.setHeight(std::max(1, static_cast<int>(std::llround(static_cast<double>(anchorInfo.heightEmu) / 9525.0))));
                const auto nonVisual = drawingTags(frame, "xdr:cNvPr", "cNvPr");
                std::string objectId;
                std::string objectName = "Chart";
                if (!nonVisual.empty()) {
                    objectId = xlpp::internal::attribute(nonVisual.front(), "id");
                    const auto parsedName = xlpp::internal::attribute(nonVisual.front(), "name");
                    if (!parsedName.empty()) objectName = parsedName;
                }
                chart.setStableId(drawingPart + "#" + (objectId.empty() ? chartRelationshipId : objectId));
                chart.setSourceDrawingPart(drawingPart);
                chart.setSourceChartPart(chartPart);
                chart.setSourceRelationshipId(chartRelationshipId);
                chart.setDrawingObjectName(objectName);
                chart.setImported(true);
                ws.addLoadedChart(std::move(chart));
            }
        };

        loadAnchors("xdr:oneCellAnchor", "oneCellAnchor", xlpp::DrawingAnchorType::OneCell);
        loadAnchors("xdr:twoCellAnchor", "twoCellAnchor", xlpp::DrawingAnchorType::TwoCell);
        loadAnchors("xdr:absoluteAnchor", "absoluteAnchor", xlpp::DrawingAnchorType::Absolute);
    }
}

void parseSheet(xlpp::Worksheet& ws, const std::string& xml, const xlpp::internal::ZipArchive& z, const std::string& target, const StyleCatalog& styleCatalog, const std::vector<xlpp::Style>& dxfStyles, const std::vector<LoadedSharedString>& shared, bool date1904) {
    using namespace xlpp;
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

loadImages(ws, xml, z, target);
loadCharts(ws, xml, z, target);

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
}namespace xlpp {
NamedStyle& Workbook::addNamedStyle(NamedStyle style){if(style.name().empty())throw std::invalid_argument("Named style name cannot be empty");if(namedStyle(style.name()))throw std::invalid_argument("Named style already exists: "+style.name());namedStyles_.push_back(std::move(style));return namedStyles_.back();}
NamedStyle* Workbook::namedStyle(const std::string& name) noexcept{for(auto& style:namedStyles_)if(style.name()==name)return &style;return nullptr;}
const NamedStyle* Workbook::namedStyle(const std::string& name) const noexcept{for(const auto& style:namedStyles_)if(style.name()==name)return &style;return nullptr;}
void Workbook::applyNamedStyle(Cell& cell,const std::string& name) const{const auto* style=namedStyle(name);if(!style)throw std::out_of_range("Unknown named style: "+name);cell.style()=style->style();cell.setNamedStyle(name);}
void Workbook::addVbaProject(const std::filesystem::path& path) {
    std::ifstream stream(path, std::ios::binary);
    if (!stream) throw std::runtime_error("Cannot open VBA project file: " + path.string());
    std::vector<unsigned char> bytes((std::istreambuf_iterator<char>(stream)), std::istreambuf_iterator<char>());
    if (bytes.empty()) throw std::invalid_argument("VBA project file is empty");
    setVbaProject(std::move(bytes));
}

void Workbook::setVbaProject(std::vector<unsigned char> bytes) {
    if (bytes.empty()) throw std::invalid_argument("VBA project bytes cannot be empty");
    removeVbaProject();
    generatedVbaProject_ = false;
    PreservedPart part;
    part.name = "xl/vbaProject.bin";
    part.data.assign(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    part.overrideType = "application/vnd.ms-office.vbaProject";
    part.extension = "bin";
    part.defaultType = "application/vnd.ms-office.vbaProject";
    part.compress = false;
    preservedParts_.push_back(std::move(part));
}

bool Workbook::hasVbaProject() const noexcept {
    return std::any_of(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
}

bool Workbook::removeVbaProject() noexcept {
    const auto oldSize = preservedParts_.size();
    preservedParts_.erase(std::remove_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin"
            || part.name == "xl/vbaProjectSignature.bin"
            || part.name == "xl/_rels/vbaProject.bin.rels";
    }), preservedParts_.end());
    const bool removed = preservedParts_.size() != oldSize;
    if (removed) generatedVbaProject_ = false;
    return removed;
}

std::vector<VbaModule> Workbook::vbaModules() const {
    const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
    if (it == preservedParts_.end()) return {};
    const std::vector<unsigned char> bytes(it->data.begin(), it->data.end());
    return internal::readVbaProjectBinary(bytes);
}

std::optional<std::string> Workbook::vbaModuleText(const std::string& moduleName) const {
    const auto target = [&] {
        std::string value = moduleName;
        std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
        return value;
    }();
    for (const auto& module : vbaModules()) {
        std::string name = module.name;
        std::transform(name.begin(), name.end(), name.begin(), [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
        if (name == target) return module.source;
    }
    return std::nullopt;
}

void Workbook::setVbaModuleText(std::string moduleName, std::string source) {
    internal::validateVbaModuleName(moduleName);
    std::vector<VbaModule> modules;
    if (hasVbaProject()) {
        try { modules = vbaModules(); }
        catch (const std::exception&) {
            throw std::runtime_error("Cannot edit VBA source in an unsupported existing vbaProject.bin; attach a generated project or replace it first");
        }
    }
    modules.erase(std::remove_if(modules.begin(), modules.end(), [](const VbaModule& module) {
        return module.type == VbaModuleType::Document;
    }), modules.end());
    auto sameName = [&](const VbaModule& module) {
        if (module.name.size() != moduleName.size()) return false;
        return std::equal(module.name.begin(), module.name.end(), moduleName.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
    };
    const auto it = std::find_if(modules.begin(), modules.end(), sameName);
    if (it == modules.end()) modules.push_back({std::move(moduleName), internal::normalizeVbaSource(std::move(source)), VbaModuleType::Standard});
    else {
        it->source = internal::normalizeVbaSource(std::move(source));
        it->type = VbaModuleType::Standard;
    }
    setVbaProject(internal::buildVbaProjectBinary(modules, sheets_.size()));
    generatedVbaProject_ = true;
}

bool Workbook::removeVbaModule(const std::string& moduleName) {
    if (!hasVbaProject()) return false;
    auto modules = vbaModules();
    modules.erase(std::remove_if(modules.begin(), modules.end(), [](const VbaModule& module) {
        return module.type == VbaModuleType::Document;
    }), modules.end());
    const auto oldSize = modules.size();
    modules.erase(std::remove_if(modules.begin(), modules.end(), [&](const VbaModule& module) {
        if (module.name.size() != moduleName.size()) return false;
        return std::equal(module.name.begin(), module.name.end(), moduleName.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
    }), modules.end());
    if (modules.size() == oldSize) return false;
    if (modules.empty()) {
        removeVbaProject();
    } else {
        setVbaProject(internal::buildVbaProjectBinary(modules, sheets_.size()));
        generatedVbaProject_ = true;
    }
    return true;
}

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
Worksheet& Workbook::copyWorksheet(const Worksheet& source, std::string newName){if(newName.empty())throw std::invalid_argument("Worksheet name cannot be empty");if(worksheet(newName))throw std::invalid_argument("Duplicate worksheet name");Worksheet copy = source;copy.rename(std::move(newName));sheets_.push_back(std::move(copy));return sheets_.back();}

ChartCacheSyncReport Workbook::synchronizeChartCaches(const ChartCacheSyncOptions& options) {
    ChartCacheSyncReport report;
    const auto cacheEqual = [](const ChartSeriesCache& a, const ChartSeriesCache& b) {
        if (a.present != b.present || a.numeric != b.numeric || a.formatCode != b.formatCode || a.pointCount != b.pointCount || a.points.size() != b.points.size()) return false;
        for (std::size_t i = 0; i < a.points.size(); ++i)
            if (a.points[i].index != b.points[i].index || a.points[i].value != b.points[i].value) return false;
        return true;
    };
    for (auto& sheet : sheets_) {
        for (auto& chart : sheet.charts_) {
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
                    if (kind == ChartCacheKind::String && (parsed.first.row != parsed.last.row || parsed.first.column != parsed.last.column)) {
                        ++report.referencesSkipped; unsupportedSelectedReference = true;
                        report.warnings.push_back(sheet.name() + ": chart series title reference must resolve to one cell: " + reference);
                        return;
                    }
                    auto rebuilt = buildChartCache(parsed, kind, date1904_, existing, &report.warnings);
                    if (!rebuilt.valid(true)) {
                        ++report.referencesSkipped; unsupportedSelectedReference = true;
                        report.warnings.push_back(sheet.name() + ": rebuilt " + label + " cache failed validation for " + reference);
                        return;
                    }
                    if (cacheEqual(rebuilt, existing)) return;
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
                    if (accepted) ++report.cachesUpdated;
                    else report.warnings.push_back(sheet.name() + ": failed to apply rebuilt " + label + " cache for series " + std::to_string(seriesIndex));
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

void Workbook::save(const std::filesystem::path& p) const { save(p, SaveOptions{}); }
void Workbook::save(const std::filesystem::path& p, const SaveOptions& options) const {
    if (sheets_.empty()) throw std::runtime_error("Workbook needs at least one worksheet");
    const bool strict = options.strictNamespace;
    StyleCatalog styleCatalog;
    DxfCatalog dxfCatalog;
    std::size_t tableCount = 0;
    std::size_t commentCount = 0;
    std::unordered_map<std::string, std::size_t> sstIndex;
    std::vector<std::string> sstStrings;
    std::size_t sstOccurrences = 0;

    constexpr auto noSourceSheet = std::numeric_limits<std::size_t>::max();
    std::vector<std::size_t> sourceSheetIndices(sheets_.size(), noSourceSheet);
    std::vector<bool> preserveDrawing(sheets_.size(), false);
    std::vector<bool> preservePivot(sheets_.size(), false);
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        const auto sourceName = std::find(sourceSheetNames_.begin(), sourceSheetNames_.end(), sheets_[i].name());
        if (sourceName == sourceSheetNames_.end()) continue;
        const auto sourceIndex = static_cast<std::size_t>(std::distance(sourceSheetNames_.begin(), sourceName));
        if (sourceIndex >= sourceSheetXml_.size() || sourceIndex >= sourceSheetParts_.size()) continue;
        sourceSheetIndices[i] = sourceIndex;
        preserveDrawing[i] = !sheets_[i].drawingsDirty()
            && !internal::tags(sourceSheetXml_[sourceIndex], "drawing").empty();
        const auto sourceRelationships = relationshipsForSource(preservedRelationships_, sourceSheetParts_[sourceIndex]);
        preservePivot[i] = !internal::tags(sourceSheetXml_[sourceIndex], "pivotTableParts").empty()
            || std::any_of(sourceRelationships.begin(), sourceRelationships.end(), [](const auto& relationship) {
                return relationshipKind(relationship) == "pivotTable";
            });
    }

    const auto firstDrawingId = nextAvailablePartId(preservedParts_, "xl/drawings/drawing", ".xml");
    const auto firstChartId = nextAvailablePartId(preservedParts_, "xl/charts/chart", ".xml");
    const auto firstPivotId = nextAvailablePartId(preservedParts_, "xl/pivotTables/pivotTable", ".xml");
    std::vector<std::size_t> generatedDrawingIds;
    std::vector<std::size_t> generatedChartIds;
    std::vector<std::size_t> generatedPivotIds;
    std::unordered_map<std::size_t, std::size_t> generatedPivotCacheIds;
    std::size_t nextDrawingId = firstDrawingId;
    std::size_t nextChartId = firstChartId;
    std::size_t nextPivotId = firstPivotId;
    std::size_t nextPivotCacheId = nextAvailablePivotCacheId(sourceWorkbookXml_);

    for (std::size_t sheetIndex = 0; sheetIndex < sheets_.size(); ++sheetIndex) {
        const auto& sheet = sheets_[sheetIndex];
        tableCount += sheet.tables().size();
        bool hasComments = false;
        for (const auto& pair : sheet.cells()) if (pair.second.hasComment()) { hasComments = true; break; }
        if (hasComments) ++commentCount;
        if (!preserveDrawing[sheetIndex] && (!sheet.images().empty() || sheet.chartCount() > 0)) {
            generatedDrawingIds.push_back(nextDrawingId++);
            for (std::size_t chartIndex = 0; chartIndex < sheet.chartCount(); ++chartIndex)
                generatedChartIds.push_back(nextChartId++);
        } else if (preserveDrawing[sheetIndex] && sheet.appendedChartCount() > 0) {
            for (std::size_t chartIndex = 0; chartIndex < sheet.appendedChartCount(); ++chartIndex)
                generatedChartIds.push_back(nextChartId++);
        }
        for (std::size_t pivotIndex = 0; pivotIndex < sheet.pivotTables().size(); ++pivotIndex) {
            const auto pivotPartId = nextPivotId++;
            generatedPivotIds.push_back(pivotPartId);
            generatedPivotCacheIds.emplace(pivotPartId, nextPivotCacheId++);
        }
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
    const bool macroEnabled = hasVbaProject();
    auto sheetXmls = serializeSheets(sheets_, styleCatalog, dxfCatalog, date1904_, strict, macroEnabled,
                                           options.parallelSheets ? options.parallelWorkers : 0,
                                           options.parallelRows, &sstIndex, &cachedSheetXml_,
                                           cachedSheetXmlStrict_, cachedSheetXmlDate1904_);
    for (auto& sheet : sheets_) sheet.clearDirty();
    internal::ZipArchive z;
    std::set<std::string> suppressedPreservedParts;
    z.setCompressionLevel(zlibLevel(options.compressionLevel));
    z.setCompressionStrategy(zlibStrategy(options.compressionStrategy));
    z.setParallelWorkers(options.parallelWorkers);
    z.add("[Content_Types].xml", contentTypes(sheets_.size(), tableCount, commentCount,
        generatedDrawingIds, preservedParts_, strict, !sstStrings.empty(), generatedChartIds,
        generatedPivotIds, !customProps_.empty(), macroEnabled));
    const auto rootOriginalRelationships = relationshipsForSource(preservedRelationships_, {});
    const auto mergedRootRelationships = mergeRelationshipsXml(rootrels(strict, !customProps_.empty()), rootOriginalRelationships,
        [](const PreservedRelationship& relationship) {
            const auto kind = relationshipKind(relationship);
            return kind != "officeDocument" && kind != "core-properties"
                && kind != "extended-properties" && kind != "custom-properties";
        }, strict, nullptr);
    z.add("_rels/.rels", mergedRootRelationships);
    z.add("docProps/core.xml", corePropertiesXml(properties_, strict));
    z.add("docProps/app.xml", appPropertiesXml(strict));
    if (!customProps_.empty()) z.add("docProps/custom.xml", customPropertiesXml(customProps_));
    z.add("xl/styles.xml", stylesXml(styleCatalog, namedStyles_, dxfCatalog, strict));
    if (!sstStrings.empty()) {
        std::ostringstream sstXml;
        sstXml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"" << nsMain(strict)
               << "\" count=\"" << sstOccurrences << "\" uniqueCount=\"" << sstStrings.size() << "\">";
        for (const auto& text : sstStrings) {
            sstXml << "<si><t xml:space=\"preserve\">"; writeXmlEscaped(sstXml, text); sstXml << "</t></si>";
        }
        sstXml << "</sst>";
        z.add("xl/sharedStrings.xml", sstXml.str());
    }
    std::ostringstream wb, rels;
    wb << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"" << nsMain(strict) << "\" xmlns:r=\"" << nsRelsDoc(strict) << "\">";
    // CT_Workbook requires workbookPr before workbookProtection and sheets.
    wb << "<workbookPr date1904=\"" << (date1904_ ? 1 : 0) << "\"";
    if (macroEnabled) wb << " codeName=\"ThisWorkbook\"";
    wb << "/>";
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
    std::size_t globalDrawingId = firstDrawingId;
    std::size_t globalMediaId = nextAvailableMediaId(preservedParts_);
    std::size_t globalChartId = firstChartId;
    std::size_t globalPivotId = firstPivotId;
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        const auto& sheet = sheets_[i];
        wb << "<sheet name=\"" << xmlEscape(sheet.name()) << "\" sheetId=\"" << i+1 << "\" r:id=\"rId" << i+1 << "\"/>";
        rels << "<Relationship Id=\"rId" << i+1 << "\" Type=\"" << nsRelsDoc(strict) << "/worksheet\" Target=\"worksheets/sheet" << i+1 << ".xml\"/>";
        bool hasLinks=false; for(const auto& pair:sheet.cells()) if(pair.second.hasHyperlink() && pair.second.hyperlinkValue()->external()) hasLinks=true;
        bool hasComments=false; for(const auto& pair:sheet.cells()) if(pair.second.hasComment()) { hasComments=true; break; }
        const bool hasGeneratedImages = !preserveDrawing[i] && !sheet.images().empty();
        const bool hasSheetCharts = !preserveDrawing[i] && sheet.chartCount() > 0;
        const bool hasSheetPivots = !sheet.pivotTables().empty();
        const auto sourceSheetIndex = sourceSheetIndices[i];
        const std::string originalSheetPart = sourceSheetIndex != noSourceSheet
            ? sourceSheetParts_[sourceSheetIndex] : std::string{};
        const std::string originalSheetXml = sourceSheetIndex != noSourceSheet
            ? sourceSheetXml_[sourceSheetIndex] : std::string{};
        const auto originalSheetRelationships = relationshipsForSource(preservedRelationships_, originalSheetPart);
        const bool hasOriginalTableOwner = !extractTagBlocks(originalSheetXml, "tableParts").empty();
        const bool hasOriginalCommentVmlOwner = !extractTagBlocks(originalSheetXml, "legacyDrawing").empty();
        const bool hasOriginalTableRelationship = std::any_of(originalSheetRelationships.begin(), originalSheetRelationships.end(), [](const auto& relationship) {
            return relationshipKind(relationship) == "table";
        });
        const bool hasOriginalCommentsRelationship = std::any_of(originalSheetRelationships.begin(), originalSheetRelationships.end(), [](const auto& relationship) {
            return relationshipKind(relationship) == "comments";
        });
        const bool preserveTables = sheet.tables().empty() && hasOriginalTableOwner && hasOriginalTableRelationship;
        const bool preserveComments = !hasComments && hasOriginalCommentsRelationship;
        const bool hasPreservedSheetRelationships = std::any_of(originalSheetRelationships.begin(), originalSheetRelationships.end(), [&](const auto& relationship) {
            const auto kind = relationshipKind(relationship);
            if (kind == "drawing") return static_cast<bool>(preserveDrawing[i]);
            if (kind == "pivotTable") return static_cast<bool>(preservePivot[i]);
            if (kind == "table") return preserveTables;
            if (kind == "comments") return preserveComments;
            if (kind == "vmlDrawing") return preserveComments && hasOriginalCommentVmlOwner;
            return kind != "hyperlink";
        });
        if (!sheet.tables().empty() || hasLinks || hasComments || hasGeneratedImages || hasSheetCharts || hasSheetPivots || hasPreservedSheetRelationships) {
            std::ostringstream sheetRels;
            sheetRels << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
            for (std::size_t t = 0; t < sheet.tables().size(); ++t, ++globalTableId) {
                sheetRels << "<Relationship Id=\"rId" << t+1 << "\" Type=\"" << nsRelsDoc(strict) << "/table\" Target=\"../tables/table" << globalTableId << ".xml\"/>";
                z.add("xl/tables/table"+std::to_string(globalTableId)+".xml", tableXml(sheet.tables()[t], sheet, globalTableId, strict));
            }
            std::size_t hid=1; for(const auto& pair:sheet.cells()) if(pair.second.hasHyperlink() && pair.second.hyperlinkValue()->external()) { const auto& h=*pair.second.hyperlinkValue(); sheetRels<<"<Relationship Id=\"rIdHyperlink"<<hid++<<"\" Type=\""<<nsRelsDoc(strict)<<"/hyperlink\" Target=\""<<xmlEscape(h.target())<<"\" TargetMode=\"External\"/>"; }
            if (hasComments) {
                const auto generatedCommentsPart = "xl/comments" + std::to_string(globalCommentId) + ".xml";
                const auto generatedVmlPart = "xl/drawings/commentsDrawing" + std::to_string(globalCommentId) + ".vml";
                for (const auto& relationship : originalSheetRelationships) {
                    const auto kind = relationshipKind(relationship);
                    if (kind != "comments" && kind != "vmlDrawing") continue;
                    const auto originalPart = internal::RelationshipGraph::resolveTarget(originalSheetPart, relationship.target);
                    const auto replacementPart = kind == "comments" ? generatedCommentsPart : generatedVmlPart;
                    if (!originalPart.empty() && originalPart != replacementPart)
                        suppressedPreservedParts.insert(originalPart);
                }
                sheetRels << "<Relationship Id=\"rIdComments\" Type=\"" << nsRelsDoc(strict) << "/comments\" Target=\"../comments" << globalCommentId << ".xml\"/>";
                sheetRels << "<Relationship Id=\"rIdCommentsVml\" Type=\"" << nsRelsDoc(strict) << "/vmlDrawing\" Target=\"../drawings/commentsDrawing" << globalCommentId << ".vml\"/>";
                z.add("xl/comments" + std::to_string(globalCommentId) + ".xml", commentsXml(sheet, strict));
                z.add("xl/drawings/commentsDrawing" + std::to_string(globalCommentId) + ".vml", commentsVml(sheet));
                ++globalCommentId;
            }
            if (hasGeneratedImages || hasSheetCharts) {
                sheetRels << "<Relationship Id=\"rIdDrawing\" Type=\"" << nsRelsDoc(strict) << "/drawing\" Target=\"../drawings/drawing" << globalDrawingId << ".xml\"/>";
                const auto firstChartId = globalChartId;
                for (std::size_t ci = 0; ci < sheet.chartCount(); ++ci, ++globalChartId)
                    z.add("xl/charts/chart" + std::to_string(globalChartId) + ".xml", chartXml(sheet.chart(ci), strict));
                z.add("xl/drawings/drawing" + std::to_string(globalDrawingId) + ".xml", drawingXml(sheet, strict));
                z.add("xl/drawings/_rels/drawing" + std::to_string(globalDrawingId) + ".xml.rels", drawingRelationshipsXml(sheet, globalMediaId, firstChartId, strict));
                for (const auto& image : sheet.images()) {
                    const std::string bytes(reinterpret_cast<const char*>(image.bytes().data()), image.bytes().size());
                    z.add("xl/media/image" + std::to_string(globalMediaId++) + "." + image.extension(), bytes, false);
                }
                ++globalDrawingId;
            }
            if (preserveDrawing[i] && (sheet.appendedImageCount() > 0 ||
                !internal::WorkbookDrawingAccess::imageEdits(sheet).empty())) {
                if (!applyImageChangesToPreservedDrawing(z, sheet, originalSheetPart, originalSheetXml, preservedRelationships_, preservedParts_,
                                                         globalMediaId, suppressedPreservedParts))
                    throw std::runtime_error("Cannot apply selective image mutation to preserved drawing for worksheet: " + sheet.name());
            }
            if (preserveDrawing[i] && (sheet.appendedChartCount() > 0 ||
                !internal::WorkbookDrawingAccess::chartEdits(sheet).empty())) {
                if (!applyChartChangesToPreservedDrawing(z, sheet, originalSheetPart, originalSheetXml,
                                                         preservedRelationships_, preservedParts_, globalChartId,
                                                         suppressedPreservedParts))
                    throw std::runtime_error("Cannot apply selective chart mutation to preserved drawing for worksheet: " + sheet.name());
            }
            if (hasSheetPivots) {
                for (std::size_t pi = 0; pi < sheet.pivotTables().size(); ++pi, ++globalPivotId) {
                    sheetRels << "<Relationship Id=\"rIdPivot" << (pi + 1) << "\" Type=\"" << nsRelsDoc(strict) << "/pivotTable\" Target=\"../pivotTables/pivotTable" << globalPivotId << ".xml\"/>";

                    const auto cacheId = generatedPivotCacheIds.at(globalPivotId);
                    const auto effectivePivot = effectivePivotTable(sheet.pivotTables()[pi], sheets_, sheet, cacheId);
                    z.add("xl/pivotTables/pivotTable" + std::to_string(globalPivotId) + ".xml", pivotTableXml(effectivePivot, cacheId, strict));
                    z.add("xl/pivotTables/_rels/pivotTable" + std::to_string(globalPivotId) + ".xml.rels",
                          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + nsRelsPkg(strict) + "\"><Relationship Id=\"rId1\" Type=\"" + nsRelsDoc(strict) + "/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition" + std::to_string(globalPivotId) + ".xml\"/></Relationships>");
                    z.add("xl/pivotCache/pivotCacheDefinition" + std::to_string(globalPivotId) + ".xml", pivotCacheXml(effectivePivot, strict));
                    z.add("xl/pivotCache/pivotCacheRecords" + std::to_string(globalPivotId) + ".xml", pivotCacheRecordsXml(effectivePivot, strict));
                    z.add("xl/pivotCache/_rels/pivotCacheDefinition" + std::to_string(globalPivotId) + ".xml.rels",
                          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + nsRelsPkg(strict) + "\"><Relationship Id=\"rId1\" Type=\"" + nsRelsDoc(strict) + "/pivotCacheRecords\" Target=\"pivotCacheRecords" + std::to_string(globalPivotId) + ".xml\"/></Relationships>");
                }
            }
            sheetRels << "</Relationships>";
            auto generatedSheetXml = sheetXmls[i];
            const auto mergedSheetRelationships = mergeRelationshipsXml(sheetRels.str(), originalSheetRelationships,
                [&](const PreservedRelationship& relationship) {
                    const auto kind = relationshipKind(relationship);
                    if (kind == "drawing") return static_cast<bool>(preserveDrawing[i]);
                    if (kind == "pivotTable") return static_cast<bool>(preservePivot[i]);
                    if (kind == "table") return preserveTables;
                    if (kind == "comments") return preserveComments;
                    if (kind == "vmlDrawing") return preserveComments && hasOriginalCommentVmlOwner;
                    return kind != "hyperlink";
                }, strict, &generatedSheetXml);
            generatedSheetXml = rebuildWorksheetTail(std::move(generatedSheetXml), originalSheetXml,
                                                      preserveDrawing[i], preservePivot[i], preserveTables, preserveComments);
            z.add("xl/worksheets/sheet"+std::to_string(i+1)+".xml", generatedSheetXml);
            z.add("xl/worksheets/_rels/sheet"+std::to_string(i+1)+".xml.rels", mergedSheetRelationships);
        } else {
            auto generatedSheetXml = rebuildWorksheetTail(sheetXmls[i], originalSheetXml,
                                                          preserveDrawing[i], preservePivot[i], preserveTables, preserveComments);
            z.add("xl/worksheets/sheet"+std::to_string(i+1)+".xml", generatedSheetXml);
        }
    }
    wb << "</sheets>";
    std::size_t nextWorkbookRelId = sheets_.size() + 1;
    std::ostringstream generatedPivotCaches;
    if (!generatedPivotIds.empty()) {
        generatedPivotCaches << "<pivotCaches>";
        for (const auto pivotPartId : generatedPivotIds) {
            const auto cacheId = generatedPivotCacheIds.at(pivotPartId);
            generatedPivotCaches << "<pivotCache cacheId=\"" << cacheId << "\" r:id=\"rId" << nextWorkbookRelId << "\"/>";
            rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition" << pivotPartId << ".xml\"/>";
        }
        generatedPivotCaches << "</pivotCaches>";
    }
    if (!sstStrings.empty())
        rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/sharedStrings\" Target=\"sharedStrings.xml\"/>";
    rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/styles\" Target=\"styles.xml\"/>";
    if (macroEnabled)
        rels << "<Relationship Id=\"rId" << nextWorkbookRelId++ << "\" Type=\"" << nsRelsDoc(strict) << "/vbaProject\" Target=\"vbaProject.bin\"/>";
    bool hasPrintDefinedNames = false;
    for (const auto& sheet : sheets_) {
        if (!sheet.printArea().empty() || !sheet.printTitlesRows().empty() || !sheet.printTitlesCols().empty()) {
            hasPrintDefinedNames = true;
            break;
        }
    }
    if (!definedNames_.empty() || hasPrintDefinedNames) {
        wb << "<definedNames>";
        for (const auto& item : definedNames_) {
            if (item.name() == "_xlnm.Print_Area" || item.name() == "_xlnm.Print_Titles") continue;
            wb << "<definedName name=\"" << xmlEscape(item.name()) << "\"";
            if (item.localSheetId()) wb << " localSheetId=\"" << *item.localSheetId() << "\"";
            if (item.hidden()) wb << " hidden=\"1\"";
            if (!item.comment().empty()) wb << " comment=\"" << xmlEscape(item.comment()) << "\"";
            wb << ">" << xmlEscape(item.value()) << "</definedName>";
        }
        for (std::size_t sheetIndex = 0; sheetIndex < sheets_.size(); ++sheetIndex) {
            const auto& sheet = sheets_[sheetIndex];
            if (!sheet.printArea().empty()) {
                wb << "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"" << sheetIndex << "\">"
                   << xmlEscape(qualifiedPrintReference(sheet.name(), sheet.printArea())) << "</definedName>";
            }
            if (!sheet.printTitlesRows().empty() || !sheet.printTitlesCols().empty()) {
                std::string titles;
                if (!sheet.printTitlesCols().empty()) titles = qualifiedPrintReference(sheet.name(), sheet.printTitlesCols());
                if (!sheet.printTitlesRows().empty()) {
                    if (!titles.empty()) titles += ',';
                    titles += qualifiedPrintReference(sheet.name(), sheet.printTitlesRows());
                }
                wb << "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"" << sheetIndex << "\">"
                   << xmlEscape(titles) << "</definedName>";
            }
        }
        wb << "</definedNames>";
    }
    {
        const auto& cp = calcProps_;
        if (cp.calcMode() != "auto" || cp.calcId() != 191029 || !cp.fullPrecision() || cp.iterate()
            || cp.fullCalcOnLoad() || cp.calcOnSave()) {
            wb << "<calcPr calcId=\"" << cp.calcId() << "\" calcMode=\"" << xmlEscape(cp.calcMode()) << "\""
               << " fullPrecision=\"" << (cp.fullPrecision() ? 1 : 0) << "\""
               << " iterate=\"" << (cp.iterate() ? 1 : 0) << "\""
               << " iterateCount=\"" << cp.iterateCount() << "\" iterateDelta=\"" << cp.iterateDelta() << "\"";
            if (cp.fullCalcOnLoad()) wb << " fullCalcOnLoad=\"1\"";
            if (cp.calcOnSave()) wb << " calcOnSave=\"1\"";
            wb << "/>";
        }
    }
    wb << generatedPivotCaches.str();
    rels << "</Relationships>";
    wb << "</workbook>";
    const bool preserveWorkbookPivotCaches = !internal::tags(sourceWorkbookXml_, "pivotCaches").empty();
    auto workbookXml = wb.str();
    const auto workbookOriginalRelationships = relationshipsForSource(preservedRelationships_, "xl/workbook.xml");
    const auto mergedWorkbookRelationships = mergeRelationshipsXml(rels.str(), workbookOriginalRelationships,
        [&](const PreservedRelationship& relationship) {
            const auto kind = relationshipKind(relationship);
            if (kind == "pivotCacheDefinition") return preserveWorkbookPivotCaches;
            return kind != "worksheet" && kind != "styles" && kind != "sharedStrings" && kind != "vbaProject";
        }, strict, &workbookXml);
    const auto rewrittenGeneratedPivotCaches = joinBlocks(extractTagBlocks(workbookXml, "pivotCaches"));
    eraseTagBlocks(workbookXml, "pivotCaches");
    workbookXml = preserveWorkbookNodes(std::move(workbookXml), sourceWorkbookXml_, false);
    insertBefore(workbookXml, "</workbook>", mergeWorkbookPivotCaches(sourceWorkbookXml_, rewrittenGeneratedPivotCaches));
    z.add("xl/workbook.xml", workbookXml);
    z.add("xl/_rels/workbook.xml.rels", mergedWorkbookRelationships);

    // A source-generated VBA project contains one document module per worksheet.
    // Rebuild it at save time so worksheets added or removed after
    // setVbaModuleText() cannot leave vbaProject.bin out of sync with sheetPr
    // codeName attributes. Externally attached projects remain byte-for-byte
    // preserved because their document-module mapping is owned by the caller.
    std::string generatedVbaData;
    if (generatedVbaProject_) {
        auto modules = vbaModules();
        modules.erase(std::remove_if(modules.begin(), modules.end(), [](const VbaModule& module) {
            return module.type == VbaModuleType::Document;
        }), modules.end());
        const auto bytes = internal::buildVbaProjectBinary(modules, sheets_.size());
        generatedVbaData.assign(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    }
    for (const auto& part : preservedParts_) {
        if (suppressedPreservedParts.contains(part.name)) continue;
        if (generatedVbaProject_ && part.name == "xl/vbaProject.bin")
            z.addUnique(part.name, generatedVbaData, false);
        else
            z.addUnique(part.name, part.data, part.compress);
    }
    if (!suppressedPreservedParts.empty() && z.contains("[Content_Types].xml")) {
        auto contentTypesXml = z.get("[Content_Types].xml");
        for (const auto& overrideNode : internal::tags(contentTypesXml, "Override")) {
            auto partName = internal::attribute(overrideNode, "PartName");
            if (!partName.empty() && partName.front() == '/') partName.erase(partName.begin());
            if (!suppressedPreservedParts.contains(partName)) continue;
            const auto position = contentTypesXml.find(overrideNode);
            if (position != std::string::npos) contentTypesXml.erase(position, overrideNode.size());
        }
        z.replace("[Content_Types].xml", std::move(contentTypesXml));
    }
    z.save(p);
}
void Workbook::load(const std::filesystem::path& p) { load(p, LoadOptions{}); }
void Workbook::load(const std::filesystem::path& p, const LoadOptions& options) { clear(); diagnostics_ = LoadDiagnostics{}; internal::ZipOpenLimits limits; limits.maxEntries = options.maxEntries; limits.maxEntryBytes = options.maxEntryBytes; limits.maxTotalBytes = options.maxTotalBytes; limits.maxFileBytes = options.maxFileBytes; limits.cancel = options.cancel; limits.progress = options.progress; auto z = internal::ZipArchive::open(p, limits);
    const auto relationshipGraph = internal::RelationshipGraph::fromArchive(z);
    preservedRelationships_ = relationshipGraph.relationships();
    const auto relationshipValidation = relationshipGraph.validate();
    for (const auto& issue : relationshipValidation.relationshipSyntaxErrors)
        diagnostics_.warnings.push_back("Relationship syntax issue: " + issue);
    for (const auto& issue : relationshipValidation.duplicateRelationshipIds)
        diagnostics_.warnings.push_back("Duplicate relationship: " + issue);
    for (const auto& issue : relationshipValidation.danglingRelationships)
        diagnostics_.warnings.push_back("Dangling relationship: " + issue);
    for (const auto& issue : relationshipValidation.orphanedParts)
        diagnostics_.warnings.push_back("Orphaned package part: " + issue);
    for (const auto& issue : relationshipValidation.contentTypeErrors)
        diagnostics_.warnings.push_back("Content-type issue: " + issue);
    for (const auto& issue : relationshipValidation.ownerReferenceErrors)
        diagnostics_.warnings.push_back("Broken owner reference: " + issue);
    if(z.contains("docProps/core.xml")){auto cp=z.get("docProps/core.xml");properties_.setTitle(internal::tagText(cp,"dc:title"));properties_.setSubject(internal::tagText(cp,"dc:subject"));properties_.setCreator(internal::tagText(cp,"dc:creator"));properties_.setDescription(internal::tagText(cp,"dc:description"));properties_.setKeywords(internal::tagText(cp,"cp:keywords"));properties_.setCategory(internal::tagText(cp,"cp:category"));properties_.setLastModifiedBy(internal::tagText(cp,"cp:lastModifiedBy"));}if(z.contains("docProps/custom.xml")){auto cust=z.get("docProps/custom.xml");for(const auto& p:internal::tags(cust,"property")){const auto n=internal::attribute(p,"name");if(n.empty())continue;const auto vtText=internal::tagText(p,"vt:lpwstr");if(!vtText.empty()){customProps_.add(CustomProperty(std::string(n),vtText));continue;}if(const auto i4Text=internal::tagText(p,"vt:i4");!i4Text.empty()){customProps_.add(CustomProperty(std::string(n),std::stoi(i4Text)));continue;}if(const auto r8Text=internal::tagText(p,"vt:r8");!r8Text.empty()){customProps_.add(CustomProperty(std::string(n),std::stod(r8Text)));continue;}if(const auto boolText=internal::tagText(p,"vt:bool");!boolText.empty()){customProps_.add(CustomProperty(std::string(n),boolText=="true"));continue;}customProps_.add(CustomProperty(std::string(n),vtText));}}StyleCatalog styleCatalog;std::vector<Style> dxfStyles;if(z.contains("xl/styles.xml")){const auto stylesText=z.get("xl/styles.xml");styleCatalog=parseStyleCatalog(stylesText);dxfStyles=parseDifferentialStyles(stylesText);for(const auto& node:internal::tags(stylesText,"cellStyle")){const auto name=internal::attribute(node,"name");if(name.empty()||name=="Normal")continue;const auto xf=internal::attribute(node,"xfId");if(xf.empty())continue;const auto id=static_cast<std::size_t>(std::stoul(xf));if(id<styleCatalog.items.size())namedStyles_.emplace_back(name,styleCatalog.items[id]);}}std::vector<LoadedSharedString> shared;if(z.contains("xl/sharedStrings.xml")){const auto sstXml = z.get("xl/sharedStrings.xml"); for(auto&si:internal::tags(sstXml,"si")){LoadedSharedString item; item.richText=parseRichTextRuns(si); item.plainText=item.richText?item.richText->plainText():internal::tagText(si,"t"); shared.push_back(std::move(item));}}auto wb=z.get("xl/workbook.xml");
sourceWorkbookXml_ = wb;
strictNamespaces_ = wb.find("http://purl.oclc.org/ooxml/spreadsheetml/main") != std::string::npos;
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
struct PendingPrintName { std::string name; std::size_t sheetIndex; std::string value; };
std::vector<PendingPrintName> pendingPrintNames;
for(const auto& node:internal::tags(wb,"definedName")){
    const auto name = internal::attribute(node,"name");
    const auto value = internal::tagText(node,"definedName");
    const auto local=internal::attribute(node,"localSheetId");
    if ((name == "_xlnm.Print_Area" || name == "_xlnm.Print_Titles") && !local.empty()) {
        pendingPrintNames.push_back({name, static_cast<std::size_t>(std::stoul(local)), value});
        continue;
    }
    DefinedName item(name,value);
    if(!local.empty()) item.setLocalSheetId(static_cast<std::size_t>(std::stoul(local)));
    item.setHidden(internal::attribute(node,"hidden")=="1");
    item.setComment(internal::attribute(node,"comment"));
    definedNames_.push_back(std::move(item));
}
auto relxml=z.get("xl/_rels/workbook.xml.rels");
std::unordered_map<std::string,std::string> targets;
for(auto&r:internal::tags(relxml,"Relationship"))
    targets[internal::attribute(r,"Id")]=internal::attribute(r,"Target");
for(auto&s:internal::tags(wb,"sheet")) {
    const auto name = internal::attribute(s,"name");
    try {
        const auto rid = internal::attribute(s,"r:id");
        auto targetIt = targets.find(rid);
        if (targetIt == targets.end())
            throw std::runtime_error("worksheet relationship '" + rid + "' was not found");
        const auto target = internal::RelationshipGraph::resolveTarget("xl/workbook.xml", targetIt->second);
        if (target.empty() || !z.contains(target))
            throw std::runtime_error("worksheet part was not found: " + target);
        auto& ws=addWorksheet(name);
        sourceSheetNames_.push_back(name);
        sourceSheetParts_.push_back(target);
        sourceSheetXml_.push_back(z.get(target));
        parseSheet(ws, sourceSheetXml_.back(), z, target, styleCatalog, dxfStyles, shared, date1904_);
    } catch (const std::exception& e) {
        diagnostics_.errors.push_back("Sheet '" + name + "': " + e.what());
        if (!options.lenient) throw;
        continue;
    }
}
    // Apply built-in print names after all worksheets have been created.
    for (const auto& pending : pendingPrintNames) {
        if (pending.sheetIndex >= sheets_.size()) continue;
        auto& sheet = sheets_[pending.sheetIndex];
        if (pending.name == "_xlnm.Print_Area") {
            sheet.setPrintArea(localPrintReference(pending.value));
            continue;
        }
        std::string rows;
        std::string cols;
        const auto localValue = localPrintReference(pending.value);
        std::size_t areaStart = 0;
        while (areaStart <= localValue.size()) {
            const auto comma = localValue.find(',', areaStart);
            const auto area = localValue.substr(areaStart, comma == std::string::npos ? std::string::npos : comma - areaStart);
            const auto first = area.find_first_not_of(' ');
            if (first != std::string::npos && std::isdigit(static_cast<unsigned char>(area[first]))) {
                if (!rows.empty()) rows += ',';
                rows += area;
            } else if (!area.empty()) {
                if (!cols.empty()) cols += ',';
                cols += area;
            }
            if (comma == std::string::npos) break;
            areaStart = comma + 1;
        }
        sheet.setPrintTitlesRows(rows);
        sheet.setPrintTitlesCols(cols);
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
    // Source-generated XL++ projects carry a stable project ID/profile. Mark
    // them so a later sheet insertion/removal followed by save() rebuilds the
    // document-module list. Arbitrary externally authored projects remain
    // preserved verbatim.
    if (const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
            return part.name == "xl/vbaProject.bin";
        }); it != preservedParts_.end()) {
        const std::vector<unsigned char> bytes(it->data.begin(), it->data.end());
        generatedVbaProject_ = internal::isXlppGeneratedVbaProjectBinary(bytes);
    }
    // After a successful load, mark all sheets clean so the next save can
    // benefit from differential caching.
    for (auto& s : sheets_) s.clearDirty();
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
