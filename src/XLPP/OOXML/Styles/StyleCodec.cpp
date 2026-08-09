#include "OOXML/Styles/StyleCodec.h"
#include "OOXML/Common/Namespaces.h"
#include "Package/Xml/XmlUtilities.h"

#include <sstream>
#include <unordered_map>

using xlpp::internal::xmlEscape;

namespace xlpp::internal::ooxml {

void writeColor(std::ostringstream& xml, const xlpp::Color& color) {
    if (!color.empty()) xml << "<color rgb=\"" << xmlEscape(color.argb()) << "\"/>";
}

void writeBorderSide(std::ostringstream& xml, const char* name, const xlpp::BorderSide& side) {
    xml << '<' << name;
    if (!side.style().empty()) xml << " style=\"" << xmlEscape(side.style()) << "\"";
    if (side.color().empty()) xml << "/>";
    else { xml << '>'; writeColor(xml, side.color()); xml << "</" << name << '>'; }
}

std::string stylesXml(const StyleCatalog& catalog, const xlpp::StableVector<xlpp::NamedStyle>& namedStyles, const DxfCatalog& dxfs, bool strict) {
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
        if (!numFmtIdText.empty()) {
            const auto id = static_cast<std::size_t>(std::stoul(numFmtIdText));
            style.setNumFmtId(static_cast<int>(id));
            const auto it = formats.find(id);
            style.setNumberFormat(it != formats.end() ? it->second : builtinNumFmt(static_cast<int>(id)));
        }
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



} // namespace xlpp::internal::ooxml
