#include "OOXML/Common/RichTextCodec.h"
#include "Package/Xml/XmlUtilities.h"

using xlpp::internal::writeXmlEscaped;

namespace xlpp::internal::ooxml {

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


} // namespace xlpp::internal::ooxml
