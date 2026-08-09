#include "OOXML/Comments/CommentCodec.h"
#include "OOXML/Common/Namespaces.h"
#include "Package/Xml/XmlUtilities.h"
#include <algorithm>
#include <sstream>
namespace xlpp::internal::ooxml {
using xlpp::internal::xmlEscape;
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



} // namespace xlpp::internal::ooxml
