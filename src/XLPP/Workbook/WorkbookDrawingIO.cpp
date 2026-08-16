#include "WorkbookDrawingIO.h"
#include "WorkbookNamespaces.h"
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Worksheet/Tables/Table.h>
#include <XLPP/Cell/Cell.h>
#include <XLPP/Cell/CellReference.h>
#include "../XML/XmlUtilities.h"
#include <algorithm>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>

namespace xlpp {
namespace internal {

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
        xml << "<comment ref=\"" << pair.second.address() << "\" authorId=\"" << authorId << "\" shapeId=\"0\""
            << (comment.visible() ? " visible=\"1\"" : "") << "><text><t xml:space=\"preserve\">"
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
        const auto& comment = *pair.second.commentValue();
        const auto widthPx = comment.width() > 0.0 ? comment.width() * (4.0 / 3.0) : 144.0;
        const auto heightPx = comment.height() > 0.0 ? comment.height() * (4.0 / 3.0) : 79.0;
        xml << "<v:shape id=\"_x0000_s" << shapeId++ << "\" type=\"#_x0000_t202\" style=\"position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:" << widthPx << "px;height:" << heightPx << "px;z-index:1;visibility:"
            << (comment.visible() ? "visible" : "hidden") << "\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>"
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
            << "<xdr:blipFill><a:blip r:embed=\"rIdImage" << imageIndex + 1 << "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm";
        if (image.rotation() != 0.0) xml << " rot=\"" << static_cast<long long>(image.rotation() * 60000.0) << "\"";
        if (image.flipHorizontal()) xml << " flipH=\"1\"";
        if (image.flipVertical()) xml << " flipV=\"1\"";
        xml << "><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" << cx << "\" cy=\"" << cy << "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";
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
        for (const auto& column : table.columns()) {
            xml << "<tableColumn id=\"" << column.id() << "\" name=\"" << xmlEscape(column.name()) << "\"";
            if (!column.totalsRowFunction().empty()) xml << " totalsRowFunction=\"" << xmlEscape(column.totalsRowFunction()) << "\"";
            if (!column.totalsRowLabel().empty()) xml << " totalsRowLabel=\"" << xmlEscape(column.totalsRowLabel()) << "\"";
            if (!column.totalsRowFormula().empty())
                xml << "><calculatedColumnFormula>" << xmlEscape(column.totalsRowFormula()) << "</calculatedColumnFormula></tableColumn>";
            else
                xml << "/>";
        }
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

} // namespace internal
} // namespace xlpp
