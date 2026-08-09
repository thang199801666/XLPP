#include "OOXML/Drawings/DrawingWriter.h"
#include "OOXML/Common/Namespaces.h"
#include "Package/Xml/XmlUtilities.h"
#include <XLPP/Cell/CellReference.h>
#include <sstream>
namespace xlpp::internal::ooxml {
using xlpp::internal::xmlEscape;
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
            << "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" << widthEmu << "\" cy=\"" << heightEmu << "\"/></xdr:xfrm><a:graphic>";
        if (chart.modern())
            xml << "<a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chart xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" r:id=\"rIdChart" << chartIndex + 1 << "\"/></a:graphicData>";
        else
            xml << "<a:graphicData uri=\"" << chartNs << "\"><c:chart xmlns:c=\"" << chartNs << "\" r:id=\"rIdChart" << chartIndex + 1 << "\"/></a:graphicData>";
        xml << "</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>";
    }
    xml << "</xdr:wsDr>";
    return xml.str();
}
std::string drawingRelationshipsXml(const xlpp::Worksheet& sheet, std::size_t firstMediaId, std::size_t firstChartId, bool strict) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" << nsRelsPkg(strict) << "\">";
    for (std::size_t i = 0; i < sheet.images().size(); ++i)
        xml << "<Relationship Id=\"rIdImage" << i+1 << "\" Type=\"" << nsRelsDoc(strict) << "/image\" Target=\"../media/image" << firstMediaId+i << '.' << sheet.images()[i].extension() << "\"/>";
    for (std::size_t i = 0; i < sheet.chartCount(); ++i) {
        const auto relationshipType = sheet.chart(i).modern()
            ? std::string("http://schemas.microsoft.com/office/2014/relationships/chartEx")
            : nsRelsDoc(strict) + "/chart";
        xml << "<Relationship Id=\"rIdChart" << i+1 << "\" Type=\"" << relationshipType << "\" Target=\"../charts/chart" << firstChartId+i << ".xml\"/>";
    }
    xml << "</Relationships>";
    return xml.str();
}


} // namespace xlpp::internal::ooxml
