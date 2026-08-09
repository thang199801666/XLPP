#include "OOXML/Tables/TableWriter.h"
#include "OOXML/Common/Namespaces.h"
#include "Package/Xml/XmlUtilities.h"
#include <XLPP/Cell/CellReference.h>
#include <algorithm>
#include <sstream>
#include <stdexcept>
#include <variant>
namespace xlpp::internal::ooxml {
using xlpp::internal::xmlEscape;
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


} // namespace xlpp::internal::ooxml
