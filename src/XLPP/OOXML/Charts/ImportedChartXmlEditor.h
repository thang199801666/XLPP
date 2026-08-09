#pragma once
#include <XLPP/Worksheet/Worksheet.h>
#include <cstddef>
#include <string>

namespace xlpp::internal::ooxml {
bool importedChartEditRequiresXml(const xlpp::Worksheet& sheet, std::size_t editIndex) noexcept;
bool applyImportedChartXmlEdit(std::string& chartXml, const xlpp::Worksheet& sheet, std::size_t editIndex);
} // namespace xlpp::internal::ooxml
