#pragma once

#include <string>

namespace xlpp { class Worksheet; }
namespace xlpp::internal { class ZipArchive; }

namespace xlpp::internal::ooxml {
void loadWorksheetImages(xlpp::Worksheet& worksheet,
                         const std::string& sheetXml,
                         const xlpp::internal::ZipArchive& archive,
                         const std::string& sheetPart);
} // namespace xlpp::internal::ooxml
