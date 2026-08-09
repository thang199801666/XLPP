#pragma once
#include "OOXML/Common/RichTextCodec.h"
#include "OOXML/Styles/StyleCodec.h"
#include <XLPP/Styles/Style.h>
#include <string>
#include <vector>

namespace xlpp { class Worksheet; }
namespace xlpp::internal { class ZipArchive; }

namespace xlpp::internal::ooxml {
void loadWorksheetModel(xlpp::Worksheet& worksheet,
                        const std::string& worksheetXml,
                        const xlpp::internal::ZipArchive& archive,
                        const std::string& worksheetPart,
                        const StyleCatalog& styleCatalog,
                        const std::vector<xlpp::Style>& differentialStyles,
                        const std::vector<LoadedSharedString>& sharedStrings,
                        bool date1904);
} // namespace xlpp::internal::ooxml
