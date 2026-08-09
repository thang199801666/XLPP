#pragma once

#include "OOXML/Styles/StyleCodec.h"
#include <XLPP/Worksheet/Worksheet.h>

#include <cstddef>
#include <string>
#include <string_view>
#include <unordered_map>

namespace xlpp::internal::ooxml {

std::string sheetXml(const xlpp::Worksheet& sheet,
                     const StyleCatalog& styles,
                     const DxfCatalog& dxfs,
                     bool date1904,
                     bool strict,
                     const std::unordered_map<std::string, std::size_t>* sstIndex = nullptr,
                     std::size_t rowWorkers = 0,
                     std::string_view vbaCodeName = {});

} // namespace xlpp::internal::ooxml
