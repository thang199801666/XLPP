#pragma once
#include "OOXML/Styles/StyleCodec.h"
#include <XLPP/Worksheet/Worksheet.h>
#include <cstddef>
#include <deque>
#include <string>
#include <unordered_map>
#include <vector>
namespace xlpp::internal::ooxml {
std::vector<std::string> serializeSheets(const std::deque<xlpp::Worksheet>& sheets,
                                         const StyleCatalog& styles, const DxfCatalog& dxfs,
                                         bool date1904, bool strict, bool macroEnabled, std::size_t workers,
                                         bool parallelRows,
                                         const std::unordered_map<std::string, std::size_t>* sstIndex,
                                         std::vector<std::string>* cache,
                                         bool& cacheStrict, bool& cacheDate1904);
}
