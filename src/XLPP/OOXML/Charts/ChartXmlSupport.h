#pragma once
#include <string>
namespace xlpp::internal::ooxml {
std::string chartSpaceDirectSpPr(const std::string& chartXmlText);
std::string plotAreaDirectSpPr(const std::string& plotArea);
std::string seriesDirectSpPr(const std::string& seriesXml);
std::string axisDirectSpPr(const std::string& axisXml);
} // namespace xlpp::internal::ooxml
