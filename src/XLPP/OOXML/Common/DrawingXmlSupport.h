#pragma once

#include <string>
#include <vector>
#include <XLPP/Worksheet/Drawings/Image.h>

namespace xlpp::internal::ooxml {

std::vector<std::string> drawingTags(const std::string& xml, const char* prefixed, const char* local);
std::string drawingTagText(const std::string& xml, const char* prefixed, const char* local);
long long drawingInteger(const std::string& xml, const char* prefixed, const char* local, long long fallback = 0);
std::string partExtension(const std::string& part);
xlpp::DrawingMarker parseDrawingMarker(const std::string& markerXml);

} // namespace xlpp::internal::ooxml
