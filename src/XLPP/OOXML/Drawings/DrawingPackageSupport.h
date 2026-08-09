#pragma once

#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Chart/Chart.h>
#include <XLPP/Worksheet/Drawings/Image.h>
#include <cstddef>
#include <functional>
#include <string>
#include <vector>

namespace xlpp::internal::ooxml::drawing_support {

const xlpp::PreservedPart* findPreservedPart(const std::vector<xlpp::PreservedPart>& parts,
                                             const std::string& name);
std::size_t maximumDrawingObjectId(const std::string& drawingXmlText);
std::string appendedImageAnchorXml(const xlpp::Image& image, const std::string& relationshipId,
                                   std::size_t objectId, bool strict);
std::string appendedChartAnchorXml(const xlpp::Chart& chart, const std::string& relationshipId,
                                   std::size_t objectId, std::size_t placementIndex, bool strict);
bool patchImportedImageAnchor(std::string& drawingXmlText, const std::string& stableId,
                              const std::string& relationshipId, const xlpp::DrawingAnchorInfo& anchorInfo,
                              bool moved, bool resized, bool remove);
bool drawingReferencesRelationship(const std::string& drawingXmlText, const std::string& relationshipId);
bool removeImportedChartAnchor(std::string& drawingXmlText, const std::string& stableId,
                               const std::string& relationshipId);
bool patchImportedChartAnchor(std::string& drawingXmlText, const std::string& stableId,
                              const std::string& relationshipId, const xlpp::DrawingAnchorInfo& anchorInfo,
                              bool moved, bool resized);
bool drawingReferencesChartRelationship(const std::string& drawingXmlText, const std::string& relationshipId);

} // namespace xlpp::internal::ooxml::drawing_support
