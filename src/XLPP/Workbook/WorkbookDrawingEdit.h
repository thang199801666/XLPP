#pragma once
#include <XLPP/Worksheet/Drawings/Image.h>
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {
class Chart;
class Image;

namespace internal {

// Drawing-part editing helpers used by the workbook writer: appending new
// image/chart anchors, patching imported anchors, and stable-ID based
// relationship-aware edits.

std::string appendedImageAnchorXml(const xlpp::Image& image,
                                   const std::string& relationshipId,
                                   std::size_t objectId,
                                   bool strict);
std::string appendedChartAnchorXml(const xlpp::Chart& chart,
                                   const std::string& relationshipId,
                                   std::size_t objectId,
                                   std::size_t placementIndex,
                                   bool strict);
bool replaceSimpleDrawingText(std::string& xml, const char* prefixed, const char* local, long long value);
bool replaceAttributeInNode(std::string& node, const std::string& attributeName, long long value);
bool patchFirstDrawingNodeAttributes(std::string& anchor,
                                     const char* prefixed,
                                     const char* local,
                                     long long first,
                                     long long second,
                                     const char* firstAttribute,
                                     const char* secondAttribute);
bool patchDrawingMarker(std::string& anchor, const char* prefixed, const char* local,
                        const xlpp::DrawingMarker& marker, bool updateOffsets = false);
std::string drawingObjectIdFromStableId(const std::string& stableId);
bool anchorMatchesStableId(const std::string& anchor, const std::string& stableId);
bool anchorReferencesRelationship(const std::string& anchor, const std::string& relationshipId);
bool drawingReferencesRelationship(const std::string& drawingXmlText, const std::string& relationshipId);
bool patchImportedImageAnchor(std::string& drawingXmlText,
                              const std::string& stableId,
                              const std::string& relationshipId,
                              const xlpp::DrawingAnchorInfo& anchorInfo,
                              bool moved,
                              bool resized,
                              bool remove);
bool chartAnchorMatchesStableId(const std::string& anchor, const std::string& stableId);
bool chartAnchorReferencesRelationship(const std::string& anchor, const std::string& relationshipId);
bool drawingReferencesChartRelationship(const std::string& drawingXmlText, const std::string& relationshipId);
bool removeImportedChartAnchor(std::string& drawingXmlText,
                               const std::string& stableId,
                               const std::string& relationshipId);
bool patchImportedChartAnchor(std::string& drawingXmlText,
                              const std::string& stableId,
                              const std::string& relationshipId,
                              const xlpp::DrawingAnchorInfo& anchorInfo,
                              bool moved,
                              bool resized);
bool replaceSimpleElementText(std::string& xml, const char* prefixed, const char* local, const std::string& value);
void eraseChartCacheBlocks(std::string& xml);
bool patchSeriesReferenceContainer(std::string& seriesXml,
                                   const char* prefixedContainer,
                                   const char* localContainer,
                                   const std::string& reference);

} // namespace internal
} // namespace xlpp
