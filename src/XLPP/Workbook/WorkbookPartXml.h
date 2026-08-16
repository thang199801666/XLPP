#pragma once
#include <cstddef>
#include <functional>
#include <set>
#include <string>
#include <vector>

namespace xlpp {
struct PreservedRelationship;
struct PreservedPart;

namespace internal {

// Package-level relationship and XML block editing helpers shared by the
// workbook writer (save path) and the drawing/chart editing modules.

std::string relationshipKind(const xlpp::PreservedRelationship& relationship);
std::vector<xlpp::PreservedRelationship> relationshipsForSource(
    const std::vector<xlpp::PreservedRelationship>& relationships,
    const std::string& sourcePart);
bool sameRelationship(const xlpp::PreservedRelationship& lhs,
                      const xlpp::PreservedRelationship& rhs);
std::string allocateRelationshipId(const std::set<std::string>& used);
void replaceRelationshipReference(std::string& ownerXml,
                                  const std::string& oldId,
                                  const std::string& newId);
std::string mergeRelationshipsXml(
    const std::string& generatedXml,
    const std::vector<xlpp::PreservedRelationship>& original,
    const std::function<bool(const xlpp::PreservedRelationship&)>& preserve,
    bool strict,
    std::string* generatedOwnerXml = nullptr);

std::vector<std::string> extractTagBlocks(const std::string& xml, const std::string& tag);
void eraseTagBlocks(std::string& xml, const std::string& tag);
std::string joinBlocks(const std::vector<std::string>& blocks);
void insertBefore(std::string& xml, const std::string& marker, const std::string& content);

const xlpp::PreservedPart* findPreservedPart(const std::vector<xlpp::PreservedPart>& parts,
                                             const std::string& name);

std::size_t maximumDrawingObjectId(const std::string& drawingXmlText);

} // namespace internal
} // namespace xlpp
