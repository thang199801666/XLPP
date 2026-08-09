#pragma once
#include <XLPP/Workbook/Workbook.h>
#include <string>
#include <vector>
#include <functional>
#include <set>
namespace xlpp::internal::ooxml {
std::string resolvePackagePart(const std::string& basePart, std::string relativeTarget);
std::string relationshipKind(const xlpp::PreservedRelationship& relationship);
std::vector<xlpp::PreservedRelationship> relationshipsForSource(const std::vector<xlpp::PreservedRelationship>& relationships, const std::string& sourcePart);
std::string allocateRelationshipId(const std::set<std::string>& used);
std::string mergeRelationshipsXml(const std::string& generatedXml,
                                  const std::vector<xlpp::PreservedRelationship>& original,
                                  const std::function<bool(const xlpp::PreservedRelationship&)>& preserve,
                                  bool strict,
                                  std::string* generatedOwnerXml = nullptr);
}
