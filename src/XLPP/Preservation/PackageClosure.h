#pragma once
#include <XLPP/Package/Preservation.h>
#include <set>
#include <string>
#include <vector>
namespace xlpp::internal::preservation {
void suppressExclusivePartClosure(const std::string& rootPart,
                                  const std::vector<xlpp::PreservedRelationship>& relationships,
                                  std::set<std::string>& suppressedParts);
} // namespace xlpp::internal::preservation
