#include "Preservation/PackageClosure.h"
#include "OOXML/Common/PackageRelationships.h"
#include "Package/Opc/RelationshipGraph.h"
#include <algorithm>
#include <unordered_set>
#include <utility>
#include <vector>

namespace xlpp::internal::preservation {

using xlpp::internal::ooxml::resolvePackagePart;

void suppressExclusivePartClosure(const std::string& rootPart,
                                  const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                  std::set<std::string>& suppressedPreservedParts) {
    std::unordered_set<std::string> closure;
    std::vector<std::string> stack{rootPart};
    while (!stack.empty()) {
        auto part = std::move(stack.back());
        stack.pop_back();
        if (part.empty() || !closure.insert(part).second) continue;
        for (const auto& relationship : allRelationships) {
            if (relationship.sourcePart != part || relationship.targetMode == "External") continue;
            const auto target = resolvePackagePart(part, relationship.target);
            if (!target.empty() && !closure.count(target)) stack.push_back(target);
        }
    }

    std::unordered_set<std::string> protectedParts;
    for (const auto& candidate : closure) {
        if (candidate == rootPart) continue;
        const bool externallyReferenced = std::any_of(allRelationships.begin(), allRelationships.end(), [&](const auto& relationship) {
            if (relationship.targetMode == "External" || closure.count(relationship.sourcePart)) return false;
            return resolvePackagePart(relationship.sourcePart, relationship.target) == candidate;
        });
        if (externallyReferenced) protectedParts.insert(candidate);
    }
    std::vector<std::string> protectStack(protectedParts.begin(), protectedParts.end());
    while (!protectStack.empty()) {
        auto part = std::move(protectStack.back());
        protectStack.pop_back();
        for (const auto& relationship : allRelationships) {
            if (relationship.sourcePart != part || relationship.targetMode == "External") continue;
            const auto target = resolvePackagePart(part, relationship.target);
            if (closure.count(target) && protectedParts.insert(target).second) protectStack.push_back(target);
        }
    }
    for (const auto& part : closure) {
        if (protectedParts.count(part)) continue;
        suppressedPreservedParts.insert(part);
        suppressedPreservedParts.insert(xlpp::internal::RelationshipGraph::relationshipsPartForSource(part));
    }
}


} // namespace xlpp::internal::preservation
