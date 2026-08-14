#pragma once

#include "../Packaging/ZipArchive.h"
#include <string>
#include <string_view>
#include <unordered_map>

namespace xlpp::internal {

struct PackageRelationshipRecord {
    std::string target;
    std::string type;
    std::string targetMode;
};

using PackageRelationshipMap = std::unordered_map<std::string, PackageRelationshipRecord>;

// Parse the .rels part owned by sourcePart with strict OPC syntax checks.
// Missing .rels is represented by an empty map; malformed/duplicate records throw.
PackageRelationshipMap loadPackageRelationships(const ZipArchive& archive,
                                                const std::string& sourcePart);

bool relationshipTypeEndsWith(std::string_view type, std::string_view suffix) noexcept;

const PackageRelationshipRecord& requirePackageRelationship(const PackageRelationshipMap& relationships,
                                                            std::string_view id,
                                                            std::string_view typeSuffix,
                                                            std::string_view ownerLabel);

// Resolve and require an Internal relationship target to reference an existing
// package part. External relationships are rejected by this helper.
std::string requireInternalPackageTarget(const ZipArchive& archive,
                                         const std::string& sourcePart,
                                         const PackageRelationshipRecord& relationship,
                                         std::string_view ownerLabel);

struct ContentTypeCatalog {
    std::unordered_map<std::string, std::string> overrides;
    std::unordered_map<std::string, std::string> defaults;
};

// Parse [Content_Types].xml when present. Duplicate or incomplete declarations
// are rejected instead of silently applying last-wins semantics.
ContentTypeCatalog loadContentTypeCatalog(const ZipArchive& archive);

} // namespace xlpp::internal
