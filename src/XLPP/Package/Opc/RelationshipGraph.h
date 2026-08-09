#pragma once

#include "Package/Zip/ZipArchive.h"
#include <XLPP/Package/Preservation.h>
#include <cstddef>
#include <string>
#include <unordered_set>
#include <vector>

namespace xlpp::internal {

struct RelationshipValidationReport {
    std::vector<std::string> relationshipSyntaxErrors;
    std::vector<std::string> duplicateRelationshipIds;
    std::vector<std::string> danglingRelationships;
    std::vector<std::string> orphanedParts;
    std::vector<std::string> contentTypeErrors;
    std::vector<std::string> ownerReferenceErrors;

    bool ok() const noexcept {
        return relationshipSyntaxErrors.empty() && duplicateRelationshipIds.empty() && danglingRelationships.empty()
            && orphanedParts.empty() && contentTypeErrors.empty()
            && ownerReferenceErrors.empty();
    }
};

// Counts only objects that are connected through their owning XML node and a
// valid relationship chain. A chart XML part that merely remains in the ZIP is
// therefore not counted unless a worksheet drawing still references it.
struct PackageObjectInventory {
    std::size_t worksheets{0};
    std::size_t drawings{0};
    std::size_t images{0};
    std::size_t charts{0};
    std::size_t shapes{0};
    std::size_t textBoxes{0};
    std::size_t connectors{0};
    std::size_t groups{0};
    std::size_t otherDrawingObjects{0};
    std::size_t tables{0};
    std::size_t comments{0};
    std::size_t externalLinks{0};
    std::size_t pivotTables{0};
    std::size_t pivotCaches{0};
};

struct PackageDiffReport {
    std::vector<std::string> addedParts;
    std::vector<std::string> removedParts;
    std::vector<std::string> changedParts;
    RelationshipValidationReport beforeValidation;
    RelationshipValidationReport afterValidation;
    PackageObjectInventory beforeObjects;
    PackageObjectInventory afterObjects;
    std::vector<std::string> objectCountRegressions;
};

class RelationshipGraph {
public:
    static RelationshipGraph fromArchive(const ZipArchive& archive);

    const std::vector<xlpp::PreservedRelationship>& relationships() const noexcept { return relationships_; }
    std::vector<xlpp::PreservedRelationship> relationshipsFrom(const std::string& sourcePart) const;
    const PackageObjectInventory& objectInventory() const noexcept { return objectInventory_; }
    RelationshipValidationReport validate() const;

    static std::string sourcePartForRelationshipsPart(const std::string& relationshipsPart);
    static std::string relationshipsPartForSource(const std::string& sourcePart);
    static std::string resolveTarget(const std::string& sourcePart, const std::string& target);
    static std::vector<xlpp::PreservedRelationship> parseRelationshipsXml(
        const std::string& sourcePart, const std::string& xml);
    static std::string serializeRelationships(
        const std::vector<xlpp::PreservedRelationship>& relationships, bool strictNamespace);

private:
    std::vector<xlpp::PreservedRelationship> relationships_;
    std::unordered_set<std::string> entries_;
    std::string contentTypesXml_;
    PackageObjectInventory objectInventory_;
    std::vector<std::string> ownerReferenceErrors_;
};

PackageDiffReport comparePackages(const ZipArchive& before, const ZipArchive& after);

} // namespace xlpp::internal
