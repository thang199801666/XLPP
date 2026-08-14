#include "WorkbookPackageReader.h"

#include "../Packaging/RelationshipGraph.h"
#include "../XML/XmlUtilities.h"

#include <stdexcept>

namespace xlpp::internal {
namespace {

std::string ownerPrefix(std::string_view ownerLabel) {
    return ownerLabel.empty() ? std::string("relationship") : std::string(ownerLabel);
}

} // namespace

bool relationshipTypeEndsWith(std::string_view type, std::string_view suffix) noexcept {
    return type.size() >= suffix.size() &&
           type.compare(type.size() - suffix.size(), suffix.size(), suffix) == 0;
}

PackageRelationshipMap loadPackageRelationships(const ZipArchive& archive,
                                                const std::string& sourcePart) {
    PackageRelationshipMap result;
    const auto relationshipsPart = RelationshipGraph::relationshipsPartForSource(sourcePart);
    if (!archive.contains(relationshipsPart)) return result;

    for (auto relationship : RelationshipGraph::parseRelationshipsXml(sourcePart, archive.get(relationshipsPart))) {
        if (relationship.id.empty())
            throw std::runtime_error("relationship in '" + sourcePart + "' has an empty Id");
        if (relationship.target.empty())
            throw std::runtime_error("relationship '" + relationship.id + "' in '" + sourcePart + "' has an empty Target");
        if (!relationship.targetMode.empty() && relationship.targetMode != "Internal" &&
            relationship.targetMode != "External")
            throw std::runtime_error("relationship '" + relationship.id + "' in '" + sourcePart +
                                     "' has an invalid TargetMode");

        if (relationship.targetMode != "External" &&
            RelationshipGraph::resolveTarget(sourcePart, relationship.target).empty())
            throw std::runtime_error("relationship '" + relationship.id + "' in '" + sourcePart +
                                     "' has an invalid internal Target");

        PackageRelationshipRecord record{std::move(relationship.target), std::move(relationship.type),
                                         std::move(relationship.targetMode)};
        if (!result.emplace(std::move(relationship.id), std::move(record)).second)
            throw std::runtime_error("duplicate relationship Id in '" + sourcePart + "'");
    }
    return result;
}

const PackageRelationshipRecord& requirePackageRelationship(const PackageRelationshipMap& relationships,
                                                            std::string_view id,
                                                            std::string_view typeSuffix,
                                                            std::string_view ownerLabel) {
    const auto label = ownerPrefix(ownerLabel);
    if (id.empty()) throw std::runtime_error(label + " has an empty relationship Id");
    const auto it = relationships.find(std::string(id));
    if (it == relationships.end())
        throw std::runtime_error(label + " references missing relationship '" + std::string(id) + "'");
    if (!typeSuffix.empty() && !relationshipTypeEndsWith(it->second.type, typeSuffix))
        throw std::runtime_error(label + " relationship '" + std::string(id) + "' has the wrong Type");
    return it->second;
}

std::string requireInternalPackageTarget(const ZipArchive& archive,
                                         const std::string& sourcePart,
                                         const PackageRelationshipRecord& relationship,
                                         std::string_view ownerLabel) {
    const auto label = ownerPrefix(ownerLabel);
    if (relationship.targetMode == "External")
        throw std::runtime_error(label + " unexpectedly uses an External relationship");
    const auto resolved = RelationshipGraph::resolveTarget(sourcePart, relationship.target);
    if (resolved.empty()) throw std::runtime_error(label + " has an invalid internal relationship Target");
    if (!archive.contains(resolved)) throw std::runtime_error(label + " target part is missing: " + resolved);
    return resolved;
}

ContentTypeCatalog loadContentTypeCatalog(const ZipArchive& archive) {
    ContentTypeCatalog result;
    if (!archive.contains("[Content_Types].xml")) return result;

    const auto& xml = archive.get("[Content_Types].xml");
    tagsForEach(xml, "Override", [&](std::string_view node) {
        const auto part = attribute(node, "PartName");
        const auto contentType = attribute(node, "ContentType");
        if (part.empty() || part.front() != '/' || contentType.empty())
            throw std::runtime_error("malformed Override declaration in [Content_Types].xml");
        if (!result.overrides.emplace(part, contentType).second)
            throw std::runtime_error("duplicate Override PartName in [Content_Types].xml: " + part);
    });
    tagsForEach(xml, "Default", [&](std::string_view node) {
        const auto extension = attribute(node, "Extension");
        const auto contentType = attribute(node, "ContentType");
        if (extension.empty() || contentType.empty())
            throw std::runtime_error("malformed Default declaration in [Content_Types].xml");
        if (!result.defaults.emplace(extension, contentType).second)
            throw std::runtime_error("duplicate Default Extension in [Content_Types].xml: " + extension);
    });
    return result;
}

} // namespace xlpp::internal
