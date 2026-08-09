#pragma once

#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

enum class EnterpriseFeatureKind {
    PivotChart,
    Slicer,
    SlicerCache,
    Timeline,
    TimelineCache,
    OlapPivotCache,
    DataModel,
    PowerQuery,
    SmartArt,
    ActiveX,
    VbaUserForm
};

struct EnterpriseFeatureRelationship {
    std::string sourcePart;
    std::string id;
    std::string type;
    std::string targetMode;
    std::string targetPart;
    bool outgoing{false};
};

struct EnterpriseFeatureInfo {
    EnterpriseFeatureKind kind{EnterpriseFeatureKind::DataModel};
    std::string partName;
    std::string contentType;
    std::string name;
    std::string sourceName;
    std::string connectionId;
    std::string cacheId;
    std::vector<std::string> referencedPivotTables;
    std::vector<EnterpriseFeatureRelationship> relationships;
    bool binary{false};
    bool semanticEditable{false};
    bool hasRefreshOnLoad{false};
    bool refreshOnLoad{false};
};

struct EnterpriseFeatureInspection {
    std::vector<EnterpriseFeatureInfo> features;
    std::vector<std::string> warnings;

    std::size_t count(EnterpriseFeatureKind kind) const noexcept;
    bool has(EnterpriseFeatureKind kind) const noexcept { return count(kind) != 0; }
};

struct EnterpriseEditReport {
    std::size_t matched{0};
    std::size_t modified{0};
    std::vector<std::string> warnings;

    bool success() const noexcept { return matched != 0 && warnings.empty(); }
};

} // namespace xlpp
