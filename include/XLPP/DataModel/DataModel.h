#pragma once
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

struct DataModelInspection {
    bool present{false};
    bool hasOlapPivotCaches{false};
    std::vector<std::string> modelParts;
    std::vector<std::string> modelRelationships;
    std::vector<std::string> olapPivotCacheParts;
    std::vector<std::string> warnings;
};

} // namespace xlpp
