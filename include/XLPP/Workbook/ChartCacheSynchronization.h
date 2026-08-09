#pragma once

#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

struct ChartCacheSyncOptions {
    bool synchronizeTitles{true};
    bool synchronizeCategories{true};
    bool synchronizeValues{true};
    bool changedReferencesOnly{false};
    bool clearUnsupportedReferences{false};
};

struct ChartCacheSyncReport {
    std::size_t chartsVisited{0};
    std::size_t seriesVisited{0};
    std::size_t referencesChecked{0};
    std::size_t referencesUnchanged{0};
    std::size_t dependenciesRegistered{0};
    std::size_t dependenciesChanged{0};
    std::size_t cachesUpdated{0};
    std::size_t cachesCleared{0};
    std::size_t referencesSkipped{0};
    std::vector<std::string> warnings;
    bool success() const noexcept { return warnings.empty(); }
};

} // namespace xlpp
