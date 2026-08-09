#pragma once

#include <cstdint>
#include <string>

namespace xlpp {

enum class VbaModuleType {
    Standard,
    Document,
    Class
};

struct VbaModule {
    std::string name;
    std::string source;
    VbaModuleType type{VbaModuleType::Standard};
    bool readOnly{false};
    bool privateModule{false};
};

// Project-wide VBA properties represented by the PROJECT stream and the
// PROJECTINFORMATION records in the compressed dir stream. XL++ intentionally
// leaves protection/password/signing state opaque and non-destructive.
struct VbaProjectProperties {
    std::string name{"VBAProject"};
    std::string description;
    std::string helpFile;
    std::uint32_t helpContextId{0};
    std::string constants;
};

} // namespace xlpp
