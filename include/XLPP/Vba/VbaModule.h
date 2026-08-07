#pragma once

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
};

} // namespace xlpp
