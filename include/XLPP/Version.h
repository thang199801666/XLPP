#pragma once
#include <string_view>

#define XLPP_VERSION_MAJOR 1
#define XLPP_VERSION_MINOR 1
#define XLPP_VERSION_PATCH 2
#define XLPP_VERSION_STRING "1.1.2"
#define XLPP_C_ABI_VERSION 2

namespace xlpp {
inline constexpr int versionMajor = XLPP_VERSION_MAJOR;
inline constexpr int versionMinor = XLPP_VERSION_MINOR;
inline constexpr int versionPatch = XLPP_VERSION_PATCH;
inline constexpr int cAbiVersion = XLPP_C_ABI_VERSION;
inline constexpr std::string_view versionString = XLPP_VERSION_STRING;
}
