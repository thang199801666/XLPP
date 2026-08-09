#pragma once
#include <string>
namespace xlpp::internal::preservation {
bool isRegeneratedPart(const std::string& name);
std::string extensionOf(const std::string& name);
}
