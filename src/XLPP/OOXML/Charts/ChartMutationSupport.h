#pragma once
#include <string>
namespace xlpp::internal::ooxml {
bool patchValAttribute(std::string& node, const std::string& value);
bool patchOrInsertValChild(std::string& container, const char* prefixed, const char* local, const std::string& value, bool insertWhenMissing = true);
void removeDrawingChild(std::string& container, const char* prefixed, const char* local);
bool patchOpeningTagAttribute(std::string& node, const std::string& name, const std::string& value, bool removeWhenEmpty = false);
}
