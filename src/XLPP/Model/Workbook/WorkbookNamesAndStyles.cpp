#include <XLPP/Workbook/Workbook.h>

#include <algorithm>
#include <cctype>
#include <stdexcept>
#include <utility>

namespace xlpp {
namespace {

bool asciiCaseEqual(std::string_view lhs, std::string_view rhs) noexcept {
    if (lhs.size() != rhs.size()) return false;
    for (std::size_t index = 0; index < lhs.size(); ++index) {
        const auto left = static_cast<unsigned char>(lhs[index]);
        const auto right = static_cast<unsigned char>(rhs[index]);
        if (left < 0x80u && right < 0x80u) {
            if (std::tolower(left) != std::tolower(right)) return false;
        } else if (left != right) {
            return false;
        }
    }
    return true;
}

} // namespace

NamedStyle& Workbook::addNamedStyle(NamedStyle style) {
    if (style.name().empty()) throw std::invalid_argument("Named style name cannot be empty");
    if (namedStyle(style.name())) throw std::invalid_argument("Named style already exists: " + style.name());
    namedStyles_.push_back(std::move(style));
    return namedStyles_.back();
}

NamedStyle* Workbook::namedStyle(const std::string& name) noexcept {
    const auto it = std::find_if(namedStyles_.begin(), namedStyles_.end(),
                                 [&](const auto& style) { return asciiCaseEqual(style.name(), name); });
    return it == namedStyles_.end() ? nullptr : &*it;
}

const NamedStyle* Workbook::namedStyle(const std::string& name) const noexcept {
    const auto it = std::find_if(namedStyles_.begin(), namedStyles_.end(),
                                 [&](const auto& style) { return asciiCaseEqual(style.name(), name); });
    return it == namedStyles_.end() ? nullptr : &*it;
}

void Workbook::applyNamedStyle(Cell& cell, const std::string& name) const {
    const auto* style = namedStyle(name);
    if (!style) throw std::out_of_range("Unknown named style: " + name);
    cell.style() = style->style();
    cell.setNamedStyle(style->name());
}

DefinedName& Workbook::addDefinedName(DefinedName name) {
    if (name.localSheetId() && *name.localSheetId() >= sheets_.size())
        throw std::out_of_range("Defined-name localSheetId is outside the workbook");
    if (definedName(name.name(), name.localSheetId()))
        throw std::invalid_argument("Defined name already exists in the same scope: " + name.name());
    definedNames_.push_back(std::move(name));
    return definedNames_.back();
}

DefinedName* Workbook::definedName(const std::string& name,
                                   std::optional<std::size_t> localSheetId) noexcept {
    const auto it = std::find_if(definedNames_.begin(), definedNames_.end(), [&](auto& item) {
        return item.localSheetId() == localSheetId && asciiCaseEqual(item.name(), name);
    });
    return it == definedNames_.end() ? nullptr : &*it;
}

const DefinedName* Workbook::definedName(const std::string& name,
                                         std::optional<std::size_t> localSheetId) const noexcept {
    const auto it = std::find_if(definedNames_.begin(), definedNames_.end(), [&](const auto& item) {
        return item.localSheetId() == localSheetId && asciiCaseEqual(item.name(), name);
    });
    return it == definedNames_.end() ? nullptr : &*it;
}

DefinedName* Workbook::definedName(const std::string& name) noexcept {
    if (auto* global = definedName(name, std::nullopt)) return global;
    const auto it = std::find_if(definedNames_.begin(), definedNames_.end(),
                                 [&](auto& item) { return asciiCaseEqual(item.name(), name); });
    return it == definedNames_.end() ? nullptr : &*it;
}

const DefinedName* Workbook::definedName(const std::string& name) const noexcept {
    if (const auto* global = definedName(name, std::nullopt)) return global;
    const auto it = std::find_if(definedNames_.begin(), definedNames_.end(),
                                 [&](const auto& item) { return asciiCaseEqual(item.name(), name); });
    return it == definedNames_.end() ? nullptr : &*it;
}

} // namespace xlpp
