#pragma once

#include <cstdint>
#include <iomanip>
#include <sstream>
#include <stdexcept>
#include <string>
#include <string_view>

namespace xlpp {

// Legacy Excel worksheet/workbook-protection hash used by the OOXML
// `password` and `workbookPassword` attributes. This is protection against
// accidental edits, not workbook encryption.
inline std::string legacyProtectionPasswordHash(std::string_view password) {
    if (password.empty()) return {};
    if (password.size() > 15)
        throw std::invalid_argument("Legacy Excel protection passwords are limited to 15 characters");

    std::uint16_t hash = 0;
    std::size_t index = 1;
    for (const char rawCharacter : password) {
        const auto character = static_cast<unsigned char>(rawCharacter);
        const std::uint32_t shifted = static_cast<std::uint32_t>(character) << index;
        const std::uint16_t rotated = static_cast<std::uint16_t>((shifted & 0x7FFFu) | (shifted >> 15u));
        hash ^= rotated;
        ++index;
    }
    hash ^= static_cast<std::uint16_t>(password.size());
    hash ^= 0xCE4Bu;

    std::ostringstream stream;
    stream << std::uppercase << std::hex << std::setw(4) << std::setfill('0') << hash;
    return stream.str();
}

} // namespace xlpp
