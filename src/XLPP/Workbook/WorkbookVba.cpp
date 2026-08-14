#include <XLPP/Workbook/Workbook.h>
#include "../Vba/VbaProjectBinary.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <optional>
#include <set>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {
namespace {

std::vector<std::string> ensureWorksheetVbaCodeNames(std::deque<Worksheet>& sheets) {
    auto lower = [](std::string value) {
        std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
        return value;
    };
    std::set<std::string> used;
    std::vector<std::string> result(sheets.size());
    for (std::size_t i = 0; i < sheets.size(); ++i) {
        if (sheets[i].vbaCodeName().empty()) continue;
        internal::validateVbaModuleName(sheets[i].vbaCodeName());
        const auto key = lower(sheets[i].vbaCodeName());
        if (!used.insert(key).second)
            throw std::invalid_argument("Duplicate worksheet VBA code name: " + sheets[i].vbaCodeName());
        result[i] = sheets[i].vbaCodeName();
    }
    std::size_t candidate = 1;
    for (std::size_t i = 0; i < result.size(); ++i) {
        if (!result[i].empty()) continue;
        for (;;) {
            auto name = "Sheet" + std::to_string(candidate++);
            if (used.insert(lower(name)).second) {
                result[i] = std::move(name);
                break;
            }
        }
    }
    for (std::size_t i = 0; i < sheets.size(); ++i)
        if (sheets[i].vbaCodeName().empty()) sheets[i].setVbaCodeName(result[i]);
    return result;
}

} // namespace

namespace {
struct UserFormStreamLayout {
    VbaUserFormInspection inspection;
    std::size_t sectionEnd{0};
    std::size_t dataEnd{0};
    std::size_t extraStart{0};
    std::optional<std::size_t> backColor, foreColor, nextId, booleanProperties;
    std::optional<std::size_t> borderStyle, mousePointer, scrollBars, groupCount;
    std::optional<std::size_t> cycle, specialEffect, borderColor, captionCount;
    std::optional<std::size_t> zoom, pictureAlignment, pictureSizeMode, shapeCookie, drawBuffer;
    std::optional<std::size_t> displayedSize, logicalSize, scrollPosition, captionData;
    std::size_t captionBytes{0};
    bool captionCompressed{false};
};

std::size_t userFormAlign(std::size_t offset, std::size_t alignment) {
    if (alignment <= 1) return offset;
    const auto remainder = offset % alignment;
    return remainder == 0 ? offset : offset + (alignment - remainder);
}

std::uint16_t userFormU16(const std::vector<unsigned char>& bytes, std::size_t offset) {
    if (offset + 2 > bytes.size()) throw std::runtime_error("Truncated MS-OFORMS stream");
    return static_cast<std::uint16_t>(bytes[offset]) | (static_cast<std::uint16_t>(bytes[offset + 1]) << 8);
}
std::uint32_t userFormU32(const std::vector<unsigned char>& bytes, std::size_t offset) {
    if (offset + 4 > bytes.size()) throw std::runtime_error("Truncated MS-OFORMS stream");
    return static_cast<std::uint32_t>(bytes[offset]) | (static_cast<std::uint32_t>(bytes[offset + 1]) << 8)
        | (static_cast<std::uint32_t>(bytes[offset + 2]) << 16) | (static_cast<std::uint32_t>(bytes[offset + 3]) << 24);
}
std::int32_t userFormI32(const std::vector<unsigned char>& bytes, std::size_t offset) {
    return static_cast<std::int32_t>(userFormU32(bytes, offset));
}
std::int16_t userFormI16(const std::vector<unsigned char>& bytes, std::size_t offset) {
    return static_cast<std::int16_t>(userFormU16(bytes, offset));
}
void userFormPutU16(std::vector<unsigned char>& bytes, std::size_t offset, std::uint16_t value) {
    if (offset + 2 > bytes.size()) throw std::runtime_error("Truncated MS-OFORMS stream");
    bytes[offset] = static_cast<unsigned char>(value & 0xFFu);
    bytes[offset + 1] = static_cast<unsigned char>((value >> 8) & 0xFFu);
}
void userFormPutU32(std::vector<unsigned char>& bytes, std::size_t offset, std::uint32_t value) {
    if (offset + 4 > bytes.size()) throw std::runtime_error("Truncated MS-OFORMS stream");
    bytes[offset] = static_cast<unsigned char>(value & 0xFFu);
    bytes[offset + 1] = static_cast<unsigned char>((value >> 8) & 0xFFu);
    bytes[offset + 2] = static_cast<unsigned char>((value >> 16) & 0xFFu);
    bytes[offset + 3] = static_cast<unsigned char>((value >> 24) & 0xFFu);
}
void userFormPutI32(std::vector<unsigned char>& bytes, std::size_t offset, std::int32_t value) {
    userFormPutU32(bytes, offset, static_cast<std::uint32_t>(value));
}
void userFormPutI16(std::vector<unsigned char>& bytes, std::size_t offset, std::int16_t value) {
    userFormPutU16(bytes, offset, static_cast<std::uint16_t>(value));
}

std::vector<std::uint32_t> userFormUtf8Codepoints(const std::string& text) {
    std::vector<std::uint32_t> result;
    for (std::size_t i = 0; i < text.size();) {
        const auto c = static_cast<unsigned char>(text[i]);
        std::uint32_t cp = 0;
        std::size_t count = 0;
        if (c < 0x80) { cp = c; count = 1; }
        else if ((c & 0xE0) == 0xC0) { cp = c & 0x1F; count = 2; }
        else if ((c & 0xF0) == 0xE0) { cp = c & 0x0F; count = 3; }
        else if ((c & 0xF8) == 0xF0) { cp = c & 0x07; count = 4; }
        else throw std::invalid_argument("Invalid UTF-8 in UserForm caption");
        if (i + count > text.size()) throw std::invalid_argument("Truncated UTF-8 in UserForm caption");
        for (std::size_t j = 1; j < count; ++j) {
            const auto cc = static_cast<unsigned char>(text[i + j]);
            if ((cc & 0xC0) != 0x80) throw std::invalid_argument("Invalid UTF-8 continuation in UserForm caption");
            cp = (cp << 6) | (cc & 0x3F);
        }
        if (cp > 0x10FFFF || (cp >= 0xD800 && cp <= 0xDFFF)) throw std::invalid_argument("Invalid Unicode scalar in UserForm caption");
        result.push_back(cp);
        i += count;
    }
    return result;
}

std::string userFormAppendUtf8(std::string out, std::uint32_t cp) {
    if (cp <= 0x7F) out.push_back(static_cast<char>(cp));
    else if (cp <= 0x7FF) {
        out.push_back(static_cast<char>(0xC0 | (cp >> 6)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else if (cp <= 0xFFFF) {
        out.push_back(static_cast<char>(0xE0 | (cp >> 12)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else {
        out.push_back(static_cast<char>(0xF0 | (cp >> 18)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 12) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    }
    return out;
}

std::string decodeUserFormString(const std::vector<unsigned char>& bytes, std::size_t offset,
                                 std::size_t byteCount, bool compressed) {
    if (offset + byteCount > bytes.size()) throw std::runtime_error("Truncated MS-OFORMS string");
    std::string result;
    if (compressed) {
        for (std::size_t i = 0; i < byteCount; ++i) result = userFormAppendUtf8(std::move(result), bytes[offset + i]);
        return result;
    }
    if ((byteCount % 2) != 0) throw std::runtime_error("Uncompressed MS-OFORMS string has odd byte length");
    for (std::size_t i = 0; i < byteCount; i += 2) {
        const auto first = static_cast<std::uint16_t>(bytes[offset + i]) | (static_cast<std::uint16_t>(bytes[offset + i + 1]) << 8);
        std::uint32_t cp = first;
        if (first >= 0xD800 && first <= 0xDBFF) {
            if (i + 3 >= byteCount) throw std::runtime_error("Truncated UTF-16 surrogate in MS-OFORMS string");
            const auto second = static_cast<std::uint16_t>(bytes[offset + i + 2]) | (static_cast<std::uint16_t>(bytes[offset + i + 3]) << 8);
            if (second < 0xDC00 || second > 0xDFFF) throw std::runtime_error("Invalid UTF-16 surrogate in MS-OFORMS string");
            cp = 0x10000u + ((static_cast<std::uint32_t>(first - 0xD800) << 10) | (second - 0xDC00));
            i += 2;
        }
        result = userFormAppendUtf8(std::move(result), cp);
    }
    return result;
}

std::pair<std::vector<unsigned char>, bool> encodeUserFormString(const std::string& text) {
    const auto cps = userFormUtf8Codepoints(text);
    const bool compressed = std::all_of(cps.begin(), cps.end(), [](auto cp) { return cp <= 0xFF; });
    std::vector<unsigned char> bytes;
    if (compressed) {
        bytes.reserve(cps.size());
        for (const auto cp : cps) bytes.push_back(static_cast<unsigned char>(cp));
    } else {
        for (const auto cp : cps) {
            if (cp <= 0xFFFF) {
                bytes.push_back(static_cast<unsigned char>(cp & 0xFF));
                bytes.push_back(static_cast<unsigned char>((cp >> 8) & 0xFF));
            } else {
                const auto adjusted = cp - 0x10000;
                const auto hi = static_cast<std::uint16_t>(0xD800 + (adjusted >> 10));
                const auto lo = static_cast<std::uint16_t>(0xDC00 + (adjusted & 0x3FF));
                bytes.push_back(static_cast<unsigned char>(hi & 0xFF)); bytes.push_back(static_cast<unsigned char>(hi >> 8));
                bytes.push_back(static_cast<unsigned char>(lo & 0xFF)); bytes.push_back(static_cast<unsigned char>(lo >> 8));
            }
        }
    }
    return {std::move(bytes), compressed};
}

UserFormStreamLayout parseUserFormStream(const std::vector<unsigned char>& bytes) {
    UserFormStreamLayout out;
    auto fail = [&](std::string message) {
        out.inspection.valid = false;
        out.inspection.error = std::move(message);
        return out;
    };
    try {
        if (bytes.size() < 8) return fail("UserForm Form stream is shorter than the MS-OFORMS header and PropMask");
        out.inspection.properties.minorVersion = bytes[0];
        out.inspection.properties.majorVersion = bytes[1];
        const auto cbForm = userFormU16(bytes, 2);
        if (cbForm < 4) return fail("UserForm cbForm is smaller than FormPropMask");
        out.sectionEnd = 4u + cbForm;
        if (out.sectionEnd > bytes.size()) return fail("UserForm cbForm extends past the Form stream");
        const auto mask = userFormU32(bytes, 4);
        out.inspection.properties.propertyMask = mask;
        std::size_t pos = 8;
        const auto has = [&](unsigned bit) { return (mask & (1u << bit)) != 0; };
        const auto take = [&](unsigned bit, std::size_t size, std::optional<std::size_t>& target) {
            if (!has(bit)) return;
            pos = userFormAlign(pos, std::min<std::size_t>(size, 4));
            if (pos + size > out.sectionEnd) throw std::runtime_error("UserForm DataBlock exceeds cbForm");
            target = pos;
            pos += size;
        };
        take(1, 4, out.backColor); take(2, 4, out.foreColor); take(3, 4, out.nextId);
        take(6, 4, out.booleanProperties); take(7, 1, out.borderStyle); take(8, 1, out.mousePointer);
        take(9, 1, out.scrollBars); take(13, 4, out.groupCount);
        std::optional<std::size_t> ignoredMouseIcon, ignoredFont, ignoredPicture;
        take(15, 2, ignoredMouseIcon); take(16, 1, out.cycle); take(17, 1, out.specialEffect);
        take(18, 4, out.borderColor); take(19, 4, out.captionCount);
        take(20, 2, ignoredFont); take(21, 2, ignoredPicture); take(22, 4, out.zoom);
        take(23, 1, out.pictureAlignment); // bit 24 (PictureTiling) has no DataBlock payload.
        take(25, 1, out.pictureSizeMode); take(26, 4, out.shapeCookie); take(27, 4, out.drawBuffer);
        out.dataEnd = userFormAlign(pos, 4);
        if (out.dataEnd > out.sectionEnd) return fail("UserForm DataBlock padding exceeds cbForm");
        out.extraStart = out.dataEnd;
        pos = out.extraStart;
        const auto takeExtra = [&](unsigned bit, std::size_t size, std::optional<std::size_t>& target) {
            if (!has(bit)) return;
            if (pos + size > out.sectionEnd) throw std::runtime_error("UserForm ExtraDataBlock exceeds cbForm");
            target = pos;
            pos += size;
        };
        takeExtra(10, 8, out.displayedSize); takeExtra(11, 8, out.logicalSize); takeExtra(12, 8, out.scrollPosition);
        if (has(19)) {
            if (!out.captionCount) throw std::runtime_error("UserForm caption flag has no LengthAndCompression field");
            const auto count = userFormU32(bytes, *out.captionCount);
            out.captionCompressed = (count & 0x80000000u) != 0;
            out.captionBytes = count & 0x7FFFFFFFu;
            if (pos + out.captionBytes > out.sectionEnd) throw std::runtime_error("UserForm caption extends past cbForm");
            out.captionData = pos;
            pos += out.captionBytes;
        }
        auto& props = out.inspection.properties;
        if (out.backColor) props.backColor = userFormU32(bytes, *out.backColor);
        if (out.foreColor) props.foreColor = userFormU32(bytes, *out.foreColor);
        if (out.nextId) props.nextAvailableId = userFormU32(bytes, *out.nextId);
        if (out.booleanProperties) props.booleanProperties = userFormU32(bytes, *out.booleanProperties);
        if (out.borderStyle) props.borderStyle = bytes[*out.borderStyle];
        if (out.mousePointer) props.mousePointer = bytes[*out.mousePointer];
        if (out.scrollBars) props.scrollBars = bytes[*out.scrollBars];
        if (out.groupCount) props.groupCount = userFormI32(bytes, *out.groupCount);
        if (out.cycle) props.cycle = bytes[*out.cycle];
        if (out.specialEffect) props.specialEffect = bytes[*out.specialEffect];
        if (out.borderColor) props.borderColor = userFormU32(bytes, *out.borderColor);
        if (out.captionData) props.caption = decodeUserFormString(bytes, *out.captionData, out.captionBytes, out.captionCompressed);
        if (out.zoom) props.zoom = userFormU32(bytes, *out.zoom);
        if (out.pictureAlignment) props.pictureAlignment = bytes[*out.pictureAlignment];
        if (out.pictureSizeMode) props.pictureSizeMode = bytes[*out.pictureSizeMode];
        if (out.shapeCookie) props.shapeCookie = userFormU32(bytes, *out.shapeCookie);
        if (out.drawBuffer) props.drawBuffer = userFormU32(bytes, *out.drawBuffer);
        if (out.displayedSize) { props.displayedWidth = userFormI32(bytes, *out.displayedSize); props.displayedHeight = userFormI32(bytes, *out.displayedSize + 4); }
        if (out.logicalSize) { props.logicalWidth = userFormI32(bytes, *out.logicalSize); props.logicalHeight = userFormI32(bytes, *out.logicalSize + 4); }
        if (out.scrollPosition) { props.scrollLeft = userFormI32(bytes, *out.scrollPosition); props.scrollTop = userFormI32(bytes, *out.scrollPosition + 4); }
        out.inspection.trailingBytes = bytes.size() - out.sectionEnd;
        out.inspection.valid = true;
        return out;
    } catch (const std::exception& e) {
        return fail(e.what());
    }
}

void requireUserFormField(bool requested, const std::optional<std::size_t>& offset, const char* name) {
    if (requested && !offset) throw std::invalid_argument(std::string("UserForm property is not materialized in FormPropMask: ") + name);
}

struct UserFormSiteStringLayout {
    std::optional<std::size_t> countOffset;
    std::optional<std::size_t> dataOffset;
    std::size_t byteCount{0};
    bool compressed{false};
};

struct UserFormControlSiteLayout {
    VbaUserFormControlSite site;
    std::size_t begin{0};
    std::size_t end{0};
    std::optional<std::size_t> helpContextId, bitFlags, objectStreamSize, tabIndex, clsidCacheIndex, groupId;
    std::optional<std::size_t> position;
    UserFormSiteStringLayout name, tag, controlTipText, runtimeLicenseKey, controlSource, rowSource;
};

struct UserFormSiteDataLayout {
    VbaUserFormControlsInspection inspection;
    std::size_t countOfBytesOffset{0};
    std::vector<UserFormControlSiteLayout> sites;
};

void parseUserFormSiteStringCount(const std::vector<unsigned char>& bytes, std::size_t offset,
                                  UserFormSiteStringLayout& out) {
    const auto count = userFormU32(bytes, offset);
    out.countOffset = offset;
    out.compressed = (count & 0x80000000u) != 0;
    out.byteCount = count & 0x7FFFFFFFu;
}

VbaUserFormControlKind userFormControlKindFromCacheIndex(std::uint16_t index) {
    if (index >= 0x8000u) return VbaUserFormControlKind::CustomClass;
    switch (index) {
        case 7: return VbaUserFormControlKind::Form;
        case 12: return VbaUserFormControlKind::Image;
        case 14: return VbaUserFormControlKind::Frame;
        case 15: return VbaUserFormControlKind::MorphData;
        case 16: return VbaUserFormControlKind::SpinButton;
        case 17: return VbaUserFormControlKind::CommandButton;
        case 18: return VbaUserFormControlKind::TabStrip;
        case 21: return VbaUserFormControlKind::Label;
        case 23: return VbaUserFormControlKind::TextBox;
        case 24: return VbaUserFormControlKind::ListBox;
        case 25: return VbaUserFormControlKind::ComboBox;
        case 26: return VbaUserFormControlKind::CheckBox;
        case 27: return VbaUserFormControlKind::OptionButton;
        case 28: return VbaUserFormControlKind::ToggleButton;
        case 47: return VbaUserFormControlKind::ScrollBar;
        case 57: return VbaUserFormControlKind::MultiPage;
        default: return VbaUserFormControlKind::Unknown;
    }
}

bool parseOneUserFormSite(const std::vector<unsigned char>& bytes, std::size_t begin, std::size_t hardEnd,
                          std::uint8_t depth, std::uint8_t siteType, UserFormControlSiteLayout& out) {
    if (begin + 8 > hardEnd) return false;
    out.begin = begin;
    out.site.depth = depth;
    out.site.siteType = siteType;
    out.site.version = userFormU16(bytes, begin);
    const auto cbSite = userFormU16(bytes, begin + 2);
    if (cbSite < 4) return false;
    out.end = begin + 4u + cbSite;
    if (out.end > hardEnd) return false;
    const auto mask = userFormU32(bytes, begin + 4);
    if ((mask & 0xFFFF8000u) != 0) return false; // reserved SitePropMask bits must be zero
    out.site.propertyMask = mask;
    const auto has = [&](unsigned bit) { return (mask & (1u << bit)) != 0; };
    const auto dataStart = begin + 8;
    std::size_t rel = 0;
    const auto take = [&](unsigned bit, std::size_t size, std::optional<std::size_t>& target) {
        if (!has(bit)) return true;
        rel = userFormAlign(rel, std::min<std::size_t>(size, 4));
        if (dataStart + rel + size > out.end) return false;
        target = dataStart + rel;
        rel += size;
        return true;
    };
    if (has(0)) { rel = userFormAlign(rel, 4); if (dataStart + rel + 4 > out.end) return false; parseUserFormSiteStringCount(bytes, dataStart + rel, out.name); rel += 4; }
    if (has(1)) { rel = userFormAlign(rel, 4); if (dataStart + rel + 4 > out.end) return false; parseUserFormSiteStringCount(bytes, dataStart + rel, out.tag); rel += 4; }
    std::optional<std::size_t> id;
    if (!take(2, 4, id) || !take(3, 4, out.helpContextId) || !take(4, 4, out.bitFlags)
        || !take(5, 4, out.objectStreamSize) || !take(6, 2, out.tabIndex)
        || !take(7, 2, out.clsidCacheIndex)) return false;
    // bit 8 Position lives in ExtraDataBlock.
    if (!take(9, 2, out.groupId)) return false;
    if (has(11)) { rel = userFormAlign(rel, 4); if (dataStart + rel + 4 > out.end) return false; parseUserFormSiteStringCount(bytes, dataStart + rel, out.controlTipText); rel += 4; }
    if (has(12)) { rel = userFormAlign(rel, 4); if (dataStart + rel + 4 > out.end) return false; parseUserFormSiteStringCount(bytes, dataStart + rel, out.runtimeLicenseKey); rel += 4; }
    if (has(13)) { rel = userFormAlign(rel, 4); if (dataStart + rel + 4 > out.end) return false; parseUserFormSiteStringCount(bytes, dataStart + rel, out.controlSource); rel += 4; }
    if (has(14)) { rel = userFormAlign(rel, 4); if (dataStart + rel + 4 > out.end) return false; parseUserFormSiteStringCount(bytes, dataStart + rel, out.rowSource); rel += 4; }
    const auto extraStart = dataStart + userFormAlign(rel, 4);
    if (extraStart > out.end) return false;
    std::size_t extra = extraStart;
    const auto takeString = [&](unsigned bit, UserFormSiteStringLayout& layout, std::optional<std::string>& target) {
        if (!has(bit)) return true;
        if (extra + layout.byteCount > out.end) return false;
        layout.dataOffset = extra;
        target = decodeUserFormString(bytes, extra, layout.byteCount, layout.compressed);
        extra += layout.byteCount;
        return true;
    };
    if (!takeString(0, out.name, out.site.name) || !takeString(1, out.tag, out.site.tag)) return false;
    if (id) out.site.id = userFormI32(bytes, *id);
    if (out.helpContextId) out.site.helpContextId = userFormI32(bytes, *out.helpContextId);
    if (out.bitFlags) out.site.bitFlags = userFormU32(bytes, *out.bitFlags);
    if (out.objectStreamSize) out.site.objectStreamSize = userFormU32(bytes, *out.objectStreamSize);
    if (out.tabIndex) out.site.tabIndex = userFormI16(bytes, *out.tabIndex);
    if (out.clsidCacheIndex) {
        out.site.clsidCacheIndex = userFormU16(bytes, *out.clsidCacheIndex);
        out.site.kind = userFormControlKindFromCacheIndex(*out.site.clsidCacheIndex);
    }
    if (out.groupId) out.site.groupId = userFormU16(bytes, *out.groupId);
    if (has(8)) {
        if (extra + 8 > out.end) return false;
        out.position = extra;
        out.site.top = userFormI32(bytes, extra);
        out.site.left = userFormI32(bytes, extra + 4);
        extra += 8;
    }
    if (!takeString(11, out.controlTipText, out.site.controlTipText)
        || !takeString(12, out.runtimeLicenseKey, out.site.runtimeLicenseKey)
        || !takeString(13, out.controlSource, out.site.controlSource)
        || !takeString(14, out.rowSource, out.site.rowSource)) return false;
    return extra == out.end;
}

UserFormSiteDataLayout parseUserFormSiteData(const std::vector<unsigned char>& bytes, std::size_t minimumOffset) {
    UserFormSiteDataLayout best;
    auto fail = [&](std::string message) {
        best.inspection.valid = false;
        best.inspection.error = std::move(message);
        return best;
    };
    if (minimumOffset > bytes.size()) return fail("UserForm semantic Form block extends beyond Form stream");
    try {
        for (std::size_t candidate = minimumOffset; candidate + 8 <= bytes.size(); ++candidate) {
            for (const bool classCountStored : {true, false}) {
                std::size_t pos = candidate;
                std::size_t classInfoCount = 0;
                if (classCountStored) {
                    if (pos + 2 > bytes.size()) continue;
                    classInfoCount = userFormU16(bytes, pos); pos += 2;
                    if (classInfoCount > 1024) continue;
                    bool badClass = false;
                    for (std::size_t i = 0; i < classInfoCount; ++i) {
                        if (pos + 4 > bytes.size()) { badClass = true; break; }
                        const auto cbClass = userFormU16(bytes, pos + 2);
                        if (cbClass < 4 || pos + 4u + cbClass > bytes.size()) { badClass = true; break; }
                        pos += 4u + cbClass;
                    }
                    if (badClass) continue;
                }
                if (pos + 8 > bytes.size()) continue;
                const auto countSites = userFormU32(bytes, pos);
                const auto countBytes = userFormU32(bytes, pos + 4);
                if (countSites > 4096) continue;
                const auto countOfBytesOffset = pos + 4;
                const auto bodyStart = pos + 8;
                if (countBytes > bytes.size() - bodyStart || bodyStart + countBytes != bytes.size()) continue;
                std::vector<std::pair<std::uint8_t, std::uint8_t>> depthTypes;
                depthTypes.reserve(countSites);
                std::size_t depthPos = bodyStart;
                while (depthTypes.size() < countSites) {
                    if (depthPos + 2 > bytes.size()) { depthTypes.clear(); break; }
                    const auto depth = bytes[depthPos++];
                    const auto encoded = bytes[depthPos++];
                    const bool isCount = (encoded & 0x80u) != 0;
                    const auto typeOrCount = static_cast<std::uint8_t>(encoded & 0x7Fu);
                    if (isCount) {
                        if (typeOrCount == 0 || depthPos >= bytes.size()) { depthTypes.clear(); break; }
                        const auto type = bytes[depthPos++];
                        if (type != 1 || depthTypes.size() + typeOrCount > countSites) { depthTypes.clear(); break; }
                        for (std::size_t i = 0; i < typeOrCount; ++i) depthTypes.emplace_back(depth, type);
                    } else {
                        if (typeOrCount != 1) { depthTypes.clear(); break; }
                        depthTypes.emplace_back(depth, typeOrCount);
                    }
                }
                if (depthTypes.size() != countSites) continue;
                const auto depthBytes = depthPos - bodyStart;
                std::size_t sitePos = bodyStart + userFormAlign(depthBytes, 4);
                if (sitePos > bytes.size()) continue;
                std::vector<UserFormControlSiteLayout> sites;
                sites.reserve(countSites);
                bool badSite = false;
                std::size_t objectBytes = 0;
                for (std::size_t i = 0; i < countSites; ++i) {
                    UserFormControlSiteLayout site;
                    if (!parseOneUserFormSite(bytes, sitePos, bytes.size(), depthTypes[i].first, depthTypes[i].second, site)) {
                        badSite = true; break;
                    }
                    if (site.site.objectStreamSize) objectBytes += *site.site.objectStreamSize;
                    sitePos = site.end;
                    sites.push_back(std::move(site));
                }
                if (badSite || sitePos != bytes.size()) continue;
                best.inspection.valid = true;
                best.inspection.error.clear();
                best.inspection.siteDataOffset = candidate;
                best.inspection.classInfoCount = classInfoCount;
                best.inspection.totalObjectStreamBytes = objectBytes;
                for (const auto& site : sites) best.inspection.controls.push_back(site.site);
                best.countOfBytesOffset = countOfBytesOffset;
                best.sites = std::move(sites);
                return best;
            }
        }
        return fail("MS-OFORMS FormSiteData could not be located or validated at the end of the Form stream");
    } catch (const std::exception& e) {
        return fail(e.what());
    }
}

struct UserFormControlObjectLayout {
    VbaUserFormControlObjectInspection inspection;
    std::optional<std::size_t> foreColor, backColor, variousPropertyBits, captionCount,
        picturePosition, mousePointer, borderColor, borderStyle, specialEffect,
        accelerator, size;
    std::optional<std::size_t> captionData;
    std::size_t captionBytes{0};
    bool captionCompressed{false};
    std::size_t semanticEnd{0};
};

UserFormControlObjectLayout parseSimpleUserFormControlObject(const std::vector<unsigned char>& bytes,
                                                              VbaUserFormControlKind kind) {
    UserFormControlObjectLayout out;
    out.inspection.properties.kind = kind;
    out.inspection.objectBytes = bytes.size();
    auto fail = [&](std::string message) {
        out.inspection.valid = false;
        out.inspection.error = std::move(message);
        return out;
    };
    if (bytes.size() < 4) return fail("UserForm control object stream is shorter than its common header");
    out.inspection.properties.minorVersion = bytes[0];
    out.inspection.properties.majorVersion = bytes[1];
    out.inspection.properties.cbControl = userFormU16(bytes, 2);
    out.semanticEnd = 4u + out.inspection.properties.cbControl;
    if (out.semanticEnd > bytes.size()) return fail("UserForm control cbControl extends past its object-stream slice");
    out.inspection.semanticSectionBytes = out.semanticEnd;
    out.inspection.trailingBytes = bytes.size() - out.semanticEnd;

    if (kind != VbaUserFormControlKind::CommandButton && kind != VbaUserFormControlKind::Label) {
        // The other built-in families have different PropMask widths/layouts
        // (notably MorphData uses a 64-bit mask). P1F validates the common
        // header and exposes the type without pretending those layouts match.
        out.inspection.valid = true;
        return out;
    }
    if (bytes.size() < 8) return fail("CommandButton/Label object stream is shorter than PropMask");
    const auto mask = userFormU32(bytes, 4);
    out.inspection.properties.propertyMask = mask;
    out.inspection.properties.semanticPropertiesSupported = true;
    const auto has = [&](unsigned bit) { return (mask & (1u << bit)) != 0; };
    const std::size_t dataStart = 8;
    std::size_t rel = 0;
    const auto take = [&](unsigned bit, std::size_t size, std::optional<std::size_t>& target) {
        if (!has(bit)) return true;
        rel = userFormAlign(rel, std::min<std::size_t>(size, 4));
        if (dataStart + rel + size > out.semanticEnd) return false;
        target = dataStart + rel;
        rel += size;
        return true;
    };
    if (!take(0, 4, out.foreColor) || !take(1, 4, out.backColor)
        || !take(2, 4, out.variousPropertyBits) || !take(3, 4, out.captionCount)
        || !take(4, 4, out.picturePosition) || !take(6, 1, out.mousePointer))
        return fail("CommandButton/Label DataBlock exceeds cbControl");

    std::optional<std::size_t> picture, mouseIcon;
    if (kind == VbaUserFormControlKind::CommandButton) {
        if (!take(7, 2, picture) || !take(8, 2, out.accelerator) || !take(10, 2, mouseIcon))
            return fail("CommandButton DataBlock exceeds cbControl");
    } else {
        if (!take(7, 4, out.borderColor) || !take(8, 2, out.borderStyle)
            || !take(9, 2, out.specialEffect) || !take(10, 2, picture)
            || !take(11, 2, out.accelerator) || !take(12, 2, mouseIcon))
            return fail("Label DataBlock exceeds cbControl");
    }
    const auto extraStart = dataStart + userFormAlign(rel, 4);
    if (extraStart > out.semanticEnd) return fail("CommandButton/Label DataBlock padding exceeds cbControl");
    std::size_t extra = extraStart;
    if (has(3)) {
        if (!out.captionCount) return fail("Control caption flag has no length/compression field");
        const auto count = userFormU32(bytes, *out.captionCount);
        out.captionCompressed = (count & 0x80000000u) != 0;
        out.captionBytes = count & 0x7FFFFFFFu;
        if (extra + out.captionBytes > out.semanticEnd) return fail("Control caption extends past cbControl");
        out.captionData = extra;
        out.inspection.properties.caption = decodeUserFormString(bytes, extra, out.captionBytes, out.captionCompressed);
        extra += out.captionBytes;
    }
    if (has(5)) {
        extra = userFormAlign(extra, 4);
        if (extra + 8 > out.semanticEnd) return fail("Control Size extends past cbControl");
        out.size = extra;
        out.inspection.properties.width = userFormI32(bytes, extra);
        out.inspection.properties.height = userFormI32(bytes, extra + 4);
        extra += 8;
    }
    // StreamData/TextProps begin after cbControl and remain opaque. Within the
    // semantic section only undefined alignment bytes may remain.
    if (extra > out.semanticEnd) return fail("Control ExtraDataBlock exceeds cbControl");
    if (out.foreColor) out.inspection.properties.foreColor = userFormU32(bytes, *out.foreColor);
    if (out.backColor) out.inspection.properties.backColor = userFormU32(bytes, *out.backColor);
    if (out.variousPropertyBits) out.inspection.properties.variousPropertyBits = userFormU32(bytes, *out.variousPropertyBits);
    if (out.picturePosition) out.inspection.properties.picturePosition = userFormU32(bytes, *out.picturePosition);
    if (out.mousePointer) out.inspection.properties.mousePointer = bytes[*out.mousePointer];
    if (out.borderColor) out.inspection.properties.borderColor = userFormU32(bytes, *out.borderColor);
    if (out.borderStyle) out.inspection.properties.borderStyle = userFormU16(bytes, *out.borderStyle);
    if (out.specialEffect) out.inspection.properties.specialEffect = userFormU16(bytes, *out.specialEffect);
    if (out.accelerator) out.inspection.properties.accelerator = userFormU16(bytes, *out.accelerator);
    out.inspection.valid = true;
    return out;
}

void requireUserFormObjectField(bool requested, const std::optional<std::size_t>& offset, const char* name) {
    if (requested && !offset)
        throw std::invalid_argument(std::string("UserForm control object property is not materialized in PropMask: ") + name);
}

void requireUserFormSiteField(bool requested, const std::optional<std::size_t>& offset, const char* name) {
    if (requested && !offset) throw std::invalid_argument(std::string("UserForm control-site property is not materialized in SitePropMask: ") + name);
}
} // namespace

VbaUserFormInspection Workbook::inspectVbaUserForm(const std::string& storageName) const {
    const auto storages = vbaDesignerStorages();
    const auto it = std::find_if(storages.begin(), storages.end(), [&](const auto& storage) { return storage.name == storageName; });
    if (it == storages.end()) return {false, "Designer storage not found: " + storageName, {}, 0};
    const auto* stream = it->findStream("f");
    if (!stream) return {false, "Designer storage does not contain an MS-OFORMS Form stream named f", {}, 0};
    return parseUserFormStream(stream->data).inspection;
}

bool Workbook::updateVbaUserFormProperties(const std::string& storageName,
                                           const VbaUserFormPropertiesPatch& patch) {
    auto info = vbaProjectInfo();
    auto storageIt = std::find_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
        return storage.name == storageName;
    });
    if (storageIt == info.designerStorages.end()) return false;
    auto streamIt = std::find_if(storageIt->streams.begin(), storageIt->streams.end(), [](const auto& stream) { return stream.path == "f"; });
    if (streamIt == storageIt->streams.end()) return false;
    auto layout = parseUserFormStream(streamIt->data);
    if (!layout.inspection.valid) throw std::runtime_error("Malformed UserForm Form stream: " + layout.inspection.error);
    auto& bytes = streamIt->data;

    requireUserFormField(patch.backColor.has_value(), layout.backColor, "BackColor");
    requireUserFormField(patch.foreColor.has_value(), layout.foreColor, "ForeColor");
    requireUserFormField(patch.nextAvailableId.has_value(), layout.nextId, "NextAvailableID");
    requireUserFormField(patch.booleanProperties.has_value(), layout.booleanProperties, "BooleanProperties");
    requireUserFormField(patch.borderStyle.has_value(), layout.borderStyle, "BorderStyle");
    requireUserFormField(patch.mousePointer.has_value(), layout.mousePointer, "MousePointer");
    requireUserFormField(patch.scrollBars.has_value(), layout.scrollBars, "ScrollBars");
    requireUserFormField(patch.groupCount.has_value(), layout.groupCount, "GroupCount");
    requireUserFormField(patch.cycle.has_value(), layout.cycle, "Cycle");
    requireUserFormField(patch.specialEffect.has_value(), layout.specialEffect, "SpecialEffect");
    requireUserFormField(patch.borderColor.has_value(), layout.borderColor, "BorderColor");
    requireUserFormField(patch.caption.has_value(), layout.captionCount, "Caption");
    requireUserFormField(patch.zoom.has_value(), layout.zoom, "Zoom");
    requireUserFormField(patch.pictureAlignment.has_value(), layout.pictureAlignment, "PictureAlignment");
    requireUserFormField(patch.pictureSizeMode.has_value(), layout.pictureSizeMode, "PictureSizeMode");
    requireUserFormField(patch.shapeCookie.has_value(), layout.shapeCookie, "ShapeCookie");
    requireUserFormField(patch.drawBuffer.has_value(), layout.drawBuffer, "DrawBuffer");
    requireUserFormField(patch.displayedWidth.has_value() || patch.displayedHeight.has_value(), layout.displayedSize, "DisplayedSize");
    requireUserFormField(patch.logicalWidth.has_value() || patch.logicalHeight.has_value(), layout.logicalSize, "LogicalSize");
    requireUserFormField(patch.scrollLeft.has_value() || patch.scrollTop.has_value(), layout.scrollPosition, "ScrollPosition");

    if (patch.backColor) userFormPutU32(bytes, *layout.backColor, *patch.backColor);
    if (patch.foreColor) userFormPutU32(bytes, *layout.foreColor, *patch.foreColor);
    if (patch.nextAvailableId) userFormPutU32(bytes, *layout.nextId, *patch.nextAvailableId);
    if (patch.booleanProperties) userFormPutU32(bytes, *layout.booleanProperties, *patch.booleanProperties);
    if (patch.borderStyle) bytes[*layout.borderStyle] = *patch.borderStyle;
    if (patch.mousePointer) bytes[*layout.mousePointer] = *patch.mousePointer;
    if (patch.scrollBars) bytes[*layout.scrollBars] = *patch.scrollBars;
    if (patch.groupCount) userFormPutI32(bytes, *layout.groupCount, *patch.groupCount);
    if (patch.cycle) bytes[*layout.cycle] = *patch.cycle;
    if (patch.specialEffect) bytes[*layout.specialEffect] = *patch.specialEffect;
    if (patch.borderColor) userFormPutU32(bytes, *layout.borderColor, *patch.borderColor);
    if (patch.zoom) userFormPutU32(bytes, *layout.zoom, *patch.zoom);
    if (patch.pictureAlignment) bytes[*layout.pictureAlignment] = *patch.pictureAlignment;
    if (patch.pictureSizeMode) bytes[*layout.pictureSizeMode] = *patch.pictureSizeMode;
    if (patch.shapeCookie) userFormPutU32(bytes, *layout.shapeCookie, *patch.shapeCookie);
    if (patch.drawBuffer) userFormPutU32(bytes, *layout.drawBuffer, *patch.drawBuffer);
    if (patch.displayedWidth) userFormPutI32(bytes, *layout.displayedSize, *patch.displayedWidth);
    if (patch.displayedHeight) userFormPutI32(bytes, *layout.displayedSize + 4, *patch.displayedHeight);
    if (patch.logicalWidth) userFormPutI32(bytes, *layout.logicalSize, *patch.logicalWidth);
    if (patch.logicalHeight) userFormPutI32(bytes, *layout.logicalSize + 4, *patch.logicalHeight);
    if (patch.scrollLeft) userFormPutI32(bytes, *layout.scrollPosition, *patch.scrollLeft);
    if (patch.scrollTop) userFormPutI32(bytes, *layout.scrollPosition + 4, *patch.scrollTop);

    if (patch.caption) {
        auto [encoded, compressed] = encodeUserFormString(*patch.caption);
        const auto oldBegin = *layout.captionData;
        const auto oldEnd = oldBegin + layout.captionBytes;
        std::vector<unsigned char> rebuilt;
        rebuilt.reserve(bytes.size() - layout.captionBytes + encoded.size());
        rebuilt.insert(rebuilt.end(), bytes.begin(), bytes.begin() + static_cast<std::ptrdiff_t>(oldBegin));
        rebuilt.insert(rebuilt.end(), encoded.begin(), encoded.end());
        rebuilt.insert(rebuilt.end(), bytes.begin() + static_cast<std::ptrdiff_t>(oldEnd), bytes.end());
        const auto newCount = static_cast<std::uint32_t>(encoded.size()) | (compressed ? 0x80000000u : 0u);
        userFormPutU32(rebuilt, *layout.captionCount, newCount);
        const auto delta = static_cast<std::ptrdiff_t>(encoded.size()) - static_cast<std::ptrdiff_t>(layout.captionBytes);
        const auto newCb = static_cast<std::ptrdiff_t>(userFormU16(rebuilt, 2)) + delta;
        if (newCb < 4 || newCb > 0xFFFF) throw std::overflow_error("UserForm cbForm exceeds 16-bit storage after caption edit");
        userFormPutU16(rebuilt, 2, static_cast<std::uint16_t>(newCb));
        bytes = std::move(rebuilt);
    }

    setVbaProjectInfo(std::move(info));
    return true;
}

VbaUserFormControlsInspection Workbook::inspectVbaUserFormControls(const std::string& storageName) const {
    const auto storages = vbaDesignerStorages();
    const auto it = std::find_if(storages.begin(), storages.end(), [&](const auto& storage) { return storage.name == storageName; });
    if (it == storages.end()) return {false, "Designer storage not found: " + storageName};
    const auto* stream = it->findStream("f");
    if (!stream) return {false, "Designer storage does not contain an MS-OFORMS Form stream named f"};
    const auto form = parseUserFormStream(stream->data);
    if (!form.inspection.valid) return {false, form.inspection.error};
    auto inspection = parseUserFormSiteData(stream->data, form.sectionEnd).inspection;
    if (!inspection.valid) return inspection;

    const auto* objectStream = it->findStream("o");
    if (inspection.totalObjectStreamBytes != 0 && !objectStream) {
        inspection.valid = false;
        inspection.error = "UserForm control sites reference object-stream bytes but Designer Storage has no stream named o";
        return inspection;
    }
    if (objectStream) {
        inspection.objectStreamBytes = objectStream->data.size();
        if (objectStream->data.size() < inspection.totalObjectStreamBytes) {
            inspection.valid = false;
            inspection.error = "UserForm object stream is shorter than the sum of OleSite ObjectStreamSize values";
            return inspection;
        }
        std::size_t offset = 0;
        for (auto& control : inspection.controls) {
            const auto size = static_cast<std::size_t>(control.objectStreamSize.value_or(0));
            control.objectStreamOffset = offset;
            control.objectData.assign(objectStream->data.begin() + static_cast<std::ptrdiff_t>(offset),
                                      objectStream->data.begin() + static_cast<std::ptrdiff_t>(offset + size));
            offset += size;
        }
        inspection.unassignedObjectStreamBytes = objectStream->data.size() - offset;
    }
    return inspection;
}

VbaUserFormControlObjectInspection Workbook::inspectVbaUserFormControlObject(const std::string& storageName,
                                                                                 std::size_t controlIndex) const {
    const auto controls = inspectVbaUserFormControls(storageName);
    if (!controls.valid) return {false, controls.error};
    if (controlIndex >= controls.controls.size()) return {false, "UserForm control-site index is out of range"};
    const auto& control = controls.controls[controlIndex];
    return parseSimpleUserFormControlObject(control.objectData, control.kind).inspection;
}

bool Workbook::updateVbaUserFormControlObject(const std::string& storageName,
                                              std::size_t controlIndex,
                                              const VbaUserFormControlObjectPatch& patch) {
    auto info = vbaProjectInfo();
    auto storageIt = std::find_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
        return storage.name == storageName;
    });
    if (storageIt == info.designerStorages.end()) return false;
    auto fIt = std::find_if(storageIt->streams.begin(), storageIt->streams.end(), [](const auto& stream) { return stream.path == "f"; });
    auto oIt = std::find_if(storageIt->streams.begin(), storageIt->streams.end(), [](const auto& stream) { return stream.path == "o"; });
    if (fIt == storageIt->streams.end() || oIt == storageIt->streams.end()) return false;

    auto form = parseUserFormStream(fIt->data);
    if (!form.inspection.valid) throw std::runtime_error("Cannot edit malformed UserForm Form stream: " + form.inspection.error);
    auto sites = parseUserFormSiteData(fIt->data, form.sectionEnd);
    if (!sites.inspection.valid) throw std::runtime_error("Cannot edit UserForm control object: " + sites.inspection.error);
    if (controlIndex >= sites.sites.size()) throw std::out_of_range("UserForm control-site index is out of range");
    auto& site = sites.sites[controlIndex];
    const auto objectSize = static_cast<std::size_t>(site.site.objectStreamSize.value_or(0));
    std::size_t objectOffset = 0;
    for (std::size_t i = 0; i < controlIndex; ++i)
        objectOffset += static_cast<std::size_t>(sites.sites[i].site.objectStreamSize.value_or(0));
    if (objectOffset + objectSize > oIt->data.size())
        throw std::runtime_error("UserForm object stream is shorter than the control's ObjectStreamSize slice");
    std::vector<unsigned char> object(oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset),
                                      oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset + objectSize));
    auto layout = parseSimpleUserFormControlObject(object, site.site.kind);
    if (!layout.inspection.valid) throw std::runtime_error("Cannot edit malformed UserForm control object: " + layout.inspection.error);
    if (!layout.inspection.properties.semanticPropertiesSupported)
        throw std::invalid_argument("Semantic object editing is currently supported only for built-in CommandButton and Label controls");

    requireUserFormObjectField(patch.foreColor.has_value(), layout.foreColor, "ForeColor");
    requireUserFormObjectField(patch.backColor.has_value(), layout.backColor, "BackColor");
    requireUserFormObjectField(patch.variousPropertyBits.has_value(), layout.variousPropertyBits, "VariousPropertyBits");
    requireUserFormObjectField(patch.caption.has_value(), layout.captionCount, "Caption");
    requireUserFormObjectField(patch.picturePosition.has_value(), layout.picturePosition, "PicturePosition");
    requireUserFormObjectField(patch.mousePointer.has_value(), layout.mousePointer, "MousePointer");
    requireUserFormObjectField(patch.borderColor.has_value(), layout.borderColor, "BorderColor");
    requireUserFormObjectField(patch.borderStyle.has_value(), layout.borderStyle, "BorderStyle");
    requireUserFormObjectField(patch.specialEffect.has_value(), layout.specialEffect, "SpecialEffect");
    requireUserFormObjectField(patch.accelerator.has_value(), layout.accelerator, "Accelerator");
    requireUserFormObjectField(patch.width.has_value() || patch.height.has_value(), layout.size, "Size");

    if (patch.caption) {
        requireUserFormObjectField(true, layout.captionData, "Caption");
        auto [encoded, compressed] = encodeUserFormString(*patch.caption);
        const auto oldBegin = *layout.captionData;
        const auto oldEnd = oldBegin + layout.captionBytes;
        std::vector<unsigned char> rebuilt;
        rebuilt.reserve(object.size() - layout.captionBytes + encoded.size());
        rebuilt.insert(rebuilt.end(), object.begin(), object.begin() + static_cast<std::ptrdiff_t>(oldBegin));
        rebuilt.insert(rebuilt.end(), encoded.begin(), encoded.end());
        rebuilt.insert(rebuilt.end(), object.begin() + static_cast<std::ptrdiff_t>(oldEnd), object.end());
        userFormPutU32(rebuilt, *layout.captionCount,
                       static_cast<std::uint32_t>(encoded.size()) | (compressed ? 0x80000000u : 0u));
        const auto delta = static_cast<std::ptrdiff_t>(encoded.size()) - static_cast<std::ptrdiff_t>(layout.captionBytes);
        const auto newCb = static_cast<std::ptrdiff_t>(userFormU16(rebuilt, 2)) + delta;
        if (newCb < 4 || newCb > 0xFFFF) throw std::overflow_error("UserForm control cbControl exceeds 16-bit storage after caption edit");
        userFormPutU16(rebuilt, 2, static_cast<std::uint16_t>(newCb));
        object = std::move(rebuilt);
        layout = parseSimpleUserFormControlObject(object, site.site.kind);
        if (!layout.inspection.valid) throw std::runtime_error("UserForm control object became malformed after caption edit: " + layout.inspection.error);
    }

    if (patch.foreColor) userFormPutU32(object, *layout.foreColor, *patch.foreColor);
    if (patch.backColor) userFormPutU32(object, *layout.backColor, *patch.backColor);
    if (patch.variousPropertyBits) userFormPutU32(object, *layout.variousPropertyBits, *patch.variousPropertyBits);
    if (patch.picturePosition) userFormPutU32(object, *layout.picturePosition, *patch.picturePosition);
    if (patch.mousePointer) object[*layout.mousePointer] = *patch.mousePointer;
    if (patch.borderColor) userFormPutU32(object, *layout.borderColor, *patch.borderColor);
    if (patch.borderStyle) userFormPutU16(object, *layout.borderStyle, *patch.borderStyle);
    if (patch.specialEffect) userFormPutU16(object, *layout.specialEffect, *patch.specialEffect);
    if (patch.accelerator) userFormPutU16(object, *layout.accelerator, *patch.accelerator);
    if (patch.width) userFormPutI32(object, *layout.size, *patch.width);
    if (patch.height) userFormPutI32(object, *layout.size + 4, *patch.height);

    if (object.size() != objectSize) {
        requireUserFormSiteField(true, site.objectStreamSize, "ObjectStreamSize");
        if (object.size() > 0xFFFFFFFFu) throw std::overflow_error("UserForm control object exceeds 32-bit ObjectStreamSize");
        userFormPutU32(fIt->data, *site.objectStreamSize, static_cast<std::uint32_t>(object.size()));
    }
    oIt->data.erase(oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset),
                    oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset + objectSize));
    oIt->data.insert(oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset), object.begin(), object.end());
    setVbaProjectInfo(std::move(info));
    return true;
}

bool Workbook::updateVbaUserFormControlSite(const std::string& storageName,
                                            std::size_t controlIndex,
                                            const VbaUserFormControlSitePatch& patch) {
    auto info = vbaProjectInfo();
    auto storageIt = std::find_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
        return storage.name == storageName;
    });
    if (storageIt == info.designerStorages.end()) return false;
    auto streamIt = std::find_if(storageIt->streams.begin(), storageIt->streams.end(), [](const auto& stream) { return stream.path == "f"; });
    if (streamIt == storageIt->streams.end()) return false;
    auto& bytes = streamIt->data;
    const auto form = parseUserFormStream(bytes);
    if (!form.inspection.valid) throw std::runtime_error("Cannot edit malformed UserForm Form stream: " + form.inspection.error);

    auto parseSites = [&]() {
        auto parsed = parseUserFormSiteData(bytes, parseUserFormStream(bytes).sectionEnd);
        if (!parsed.inspection.valid) throw std::runtime_error("Cannot edit UserForm controls: " + parsed.inspection.error);
        if (controlIndex >= parsed.sites.size()) throw std::out_of_range("UserForm control-site index is out of range");
        return parsed;
    };

    auto patchString = [&](const std::optional<std::string>& requested, auto member, const char* name) {
        if (!requested) return;
        auto parsed = parseSites();
        auto& site = parsed.sites[controlIndex];
        const auto& layout = site.*member;
        requireUserFormSiteField(true, layout.countOffset, name);
        requireUserFormSiteField(true, layout.dataOffset, name);
        auto [encoded, compressed] = encodeUserFormString(*requested);
        const auto oldBytes = layout.byteCount;
        const auto dataOffset = *layout.dataOffset;
        bytes.erase(bytes.begin() + static_cast<std::ptrdiff_t>(dataOffset),
                    bytes.begin() + static_cast<std::ptrdiff_t>(dataOffset + oldBytes));
        bytes.insert(bytes.begin() + static_cast<std::ptrdiff_t>(dataOffset), encoded.begin(), encoded.end());
        const auto newCount = static_cast<std::uint32_t>(encoded.size()) | (compressed ? 0x80000000u : 0u);
        userFormPutU32(bytes, *layout.countOffset, newCount);
        const auto delta = static_cast<std::ptrdiff_t>(encoded.size()) - static_cast<std::ptrdiff_t>(oldBytes);
        const auto oldCbSite = static_cast<std::ptrdiff_t>(userFormU16(bytes, site.begin + 2));
        const auto newCbSite = oldCbSite + delta;
        if (newCbSite < 4 || newCbSite > 0xFFFF) throw std::overflow_error("UserForm control cbSite exceeds 16-bit storage after string edit");
        userFormPutU16(bytes, site.begin + 2, static_cast<std::uint16_t>(newCbSite));
        const auto oldCountBytes = static_cast<std::int64_t>(userFormU32(bytes, parsed.countOfBytesOffset));
        const auto newCountBytes = oldCountBytes + delta;
        if (newCountBytes < 0 || newCountBytes > 0xFFFFFFFFll) throw std::overflow_error("UserForm FormSiteData CountOfBytes overflow");
        userFormPutU32(bytes, parsed.countOfBytesOffset, static_cast<std::uint32_t>(newCountBytes));
    };

    patchString(patch.name, &UserFormControlSiteLayout::name, "Name");
    patchString(patch.tag, &UserFormControlSiteLayout::tag, "Tag");
    patchString(patch.controlTipText, &UserFormControlSiteLayout::controlTipText, "ControlTipText");
    patchString(patch.controlSource, &UserFormControlSiteLayout::controlSource, "ControlSource");
    patchString(patch.rowSource, &UserFormControlSiteLayout::rowSource, "RowSource");

    auto parsed = parseSites();
    auto& site = parsed.sites[controlIndex];
    requireUserFormSiteField(patch.helpContextId.has_value(), site.helpContextId, "HelpContextID");
    requireUserFormSiteField(patch.bitFlags.has_value(), site.bitFlags, "BitFlags");
    requireUserFormSiteField(patch.tabIndex.has_value(), site.tabIndex, "TabIndex");
    requireUserFormSiteField(patch.groupId.has_value(), site.groupId, "GroupID");
    requireUserFormSiteField(patch.top.has_value() || patch.left.has_value(), site.position, "Position");
    if (patch.helpContextId) userFormPutI32(bytes, *site.helpContextId, *patch.helpContextId);
    if (patch.bitFlags) userFormPutU32(bytes, *site.bitFlags, *patch.bitFlags);
    if (patch.tabIndex) userFormPutI16(bytes, *site.tabIndex, *patch.tabIndex);
    if (patch.groupId) userFormPutU16(bytes, *site.groupId, *patch.groupId);
    if (patch.top) userFormPutI32(bytes, *site.position, *patch.top);
    if (patch.left) userFormPutI32(bytes, *site.position + 4, *patch.left);

    setVbaProjectInfo(std::move(info));
    return true;
}

VbaDesignerValidationReport Workbook::validateVbaDesignerProject() const {
    VbaDesignerValidationReport report;
    const auto modules = vbaModules();
    const auto storages = vbaDesignerStorages();
    std::set<std::string> moduleNames;
    std::set<std::string> storageNames;
    for (const auto& module : modules) {
        if (module.type != VbaModuleType::Designer) continue;
        ++report.designerModules;
        if (!moduleNames.insert(module.name).second) {
            report.issues.push_back({module.name, "Duplicate VBA Designer module name"});
            continue;
        }
        const auto storage = std::find_if(storages.begin(), storages.end(), [&](const auto& item) { return item.name == module.name; });
        if (storage == storages.end()) {
            report.issues.push_back({module.name, "Designer module has no matching root Designer Storage"});
            continue;
        }
        const auto* form = storage->findStream("f");
        const bool looksLikeUserForm = module.designerClassId.find("AC9F2F90-E877-11CE-9F68-00AA00574A4F") != std::string::npos;
        if (!form) {
            if (looksLikeUserForm) report.issues.push_back({module.name, "UserForm Designer Storage has no Form stream named f"});
            continue;
        }
        const auto parsed = parseUserFormStream(form->data);
        if (!parsed.inspection.valid) report.issues.push_back({module.name, "Malformed MS-OFORMS Form stream: " + parsed.inspection.error});
        else ++report.validUserFormStreams;
    }
    for (const auto& storage : storages) {
        ++report.designerStorages;
        if (!storageNames.insert(storage.name).second) {
            report.issues.push_back({storage.name, "Duplicate root Designer Storage name"});
            continue;
        }
        const auto module = std::find_if(modules.begin(), modules.end(), [&](const auto& item) {
            return item.type == VbaModuleType::Designer && item.name == storage.name;
        });
        if (module == modules.end()) report.issues.push_back({storage.name, "Designer Storage is orphaned from VBA Designer modules"});
    }
    return report;
}

void Workbook::addVbaProject(const std::filesystem::path& path) {
    std::ifstream stream(path, std::ios::binary);
    if (!stream) throw std::runtime_error("Cannot open VBA project file: " + path.string());
    std::vector<unsigned char> bytes((std::istreambuf_iterator<char>(stream)), std::istreambuf_iterator<char>());
    if (bytes.empty()) throw std::invalid_argument("VBA project file is empty");
    setVbaProject(std::move(bytes));
}

void Workbook::setVbaProject(std::vector<unsigned char> bytes) {
    if (bytes.empty()) throw std::invalid_argument("VBA project bytes cannot be empty");
    removeVbaProject();
    generatedVbaProject_ = false;
    PreservedPart part;
    part.name = "xl/vbaProject.bin";
    part.data.assign(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    part.overrideType = "application/vnd.ms-office.vbaProject";
    part.extension = "bin";
    part.defaultType = "application/vnd.ms-office.vbaProject";
    part.compress = false;
    preservedParts_.push_back(std::move(part));
}

bool Workbook::hasVbaProject() const noexcept {
    return std::any_of(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
}

bool Workbook::removeVbaProject() noexcept {
    const auto oldSize = preservedParts_.size();
    preservedParts_.erase(std::remove_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin"
            || part.name == "xl/vbaProjectSignature.bin"
            || part.name == "xl/_rels/vbaProject.bin.rels";
    }), preservedParts_.end());
    const bool removed = preservedParts_.size() != oldSize;
    if (removed) generatedVbaProject_ = false;
    return removed;
}

std::vector<VbaModule> Workbook::vbaModules() const {
    const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
    if (it == preservedParts_.end()) return {};
    const std::vector<unsigned char> bytes(it->data.begin(), it->data.end());
    return internal::readVbaProjectBinary(bytes);
}

std::optional<std::string> Workbook::vbaModuleText(const std::string& moduleName) const {
    const auto target = [&] {
        std::string value = moduleName;
        std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
        return value;
    }();
    for (const auto& module : vbaModules()) {
        std::string name = module.name;
        std::transform(name.begin(), name.end(), name.begin(), [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
        if (name == target) return module.source;
    }
    return std::nullopt;
}

void Workbook::setVbaModule(VbaModule module) {
    if (module.type == VbaModuleType::Document) {
        const auto lower = [](std::string value) {
            std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
            return value;
        };
        const auto name = lower(module.name);
        bool validDocument = name == "thisworkbook";
        const auto codeNames = ensureWorksheetVbaCodeNames(sheets_);
        if (!validDocument) validDocument = std::any_of(codeNames.begin(), codeNames.end(), [&](const auto& codeName) {
            return lower(codeName) == name;
        });
        if (!validDocument) throw std::invalid_argument("Unknown VBA document module: " + module.name);
    } else {
        internal::validateVbaModuleName(module.name);
    }
    module.source = internal::normalizeVbaSource(std::move(module.source));

    std::vector<VbaModule> modules;
    VbaProjectInfo info;
    if (hasVbaProject()) {
        try {
            modules = vbaModules();
            info = vbaProjectInfo();
        } catch (const std::exception&) {
            throw std::runtime_error("Cannot edit VBA source in an unsupported existing vbaProject.bin; attach a generated project or replace it first");
        }
    }
    const auto sameName = [&](const VbaModule& existing) {
        if (existing.name.size() != module.name.size()) return false;
        return std::equal(existing.name.begin(), existing.name.end(), module.name.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
    };
    const auto it = std::find_if(modules.begin(), modules.end(), sameName);
    if (it == modules.end()) modules.push_back(std::move(module));
    else {
        if (it->type == VbaModuleType::Document && module.type != VbaModuleType::Document)
            throw std::invalid_argument("Cannot replace a VBA document module with a standard/class module: " + module.name);
        if (it->type != VbaModuleType::Document && module.type == VbaModuleType::Document)
            throw std::invalid_argument("Cannot replace a VBA standard/class module with a document module: " + module.name);
        *it = std::move(module);
    }
    setVbaProject(internal::buildVbaProjectBinary(modules, ensureWorksheetVbaCodeNames(sheets_), info));
    generatedVbaProject_ = true;
}

void Workbook::setVbaModuleText(std::string moduleName, std::string source) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Standard});
}

void Workbook::setVbaClassModuleText(std::string moduleName, std::string source) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Class});
}

void Workbook::setVbaDesignerModule(std::string moduleName, std::string source, VbaDesignerStorage storage,
                                    std::string designerClassId) {
    if (storage.name.empty()) storage.name = moduleName;
    if (storage.name != moduleName)
        throw std::invalid_argument("VBA designer storage name must match the designer module name");
    VbaModule module;
    module.name = std::move(moduleName);
    module.source = std::move(source);
    module.type = VbaModuleType::Designer;
    module.designerClassId = std::move(designerClassId);
    setVbaModule(std::move(module));
    setVbaDesignerStorage(std::move(storage));
}

std::vector<VbaDesignerStorage> Workbook::vbaDesignerStorages() const {
    if (!hasVbaProject()) return {};
    return vbaProjectInfo().designerStorages;
}

void Workbook::setVbaDesignerStorage(VbaDesignerStorage storage) {
    if (storage.name.empty()) throw std::invalid_argument("VBA designer storage name cannot be empty");
    if (!hasVbaProject()) throw std::logic_error("A VBA project is required before adding designer storage");
    auto info = vbaProjectInfo();
    const auto same = [&](const VbaDesignerStorage& existing) {
        if (existing.name.size() != storage.name.size()) return false;
        return std::equal(existing.name.begin(), existing.name.end(), storage.name.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
    };
    const auto it = std::find_if(info.designerStorages.begin(), info.designerStorages.end(), same);
    if (it == info.designerStorages.end()) info.designerStorages.push_back(std::move(storage));
    else *it = std::move(storage);
    setVbaProjectInfo(std::move(info));
}

bool Workbook::removeVbaDesignerStorage(const std::string& storageName) {
    if (!hasVbaProject()) return false;
    auto info = vbaProjectInfo();
    const auto oldSize = info.designerStorages.size();
    info.designerStorages.erase(std::remove_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
        if (storage.name.size() != storageName.size()) return false;
        return std::equal(storage.name.begin(), storage.name.end(), storageName.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
    }), info.designerStorages.end());
    if (info.designerStorages.size() == oldSize) return false;
    setVbaProjectInfo(std::move(info));
    return true;
}

void Workbook::setVbaDocumentModuleText(std::string moduleName, std::string source) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Document});
}

VbaProjectInfo Workbook::vbaProjectInfo() const {
    const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
    if (it == preservedParts_.end()) return {};
    const std::vector<unsigned char> bytes(it->data.begin(), it->data.end());
    return internal::readVbaProjectInfoBinary(bytes);
}

void Workbook::setVbaProjectInfo(VbaProjectInfo info) {
    if (info.name.empty()) throw std::invalid_argument("VBA project name cannot be empty");
    if (info.name.size() > 31) throw std::invalid_argument("VBA project name must contain at most 31 characters");
    auto modules = hasVbaProject() ? vbaModules() : std::vector<VbaModule>{};
    setVbaProject(internal::buildVbaProjectBinary(modules, ensureWorksheetVbaCodeNames(sheets_), info));
    generatedVbaProject_ = true;
}

bool Workbook::removeVbaModule(const std::string& moduleName) {
    if (!hasVbaProject()) return false;
    auto modules = vbaModules();
    bool removedDesigner = false;
    const auto oldSize = modules.size();
    modules.erase(std::remove_if(modules.begin(), modules.end(), [&](const VbaModule& module) {
        if (module.type == VbaModuleType::Document || module.name.size() != moduleName.size()) return false;
        const bool same = std::equal(module.name.begin(), module.name.end(), moduleName.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
        if (same && module.type == VbaModuleType::Designer) removedDesigner = true;
        return same;
    }), modules.end());
    if (modules.size() == oldSize) return false;
    const bool hasUserSource = std::any_of(modules.begin(), modules.end(), [](const VbaModule& module) {
        return module.type != VbaModuleType::Document || !module.source.empty();
    });
    if (!hasUserSource) {
        removeVbaProject();
    } else {
        VbaProjectInfo info;
        try { info = vbaProjectInfo(); } catch (...) {}
        if (removedDesigner) {
            info.designerStorages.erase(std::remove_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
                if (storage.name.size() != moduleName.size()) return false;
                return std::equal(storage.name.begin(), storage.name.end(), moduleName.begin(), [](unsigned char a, unsigned char b) {
                    return std::tolower(a) == std::tolower(b);
                });
            }), info.designerStorages.end());
        }
        setVbaProject(internal::buildVbaProjectBinary(modules, ensureWorksheetVbaCodeNames(sheets_), info));
        generatedVbaProject_ = true;
    }
    return true;
}

} // namespace xlpp
