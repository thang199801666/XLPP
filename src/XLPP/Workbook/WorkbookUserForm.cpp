#include "WorkbookUserForm.h"
#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <optional>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {
namespace internal {

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


// Reads a string LengthAndCompression field plus its ExtraDataBlock payload.
// MS-OFORMS aligns every string payload to a 4-byte boundary in the
// ExtraDataBlock, so `extra` is aligned before the payload is claimed.
bool takeUserFormObjectString(std::optional<std::size_t>& countOffset, std::optional<std::size_t>& dataOffset,
                              std::size_t& byteCount, bool& compressed,
                              const std::vector<unsigned char>& bytes, bool present,
                              std::size_t& extra, std::size_t semanticEnd) {
    if (!present) return true;
    if (!countOffset) return false;
    const auto count = userFormU32(bytes, *countOffset);
    compressed = (count & 0x80000000u) != 0;
    byteCount = count & 0x7FFFFFFFu;
    extra = userFormAlign(extra, 4);
    if (extra + byteCount > semanticEnd) return false;
    dataOffset = extra;
    extra += byteCount;
    return true;
}

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

    // Only ComboBox, ListBox, MultiPage and TabStrip use a 64-bit PropMask in
    // MS-OFORMS 2.x. ScrollBar/SpinButton and the toggle-family controls keep
    // a single 32-bit flag word.
    out.maskIs64Bit = kind == VbaUserFormControlKind::ComboBox || kind == VbaUserFormControlKind::ListBox ||
                      kind == VbaUserFormControlKind::MultiPage || kind == VbaUserFormControlKind::TabStrip;
    const std::size_t maskBytes = out.maskIs64Bit ? 8u : 4u;
    if (bytes.size() < 4u + maskBytes) return fail("UserForm control object stream is shorter than PropMask");
    const auto mask = userFormU32(bytes, 4);
    out.inspection.properties.propertyMask = mask;
    if (out.maskIs64Bit) out.inspection.properties.propertyMaskHigh = userFormU32(bytes, 8);
    out.inspection.properties.semanticPropertiesSupported = true;

    const auto has = [&](unsigned bit) {
        if (bit < 32) return (mask & (1u << bit)) != 0;
        return (out.inspection.properties.propertyMaskHigh & (1u << (bit - 32))) != 0;
    };
    const std::size_t dataStart = 4u + maskBytes;
    std::size_t rel = 0;
    const auto take = [&](unsigned bit, std::size_t size, std::optional<std::size_t>& target) {
        if (!has(bit)) return true;
        rel = userFormAlign(rel, std::min<std::size_t>(size, 4));
        if (dataStart + rel + size > out.semanticEnd) return false;
        target = dataStart + rel;
        rel += size;
        return true;
    };

    // DataBlock fields are serialized in ascending flag-bit order for the
    // owning control class. Common bits 0-2 (ForeColor/BackColor/
    // VariousPropertyBits) are identical across families; everything from bit 3
    // on is class-specific and is decoded in the per-class switch below.
    if (!take(0, 4, out.foreColor) || !take(1, 4, out.backColor) || !take(2, 4, out.variousPropertyBits))
        return fail("Control DataBlock exceeds cbControl");

    std::optional<std::size_t> picture, mouseIcon;
    switch (kind) {
        case VbaUserFormControlKind::CommandButton:
            // bits: 3 Caption, 4 PicturePosition, 5 Size, 6 MousePointer,
            // 7 Picture, 8 Accelerator, 10 MouseIcon.
            if (!take(3, 4, out.captionCount) || !take(4, 4, out.picturePosition) || !take(6, 1, out.mousePointer)
                || !take(7, 2, picture) || !take(8, 2, out.accelerator) || !take(10, 2, mouseIcon))
                return fail("CommandButton DataBlock exceeds cbControl");
            break;
        case VbaUserFormControlKind::Label:
            // bits: 3 Caption, 5 Size, 6 MousePointer, 7 BorderColor,
            // 8 BorderStyle, 9 SpecialEffect, 10 Picture, 11 Accelerator,
            // 12 MouseIcon.
            if (!take(3, 4, out.captionCount) || !take(6, 1, out.mousePointer)
                || !take(7, 4, out.borderColor) || !take(8, 2, out.borderStyle)
                || !take(9, 2, out.specialEffect) || !take(10, 2, picture)
                || !take(11, 2, out.accelerator) || !take(12, 2, mouseIcon))
                return fail("Label DataBlock exceeds cbControl");
            break;
        case VbaUserFormControlKind::TextBox:
            // bits: 3 Caption, 4 PicturePosition, 5 SpecialEffect,
            // 6 MousePointer, 7 Picture, 8 BorderColor, 9 BorderStyle,
            // 10 ScrollBars, 11 DisplayStyle, 12 MouseIcon, 13 EnterKeyBehavior,
            // 14 TabKeyBehavior, 15 MaxLength, 16 WordWrap, 17 Text,
            // 18 IMEMode, 19 IMEStatus, 20 AutoWordSelect, 21 IntegralHeight,
            // 22 PasswordChar, 23 Value, 24 LineCount, 25 MultiLine,
            // 26 MultiSelect, 27 HideSelection, 28 DataEntry, 29 DragBehavior,
            // 30 Size, 31 ListRows.
            if (!take(3, 4, out.captionCount) || !take(4, 4, out.picturePosition) || !take(5, 2, out.specialEffect)
                || !take(6, 1, out.mousePointer) || !take(7, 2, picture)
                || !take(8, 4, out.borderColor) || !take(9, 2, out.borderStyle) || !take(10, 2, out.scrollBars)
                || !take(11, 1, out.displayStyle) || !take(12, 2, mouseIcon)
                || !take(13, 1, out.enterKeyBehavior) || !take(14, 1, out.tabKeyBehavior)
                || !take(15, 4, out.maxLength) || !take(16, 2, out.wordWrap) || !take(17, 4, out.textCount)
                || !take(18, 1, out.iMEMode) || !take(19, 1, out.iMEStatus)
                || !take(20, 1, out.autoWordSelect) || !take(21, 1, out.integralHeight)
                || !take(22, 2, out.passwordChar) || !take(23, 4, out.valueCount)
                || !take(24, 4, out.lineCount) || !take(25, 2, out.multiLine) || !take(26, 2, out.multiSelect)
                || !take(27, 2, out.hideSelection) || !take(28, 1, out.dataEntry) || !take(29, 1, out.dragBehavior)
                || !take(31, 4, out.listRows))
                return fail("TextBox DataBlock exceeds cbControl");
            break;
        case VbaUserFormControlKind::CheckBox:
        case VbaUserFormControlKind::OptionButton:
        case VbaUserFormControlKind::ToggleButton:
            // bits: 3 Caption, 4 PicturePosition, 5 SpecialEffect,
            // 6 MousePointer, 7 Picture, 8 BorderColor, 9 BorderStyle,
            // 10 GroupName, 11 Accelerator, 12 MouseIcon, 13 Value,
            // 14 GroupNumber, 15 TripleState, 16 Size.
            if (!take(3, 4, out.captionCount) || !take(4, 4, out.picturePosition) || !take(5, 2, out.specialEffect)
                || !take(6, 1, out.mousePointer) || !take(7, 2, picture)
                || !take(8, 4, out.borderColor) || !take(9, 2, out.borderStyle) || !take(10, 4, out.groupNameCount)
                || !take(11, 2, out.accelerator) || !take(12, 2, mouseIcon) || !take(13, 4, out.valueCount)
                || !take(14, 2, out.groupNumber) || !take(15, 2, out.tripleState))
                return fail("Toggle-family DataBlock exceeds cbControl");
            break;
        case VbaUserFormControlKind::ScrollBar:
        case VbaUserFormControlKind::SpinButton:
            // ScrollBar/SpinButton use a 32-bit mask and store Value/Min/Max/
            // SmallChange/LargeChange as 32-bit signed integers (not strings).
            if (!take(3, 4, out.valueCount) || !take(4, 4, out.min) || !take(5, 4, out.max)
                || !take(6, 4, out.smallChange) || !take(7, 4, out.largeChange) || !take(8, 1, out.orientation)
                || !take(9, 1, out.mousePointer) || !take(10, 2, mouseIcon))
                return fail("SpinButton/ScrollBar DataBlock exceeds cbControl");
            break;
        case VbaUserFormControlKind::ComboBox:
        case VbaUserFormControlKind::ListBox:
            // bits (first flag word): 3 Caption, 4 PicturePosition,
            // 5 SpecialEffect, 6 MousePointer, 7 Picture, 8 BorderColor,
            // 9 BorderStyle, 10 ScrollBars, 11 DisplayStyle, 12 MouseIcon,
            // 15 ListRows, 16 ListWidth, 17 BoundColumn, 18 TextColumn,
            // 19 ColumnCount, 20 ColumnWidths, 21 Style, 22 ListStyle,
            // 23 MatchEntry, 24 ShowDropButtonWhen, 25 DropButtonStyle,
            // 26 MultiSelect, 27 Value, 28 MatchFound, 29 IMEMode,
            // 30 IMEStatus, 31 Size.
            if (!take(3, 4, out.captionCount) || !take(4, 4, out.picturePosition) || !take(5, 2, out.specialEffect)
                || !take(6, 1, out.mousePointer) || !take(7, 2, picture)
                || !take(8, 4, out.borderColor) || !take(9, 2, out.borderStyle) || !take(10, 2, out.scrollBars)
                || !take(11, 1, out.displayStyle) || !take(12, 2, mouseIcon)
                || !take(15, 4, out.listRows) || !take(16, 4, out.listWidth) || !take(17, 4, out.boundColumn)
                || !take(18, 4, out.textColumn) || !take(19, 4, out.columnCount) || !take(20, 4, out.columnWidthsCount)
                || !take(21, 1, out.style) || !take(22, 1, out.listStyle) || !take(23, 1, out.matchEntry)
                || !take(24, 1, out.showDropButtonWhen) || !take(25, 1, out.dropButtonStyle)
                || !take(26, 2, out.multiSelect) || !take(27, 4, out.valueCount)
                || !take(28, 1, out.matchFound) || !take(29, 1, out.iMEMode) || !take(30, 1, out.iMEStatus))
                return fail("ComboBox/ListBox DataBlock exceeds cbControl");
            break;
        default:
            // Unsupported built-in families expose only the common header and
            // PropMask without pretending the class-specific layout matches.
            out.inspection.properties.semanticPropertiesSupported = false;
            out.inspection.valid = true;
            return out;
    }

    const auto extraStart = dataStart + userFormAlign(rel, 4);
    if (extraStart > out.semanticEnd) return fail("Control DataBlock padding exceeds cbControl");
    std::size_t extra = extraStart;

    const auto decodeString = [&](std::optional<std::size_t>& dataOff, std::size_t byteCount, bool compressed,
                                  std::optional<std::string>& target) {
        if (!dataOff) return;
        target = decodeUserFormString(bytes, *dataOff, byteCount, compressed);
    };
    // ExtraDataBlock strings are emitted in ascending flag-bit order for the
    // owning control class. MS-OFORMS always serializes string payloads after
    // the fixed-size DataBlock, in the same relative order as their flags.
    auto failString = [&](const char* what) { return fail(std::string("Control ") + what + " extends past cbControl"); };
    switch (kind) {
        case VbaUserFormControlKind::TextBox:
            // bits: 3 Caption, 17 Text, 23 Value.
            if (!takeUserFormObjectString(out.captionCount, out.captionData, out.captionBytes, out.captionCompressed,
                                          bytes, has(3), extra, out.semanticEnd)) return failString("caption");
            if (!takeUserFormObjectString(out.textCount, out.textData, out.textBytes, out.textCompressed,
                                          bytes, has(17), extra, out.semanticEnd)) return failString("TextBox Text");
            if (!takeUserFormObjectString(out.valueCount, out.valueData, out.valueBytes, out.valueCompressed,
                                          bytes, has(23), extra, out.semanticEnd)) return failString("Value");
            break;
        case VbaUserFormControlKind::CheckBox:
        case VbaUserFormControlKind::OptionButton:
        case VbaUserFormControlKind::ToggleButton:
            // bits: 3 Caption, 10 GroupName, 13 Value.
            if (!takeUserFormObjectString(out.captionCount, out.captionData, out.captionBytes, out.captionCompressed,
                                          bytes, has(3), extra, out.semanticEnd)) return failString("caption");
            if (!takeUserFormObjectString(out.groupNameCount, out.groupNameData, out.groupNameBytes, out.groupNameCompressed,
                                          bytes, has(10), extra, out.semanticEnd)) return failString("GroupName");
            if (!takeUserFormObjectString(out.valueCount, out.valueData, out.valueBytes, out.valueCompressed,
                                          bytes, has(13), extra, out.semanticEnd)) return failString("Value");
            break;
        case VbaUserFormControlKind::ComboBox:
        case VbaUserFormControlKind::ListBox:
            // bits: 3 Caption, 20 ColumnWidths, 27 Value.
            if (!takeUserFormObjectString(out.captionCount, out.captionData, out.captionBytes, out.captionCompressed,
                                          bytes, has(3), extra, out.semanticEnd)) return failString("caption");
            if (!takeUserFormObjectString(out.columnWidthsCount, out.columnWidthsData, out.columnWidthsBytes, out.columnWidthsCompressed,
                                          bytes, has(20), extra, out.semanticEnd)) return failString("ColumnWidths");
            if (!takeUserFormObjectString(out.valueCount, out.valueData, out.valueBytes, out.valueCompressed,
                                          bytes, has(27), extra, out.semanticEnd)) return failString("Value");
            break;
        case VbaUserFormControlKind::ScrollBar:
        case VbaUserFormControlKind::SpinButton:
            // bits: 3 Value is a 32-bit integer, not a string. Skip string payloads.
            break;
        default:
            if (!takeUserFormObjectString(out.captionCount, out.captionData, out.captionBytes, out.captionCompressed,
                                          bytes, has(3), extra, out.semanticEnd)) return failString("caption");
            break;
    }

    // Size field bit depends on the class per MS-OFORMS 2.x:
    //   CommandButton/Label ...... bit 5
    //   TextBox ................. bit 30
    //   CheckBox/OptionButton/ToggleButton ... bit 16
    //   ScrollBar/SpinButton ..... bit 11
    //   ComboBox/ListBox ......... bit 31
    bool hasSize = false;
    switch (kind) {
        case VbaUserFormControlKind::CommandButton:
        case VbaUserFormControlKind::Label:
            hasSize = has(5);
            break;
        case VbaUserFormControlKind::TextBox:
            hasSize = has(30);
            break;
        case VbaUserFormControlKind::CheckBox:
        case VbaUserFormControlKind::OptionButton:
        case VbaUserFormControlKind::ToggleButton:
            hasSize = has(16);
            break;
        case VbaUserFormControlKind::ScrollBar:
        case VbaUserFormControlKind::SpinButton:
            hasSize = has(11);
            break;
        case VbaUserFormControlKind::ComboBox:
        case VbaUserFormControlKind::ListBox:
            hasSize = has(31);
            break;
        default:
            hasSize = has(5);
            break;
    }
    if (hasSize) {
        extra = userFormAlign(extra, 4);
        if (extra + 8 > out.semanticEnd) return fail("Control Size extends past cbControl");
        out.size = extra;
        out.inspection.properties.width = userFormI32(bytes, extra);
        out.inspection.properties.height = userFormI32(bytes, extra + 4);
        extra += 8;
    }
    if (extra > out.semanticEnd) return fail("Control ExtraDataBlock exceeds cbControl");

    auto& props = out.inspection.properties;
    if (out.foreColor) props.foreColor = userFormU32(bytes, *out.foreColor);
    if (out.backColor) props.backColor = userFormU32(bytes, *out.backColor);
    if (out.variousPropertyBits) props.variousPropertyBits = userFormU32(bytes, *out.variousPropertyBits);
    if (out.picturePosition) props.picturePosition = userFormU32(bytes, *out.picturePosition);
    if (out.mousePointer) props.mousePointer = bytes[*out.mousePointer];
    if (out.borderColor) props.borderColor = userFormU32(bytes, *out.borderColor);
    if (out.borderStyle) props.borderStyle = userFormU16(bytes, *out.borderStyle);
    if (out.specialEffect) props.specialEffect = userFormU16(bytes, *out.specialEffect);
    if (out.accelerator) props.accelerator = userFormU16(bytes, *out.accelerator);
    decodeString(out.captionData, out.captionBytes, out.captionCompressed, props.caption);
    decodeString(out.textData, out.textBytes, out.textCompressed, props.text);
    decodeString(out.valueData, out.valueBytes, out.valueCompressed, props.value);
    decodeString(out.groupNameData, out.groupNameBytes, out.groupNameCompressed, props.groupName);
    decodeString(out.columnWidthsData, out.columnWidthsBytes, out.columnWidthsCompressed, props.columnWidths);
    if (out.scrollBars) props.scrollBars = userFormU16(bytes, *out.scrollBars);
    if (out.displayStyle) props.displayStyle = bytes[*out.displayStyle];
    if (out.enterKeyBehavior) props.enterKeyBehavior = bytes[*out.enterKeyBehavior];
    if (out.tabKeyBehavior) props.tabKeyBehavior = bytes[*out.tabKeyBehavior];
    if (out.maxLength) props.maxLength = userFormU32(bytes, *out.maxLength);
    if (out.wordWrap) props.wordWrap = userFormU16(bytes, *out.wordWrap);
    if (out.autoWordSelect) props.autoWordSelect = bytes[*out.autoWordSelect];
    if (out.integralHeight) props.integralHeight = bytes[*out.integralHeight];
    if (out.passwordChar) props.passwordChar = userFormU16(bytes, *out.passwordChar);
    if (out.lineCount) props.lineCount = userFormU32(bytes, *out.lineCount);
    if (out.multiLine) props.multiLine = userFormU16(bytes, *out.multiLine);
    if (out.multiSelect) props.multiSelect = userFormU16(bytes, *out.multiSelect);
    if (out.hideSelection) props.hideSelection = userFormU16(bytes, *out.hideSelection);
    if (out.dataEntry) props.dataEntry = bytes[*out.dataEntry];
    if (out.dragBehavior) props.dragBehavior = bytes[*out.dragBehavior];
    if (out.listRows) props.listRows = userFormU32(bytes, *out.listRows);
    if (out.groupNumber) props.groupNumber = userFormU16(bytes, *out.groupNumber);
    if (out.tripleState) props.tripleState = userFormU16(bytes, *out.tripleState);
    if (out.min) props.min = userFormU32(bytes, *out.min);
    if (out.max) props.max = userFormU32(bytes, *out.max);
    if (out.smallChange) props.smallChange = userFormU32(bytes, *out.smallChange);
    if (out.largeChange) props.largeChange = userFormU32(bytes, *out.largeChange);
    if (out.orientation) props.orientation = bytes[*out.orientation];
    if (out.listWidth) props.listWidth = userFormU32(bytes, *out.listWidth);
    if (out.boundColumn) props.boundColumn = userFormU32(bytes, *out.boundColumn);
    if (out.textColumn) props.textColumn = userFormU32(bytes, *out.textColumn);
    if (out.columnCount) props.columnCount = userFormU32(bytes, *out.columnCount);
    if (out.style) props.style = bytes[*out.style];
    if (out.listStyle) props.listStyle = bytes[*out.listStyle];
    if (out.matchEntry) props.matchEntry = bytes[*out.matchEntry];
    if (out.showDropButtonWhen) props.showDropButtonWhen = bytes[*out.showDropButtonWhen];
    if (out.dropButtonStyle) props.dropButtonStyle = bytes[*out.dropButtonStyle];
    if (out.matchFound) props.matchFound = bytes[*out.matchFound];
    if (out.iMEMode) props.iMEMode = bytes[*out.iMEMode];
    if (out.iMEStatus) props.iMEStatus = bytes[*out.iMEStatus];
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

namespace {

void ufPutU8(std::vector<unsigned char>& out, std::uint8_t value) { out.push_back(value); }
void ufPutU16(std::vector<unsigned char>& out, std::uint16_t value) {
    out.push_back(static_cast<unsigned char>(value & 0xFFu));
    out.push_back(static_cast<unsigned char>((value >> 8) & 0xFFu));
}
void ufPutU32(std::vector<unsigned char>& out, std::uint32_t value) {
    ufPutU16(out, static_cast<std::uint16_t>(value & 0xFFFFu));
    ufPutU16(out, static_cast<std::uint16_t>((value >> 16) & 0xFFFFu));
}
void ufPutI32(std::vector<unsigned char>& out, std::int32_t value) { ufPutU32(out, static_cast<std::uint32_t>(value)); }

struct UfString {
    bool present{false};
    std::vector<unsigned char> bytes;
    std::uint32_t count{0};
};
UfString ufEncodeString(const std::string& text) {
    UfString field;
    auto [payload, compressed] = encodeUserFormString(text);
    field.bytes = std::move(payload);
    field.count = (compressed ? 0x80000000u : 0u) | static_cast<std::uint32_t>(field.bytes.size());
    field.present = true;
    return field;
}

std::uint16_t ufClsidCacheIndex(VbaUserFormControlKind kind) {
    switch (kind) {
        case VbaUserFormControlKind::Form: return 7;
        case VbaUserFormControlKind::Image: return 12;
        case VbaUserFormControlKind::Frame: return 14;
        case VbaUserFormControlKind::MorphData: return 15;
        case VbaUserFormControlKind::SpinButton: return 16;
        case VbaUserFormControlKind::CommandButton: return 17;
        case VbaUserFormControlKind::TabStrip: return 18;
        case VbaUserFormControlKind::Label: return 21;
        case VbaUserFormControlKind::TextBox: return 23;
        case VbaUserFormControlKind::ListBox: return 24;
        case VbaUserFormControlKind::ComboBox: return 25;
        case VbaUserFormControlKind::CheckBox: return 26;
        case VbaUserFormControlKind::OptionButton: return 27;
        case VbaUserFormControlKind::ToggleButton: return 28;
        case VbaUserFormControlKind::ScrollBar: return 47;
        case VbaUserFormControlKind::MultiPage: return 57;
        default: return 0;
    }
}

// Serialize the control-object stream for one control (the "o" stream slice).
std::vector<unsigned char> ufBuildControlObject(const VbaUserFormControlDesign& c) {
    const auto kind = c.kind;
    const bool mask64 = kind == VbaUserFormControlKind::ComboBox || kind == VbaUserFormControlKind::ListBox;
    const auto caption = ufEncodeString(c.caption);
    const auto text = ufEncodeString(c.text);
    const auto value = ufEncodeString(c.value ? "True" : "False");
    const auto groupName = ufEncodeString("");
    const auto columnWidths = ufEncodeString("");

    std::vector<unsigned char> body;
    std::uint32_t mask = 0;
    std::uint32_t maskHigh = 0;
    auto setBit = [&](unsigned bit) {
        if (bit < 32) mask |= (1u << bit);
        else maskHigh |= (1u << (bit - 32));
    };
    setBit(0); setBit(1); setBit(2);
    ufPutU32(body, mask);                 // placeholder; patched after the switch
    if (mask64) ufPutU32(body, maskHigh);

    // Declare every field this class serializes; the emitter below follows the
    // same ascending flag-bit order as the reader's `take` calls.
    auto align = [&](std::size_t size) {
        const std::size_t pad = std::min<std::size_t>(size, 4);
        while (body.size() % pad != 0) body.push_back(0);
    };
    auto rel = [&]() -> std::size_t { return body.size(); };

    switch (kind) {
        case VbaUserFormControlKind::Label:
            setBit(3); setBit(6); setBit(7); setBit(8); setBit(9); setBit(11);
            align(4); ufPutU32(body, c.foreColor);
            align(4); ufPutU32(body, c.backColor);
            align(4); ufPutU32(body, c.variousPropertyBits);
            align(4); ufPutU32(body, caption.count);
            align(1); ufPutU8(body, 0);
            align(4); ufPutU32(body, c.borderColor);
            align(2); ufPutU16(body, 0);
            align(2); ufPutU16(body, 0);
            align(2); ufPutU16(body, 0);
            break;
        case VbaUserFormControlKind::CommandButton:
            setBit(3); setBit(4); setBit(6); setBit(8);
            align(4); ufPutU32(body, c.foreColor);
            align(4); ufPutU32(body, c.backColor);
            align(4); ufPutU32(body, c.variousPropertyBits);
            align(4); ufPutU32(body, caption.count);
            align(4); ufPutU32(body, 0);                   // bit4 picturePosition
            align(1); ufPutU8(body, 0);                    // bit6 mousePointer
            align(2); ufPutU16(body, 0);                   // bit8 accelerator
            break;
        case VbaUserFormControlKind::TextBox:
            setBit(3); setBit(4); setBit(5); setBit(6); setBit(8); setBit(9); setBit(10);
            setBit(11); setBit(13); setBit(14); setBit(15); setBit(16); setBit(17);
            setBit(20); setBit(21); setBit(22); setBit(23); setBit(24); setBit(25); setBit(26);
            setBit(27); setBit(28); setBit(29);
            align(4); ufPutU32(body, c.foreColor);
            align(4); ufPutU32(body, c.backColor);
            align(4); ufPutU32(body, c.variousPropertyBits);
            align(4); ufPutU32(body, caption.count);
            align(4); ufPutU32(body, 0);                   // bit4 picturePosition
            align(2); ufPutU16(body, 0);                   // bit5 specialEffect
            align(1); ufPutU8(body, 0);                    // bit6 mousePointer
            align(4); ufPutU32(body, c.borderColor);       // bit8 borderColor
            align(2); ufPutU16(body, 0);                   // bit9 borderStyle
            align(2); ufPutU16(body, 0);                   // bit10 scrollBars
            align(1); ufPutU8(body, 0);                    // bit11 displayStyle
            align(1); ufPutU8(body, 0);                    // bit13 enterKeyBehavior
            align(1); ufPutU8(body, 0);                    // bit14 tabKeyBehavior
            align(4); ufPutU32(body, c.maxLength);         // bit15 maxLength
            align(2); ufPutU16(body, c.wordWrap ? 1 : 0);  // bit16 wordWrap
            align(4); ufPutU32(body, text.count);          // bit17 Text
            align(1); ufPutU8(body, 0);                    // bit20 autoWordSelect
            align(1); ufPutU8(body, 0);                    // bit21 integralHeight
            align(2); ufPutU16(body, c.passwordChar ? 42 : 0); // bit22 passwordChar
            align(4); ufPutU32(body, value.count);         // bit23 Value
            align(4); ufPutU32(body, 0);                   // bit24 lineCount
            align(2); ufPutU16(body, c.multiLine ? 1 : 0); // bit25 multiLine
            align(2); ufPutU16(body, 0);                   // bit26 multiSelect
            align(2); ufPutU16(body, 1);                   // bit27 hideSelection
            align(1); ufPutU8(body, 0);                    // bit28 dataEntry
            align(1); ufPutU8(body, 0);                    // bit29 dragBehavior
            break;
        case VbaUserFormControlKind::CheckBox:
        case VbaUserFormControlKind::OptionButton:
        case VbaUserFormControlKind::ToggleButton:
            setBit(3); setBit(4); setBit(5); setBit(6); setBit(8); setBit(9); setBit(10);
            setBit(11); setBit(13); setBit(14); setBit(15);
            align(4); ufPutU32(body, c.foreColor);
            align(4); ufPutU32(body, c.backColor);
            align(4); ufPutU32(body, c.variousPropertyBits);
            align(4); ufPutU32(body, caption.count);
            align(4); ufPutU32(body, 0);                   // bit4 picturePosition
            align(2); ufPutU16(body, 0);                   // bit5 specialEffect
            align(1); ufPutU8(body, 0);                    // bit6 mousePointer
            align(4); ufPutU32(body, c.borderColor);       // bit8 borderColor
            align(2); ufPutU16(body, 0);                   // bit9 borderStyle
            align(4); ufPutU32(body, groupName.count);     // bit10 GroupName
            align(2); ufPutU16(body, 0);                   // bit11 accelerator
            align(4); ufPutU32(body, value.count);         // bit13 Value
            align(2); ufPutU16(body, c.groupName ? 1 : 0); // bit14 groupNumber
            align(2); ufPutU16(body, 0);                   // bit15 tripleState
            break;
        case VbaUserFormControlKind::ComboBox:
        case VbaUserFormControlKind::ListBox:
            setBit(3); setBit(4); setBit(5); setBit(6); setBit(8); setBit(9); setBit(10);
            setBit(11); setBit(15); setBit(16); setBit(17); setBit(18); setBit(19);
            setBit(20); setBit(21); setBit(22); setBit(23); setBit(24); setBit(25);
            setBit(26); setBit(27); setBit(28); setBit(29); setBit(30);
            align(4); ufPutU32(body, c.foreColor);
            align(4); ufPutU32(body, c.backColor);
            align(4); ufPutU32(body, c.variousPropertyBits);
            align(4); ufPutU32(body, caption.count);
            align(4); ufPutU32(body, 0);                   // bit4 picturePosition
            align(2); ufPutU16(body, 0);                   // bit5 specialEffect
            align(1); ufPutU8(body, 0);                    // bit6 mousePointer
            align(4); ufPutU32(body, c.borderColor);       // bit8 borderColor
            align(2); ufPutU16(body, 0);                   // bit9 borderStyle
            align(2); ufPutU16(body, 0);                   // bit10 scrollBars
            align(1); ufPutU8(body, 0);                    // bit11 displayStyle
            align(4); ufPutU32(body, 0);                   // bit15 listRows
            align(4); ufPutU32(body, 0);                   // bit16 listWidth
            align(4); ufPutU32(body, 0);                   // bit17 boundColumn
            align(4); ufPutU32(body, 0);                   // bit18 textColumn
            align(4); ufPutU32(body, 0);                   // bit19 columnCount
            align(4); ufPutU32(body, columnWidths.count);  // bit20 columnWidths
            align(1); ufPutU8(body, 0);                    // bit21 style
            align(1); ufPutU8(body, 0);                    // bit22 listStyle
            align(1); ufPutU8(body, 0);                    // bit23 matchEntry
            align(1); ufPutU8(body, 0);                    // bit24 showDropButtonWhen
            align(1); ufPutU8(body, 0);                    // bit25 dropButtonStyle
            align(2); ufPutU16(body, 0);                   // bit26 multiSelect
            align(4); ufPutU32(body, value.count);         // bit27 Value
            align(1); ufPutU8(body, 0);                    // bit28 matchFound
            align(1); ufPutU8(body, 0);                    // bit29 iMEMode
            align(1); ufPutU8(body, 0);                    // bit30 iMEStatus
            break;
        default:
            break;
    }

    // ExtraDataBlock: string payloads in ascending flag-bit order, aligned.
    userFormPutU32(body, 0, mask);
    if (mask64) userFormPutU32(body, 4, maskHigh);
    auto extraStart = rel();
    const auto pad = extraStart % 4;
    extraStart += pad == 0 ? 0 : (4 - pad);
    body.resize(extraStart);
    auto emitPayload = [&](const UfString& field) {
        if (!field.present) return;
        while (body.size() % 4 != 0) body.push_back(0);
        body.insert(body.end(), field.bytes.begin(), field.bytes.end());
    };
    switch (kind) {
        case VbaUserFormControlKind::TextBox:
            emitPayload(caption); emitPayload(text); emitPayload(value);
            break;
        case VbaUserFormControlKind::CheckBox:
        case VbaUserFormControlKind::OptionButton:
        case VbaUserFormControlKind::ToggleButton:
            emitPayload(caption); emitPayload(groupName); emitPayload(value);
            break;
        case VbaUserFormControlKind::ComboBox:
        case VbaUserFormControlKind::ListBox:
            emitPayload(caption); emitPayload(columnWidths); emitPayload(value);
            break;
        default:
            emitPayload(caption);
            break;
    }

    std::vector<unsigned char> out;
    out.push_back(0); out.push_back(2);                    // minor/major
    ufPutU16(out, static_cast<std::uint16_t>(body.size())); // cbControl
    out.insert(out.end(), body.begin(), body.end());
    return out;
}

// Serialize the FormControl + FormSiteData payload of stream "f".
std::vector<unsigned char> ufBuildFormStream(const VbaUserFormDesign& design,
                                             const std::vector<std::vector<unsigned char>>& objects) {
    const auto caption = ufEncodeString(design.caption);
    const std::uint32_t nextId = static_cast<std::uint32_t>(design.controls.size() + 1);

    std::vector<unsigned char> body;
    std::uint32_t mask = 0;
    auto setBit = [&](unsigned bit) { mask |= (1u << bit); };
    setBit(1); setBit(2); setBit(3); setBit(6); setBit(7); setBit(8); setBit(9);
    setBit(10); setBit(11); setBit(16); setBit(17); setBit(18); setBit(19);
    ufPutU32(body, mask);

    auto align = [&](std::size_t size) {
        const std::size_t pad = std::min<std::size_t>(size, 4);
        while (body.size() % pad != 0) body.push_back(0);
    };
    align(4); ufPutU32(body, design.backColor);      // bit1
    align(4); ufPutU32(body, design.foreColor);      // bit2
    align(4); ufPutU32(body, nextId);                // bit3
    align(4); ufPutU32(body, design.booleanProperties); // bit6
    align(1); ufPutU8(body, design.borderStyle);     // bit7
    align(1); ufPutU8(body, 0);                      // bit8 mousePointer
    align(1); ufPutU8(body, design.scrollBars);      // bit9
    align(1); ufPutU8(body, design.cycle);           // bit16
    align(1); ufPutU8(body, design.specialEffect);   // bit17
    align(4); ufPutU32(body, design.borderColor);    // bit18
    align(4); ufPutU32(body, caption.count);         // bit19
    // ExtraDataBlock (fixed-size fields, no alignment between them).
    const auto extraStart = body.size();
    while (body.size() % 4 != 0) body.push_back(0);
    ufPutI32(body, design.width);                    // bit10 displayedSize
    ufPutI32(body, design.height);
    ufPutI32(body, design.width);                    // bit11 logicalSize
    ufPutI32(body, design.height);
    body.insert(body.end(), caption.bytes.begin(), caption.bytes.end());

    std::vector<unsigned char> stream;
    stream.push_back(0); stream.push_back(2);
    const auto cbForm = static_cast<std::size_t>(body.size());
    ufPutU16(stream, static_cast<std::uint16_t>(cbForm));
    stream.insert(stream.end(), body.begin(), body.end());

    // FormSiteData. The reader computes bodyStart as countSites + 8 and
    // requires bodyStart + countBytes == stream size, so countBytes covers the
    // depth bytes (aligned) plus every site record.
    const auto formSiteDataStart = stream.size();
    ufPutU16(stream, 0);                             // classInfoCount (none)
    ufPutU32(stream, static_cast<std::uint32_t>(objects.size()));
    const auto countBytesPos = stream.size();
    ufPutU32(stream, 0);                             // patched below
    const auto depthStart = stream.size();
    for (std::size_t i = 0; i < objects.size(); ++i) {
        stream.push_back(1);                         // depth
        stream.push_back(1);                         // type 1 (single site)
    }
    const auto bodyStart = countBytesPos + 4;
    while ((stream.size() - bodyStart) % 4 != 0) stream.push_back(0);
    for (std::size_t i = 0; i < objects.size(); ++i) {
        const auto& control = design.controls[i];
        const auto& object = objects[i];
        const auto name = ufEncodeString(control.name);
        const auto tip = ufEncodeString(control.controlTipText);
        const auto clsid = ufClsidCacheIndex(control.kind);

        std::vector<unsigned char> site;
        std::uint32_t siteMask = 0;
        auto setBit = [&](unsigned bit) { siteMask |= (1u << bit); };
        setBit(0); setBit(2); setBit(3); setBit(4); setBit(5); setBit(6); setBit(7);
        setBit(8); setBit(9); setBit(11);
        ufPutU32(site, siteMask);
        auto align = [&](std::size_t size) {
            const std::size_t pad = std::min<std::size_t>(size, 4);
            while (site.size() % pad != 0) site.push_back(0);
        };
        align(4); ufPutU32(site, name.count);        // bit0 Name
        align(4); ufPutI32(site, static_cast<std::int32_t>(i + 1)); // bit2 Id
        align(4); ufPutU32(site, 0);                 // bit3 HelpContextId
        align(4); ufPutU32(site, 0);                 // bit4 BitFlags
        align(4); ufPutU32(site, static_cast<std::uint32_t>(object.size())); // bit5 ObjectStreamSize
        align(2); ufPutU16(site, static_cast<std::uint16_t>(control.tabIndex)); // bit6
        align(2); ufPutU16(site, clsid);             // bit7 ClsidCacheIndex
        align(2); ufPutU16(site, 0);                 // bit9 GroupId
        align(4); ufPutU32(site, tip.count);         // bit11 ControlTipText
        while (site.size() % 4 != 0) site.push_back(0);
        site.insert(site.end(), name.bytes.begin(), name.bytes.end());
        ufPutI32(site, control.top);                 // bit8 Position
        ufPutI32(site, control.left);
        site.insert(site.end(), tip.bytes.begin(), tip.bytes.end());

        std::vector<unsigned char> siteRecord;
        ufPutU16(siteRecord, 0);                     // version
        ufPutU16(siteRecord, static_cast<std::uint16_t>(site.size())); // cbSite
        siteRecord.insert(siteRecord.end(), site.begin(), site.end());
        stream.insert(stream.end(), siteRecord.begin(), siteRecord.end());
    }
    userFormPutU32(stream, countBytesPos, static_cast<std::uint32_t>(stream.size() - (countBytesPos + 4)));
    return stream;
}

} // namespace

VbaDesignerStorage buildUserFormDesign(const VbaUserFormDesign& design) {
    if (design.name.empty()) throw std::invalid_argument("UserForm design requires a module name");
    std::vector<std::vector<unsigned char>> objects;
    objects.reserve(design.controls.size());
    for (const auto& control : design.controls) objects.push_back(ufBuildControlObject(control));
    const auto formStream = ufBuildFormStream(design, objects);

    std::vector<unsigned char> objectStream;
    for (const auto& object : objects) objectStream.insert(objectStream.end(), object.begin(), object.end());

    // "f3" FormControlEx: header + empty FormControlEx body.
    std::vector<unsigned char> f3;
    f3.push_back(0); f3.push_back(2);
    ufPutU16(f3, 0x1C); // cbHeader
    ufPutU16(f3, 4);    // cbData
    ufPutU32(f3, 0);    // FormControlEx PropMask (no optional properties)

    // "Designer" Designer_CompData: FormStreamId + CompressedName + flags.
    const auto formName = ufEncodeString(design.name);
    std::vector<unsigned char> designer;
    ufPutU32(designer, 0x000F0301u);                 // Designer_FormStreamId
    ufPutU32(designer, formName.count);
    designer.insert(designer.end(), formName.bytes.begin(), formName.bytes.end());
    ufPutU32(designer, 0);                           // Designer_Flags
    ufPutU32(designer, 0);                           // UserFormProperties: none

    VbaDesignerStorage storage;
    storage.name = design.name;
    storage.streams.push_back({"f", std::move(formStream)});
    storage.streams.push_back({"o", std::move(objectStream)});
    storage.streams.push_back({"f3", std::move(f3)});
    storage.streams.push_back({"Designer", std::move(designer)});
    return storage;
}

} // namespace internal
} // namespace xlpp

