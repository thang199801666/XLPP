#include "VbaProjectBinary.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cstdint>
#include <cstring>
#include <limits>
#include <map>
#include <stdexcept>
#include <string_view>
#include <unordered_map>
#include <utility>

namespace xlpp::internal {
namespace {

constexpr std::uint32_t kFreeSect = 0xFFFFFFFFu;
constexpr std::uint32_t kEndOfChain = 0xFFFFFFFEu;
constexpr std::uint32_t kFatSect = 0xFFFFFFFDu;
constexpr std::uint32_t kDifSect = 0xFFFFFFFCu;
constexpr std::size_t kSectorSize = 512;
constexpr std::size_t kMiniSectorSize = 64;
constexpr std::size_t kMiniCutoff = 4096;

void putU16(std::vector<unsigned char>& out, std::uint16_t value) {
    out.push_back(static_cast<unsigned char>(value & 0xFFu));
    out.push_back(static_cast<unsigned char>((value >> 8u) & 0xFFu));
}

void putU32(std::vector<unsigned char>& out, std::uint32_t value) {
    for (unsigned shift = 0; shift < 32; shift += 8)
        out.push_back(static_cast<unsigned char>((value >> shift) & 0xFFu));
}


std::uint16_t getU16(const unsigned char* p) {
    return static_cast<std::uint16_t>(p[0]) | (static_cast<std::uint16_t>(p[1]) << 8u);
}

std::uint32_t getU32(const unsigned char* p) {
    return static_cast<std::uint32_t>(p[0]) |
           (static_cast<std::uint32_t>(p[1]) << 8u) |
           (static_cast<std::uint32_t>(p[2]) << 16u) |
           (static_cast<std::uint32_t>(p[3]) << 24u);
}

std::uint64_t getU64(const unsigned char* p) {
    std::uint64_t value = 0;
    for (unsigned shift = 0; shift < 64; shift += 8)
        value |= static_cast<std::uint64_t>(p[shift / 8]) << shift;
    return value;
}


int hexNibble(char ch) {
    if (ch >= '0' && ch <= '9') return ch - '0';
    if (ch >= 'a' && ch <= 'f') return ch - 'a' + 10;
    if (ch >= 'A' && ch <= 'F') return ch - 'A' + 10;
    return -1;
}

std::array<unsigned char, 16> parseGuidBytes(std::string text) {
    if (!text.empty() && text.front() == '{' && text.back() == '}') text = text.substr(1, text.size() - 2);
    std::string digits;
    digits.reserve(32);
    for (const auto ch : text) if (ch != '-') digits.push_back(ch);
    if (digits.size() != 32) throw std::invalid_argument("VBA control reference originalTypeLib must be a GUID");
    std::array<unsigned char, 16> canonical{};
    for (std::size_t i = 0; i < canonical.size(); ++i) {
        const auto hi = hexNibble(digits[i * 2]); const auto lo = hexNibble(digits[i * 2 + 1]);
        if (hi < 0 || lo < 0) throw std::invalid_argument("VBA control reference originalTypeLib contains invalid GUID hex");
        canonical[i] = static_cast<unsigned char>((hi << 4) | lo);
    }
    // GUID wire/storage layout is little-endian for Data1/Data2/Data3.
    return {canonical[3], canonical[2], canonical[1], canonical[0], canonical[5], canonical[4], canonical[7], canonical[6],
            canonical[8], canonical[9], canonical[10], canonical[11], canonical[12], canonical[13], canonical[14], canonical[15]};
}

std::string guidTextFromBytes(const unsigned char* p) {
    const std::array<unsigned char, 16> canonical{p[3], p[2], p[1], p[0], p[5], p[4], p[7], p[6],
                                                  p[8], p[9], p[10], p[11], p[12], p[13], p[14], p[15]};
    static constexpr char hex[] = "0123456789ABCDEF";
    std::string out = "{";
    for (std::size_t i = 0; i < canonical.size(); ++i) {
        if (i == 4 || i == 6 || i == 8 || i == 10) out.push_back('-');
        out.push_back(hex[(canonical[i] >> 4) & 0xF]); out.push_back(hex[canonical[i] & 0xF]);
    }
    out.push_back('}');
    return out;
}

void overwriteU16(std::vector<unsigned char>& out, std::size_t offset, std::uint16_t value) {
    out.at(offset) = static_cast<unsigned char>(value & 0xFFu);
    out.at(offset + 1) = static_cast<unsigned char>((value >> 8u) & 0xFFu);
}

void overwriteU32(std::vector<unsigned char>& out, std::size_t offset, std::uint32_t value) {
    for (unsigned shift = 0; shift < 32; shift += 8)
        out.at(offset + shift / 8) = static_cast<unsigned char>((value >> shift) & 0xFFu);
}

void appendBytes(std::vector<unsigned char>& out, std::string_view value) {
    out.insert(out.end(), value.begin(), value.end());
}

std::vector<unsigned char> utf16Le(std::string_view ascii, bool terminatingNull = false) {
    std::vector<unsigned char> result;
    result.reserve((ascii.size() + (terminatingNull ? 1u : 0u)) * 2u);
    for (unsigned char ch : ascii) {
        result.push_back(ch);
        result.push_back(0);
    }
    if (terminatingNull) {
        result.push_back(0);
        result.push_back(0);
    }
    return result;
}

std::string asciiLower(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    return value;
}

void appendFixedRecord(std::vector<unsigned char>& out, std::uint16_t id, std::uint32_t value) {
    putU16(out, id);
    putU32(out, 4);
    putU32(out, value);
}

void appendSizedStringRecord(std::vector<unsigned char>& out, std::uint16_t id, std::string_view value) {
    putU16(out, id);
    putU32(out, static_cast<std::uint32_t>(value.size()));
    appendBytes(out, value);
}

void appendDualStringRecord(std::vector<unsigned char>& out, std::uint16_t id, std::uint16_t reserved,
                            std::string_view value) {
    putU16(out, id);
    putU32(out, static_cast<std::uint32_t>(value.size()));
    appendBytes(out, value);
    putU16(out, reserved);
    const auto wide = utf16Le(value);
    putU32(out, static_cast<std::uint32_t>(wide.size()));
    out.insert(out.end(), wide.begin(), wide.end());
}

std::vector<unsigned char> compressOvba(std::string_view source) {
    std::vector<unsigned char> out;
    out.push_back(0x01);
    std::size_t offset = 0;
    constexpr std::size_t kLiteralChunk = 3000;
    while (offset < source.size()) {
        const auto count = std::min(kLiteralChunk, source.size() - offset);
        std::vector<unsigned char> payload;
        payload.reserve(count + (count + 7) / 8);
        std::size_t local = 0;
        while (local < count) {
            payload.push_back(0); // eight literal tokens
            const auto group = std::min<std::size_t>(8, count - local);
            for (std::size_t i = 0; i < group; ++i)
                payload.push_back(static_cast<unsigned char>(source[offset + local + i]));
            local += group;
        }
        const auto totalSize = payload.size() + 2;
        if (totalSize > 4095) throw std::runtime_error("VBA compressed chunk overflow");
        const auto header = static_cast<std::uint16_t>(0xB000u | static_cast<std::uint16_t>(totalSize - 3));
        putU16(out, header);
        out.insert(out.end(), payload.begin(), payload.end());
        offset += count;
    }
    if (source.empty()) {
        // A valid empty compressed container still carries one compressed chunk.
        putU16(out, 0xB000u);
    }
    return out;
}

std::string decompressOvba(const std::vector<unsigned char>& bytes, std::size_t start = 0) {
    if (start >= bytes.size() || bytes[start] != 0x01)
        throw std::runtime_error("Invalid VBA compressed-container signature");
    std::string out;
    std::size_t pos = start + 1;
    while (pos + 2 <= bytes.size()) {
        const auto header = getU16(bytes.data() + pos);
        const auto signature = static_cast<unsigned>((header >> 12u) & 0x7u);
        if (signature != 0x3u) throw std::runtime_error("Invalid VBA compressed-chunk signature");
        const auto chunkSize = static_cast<std::size_t>(header & 0x0FFFu) + 3u;
        if (chunkSize < 3 || pos + chunkSize > bytes.size())
            throw std::runtime_error("Truncated VBA compressed chunk");
        const auto chunkEnd = pos + chunkSize;
        const bool compressed = (header & 0x8000u) != 0;
        pos += 2;
        const auto chunkOutputStart = out.size();
        if (!compressed) {
            out.append(reinterpret_cast<const char*>(bytes.data() + pos), chunkEnd - pos);
            pos = chunkEnd;
            continue;
        }
        while (pos < chunkEnd) {
            const auto flags = bytes[pos++];
            for (unsigned bit = 0; bit < 8 && pos < chunkEnd; ++bit) {
                if ((flags & (1u << bit)) == 0) {
                    out.push_back(static_cast<char>(bytes[pos++]));
                    continue;
                }
                if (pos + 2 > chunkEnd) throw std::runtime_error("Truncated VBA copy token");
                const auto token = getU16(bytes.data() + pos);
                pos += 2;
                const auto difference = out.size() - chunkOutputStart;
                unsigned bitCount = 4;
                while (bitCount < 12 && (std::size_t{1} << bitCount) < difference) ++bitCount;
                const auto lengthMask = static_cast<std::uint16_t>(0xFFFFu >> bitCount);
                const auto offsetMask = static_cast<std::uint16_t>(~lengthMask);
                const auto length = static_cast<std::size_t>(token & lengthMask) + 3u;
                const auto copyOffset = static_cast<std::size_t>((token & offsetMask) >> (16u - bitCount)) + 1u;
                if (copyOffset == 0 || copyOffset > out.size() - chunkOutputStart)
                    throw std::runtime_error("Invalid VBA copy-token offset");
                for (std::size_t i = 0; i < length; ++i)
                    out.push_back(out[out.size() - copyOffset]);
            }
        }
        pos = chunkEnd;
    }
    return out;
}

struct CfbNode {
    std::string name;
    std::uint8_t type{2};
    int parent{-1};
    std::vector<unsigned char> data;
    std::uint32_t start{kEndOfChain};
    std::uint64_t size{0};
    std::uint32_t left{kFreeSect};
    std::uint32_t right{kFreeSect};
    std::uint32_t child{kFreeSect};
    std::uint8_t color{1};
};

bool cfbNameLess(const CfbNode& lhs, const CfbNode& rhs) {
    const auto left = asciiLower(lhs.name);
    const auto right = asciiLower(rhs.name);
    if (left.size() != right.size()) return left.size() < right.size();
    return left < right;
}

void assignChildTree(std::vector<CfbNode>& nodes, std::uint32_t storageId) {
    std::vector<std::uint32_t> children;
    for (std::uint32_t i = 1; i < nodes.size(); ++i)
        if (nodes[i].parent == static_cast<int>(storageId)) children.push_back(i);

    for (const auto id : children) {
        nodes[id].left = kFreeSect;
        nodes[id].right = kFreeSect;
        nodes[id].color = 0; // newly inserted red node
    }
    if (children.empty()) {
        nodes[storageId].child = kFreeSect;
        return;
    }

    std::vector<std::uint32_t> treeParent(nodes.size(), kFreeSect);
    std::uint32_t root = kFreeSect;

    auto rotateLeft = [&](std::uint32_t x) {
        const auto y = nodes[x].right;
        nodes[x].right = nodes[y].left;
        if (nodes[y].left != kFreeSect) treeParent[nodes[y].left] = x;
        treeParent[y] = treeParent[x];
        if (treeParent[x] == kFreeSect) root = y;
        else if (x == nodes[treeParent[x]].left) nodes[treeParent[x]].left = y;
        else nodes[treeParent[x]].right = y;
        nodes[y].left = x;
        treeParent[x] = y;
    };
    auto rotateRight = [&](std::uint32_t y) {
        const auto x = nodes[y].left;
        nodes[y].left = nodes[x].right;
        if (nodes[x].right != kFreeSect) treeParent[nodes[x].right] = y;
        treeParent[x] = treeParent[y];
        if (treeParent[y] == kFreeSect) root = x;
        else if (y == nodes[treeParent[y]].right) nodes[treeParent[y]].right = x;
        else nodes[treeParent[y]].left = x;
        nodes[x].right = y;
        treeParent[y] = x;
    };

    for (const auto zInitial : children) {
        std::uint32_t parent = kFreeSect;
        auto cursor = root;
        while (cursor != kFreeSect) {
            parent = cursor;
            cursor = cfbNameLess(nodes[zInitial], nodes[cursor]) ? nodes[cursor].left : nodes[cursor].right;
        }
        treeParent[zInitial] = parent;
        if (parent == kFreeSect) root = zInitial;
        else if (cfbNameLess(nodes[zInitial], nodes[parent])) nodes[parent].left = zInitial;
        else nodes[parent].right = zInitial;

        auto z = zInitial;
        while (treeParent[z] != kFreeSect && nodes[treeParent[z]].color == 0) {
            const auto parentId = treeParent[z];
            const auto grandparent = treeParent[parentId];
            if (grandparent == kFreeSect) break;
            if (parentId == nodes[grandparent].left) {
                const auto uncle = nodes[grandparent].right;
                if (uncle != kFreeSect && nodes[uncle].color == 0) {
                    nodes[parentId].color = 1;
                    nodes[uncle].color = 1;
                    nodes[grandparent].color = 0;
                    z = grandparent;
                } else {
                    if (z == nodes[parentId].right) {
                        z = parentId;
                        rotateLeft(z);
                    }
                    const auto newParent = treeParent[z];
                    const auto newGrandparent = treeParent[newParent];
                    nodes[newParent].color = 1;
                    if (newGrandparent != kFreeSect) {
                        nodes[newGrandparent].color = 0;
                        rotateRight(newGrandparent);
                    }
                }
            } else {
                const auto uncle = nodes[grandparent].left;
                if (uncle != kFreeSect && nodes[uncle].color == 0) {
                    nodes[parentId].color = 1;
                    nodes[uncle].color = 1;
                    nodes[grandparent].color = 0;
                    z = grandparent;
                } else {
                    if (z == nodes[parentId].left) {
                        z = parentId;
                        rotateRight(z);
                    }
                    const auto newParent = treeParent[z];
                    const auto newGrandparent = treeParent[newParent];
                    nodes[newParent].color = 1;
                    if (newGrandparent != kFreeSect) {
                        nodes[newGrandparent].color = 0;
                        rotateLeft(newGrandparent);
                    }
                }
            }
        }
        nodes[root].color = 1;
    }
    nodes[storageId].child = root;
}

std::array<unsigned char, 128> directoryEntry(const CfbNode& node) {
    std::array<unsigned char, 128> out{};
    const auto wide = utf16Le(node.name, true);
    if (wide.size() > 64) throw std::invalid_argument("CFB directory name is too long: " + node.name);
    std::copy(wide.begin(), wide.end(), out.begin());
    out[64] = static_cast<unsigned char>(wide.size() & 0xFFu);
    out[65] = static_cast<unsigned char>((wide.size() >> 8u) & 0xFFu);
    out[66] = node.type;
    out[67] = node.color;
    auto write32 = [&](std::size_t offset, std::uint32_t value) {
        for (unsigned shift = 0; shift < 32; shift += 8)
            out[offset + shift / 8] = static_cast<unsigned char>((value >> shift) & 0xFFu);
    };
    auto write64 = [&](std::size_t offset, std::uint64_t value) {
        for (unsigned shift = 0; shift < 64; shift += 8)
            out[offset + shift / 8] = static_cast<unsigned char>((value >> shift) & 0xFFu);
    };
    write32(68, node.left);
    write32(72, node.right);
    write32(76, node.child);
    write32(116, node.start);
    write64(120, node.size);
    return out;
}

std::vector<unsigned char> buildCfb(const std::map<std::string, std::vector<unsigned char>>& rootStreams,
                                    const std::map<std::string, std::vector<unsigned char>>& vbaStreams,
                                    const std::vector<xlpp::VbaDesignerStorage>& designerStorages = {},
                                    bool includeVbaStorage = true) {
    std::vector<CfbNode> nodes;
    nodes.push_back({"Root Entry", 5, -1, {}});
    std::uint32_t vbaStorage = kFreeSect;
    if (includeVbaStorage) {
        vbaStorage = static_cast<std::uint32_t>(nodes.size());
        nodes.push_back({"VBA", 1, 0, {}});
    } else if (!vbaStreams.empty() || !designerStorages.empty()) {
        throw std::invalid_argument("Root-only CFB cannot contain VBA/designer streams");
    }
    for (const auto& [name, data] : rootStreams) nodes.push_back({name, 2, 0, data});
    for (const auto& [name, data] : vbaStreams) nodes.push_back({name, 2, static_cast<int>(vbaStorage), data});

    auto childStorage = [&](int parent, const std::string& name) -> int {
        for (std::size_t i = 0; i < nodes.size(); ++i)
            if (nodes[i].parent == parent && nodes[i].type == 1 && asciiLower(nodes[i].name) == asciiLower(name))
                return static_cast<int>(i);
        nodes.push_back({name, 1, parent, {}});
        return static_cast<int>(nodes.size() - 1);
    };
    auto addHierarchicalStream = [&](int rootStorage, const std::string& relativePath, const std::vector<unsigned char>& data) {
        if (relativePath.empty()) throw std::invalid_argument("VBA designer stream path cannot be empty");
        int parent = rootStorage;
        std::size_t begin = 0;
        while (true) {
            const auto slash = relativePath.find('/', begin);
            const auto component = relativePath.substr(begin, slash == std::string::npos ? std::string::npos : slash - begin);
            if (component.empty() || component == "." || component == "..")
                throw std::invalid_argument("VBA designer stream path contains an invalid component: " + relativePath);
            if (slash == std::string::npos) {
                for (const auto& node : nodes)
                    if (node.parent == parent && asciiLower(node.name) == asciiLower(component))
                        throw std::invalid_argument("Duplicate VBA designer CFB entry: " + relativePath);
                nodes.push_back({component, 2, parent, data});
                break;
            }
            parent = childStorage(parent, component);
            begin = slash + 1;
        }
    };

    for (const auto& storage : designerStorages) {
        if (storage.name.empty()) throw std::invalid_argument("VBA designer storage name cannot be empty");
        if (asciiLower(storage.name) == "vba") throw std::invalid_argument("VBA designer storage cannot use reserved name VBA");
        int storageId = childStorage(0, storage.name);
        for (const auto& stream : storage.streams) addHierarchicalStream(storageId, stream.path, stream.data);
    }

    for (std::size_t i = 0; i < nodes.size(); ++i)
        if (nodes[i].type == 1 || nodes[i].type == 5) assignChildTree(nodes, static_cast<std::uint32_t>(i));

    std::vector<std::uint32_t> miniFat;
    std::vector<unsigned char> miniStream;
    struct RegularBinding { std::uint32_t node; std::vector<unsigned char> data; };
    std::vector<RegularBinding> regular;

    for (std::uint32_t i = 1; i < nodes.size(); ++i) {
        auto& node = nodes[i];
        if (node.type != 2) continue;
        node.size = node.data.size();
        if (node.data.empty()) {
            node.start = kEndOfChain;
            continue;
        }
        if (node.data.size() < kMiniCutoff) {
            const auto first = static_cast<std::uint32_t>(miniFat.size());
            const auto sectors = (node.data.size() + kMiniSectorSize - 1) / kMiniSectorSize;
            node.start = first;
            for (std::size_t s = 0; s < sectors; ++s) {
                miniFat.push_back(s + 1 == sectors ? kEndOfChain : first + static_cast<std::uint32_t>(s + 1));
                const auto from = s * kMiniSectorSize;
                const auto count = std::min(kMiniSectorSize, node.data.size() - from);
                miniStream.insert(miniStream.end(), node.data.begin() + static_cast<std::ptrdiff_t>(from),
                                  node.data.begin() + static_cast<std::ptrdiff_t>(from + count));
                miniStream.resize(miniStream.size() + (kMiniSectorSize - count), 0);
            }
        } else {
            regular.push_back({i, node.data});
        }
    }

    const auto rootMiniSize = miniStream.size();
    nodes[0].size = rootMiniSize;

    const auto dirBytes = ((nodes.size() * 128 + kSectorSize - 1) / kSectorSize) * kSectorSize;
    const auto dirSectors = dirBytes / kSectorSize;
    const auto rootMiniSectors = (miniStream.size() + kSectorSize - 1) / kSectorSize;
    const auto miniFatSectors = miniFat.empty() ? 0 : (miniFat.size() * 4 + kSectorSize - 1) / kSectorSize;
    std::size_t regularSectors = 0;
    for (const auto& binding : regular) regularSectors += (binding.data.size() + kSectorSize - 1) / kSectorSize;
    const auto nonFatSectors = regularSectors + rootMiniSectors + dirSectors + miniFatSectors;
    std::size_t fatSectors = 1;
    std::size_t difatSectors = 0;
    for (;;) {
        const auto nextDifat = fatSectors > 109 ? (fatSectors - 109 + 126) / 127 : 0;
        const auto nextFat = (nonFatSectors + fatSectors + nextDifat + 127) / 128;
        if (nextFat == fatSectors && nextDifat == difatSectors) break;
        fatSectors = nextFat;
        difatSectors = nextDifat;
    }
    const auto totalSectors = nonFatSectors + fatSectors + difatSectors;

    std::vector<std::uint32_t> fat(fatSectors * 128, kFreeSect);
    std::vector<std::array<unsigned char, kSectorSize>> sectors(totalSectors);
    std::size_t nextSector = 0;
    auto placeChain = [&](const std::vector<unsigned char>& data, std::size_t sectorCount) -> std::uint32_t {
        if (sectorCount == 0) return kEndOfChain;
        const auto first = static_cast<std::uint32_t>(nextSector);
        for (std::size_t s = 0; s < sectorCount; ++s) {
            auto& target = sectors[nextSector];
            const auto from = s * kSectorSize;
            const auto count = std::min(kSectorSize, data.size() > from ? data.size() - from : 0u);
            if (count) std::copy_n(data.begin() + static_cast<std::ptrdiff_t>(from), count, target.begin());
            fat[nextSector] = s + 1 == sectorCount ? kEndOfChain : static_cast<std::uint32_t>(nextSector + 1);
            ++nextSector;
        }
        return first;
    };

    for (auto& binding : regular) {
        const auto count = (binding.data.size() + kSectorSize - 1) / kSectorSize;
        nodes[binding.node].start = placeChain(binding.data, count);
    }
    nodes[0].start = placeChain(miniStream, rootMiniSectors);

    std::vector<unsigned char> directoryData(dirBytes, 0);
    for (std::size_t i = 0; i < nodes.size(); ++i) {
        const auto entry = directoryEntry(nodes[i]);
        std::copy(entry.begin(), entry.end(), directoryData.begin() + static_cast<std::ptrdiff_t>(i * 128));
    }
    const auto firstDirectorySector = placeChain(directoryData, dirSectors);

    std::vector<unsigned char> miniFatData(miniFatSectors * kSectorSize, 0xFF);
    for (std::size_t i = 0; i < miniFat.size(); ++i) {
        for (unsigned shift = 0; shift < 32; shift += 8)
            miniFatData[i * 4 + shift / 8] = static_cast<unsigned char>((miniFat[i] >> shift) & 0xFFu);
    }
    const auto firstMiniFatSector = placeChain(miniFatData, miniFatSectors);

    const auto firstFatSector = nextSector;
    for (std::size_t i = 0; i < fatSectors; ++i) fat[nextSector + i] = kFatSect;
    nextSector += fatSectors;
    const auto firstDifatSector = difatSectors ? static_cast<std::uint32_t>(nextSector) : kEndOfChain;
    for (std::size_t i = 0; i < difatSectors; ++i) fat[nextSector + i] = kDifSect;

    // FAT sector contents include the FAT/DIFAT sector markers above.
    for (std::size_t f = 0; f < fatSectors; ++f) {
        auto& target = sectors[firstFatSector + f];
        for (std::size_t j = 0; j < 128; ++j) {
            const auto value = fat[f * 128 + j];
            for (unsigned shift = 0; shift < 32; shift += 8)
                target[j * 4 + shift / 8] = static_cast<unsigned char>((value >> shift) & 0xFFu);
        }
    }

    // Header DIFAT holds the first 109 FAT sector IDs. Additional FAT IDs are
    // chained through DIFAT sectors, 127 IDs plus one next-sector pointer each.
    for (std::size_t d = 0; d < difatSectors; ++d) {
        auto& target = sectors[firstDifatSector + d];
        const auto base = 109 + d * 127;
        for (std::size_t j = 0; j < 127; ++j) {
            const auto index = base + j;
            const auto value = index < fatSectors ? static_cast<std::uint32_t>(firstFatSector + index) : kFreeSect;
            for (unsigned shift = 0; shift < 32; shift += 8)
                target[j * 4 + shift / 8] = static_cast<unsigned char>((value >> shift) & 0xFFu);
        }
        const auto next = d + 1 < difatSectors ? static_cast<std::uint32_t>(firstDifatSector + d + 1) : kEndOfChain;
        for (unsigned shift = 0; shift < 32; shift += 8)
            target[508 + shift / 8] = static_cast<unsigned char>((next >> shift) & 0xFFu);
    }

    std::vector<unsigned char> output(kSectorSize, 0);
    const unsigned char signature[8] = {0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1};
    std::copy(std::begin(signature), std::end(signature), output.begin());
    overwriteU16(output, 24, 0x003E);
    overwriteU16(output, 26, 0x0003);
    overwriteU16(output, 28, 0xFFFE);
    overwriteU16(output, 30, 9);
    overwriteU16(output, 32, 6);
    overwriteU32(output, 40, 0); // v3 directory-sector count
    overwriteU32(output, 44, static_cast<std::uint32_t>(fatSectors));
    overwriteU32(output, 48, firstDirectorySector);
    overwriteU32(output, 52, 0);
    overwriteU32(output, 56, static_cast<std::uint32_t>(kMiniCutoff));
    overwriteU32(output, 60, firstMiniFatSector);
    overwriteU32(output, 64, static_cast<std::uint32_t>(miniFatSectors));
    overwriteU32(output, 68, firstDifatSector);
    overwriteU32(output, 72, static_cast<std::uint32_t>(difatSectors));
    for (std::size_t i = 0; i < 109; ++i)
        overwriteU32(output, 76 + i * 4, i < fatSectors ? static_cast<std::uint32_t>(firstFatSector + i) : kFreeSect);
    for (const auto& sector : sectors) output.insert(output.end(), sector.begin(), sector.end());
    return output;
}

struct ParsedDirectoryEntry {
    std::string name;
    std::uint8_t type{0};
    std::uint32_t left{kFreeSect};
    std::uint32_t right{kFreeSect};
    std::uint32_t child{kFreeSect};
    std::uint32_t start{kEndOfChain};
    std::uint64_t size{0};
};

class CfbReader {
public:
    explicit CfbReader(const std::vector<unsigned char>& bytes) : bytes_(bytes) {
        parse();
    }

    std::vector<unsigned char> stream(const std::string& path) const {
        const auto it = paths_.find(asciiLower(path));
        if (it == paths_.end()) throw std::runtime_error("CFB stream not found: " + path);
        const auto& entry = directory_.at(it->second);
        if (entry.type != 2) throw std::runtime_error("CFB path is not a stream: " + path);
        if (entry.size == 0) return {};
        if (entry.size < miniCutoff_) return readMiniChain(entry.start, entry.size);
        return readRegularChain(entry.start, entry.size);
    }

    std::vector<std::string> rootStorageNames() const {
        std::vector<std::string> result;
        for (std::size_t id = 1; id < directory_.size(); ++id) {
            if (directory_[id].type != 1 || id >= originalPaths_.size()) continue;
            const auto& path = originalPaths_[id];
            if (!path.empty() && path.find('/') == std::string::npos) result.push_back(path);
        }
        std::sort(result.begin(), result.end());
        return result;
    }

    std::vector<std::string> streamPathsUnder(const std::string& storage) const {
        std::vector<std::string> result;
        const auto prefix = storage + "/";
        for (std::size_t id = 1; id < directory_.size(); ++id) {
            if (directory_[id].type != 2 || id >= originalPaths_.size()) continue;
            const auto& path = originalPaths_[id];
            if (path.rfind(prefix, 0) == 0) result.push_back(path);
        }
        std::sort(result.begin(), result.end());
        return result;
    }

private:
    const std::vector<unsigned char>& bytes_;
    std::size_t sectorSize_{512};
    std::size_t miniSectorSize_{64};
    std::size_t miniCutoff_{4096};
    std::vector<std::uint32_t> fat_;
    std::vector<std::uint32_t> miniFat_;
    std::vector<ParsedDirectoryEntry> directory_;
    std::vector<unsigned char> rootMiniStream_;
    std::unordered_map<std::string, std::size_t> paths_;
    std::vector<std::string> originalPaths_;

    const unsigned char* sector(std::uint32_t id) const {
        if (sectorSize_ == 0 || static_cast<std::size_t>(id) >
                (std::numeric_limits<std::size_t>::max() - kSectorSize) / sectorSize_)
            throw std::runtime_error("CFB sector offset overflow");
        const auto offset = kSectorSize + static_cast<std::size_t>(id) * sectorSize_;
        if (offset > bytes_.size() || sectorSize_ > bytes_.size() - offset)
            throw std::runtime_error("CFB sector is outside the file");
        return bytes_.data() + offset;
    }

    std::vector<unsigned char> readRegularChain(std::uint32_t start, std::uint64_t expected) const {
        std::vector<unsigned char> out;
        const bool exactSize = expected != std::numeric_limits<std::uint64_t>::max();
        if (start == kEndOfChain) {
            if (exactSize && expected != 0) throw std::runtime_error("Truncated CFB regular stream");
            return out;
        }
        std::uint32_t id = start;
        std::size_t guard = 0;
        while (id != kEndOfChain && id != kFreeSect) {
            if (id >= fat_.size()) throw std::runtime_error("Invalid CFB FAT chain");
            const auto* p = sector(id);
            out.insert(out.end(), p, p + sectorSize_);
            id = fat_[id];
            if (++guard > fat_.size()) throw std::runtime_error("Cyclic CFB FAT chain");
        }
        if (exactSize) {
            if (expected > static_cast<std::uint64_t>(out.size()))
                throw std::runtime_error("Truncated CFB regular stream");
            if (expected < out.size()) out.resize(static_cast<std::size_t>(expected));
        }
        return out;
    }

    std::vector<unsigned char> readMiniChain(std::uint32_t start, std::uint64_t expected) const {
        std::vector<unsigned char> out;
        if (start == kEndOfChain) {
            if (expected != 0) throw std::runtime_error("Truncated CFB mini stream");
            return out;
        }
        std::uint32_t id = start;
        std::size_t guard = 0;
        while (id != kEndOfChain && id != kFreeSect && out.size() < expected) {
            if (id >= miniFat_.size()) throw std::runtime_error("Invalid CFB miniFAT chain");
            const auto offset = static_cast<std::size_t>(id) * miniSectorSize_;
            if (offset + miniSectorSize_ > rootMiniStream_.size())
                throw std::runtime_error("CFB mini sector is outside the root mini stream");
            out.insert(out.end(), rootMiniStream_.begin() + static_cast<std::ptrdiff_t>(offset),
                       rootMiniStream_.begin() + static_cast<std::ptrdiff_t>(offset + miniSectorSize_));
            id = miniFat_[id];
            if (++guard > miniFat_.size()) throw std::runtime_error("Cyclic CFB miniFAT chain");
        }
        if (expected > static_cast<std::uint64_t>(out.size()))
            throw std::runtime_error("Truncated CFB mini stream");
        if (expected < out.size()) out.resize(static_cast<std::size_t>(expected));
        return out;
    }

    void walkTree(std::uint32_t id, const std::string& prefix, std::vector<bool>& visited) {
        if (id == kFreeSect || id == kEndOfChain) return;
        if (id >= directory_.size()) throw std::runtime_error("CFB directory tree references an invalid entry");
        if (visited[id]) throw std::runtime_error("Cyclic CFB directory tree");
        visited[id] = true;
        const auto entry = directory_[id];
        if (entry.type != 1 && entry.type != 2 && entry.type != 5)
            throw std::runtime_error("CFB directory tree references an invalid entry type");
        if (entry.name.empty()) throw std::runtime_error("CFB directory entry has an empty name");
        walkTree(entry.left, prefix, visited);
        const auto path = prefix.empty() ? entry.name : prefix + "/" + entry.name;
        if (!paths_.emplace(asciiLower(path), id).second)
            throw std::runtime_error("Duplicate CFB directory path: " + path);
        if (originalPaths_.size() < directory_.size()) originalPaths_.resize(directory_.size());
        originalPaths_[id] = path;
        if (entry.type == 1 || entry.type == 5) walkTree(entry.child, path == "Root Entry" ? "" : path, visited);
        walkTree(entry.right, prefix, visited);
    }

    void parse() {
        if (bytes_.size() < 512) throw std::runtime_error("CFB file is truncated");
        const unsigned char signature[8] = {0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1};
        if (!std::equal(std::begin(signature), std::end(signature), bytes_.begin()))
            throw std::runtime_error("Invalid CFB signature");
        sectorSize_ = std::size_t{1} << getU16(bytes_.data() + 30);
        miniSectorSize_ = std::size_t{1} << getU16(bytes_.data() + 32);
        miniCutoff_ = getU32(bytes_.data() + 56);
        const auto fatCount = getU32(bytes_.data() + 44);
        const auto firstDirectory = getU32(bytes_.data() + 48);
        const auto firstMiniFat = getU32(bytes_.data() + 60);
        const auto miniFatCount = getU32(bytes_.data() + 64);
        const auto firstDifat = getU32(bytes_.data() + 68);
        const auto difatCount = getU32(bytes_.data() + 72);
        if (sectorSize_ != 512 || miniSectorSize_ != 64)
            throw std::runtime_error("Unsupported CFB sector geometry");

        std::vector<std::uint32_t> fatSectors;
        for (std::size_t i = 0; i < 109 && fatSectors.size() < fatCount; ++i) {
            const auto id = getU32(bytes_.data() + 76 + i * 4);
            if (id != kFreeSect) fatSectors.push_back(id);
        }
        std::uint32_t difat = firstDifat;
        for (std::uint32_t d = 0; d < difatCount && difat != kEndOfChain; ++d) {
            const auto* p = sector(difat);
            for (std::size_t i = 0; i < sectorSize_ / 4 - 1 && fatSectors.size() < fatCount; ++i) {
                const auto id = getU32(p + i * 4);
                if (id != kFreeSect) fatSectors.push_back(id);
            }
            difat = getU32(p + sectorSize_ - 4);
        }
        if (fatSectors.size() < fatCount) throw std::runtime_error("CFB FAT sector list is incomplete");
        for (auto id : fatSectors) {
            const auto* p = sector(id);
            for (std::size_t i = 0; i < sectorSize_ / 4; ++i) fat_.push_back(getU32(p + i * 4));
        }

        const auto directoryBytes = readRegularChain(firstDirectory, std::numeric_limits<std::uint64_t>::max());
        for (std::size_t offset = 0; offset + 128 <= directoryBytes.size(); offset += 128) {
            const auto* p = directoryBytes.data() + offset;
            const auto nameBytes = getU16(p + 64);
            ParsedDirectoryEntry entry;
            if (nameBytes >= 2 && nameBytes <= 64) {
                for (std::size_t i = 0; i + 1 < static_cast<std::size_t>(nameBytes - 2); i += 2)
                    entry.name.push_back(static_cast<char>(p[i]));
            }
            entry.type = p[66];
            entry.left = getU32(p + 68);
            entry.right = getU32(p + 72);
            entry.child = getU32(p + 76);
            entry.start = getU32(p + 116);
            entry.size = getU64(p + 120);
            if (entry.type != 0) {
                if (nameBytes < 2 || nameBytes > 64 || (nameBytes & 1u) != 0)
                    throw std::runtime_error("Invalid CFB directory-entry name length");
                if (p[nameBytes - 2] != 0 || p[nameBytes - 1] != 0)
                    throw std::runtime_error("CFB directory-entry name is not null terminated");
            }
            directory_.push_back(std::move(entry));
        }
        if (directory_.empty() || directory_[0].type != 5) throw std::runtime_error("CFB root entry is missing");
        rootMiniStream_ = readRegularChain(directory_[0].start, directory_[0].size);
        if (miniFatCount && firstMiniFat != kEndOfChain) {
            const auto bytes = readRegularChain(firstMiniFat, static_cast<std::uint64_t>(miniFatCount) * sectorSize_);
            for (std::size_t i = 0; i + 4 <= bytes.size(); i += 4) miniFat_.push_back(getU32(bytes.data() + i));
        }
        std::vector<bool> visited(directory_.size(), false);
        paths_["root entry"] = 0;
        originalPaths_.resize(directory_.size());
        originalPaths_[0] = "Root Entry";
        walkTree(directory_[0].child, "", visited);
    }
};

std::string stripGeneratedAttributes(std::string source) {
    std::size_t pos = 0;
    while (pos < source.size()) {
        const auto end = source.find("\r\n", pos);
        const auto lineEnd = end == std::string::npos ? source.size() : end;
        const std::string_view line(source.data() + pos, lineEnd - pos);
        if (line.rfind("Attribute ", 0) != 0) break;
        pos = end == std::string::npos ? source.size() : end + 2;
    }
    return source.substr(pos);
}

std::vector<unsigned char> buildDirStream(const std::vector<xlpp::VbaModule>& modules, const xlpp::VbaProjectInfo& info) {
    std::vector<unsigned char> dir;
    appendFixedRecord(dir, 0x0001, info.systemKind);
    appendFixedRecord(dir, 0x0002, info.lcid);
    appendFixedRecord(dir, 0x0014, info.lcidInvoke);
    putU16(dir, 0x0003); putU32(dir, 2); putU16(dir, info.codePage);
    appendSizedStringRecord(dir, 0x0004, info.name.empty() ? "VBAProject" : info.name);
    appendDualStringRecord(dir, 0x0005, 0x0040, info.description);
    appendDualStringRecord(dir, 0x0006, 0x003D, info.helpFile);
    appendFixedRecord(dir, 0x0007, static_cast<std::uint32_t>(std::max(0, info.helpContextId)));
    appendFixedRecord(dir, 0x0008, 0);
    putU16(dir, 0x0009); putU32(dir, 4); putU32(dir, 0x65BE0257u); putU16(dir, 0x0011);
    appendDualStringRecord(dir, 0x000C, 0x003C, info.constants);

    // Use the compact reference set and versions from an Excel-accepted,
    // source-only project. Excel supplies the host object library when it
    // opens the workbook and recompiles the source modules.
    std::vector<xlpp::VbaReference> references{
        {"stdole", "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"},
        {"Office", "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library"}
    };
    for (const auto& reference : info.references) {
        if (reference.name.empty() || reference.libid.empty())
            throw std::invalid_argument("VBA reference requires both name and LibId");
        const auto duplicate = std::find_if(references.begin(), references.end(), [&](const auto& existing) {
            return asciiLower(existing.name) == asciiLower(reference.name) || existing.libid == reference.libid;
        });
        if (duplicate == references.end()) references.push_back(reference);
        else if (duplicate->libid != reference.libid && asciiLower(duplicate->name) == asciiLower(reference.name))
            *duplicate = reference;
    }
    for (const auto& reference : references) {
        appendDualStringRecord(dir, 0x0016, 0x003E, reference.name);
        if (reference.kind == xlpp::VbaReferenceKind::Project) {
            if (reference.relativeLibid.empty())
                throw std::invalid_argument("VBA project reference requires a relative LibId");
            putU16(dir, 0x000E);
            putU32(dir, static_cast<std::uint32_t>(4 + reference.libid.size() + 4 + reference.relativeLibid.size() + 4 + 2));
            putU32(dir, static_cast<std::uint32_t>(reference.libid.size()));
            appendBytes(dir, reference.libid);
            putU32(dir, static_cast<std::uint32_t>(reference.relativeLibid.size()));
            appendBytes(dir, reference.relativeLibid);
            putU32(dir, reference.majorVersion);
            putU16(dir, reference.minorVersion);
        } else if (reference.kind == xlpp::VbaReferenceKind::Control) {
            const auto twiddled = reference.twiddledLibid.empty()
                ? std::string("*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##") : reference.twiddledLibid;
            if (reference.originalTypeLib.empty())
                throw std::invalid_argument("VBA control reference requires originalTypeLib GUID");
            const auto guid = parseGuidBytes(reference.originalTypeLib);
            putU16(dir, 0x002F);
            putU32(dir, static_cast<std::uint32_t>(4 + twiddled.size() + 4 + 2));
            putU32(dir, static_cast<std::uint32_t>(twiddled.size()));
            appendBytes(dir, twiddled);
            putU32(dir, 0); putU16(dir, 0);
            appendDualStringRecord(dir, 0x0016, 0x003E, reference.extendedName.empty() ? reference.name : reference.extendedName);
            putU16(dir, 0x0030);
            putU32(dir, static_cast<std::uint32_t>(4 + reference.libid.size() + 4 + 2 + 16 + 4));
            putU32(dir, static_cast<std::uint32_t>(reference.libid.size()));
            appendBytes(dir, reference.libid);
            putU32(dir, 0); putU16(dir, 0);
            dir.insert(dir.end(), guid.begin(), guid.end());
            putU32(dir, reference.controlCookie);
        } else {
            putU16(dir, 0x000D);
            putU32(dir, static_cast<std::uint32_t>(4 + reference.libid.size() + 4 + 2));
            putU32(dir, static_cast<std::uint32_t>(reference.libid.size()));
            appendBytes(dir, reference.libid);
            putU32(dir, 0);
            putU16(dir, 0);
        }
    }

    putU16(dir, 0x000F); putU32(dir, 2); putU16(dir, static_cast<std::uint16_t>(modules.size()));
    putU16(dir, 0x0013); putU32(dir, 2); putU16(dir, 0xFFFF);
    for (const auto& module : modules) {
        appendSizedStringRecord(dir, 0x0019, module.name);
        const auto wide = utf16Le(module.name);
        putU16(dir, 0x0047); putU32(dir, static_cast<std::uint32_t>(wide.size()));
        dir.insert(dir.end(), wide.begin(), wide.end());
        appendDualStringRecord(dir, 0x001A, 0x0032, module.name);
        appendDualStringRecord(dir, 0x001C, 0x0048, module.docString);
        putU16(dir, 0x0031); putU32(dir, 4); putU32(dir, 0);
        putU16(dir, 0x001E); putU32(dir, 4); putU32(dir, static_cast<std::uint32_t>(std::max(0, module.helpContextId)));
        putU16(dir, 0x002C); putU32(dir, 2); putU16(dir, 0xFFFF);
        putU16(dir, module.type == xlpp::VbaModuleType::Standard ? 0x0021 : 0x0022); putU32(dir, 0);
        if (module.readOnly) { putU16(dir, 0x0025); putU32(dir, 0); }
        if (module.isPrivate) { putU16(dir, 0x0028); putU32(dir, 0); }
        putU16(dir, 0x002B); putU32(dir, 0);
    }
    putU16(dir, 0x0010); putU32(dir, 0);
    return compressOvba(std::string_view(reinterpret_cast<const char*>(dir.data()), dir.size()));
}

std::string buildProjectText(const std::vector<xlpp::VbaModule>& modules, const xlpp::VbaProjectInfo& info) {
    std::string text;
    // These unprotected project fields match a source-only VBA project that
    // Excel accepts and recompiles. They are not a password or signature.
    text += "ID=\"" + (info.projectId.empty() ? std::string("{9E394C0B-697E-4AEE-9FA6-446F51FB30DC}") : info.projectId) + "\"\r\n";
    for (const auto& module : modules) {
        if (module.type == xlpp::VbaModuleType::Document)
            text += "Document=" + module.name + "/&H00000000\r\n";
        else if (module.type == xlpp::VbaModuleType::Class)
            text += "Class=" + module.name + "\r\n";
        else if (module.type == xlpp::VbaModuleType::Designer) {
            text += "Package=" + (module.designerClassId.empty()
                ? std::string("{AC9F2F90-E877-11CE-9F68-00AA00574A4F}") : module.designerClassId) + "\r\n";
            text += "BaseClass=" + module.name + "\r\n";
        } else
            text += "Module=" + module.name + "\r\n";
    }
    text += "Name=\"" + (info.name.empty() ? std::string("VBAProject") : info.name) + "\"\r\n";
    if (!info.description.empty()) text += "Description=\"" + info.description + "\"\r\n";
    if (!info.helpFile.empty()) text += "HelpFile=\"" + info.helpFile + "\"\r\n";
    text += "HelpContextID=\"" + std::to_string(std::max(0, info.helpContextId)) + "\"\r\n";
    text += "CMG=\"6D6F7625A5A1A9A1A9A1A9A1A9\"\r\n";
    text += "DPB=\"3D3F26A4400941094109\"\r\n";
    text += "GC=\"7E7C652AB2EA03EB03EBFC\"\r\n\r\n";
    text += "[Host Extender Info]\r\n";
    text += "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n";
    return text;
}

std::vector<unsigned char> buildProjectWm(const std::vector<xlpp::VbaModule>& modules) {
    std::vector<unsigned char> out;
    for (const auto& module : modules) {
        appendBytes(out, module.name);
        out.push_back(0);
        const auto wide = utf16Le(module.name, true);
        out.insert(out.end(), wide.begin(), wide.end());
    }
    putU16(out, 0);
    return out;
}

std::string moduleStreamSource(const xlpp::VbaModule& module) {
    std::string source = "Attribute VB_Name = \"" + module.name + "\"\r\n";
    if (module.type == xlpp::VbaModuleType::Document) {
        const bool workbookModule = asciiLower(module.name) == "thisworkbook";
        source += workbookModule
            ? "Attribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"\r\n"
            : "Attribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"\r\n";
        source += "Attribute VB_GlobalNameSpace = False\r\n";
        source += "Attribute VB_Creatable = False\r\n";
        source += "Attribute VB_PredeclaredId = True\r\n";
        source += "Attribute VB_Exposed = True\r\n";
        source += "Attribute VB_TemplateDerived = False\r\n";
        source += "Attribute VB_Customizable = True\r\n";
    } else if (module.type == xlpp::VbaModuleType::Class) {
        source += "Attribute VB_GlobalNameSpace = False\r\n";
        source += "Attribute VB_Creatable = False\r\n";
        source += "Attribute VB_PredeclaredId = False\r\n";
        source += "Attribute VB_Exposed = False\r\n";
    } else if (module.type == xlpp::VbaModuleType::Designer) {
        source += "Attribute VB_Base = \"" + (module.designerBaseClass.empty()
            ? std::string("0{842E9C5E-88B5-439A-912E-4C2D9AA0EC27}{2DC3C962-DA1C-47BA-AB63-E9D578FC2637}")
            : module.designerBaseClass) + "\"\r\n";
        source += "Attribute VB_GlobalNameSpace = False\r\n";
        source += "Attribute VB_Creatable = False\r\n";
        source += "Attribute VB_PredeclaredId = True\r\n";
        source += "Attribute VB_Exposed = False\r\n";
        source += "Attribute VB_TemplateDerived = False\r\n";
        source += "Attribute VB_Customizable = False\r\n";
    }
    source += normalizeVbaSource(module.source);
    return source;
}

struct ParsedModuleRecord {
    std::string name;
    xlpp::VbaModuleType type{xlpp::VbaModuleType::Standard};
    std::uint32_t offset{0};
    std::string docString;
    bool readOnly{false};
    bool isPrivate{false};
    int helpContextId{0};
};

std::vector<ParsedModuleRecord> parseModuleRecords(const std::string& dir) {
    const auto* p = reinterpret_cast<const unsigned char*>(dir.data());
    std::size_t modulesPos = std::string::npos;
    for (std::size_t i = 0; i + 8 <= dir.size(); ++i) {
        if (getU16(p + i) == 0x000F && getU32(p + i + 2) == 2) {
            const auto count = getU16(p + i + 6);
            if (count > 0 && count < 4096) { modulesPos = i; break; }
        }
    }
    if (modulesPos == std::string::npos) throw std::runtime_error("VBA PROJECTMODULES record not found");
    std::size_t pos = modulesPos + 8;
    if (pos + 8 > dir.size() || getU16(p + pos) != 0x0013) throw std::runtime_error("VBA project cookie is missing");
    pos += 8;
    const auto count = getU16(p + modulesPos + 6);
    std::vector<ParsedModuleRecord> modules;
    for (std::uint16_t m = 0; m < count; ++m) {
        ParsedModuleRecord module;
        if (pos + 6 > dir.size() || getU16(p + pos) != 0x0019) throw std::runtime_error("VBA module name record is missing");
        const auto nameLen = getU32(p + pos + 2); pos += 6;
        if (pos + nameLen > dir.size()) throw std::runtime_error("VBA module name is truncated");
        module.name.assign(dir.data() + static_cast<std::ptrdiff_t>(pos), nameLen); pos += nameLen;
        if (pos + 6 <= dir.size() && getU16(p + pos) == 0x0047) { const auto n=getU32(p+pos+2); pos += 6 + n; }
        if (pos + 6 > dir.size() || getU16(p + pos) != 0x001A) throw std::runtime_error("VBA stream-name record is missing");
        auto n = getU32(p + pos + 2); pos += 6 + n;
        if (pos + 6 > dir.size() || getU16(p + pos) != 0x0032) throw std::runtime_error("VBA Unicode stream-name record is missing");
        n = getU32(p + pos + 2); pos += 6 + n;
        if (pos + 6 > dir.size() || getU16(p + pos) != 0x001C) throw std::runtime_error("VBA doc-string record is missing");
        n = getU32(p + pos + 2); pos += 6;
        if (pos + n > dir.size()) throw std::runtime_error("VBA module doc-string is truncated");
        module.docString.assign(dir.data() + static_cast<std::ptrdiff_t>(pos), n); pos += n;
        if (pos + 6 > dir.size() || getU16(p + pos) != 0x0048) throw std::runtime_error("VBA Unicode doc-string record is missing");
        n = getU32(p + pos + 2); pos += 6 + n;
        if (pos + 10 > dir.size() || getU16(p + pos) != 0x0031) throw std::runtime_error("VBA module offset record is missing");
        module.offset = getU32(p + pos + 6); pos += 10;
        if (pos + 10 > dir.size() || getU16(p + pos) != 0x001E) throw std::runtime_error("VBA module help-context record is missing");
        module.helpContextId = static_cast<int>(getU32(p + pos + 6));
        pos += 10;
        if (pos + 8 > dir.size() || getU16(p + pos) != 0x002C) throw std::runtime_error("VBA module cookie record is missing");
        pos += 8;
        if (pos + 6 > dir.size()) throw std::runtime_error("VBA module type record is missing");
        const auto typeId = getU16(p + pos); pos += 6;
        module.type = typeId == 0x0021 ? xlpp::VbaModuleType::Standard : xlpp::VbaModuleType::Document;
        while (pos + 2 <= dir.size() && getU16(p + pos) != 0x002B) {
            const auto id = getU16(p + pos);
            if (id == 0x0025) { module.readOnly = true; pos += 6; }
            else if (id == 0x0028) { module.isPrivate = true; pos += 6; }
            else throw std::runtime_error("Unsupported VBA module record extension");
        }
        if (pos + 6 > dir.size()) throw std::runtime_error("VBA module terminator is missing");
        pos += 6;
        modules.push_back(std::move(module));
    }
    return modules;
}

} // namespace

std::vector<unsigned char> buildRootCompoundFile(
    const std::map<std::string, std::vector<unsigned char>>& rootStreams) {
    return buildCfb(rootStreams, {}, {}, false);
}

bool isCompoundFile(const std::vector<unsigned char>& bytes) noexcept {
    static constexpr unsigned char signature[8] = {0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1};
    return bytes.size() >= 8 && std::equal(std::begin(signature), std::end(signature), bytes.begin());
}

bool compoundFileContainsStream(const std::vector<unsigned char>& bytes, const std::string& path) noexcept {
    try { CfbReader reader(bytes); (void)reader.stream(path); return true; } catch (...) { return false; }
}

std::vector<unsigned char> readCompoundFileStream(const std::vector<unsigned char>& bytes, const std::string& path) {
    CfbReader reader(bytes);
    return reader.stream(path);
}

std::string normalizeVbaSource(std::string source) {
    std::string normalized;
    normalized.reserve(source.size() + 8);
    for (std::size_t i = 0; i < source.size(); ++i) {
        const char ch = source[i];
        if (ch == '\r') {
            normalized += "\r\n";
            if (i + 1 < source.size() && source[i + 1] == '\n') ++i;
        } else if (ch == '\n') {
            normalized += "\r\n";
        } else {
            normalized.push_back(ch);
        }
    }
    if (!normalized.empty() && normalized.size() >= 2 && normalized.substr(normalized.size() - 2) != "\r\n")
        normalized += "\r\n";
    return normalized;
}

void validateVbaModuleName(const std::string& name) {
    if (name.empty() || name.size() > 31) throw std::invalid_argument("VBA module name must contain 1-31 characters");
    const auto first = static_cast<unsigned char>(name.front());
    if (!(std::isalpha(first) || first == '_')) throw std::invalid_argument("VBA module name must begin with a letter or underscore");
    for (unsigned char ch : name)
        if (!(std::isalnum(ch) || ch == '_')) throw std::invalid_argument("VBA module name contains an invalid character");
    if (asciiLower(name) == "thisworkbook") throw std::invalid_argument("ThisWorkbook is reserved for the document module");
}

std::vector<unsigned char> buildVbaProjectBinary(const std::vector<xlpp::VbaModule>& inputModules,
                                                 const std::vector<std::string>& worksheetCodeNames,
                                                 const xlpp::VbaProjectInfo& projectInfo) {
    std::vector<xlpp::VbaModule> modules;
    modules.reserve(worksheetCodeNames.size() + inputModules.size() + 1);
    // Excel workbooks expose one document module per worksheet. The sheet
    // display name is independent; the stable VBA code names are Sheet1,
    // Sheet2, ... and are also written to worksheet sheetPr/codeName.
    for (const auto& codeName : worksheetCodeNames) {
        if (codeName.empty()) throw std::invalid_argument("Worksheet VBA code name cannot be empty");
        if (asciiLower(codeName) == "thisworkbook") throw std::invalid_argument("Worksheet VBA code name cannot be ThisWorkbook");
        if (std::any_of(modules.begin(), modules.end(), [&](const auto& existing) { return asciiLower(existing.name) == asciiLower(codeName); }))
            throw std::invalid_argument("Duplicate worksheet VBA code name: " + codeName);
        modules.push_back({codeName, "", xlpp::VbaModuleType::Document});
    }
    modules.push_back({"ThisWorkbook", "", xlpp::VbaModuleType::Document});
    // Document module source is editable, but document module identity is owned
    // by the workbook host. Copy source/metadata only into matching host modules.
    for (const auto& input : inputModules) {
        if (input.type != xlpp::VbaModuleType::Document) continue;
        const auto found = std::find_if(modules.begin(), modules.end(), [&](const auto& existing) {
            return asciiLower(existing.name) == asciiLower(input.name);
        });
        // A worksheet may have been removed since the source-generated project
        // was last built. Its stale document module is intentionally dropped;
        // public setVbaDocumentModuleText() validates unknown names earlier.
        if (found == modules.end()) continue;
        found->source = normalizeVbaSource(input.source);
        found->docString = input.docString;
        found->readOnly = input.readOnly;
        found->isPrivate = input.isPrivate;
    }
    for (auto module : inputModules) {
        if (module.type == xlpp::VbaModuleType::Document) continue;
        validateVbaModuleName(module.name);
        if (std::any_of(modules.begin(), modules.end(), [&](const auto& existing) {
            return asciiLower(existing.name) == asciiLower(module.name);
        })) throw std::invalid_argument("Duplicate VBA module name: " + module.name);
        module.source = normalizeVbaSource(std::move(module.source));
        modules.push_back(std::move(module));
    }

    std::map<std::string, std::vector<unsigned char>> rootStreams;
    const auto projectText = buildProjectText(modules, projectInfo);
    rootStreams["PROJECT"] = std::vector<unsigned char>(projectText.begin(), projectText.end());
    rootStreams["PROJECTwm"] = buildProjectWm(modules);

    std::map<std::string, std::vector<unsigned char>> vbaStreams;
    vbaStreams["_VBA_PROJECT"] = {0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x03, 0x00};
    vbaStreams["dir"] = buildDirStream(modules, projectInfo);
    for (const auto& module : modules) {
        const auto source = moduleStreamSource(module);
        vbaStreams[module.name] = compressOvba(source);
    }
    return buildCfb(rootStreams, vbaStreams, projectInfo.designerStorages);
}

bool isXlppGeneratedVbaProjectBinary(const std::vector<unsigned char>& bytes) noexcept {
    try {
        CfbReader cfb(bytes);
        const auto projectBytes = cfb.stream("PROJECT");
        const std::string project(projectBytes.begin(), projectBytes.end());
        const auto projectStub = cfb.stream("VBA/_VBA_PROJECT");
        static const std::vector<unsigned char> expectedStub{0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x03, 0x00};
        // Project name/description/ID are now user-editable metadata, so they
        // cannot serve as the XL++ ownership marker. Keep detection tied to the
        // stable source-only host scaffold that XL++ writes instead.
        return projectStub == expectedStub
            && project.find("CMG=\"6D6F7625A5A1A9A1A9A1A9A1A9\"\r\n") != std::string::npos
            && project.find("DPB=\"3D3F26A4400941094109\"\r\n") != std::string::npos
            && project.find("GC=\"7E7C652AB2EA03EB03EBFC\"\r\n") != std::string::npos
            && project.find("[Host Extender Info]\r\n") != std::string::npos
            && project.find("{3832D640-CF90-11CF-8E43-00A0C911005A}") != std::string::npos;
    } catch (...) {
        return false;
    }
}

std::vector<xlpp::VbaModule> readVbaProjectBinary(const std::vector<unsigned char>& bytes) {
    CfbReader cfb(bytes);
    const auto dirBytes = cfb.stream("VBA/dir");
    const auto dirText = decompressOvba(dirBytes);
    const auto records = parseModuleRecords(dirText);
    std::unordered_map<std::string, xlpp::VbaModuleType> projectTypes;
    std::unordered_map<std::string, std::string> designerClassIds;
    try {
        const auto projectBytes = cfb.stream("PROJECT");
        const std::string project(projectBytes.begin(), projectBytes.end());
        std::string pendingPackage;
        std::size_t pos = 0;
        while (pos <= project.size()) {
            const auto end = project.find("\r\n", pos);
            const auto line = project.substr(pos, end == std::string::npos ? std::string::npos : end - pos);
            const auto assign = line.find('=');
            if (assign != std::string::npos) {
                const auto key = line.substr(0, assign);
                auto value = line.substr(assign + 1);
                if (key == "Document") { const auto slash = value.find('/'); if (slash != std::string::npos) value.resize(slash); projectTypes[asciiLower(value)] = xlpp::VbaModuleType::Document; pendingPackage.clear(); }
                else if (key == "Class") { projectTypes[asciiLower(value)] = xlpp::VbaModuleType::Class; pendingPackage.clear(); }
                else if (key == "Module") { projectTypes[asciiLower(value)] = xlpp::VbaModuleType::Standard; pendingPackage.clear(); }
                else if (key == "Package") pendingPackage = value;
                else if (key == "BaseClass") {
                    projectTypes[asciiLower(value)] = xlpp::VbaModuleType::Designer;
                    if (!pendingPackage.empty()) designerClassIds[asciiLower(value)] = pendingPackage;
                    pendingPackage.clear();
                }
            }
            if (end == std::string::npos) break;
            pos = end + 2;
        }
    } catch (...) {}
    std::vector<xlpp::VbaModule> modules;
    modules.reserve(records.size());
    for (const auto& record : records) {
        const auto streamBytes = cfb.stream("VBA/" + record.name);
        auto source = decompressOvba(streamBytes, record.offset);
        auto type = record.type;
        const auto lowerName = asciiLower(record.name);
        if (const auto it = projectTypes.find(lowerName); it != projectTypes.end()) type = it->second;
        xlpp::VbaModule module;
        module.name = record.name;
        module.type = type;
        module.docString = record.docString;
        module.readOnly = record.readOnly;
        module.isPrivate = record.isPrivate;
        module.helpContextId = record.helpContextId;
        if (type == xlpp::VbaModuleType::Designer) {
            if (const auto it = designerClassIds.find(lowerName); it != designerClassIds.end()) module.designerClassId = it->second;
            const std::string prefix = "Attribute VB_Base = \"";
            const auto attr = source.find(prefix);
            if (attr != std::string::npos) {
                const auto begin = attr + prefix.size();
                const auto end = source.find('\"', begin);
                if (end != std::string::npos) module.designerBaseClass = source.substr(begin, end - begin);
            }
        }
        source = stripGeneratedAttributes(std::move(source));
        module.source = normalizeVbaSource(std::move(source));
        modules.push_back(std::move(module));
    }
    return modules;
}

xlpp::VbaProjectInfo readVbaProjectInfoBinary(const std::vector<unsigned char>& bytes) {
    CfbReader cfb(bytes);
    xlpp::VbaProjectInfo info;
    const auto projectBytes = cfb.stream("PROJECT");
    const std::string project(projectBytes.begin(), projectBytes.end());
    auto unquote = [](std::string value) {
        if (value.size() >= 2 && value.front() == '"' && value.back() == '"') return value.substr(1, value.size() - 2);
        return value;
    };
    std::size_t pos = 0;
    while (pos <= project.size()) {
        const auto end = project.find("\r\n", pos);
        const auto line = project.substr(pos, end == std::string::npos ? std::string::npos : end - pos);
        const auto assign = line.find('=');
        if (assign != std::string::npos) {
            const auto key = line.substr(0, assign);
            const auto value = unquote(line.substr(assign + 1));
            if (key == "ID") info.projectId = value;
            else if (key == "Name") info.name = value;
            else if (key == "Description") info.description = value;
            else if (key == "HelpFile") info.helpFile = value;
            else if (key == "HelpContextID") { try { info.helpContextId = std::stoi(value); } catch (...) {} }
        }
        if (end == std::string::npos) break;
        pos = end + 2;
    }
    // PROJECTCONSTANTS lives in the compressed dir stream rather than the
    // textual PROJECT stream. Recover the ANSI half of the dual-string record
    // so metadata can round-trip through the public API.
    try {
        const auto dir = decompressOvba(cfb.stream("VBA/dir"));
        const auto* data = reinterpret_cast<const unsigned char*>(dir.data());
        const auto readFixed = [&](std::uint16_t id, std::uint32_t size, std::uint32_t fallback) {
            for (std::size_t offset = 0; offset + 6 + size <= dir.size(); ++offset) {
                if (getU16(data + offset) == id && getU32(data + offset + 2) == size)
                    return size == 2 ? static_cast<std::uint32_t>(getU16(data + offset + 6)) : getU32(data + offset + 6);
            }
            return fallback;
        };
        info.systemKind = readFixed(0x0001, 4, info.systemKind);
        info.lcid = readFixed(0x0002, 4, info.lcid);
        info.lcidInvoke = readFixed(0x0014, 4, info.lcidInvoke);
        info.codePage = static_cast<std::uint16_t>(readFixed(0x0003, 2, info.codePage));
        for (std::size_t i = 0; i + 12 <= dir.size(); ++i) {
            if (getU16(data + i) == 0x000C) {
                const auto length = static_cast<std::size_t>(getU32(data + i + 2));
                if (i + 6 + length + 6 <= dir.size() && getU16(data + i + 6 + length) == 0x003C)
                    info.constants.assign(dir.data() + static_cast<std::ptrdiff_t>(i + 6), length);
            }
            if (getU16(data + i) != 0x0016) continue;
            const auto nameLength = static_cast<std::size_t>(getU32(data + i + 2));
            if (i + 6 + nameLength + 6 > dir.size()) continue;
            const auto unicodePos = i + 6 + nameLength;
            if (getU16(data + unicodePos) != 0x003E) continue;
            const auto unicodeLength = static_cast<std::size_t>(getU32(data + unicodePos + 2));
            const auto referencePos = unicodePos + 6 + unicodeLength;
            if (referencePos + 10 > dir.size()) continue;
            const auto referenceId = getU16(data + referencePos);
            if (referenceId != 0x000D && referenceId != 0x000E && referenceId != 0x002F) continue;
            xlpp::VbaReference reference;
            reference.name.assign(dir.data() + static_cast<std::ptrdiff_t>(i + 6), nameLength);
            if (referenceId == 0x002F) {
                reference.kind = xlpp::VbaReferenceKind::Control;
                if (referencePos + 10 > dir.size()) continue;
                const auto twiddledLength = static_cast<std::size_t>(getU32(data + referencePos + 6));
                auto cursor = referencePos + 10;
                if (cursor + twiddledLength + 6 > dir.size()) continue;
                reference.twiddledLibid.assign(dir.data() + static_cast<std::ptrdiff_t>(cursor), twiddledLength);
                cursor += twiddledLength + 6; // Reserved1 + Reserved2.
                if (cursor + 6 <= dir.size() && getU16(data + cursor) == 0x0016) {
                    const auto extNameLength = static_cast<std::size_t>(getU32(data + cursor + 2));
                    if (cursor + 6 + extNameLength + 6 > dir.size()) continue;
                    reference.extendedName.assign(dir.data() + static_cast<std::ptrdiff_t>(cursor + 6), extNameLength);
                    const auto extUnicodePos = cursor + 6 + extNameLength;
                    if (getU16(data + extUnicodePos) != 0x003E) continue;
                    const auto extUnicodeLength = static_cast<std::size_t>(getU32(data + extUnicodePos + 2));
                    cursor = extUnicodePos + 6 + extUnicodeLength;
                }
                if (cursor + 10 > dir.size() || getU16(data + cursor) != 0x0030) continue;
                cursor += 2; // Reserved3.
                const auto extendedSize = static_cast<std::size_t>(getU32(data + cursor)); cursor += 4;
                const auto extendedEnd = cursor + extendedSize;
                if (extendedEnd > dir.size() || cursor + 4 > extendedEnd) continue;
                const auto extendedLibidLength = static_cast<std::size_t>(getU32(data + cursor)); cursor += 4;
                if (cursor + extendedLibidLength + 4 + 2 + 16 + 4 > extendedEnd) continue;
                reference.libid.assign(dir.data() + static_cast<std::ptrdiff_t>(cursor), extendedLibidLength);
                cursor += extendedLibidLength + 4 + 2; // Reserved4 + Reserved5.
                reference.originalTypeLib = guidTextFromBytes(data + cursor); cursor += 16;
                reference.controlCookie = getU32(data + cursor);
            } else {
                const auto recordLength = static_cast<std::size_t>(getU32(data + referencePos + 2));
                if (referencePos + 6 + recordLength > dir.size()) continue;
                const auto libidLength = static_cast<std::size_t>(getU32(data + referencePos + 6));
                if (referencePos + 10 + libidLength > dir.size()) continue;
                reference.libid.assign(dir.data() + static_cast<std::ptrdiff_t>(referencePos + 10), libidLength);
                if (referenceId == 0x000E) {
                    reference.kind = xlpp::VbaReferenceKind::Project;
                    auto cursor = referencePos + 10 + libidLength;
                    if (cursor + 4 > referencePos + 6 + recordLength) continue;
                    const auto relativeLength = static_cast<std::size_t>(getU32(data + cursor));
                    cursor += 4;
                    if (cursor + relativeLength + 6 > referencePos + 6 + recordLength) continue;
                    reference.relativeLibid.assign(dir.data() + static_cast<std::ptrdiff_t>(cursor), relativeLength);
                    cursor += relativeLength;
                    reference.majorVersion = getU32(data + cursor); cursor += 4;
                    reference.minorVersion = getU16(data + cursor);
                }
            }
            info.references.push_back(std::move(reference));
        }
    } catch (...) {}

    // Designer modules are declared by BaseClass in the PROJECT stream and
    // own a same-named root storage. Preserve every descendant stream as a
    // relative path so unsupported MS-OFORMS/control payloads remain exact.
    std::vector<std::string> designerNames;
    {
        std::size_t projectPos = 0;
        while (projectPos <= project.size()) {
            const auto end = project.find("\r\n", projectPos);
            const auto line = project.substr(projectPos, end == std::string::npos ? std::string::npos : end - projectPos);
            if (line.rfind("BaseClass=", 0) == 0 && line.size() > 10) designerNames.push_back(line.substr(10));
            if (end == std::string::npos) break;
            projectPos = end + 2;
        }
    }
    for (const auto& storageName : cfb.rootStorageNames()) {
        if (asciiLower(storageName) == "vba") continue;
        const auto declared = std::find_if(designerNames.begin(), designerNames.end(), [&](const auto& name) {
            return asciiLower(name) == asciiLower(storageName);
        });
        if (declared == designerNames.end()) continue;
        xlpp::VbaDesignerStorage storage;
        storage.name = storageName;
        const auto prefix = storageName + "/";
        for (const auto& path : cfb.streamPathsUnder(storageName)) {
            xlpp::VbaDesignerStream stream;
            stream.path = path.substr(prefix.size());
            stream.data = cfb.stream(path);
            storage.streams.push_back(std::move(stream));
        }
        info.designerStorages.push_back(std::move(storage));
    }
    return info;
}

} // namespace xlpp::internal
