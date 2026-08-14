#pragma once
#include "../Packaging/ZipArchiveReader.h"
#include <cstddef>
#include <mutex>
#include <string>
#include <vector>

namespace xlpp::internal {

// Lazily loads and caches the workbook shared string table. The table is only
// decompressed on first access, and entries are retained afterwards.
class SharedStringsReader {
public:
    explicit SharedStringsReader(ZipArchiveReader archive);

    // Returns the string for the given 0-based index, or nullptr when the table
    // has no such index.
    const std::string* lookup(std::size_t index) const;

    // Number of shared strings in the table (triggers loading).
    std::size_t size() const;

private:
    void load() const;

    ZipArchiveReader archive_;
    mutable std::once_flag loaded_;
    mutable std::vector<std::string> strings_;
};

} // namespace xlpp::internal
