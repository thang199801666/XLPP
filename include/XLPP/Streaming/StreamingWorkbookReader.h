#pragma once
#include <XLPP/Cell/Cell.h>
#include <cstddef>
#include <filesystem>
#include <functional>
#include <iterator>
#include <memory>
#include <optional>
#include <string>
#include <vector>

namespace xlpp {

struct StreamingCell {
    std::string address;
    CellValue value;
    std::string formula;
    std::optional<std::size_t> styleIndex;
};
using StreamingRow = std::vector<StreamingCell>;

namespace internal {
class WorksheetRowSource;
}

// Forward input iterator over the rows of a streaming worksheet. Each iterator
// streams the worksheet entry independently.
class StreamingRowIterator {
public:
    using iterator_category = std::input_iterator_tag;
    using value_type = StreamingRow;
    using difference_type = std::ptrdiff_t;
    using pointer = const StreamingRow*;
    using reference = const StreamingRow&;

    StreamingRowIterator() = default;
    explicit StreamingRowIterator(std::shared_ptr<internal::WorksheetRowSource> source);

    reference operator*() const noexcept { return current_; }
    pointer operator->() const noexcept { return &current_; }
    StreamingRowIterator& operator++();
    StreamingRowIterator operator++(int) { StreamingRowIterator copy = *this; ++*this; return copy; }

    std::size_t rowNumber() const noexcept { return rowNumber_; }

    bool operator==(const StreamingRowIterator& other) const noexcept;
    bool operator!=(const StreamingRowIterator& other) const noexcept { return !(*this == other); }

private:
    void advance();
    std::shared_ptr<internal::WorksheetRowSource> source_;
    StreamingRow current_;
    std::size_t rowNumber_{0};
    bool valid_{false};
};

// Factory for StreamingRowIterator. Each iterator advances independently over
// the worksheet, so multiple iterators may be created and used in any order.
class StreamingWorksheetReader {
public:
    StreamingWorksheetReader() = default;
    StreamingWorksheetReader(const StreamingWorksheetReader&) = delete;
    StreamingWorksheetReader& operator=(const StreamingWorksheetReader&) = delete;
    StreamingWorksheetReader(StreamingWorksheetReader&&) noexcept = default;
    StreamingWorksheetReader& operator=(StreamingWorksheetReader&&) noexcept = default;

    StreamingRowIterator begin();
    StreamingRowIterator end() const noexcept { return {}; }

    // Row callback API. Return false from the callback to stop early.
    void forEachRow(const std::function<bool(std::size_t, const StreamingRow&)>& callback);

private:
    friend class StreamingWorkbookReader;
    explicit StreamingWorksheetReader(
        std::function<std::shared_ptr<internal::WorksheetRowSource>()> factory);
    std::function<std::shared_ptr<internal::WorksheetRowSource>()> factory_;
};

class StreamingWorkbookReader {
public:
    explicit StreamingWorkbookReader(const std::filesystem::path& path);

    std::vector<std::string> worksheetNames() const;
    StreamingWorksheetReader worksheet(const std::string& worksheetName) const;
    void forEachRow(const std::string& worksheetName,
                    const std::function<bool(std::size_t, const StreamingRow&)>& callback) const;

private:
    struct SharedState;
    std::shared_ptr<SharedState> state_;
};

} // namespace xlpp
