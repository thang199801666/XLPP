#pragma once
#include <XLPP/Cell/Cell.h>
#include <XLPP/Compression.h>
#include <cstddef>
#include <filesystem>
#include <fstream>
#include <memory>
#include <string>
#include <vector>

namespace xlpp {

// Shared-string deduplication strategy for the streaming writer.
enum class SharedStringMode {
    // No shared-string table; strings are written inline in each cell.
    Disabled,
    // Exact deduplication across the whole workbook via a hash table.
    Hash,
    // Deduplication bounded to an LRU cache of recent strings; evicted strings
    // get fresh indexes when they recur, bounding cache memory.
    BoundedLru,
};

class StreamingWorksheetWriter {
public:
    StreamingWorksheetWriter() = default;
    StreamingWorksheetWriter(const StreamingWorksheetWriter&) = delete;
    StreamingWorksheetWriter& operator=(const StreamingWorksheetWriter&) = delete;
    StreamingWorksheetWriter(StreamingWorksheetWriter&&) noexcept;
    StreamingWorksheetWriter& operator=(StreamingWorksheetWriter&&) noexcept;
    ~StreamingWorksheetWriter();

    void append(const std::vector<CellValue>& row);
    std::size_t rowCount() const noexcept { return rowCount_; }
    const std::string& name() const noexcept { return name_; }

private:
    friend class StreamingWorkbookWriter;
    StreamingWorksheetWriter(std::string name, std::filesystem::path spoolPath,
                             class SharedStringTable* sharedStrings, bool date1904);
    void finish();
    std::string name_;
    std::filesystem::path spoolPath_;
    std::unique_ptr<std::ofstream> stream_;
    class SharedStringTable* sharedStrings_{nullptr};
    bool date1904_{false};
    std::size_t rowCount_{0};
    bool finished_{false};
};

class StreamingWorkbookWriter {
public:
    explicit StreamingWorkbookWriter(std::filesystem::path outputPath,
                                     SharedStringMode sharedStrings = SharedStringMode::Disabled,
                                     std::size_t lruCapacity = 1024);
    StreamingWorkbookWriter(const StreamingWorkbookWriter&) = delete;
    StreamingWorkbookWriter& operator=(const StreamingWorkbookWriter&) = delete;
    ~StreamingWorkbookWriter();

    StreamingWorksheetWriter& addWorksheet(std::string name);
    std::size_t sheetCount() const noexcept { return sheets_.size(); }
    StreamingWorksheetWriter& worksheet(std::size_t index) { return sheets_.at(index); }
    void close();
    bool closed() const noexcept { return closed_; }

    void setDate1904(bool enabled) noexcept { date1904_ = enabled; }
    bool date1904() const noexcept { return date1904_; }

    void setCompressionLevel(CompressionLevel level) noexcept { compressionLevel_ = level; }
    void setCompressionStrategy(CompressionStrategy strategy) noexcept { compressionStrategy_ = strategy; }
    // 0 = sequential; N = up to N threads for entry compression on close().
    void setParallelWorkers(std::size_t workers) noexcept { parallelWorkers_ = workers; }

private:
    std::filesystem::path outputPath_;
    std::filesystem::path tempDirectory_;
    std::vector<StreamingWorksheetWriter> sheets_;
    std::unique_ptr<class SharedStringTable> sharedStrings_;
    bool date1904_{false};
    bool closed_{false};
    CompressionLevel compressionLevel_{CompressionLevel::Default};
    CompressionStrategy compressionStrategy_{CompressionStrategy::Default};
    std::size_t parallelWorkers_{0};
};
}
