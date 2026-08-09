#include <XLPP/Workbook/Workbook.h>
#include "IO/FileTransaction.h"

#include <fstream>
#include <stdexcept>

namespace xlpp {

void Workbook::save(const std::filesystem::path& path) const {
    save(path, SaveOptions{});
}

void Workbook::save(const std::filesystem::path& path, const SaveOptions& options) const {
    if (!options.atomicWrite) {
        saveInPlace(path, options);
        if (options.durableWrite) internal::syncFileToDisk(path);
        return;
    }
    if (path.empty()) throw std::invalid_argument("Workbook save path cannot be empty");

    auto stagingOptions = options;
    stagingOptions.atomicWrite = false;
    internal::ScopedTemporaryFile staging(internal::atomicSaveTemporaryPath(path));
    saveInPlace(staging.path(), stagingOptions);
    internal::replaceFileAtomically(staging.path(), path, options.durableWrite);
    staging.release();
}

void Workbook::load(const std::filesystem::path& path) {
    load(path, LoadOptions{});
}

void Workbook::load(const std::filesystem::path& path, const LoadOptions& options) {
    // Strong exception guarantee: parsing takes place in a staging model and
    // only a fully loaded workbook is committed to *this.
    Workbook staged;
    staged.loadInPlace(path, options);
    *this = std::move(staged);
}

void Workbook::load(std::istream& stream) {
    load(stream, LoadOptions{});
}

void Workbook::load(std::istream& stream, const LoadOptions& options) {
    internal::ScopedTemporaryFile temporary(internal::xlppTemporaryPath("stream_load"));
    {
        std::ofstream tmp(temporary.path(), std::ios::binary | std::ios::trunc);
        if (!tmp) throw std::runtime_error("Cannot create temporary workbook stream file");
        tmp << stream.rdbuf();
        if (!tmp) throw std::runtime_error("Failed writing temporary workbook stream file");
    }
    load(temporary.path(), options);
}

void Workbook::save(std::ostream& stream) const {
    save(stream, SaveOptions{});
}

void Workbook::save(std::ostream& stream, const SaveOptions& options) const {
    internal::ScopedTemporaryFile temporary(internal::xlppTemporaryPath("stream_save"));
    auto streamOptions = options;
    streamOptions.atomicWrite = false;
    saveInPlace(temporary.path(), streamOptions);

    std::ifstream tmp(temporary.path(), std::ios::binary);
    if (!tmp) throw std::runtime_error("Cannot reopen temporary workbook stream file");
    stream << tmp.rdbuf();
    if (!stream) throw std::runtime_error("Failed writing workbook output stream");
}

} // namespace xlpp
