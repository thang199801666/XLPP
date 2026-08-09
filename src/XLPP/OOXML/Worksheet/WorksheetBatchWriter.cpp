#include "OOXML/Worksheet/WorksheetBatchWriter.h"
#include "OOXML/Worksheet/WorksheetWriter.h"
#include "Core/Threading/ThreadPool.h"
#include <algorithm>
namespace xlpp::internal::ooxml {
// Serialize every worksheet to XML, using a ThreadPool when workers > 1.
// Output is indexed by worksheet order and is identical to the sequential result.
// A cache snapshot is retained for deterministic output, but mutable-reference APIs require re-serialization for correctness.
std::vector<std::string> serializeSheets(const std::deque<xlpp::Worksheet>& sheets,
                                         const StyleCatalog& styles, const DxfCatalog& dxfs,
                                         bool date1904, bool strict, bool macroEnabled, std::size_t workers,
                                         bool parallelRows,
                                         const std::unordered_map<std::string, std::size_t>* sstIndex,
                                         std::vector<std::string>* cache,
                                         bool& cacheStrict, bool& cacheDate1904) {
    std::vector<std::string> result(sheets.size());
    std::vector<char> needsSerialize(sheets.size(), true);
    if (cache) {
        // Worksheet exposes mutable references (Cell&, Style&, AutoFilter&, ...).
        // A caller may keep one of those references and mutate it after a save,
        // bypassing Worksheet::dirty_. Re-serializing is therefore required for
        // correctness until mutations are tracked by owning proxy objects.
        if (cacheStrict != strict || cacheDate1904 != date1904) cache->clear();
    }

    // Serialize dirty sheets in parallel
    std::vector<std::size_t> dirtyIndexes;
    for (std::size_t i = 0; i < sheets.size(); ++i)
        if (needsSerialize[i]) dirtyIndexes.push_back(i);

    auto serializeOne = [&](std::size_t i) {
        const auto codeName = macroEnabled ? sheets[i].vbaCodeName() : std::string{};
        result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, 0, codeName);
    };

    if (workers > 1 && dirtyIndexes.size() > 1) {
        xlpp::internal::ThreadPool pool(std::min(workers, dirtyIndexes.size()));
        pool.parallelFor(0, dirtyIndexes.size(), [&](std::size_t j) {
            serializeOne(dirtyIndexes[j]);
        });
    } else if (parallelRows && workers > 1 && dirtyIndexes.size() == 1) {
        for (auto i : dirtyIndexes) {
            const auto codeName = macroEnabled ? sheets[i].vbaCodeName() : std::string{};
            result[i] = sheetXml(sheets[i], styles, dxfs, date1904, strict, sstIndex, workers, codeName);
        }
    } else {
        for (auto i : dirtyIndexes) serializeOne(i);
    }

    // Update cache
    if (cache) {
        if (cache->size() < sheets.size()) cache->resize(sheets.size());
        for (std::size_t i = 0; i < sheets.size(); ++i)
            (*cache)[i] = result[i];
        cacheStrict = strict;
        cacheDate1904 = date1904;
    }
    return result;
}


} // namespace xlpp::internal::ooxml
