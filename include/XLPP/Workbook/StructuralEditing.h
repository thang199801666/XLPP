#pragma once
#include <XLPP/Formula/ReferenceTranslator.h>
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

struct StructuralEditOptions {
    // Stage on a workbook copy and commit only after all rewrites succeed.
    // Disable for very large workbooks when the caller already provides its
    // own transaction/rollback boundary.
    bool transactional{true};
    bool updateDefinedNames{true};
    bool recalculateFormulas{false};
    bool synchronizeChartCaches{true};
    bool changedChartCachesOnly{false};
    // If true, any dependency that becomes #REF! aborts a transactional edit.
    bool failOnInvalidReference{false};
};


struct WorksheetRenameOptions {
    bool recalculateFormulas{false};
    bool synchronizeChartCaches{true};
    bool changedChartCachesOnly{false};
};

struct WorksheetRenameReport {
    std::size_t worksheetsVisited{0};
    std::size_t formulasUpdated{0};
    std::size_t formulaMetadataUpdated{0};
    std::size_t definedNamesUpdated{0};
    std::size_t chartReferencesUpdated{0};
    std::size_t pivotReferencesUpdated{0};
    std::size_t hyperlinksUpdated{0};
    std::size_t referencesUpdated{0};
    std::size_t formulasCalculated{0};
    std::size_t chartCachesUpdated{0};
    std::vector<std::string> warnings;

    bool success() const noexcept { return warnings.empty(); }
};

struct StructuralEditReport {
    std::size_t worksheetsVisited{0};
    std::size_t cellsMoved{0};
    std::size_t cellsRemoved{0};
    std::size_t formulasUpdated{0};
    std::size_t formulaMetadataUpdated{0};
    std::size_t worksheetReferencesUpdated{0};
    std::size_t definedNamesUpdated{0};
    std::size_t chartReferencesUpdated{0};
    std::size_t pivotReferencesUpdated{0};
    std::size_t drawingAnchorsUpdated{0};
    std::size_t hyperlinksUpdated{0};
    std::size_t referencesInvalidated{0};
    std::size_t formulasCalculated{0};
    std::size_t chartCachesUpdated{0};
    std::vector<std::string> warnings;

    bool success() const noexcept { return warnings.empty(); }
};

} // namespace xlpp
