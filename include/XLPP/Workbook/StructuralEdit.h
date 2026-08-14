#pragma once
#include <cstddef>
#include <functional>
#include <stdexcept>
#include <string>
#include <vector>

namespace xlpp {

enum class StructuralEditKind {
    InsertRows,
    DeleteRows,
    InsertColumns,
    DeleteColumns
};

// Thrown when a caller-provided structural-edit cancellation callback requests
// cancellation. When rollbackOnFailure is enabled, the workbook model is
// restored before this exception reaches the caller.
class StructuralEditCancelled : public std::runtime_error {
public:
    StructuralEditCancelled() : std::runtime_error("Structural edit cancelled") {}
};

// Controls which dependent worksheet/workbook structures are rewritten when
// rows or columns are inserted/deleted through Workbook's structural-edit API.
// The defaults favor semantic correctness. Unsupported/ambiguous references
// are retained and surfaced through StructuralEditReport::diagnostics rather
// than being silently guessed.
struct StructuralEditOptions {
    bool updateCellFormulas{true};
    bool updateFormulaMetadata{true};
    bool updateHyperlinks{true};
    bool updateDefinedNames{true};
    bool updateMergedCells{true};
    bool updateFreezePanes{true};
    bool updateDimensions{true};
    bool updateAutoFilter{true};
    bool updateConditionalFormatting{true};
    bool updateDataValidations{true};
    bool updateTables{true};
    bool updatePrintSettings{true};
    bool updateCharts{true};
    bool synchronizeChartCaches{true};
    bool updatePivotSources{true};
    bool updateDrawings{true};

    // P1K transactional hardening. Structural edits touch multiple model
    // objects and may allocate while doing so. With this enabled, any
    // exception/cancellation restores worksheet contents, defined names and
    // calculation/cache state before the exception escapes. Worksheet object
    // identities are preserved during rollback.
    bool rollbackOnFailure{true};

    // Validate the semantic model after the edit and reject/rollback only if
    // the edit introduces a new validation ERROR compared with the pre-edit
    // model. Pre-existing errors do not make an otherwise unrelated edit fail.
    bool validateResult{true};

    // Optional cooperative cancellation hook for large edits. Returning true
    // at a checkpoint throws StructuralEditCancelled.
    std::function<bool()> cancel{};
};

struct StructuralEditReport {
    StructuralEditKind kind{StructuralEditKind::InsertRows};
    std::string worksheetName;
    std::size_t index{0};
    std::size_t amount{0};

    std::size_t cellsMoved{0};
    std::size_t cellsDeleted{0};
    std::size_t formulasRewritten{0};
    std::size_t formulaReferencesInvalidated{0};
    std::size_t referencesSkippedUnsupported{0};
    std::size_t formulaMetadataRewritten{0};
    std::size_t hyperlinksRewritten{0};
    std::size_t definedNamesRewritten{0};
    std::size_t worksheetRangesRewritten{0};
    std::size_t worksheetRangesRemoved{0};
    std::size_t tablesRewritten{0};
    std::size_t tablesRemoved{0};
    std::size_t chartReferencesRewritten{0};
    std::size_t chartCachesSynchronized{0};
    std::size_t pivotSourcesRewritten{0};
    std::size_t drawingAnchorsRewritten{0};
    std::size_t modelErrorsAfterEdit{0};
    std::size_t modelWarningsAfterEdit{0};
    std::vector<std::string> diagnostics;

    bool hasDiagnostics() const noexcept { return !diagnostics.empty(); }
};

} // namespace xlpp
