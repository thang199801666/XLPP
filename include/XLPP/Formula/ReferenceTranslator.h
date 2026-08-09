#pragma once
#include <cstddef>
#include <string>
#include <string_view>

namespace xlpp {

enum class StructuralEditKind {
    InsertRows,
    DeleteRows,
    InsertColumns,
    DeleteColumns
};

struct StructuralEdit {
    std::string sheetName;
    StructuralEditKind kind{StructuralEditKind::InsertRows};
    std::size_t index{1};
    std::size_t amount{1};
};

struct ReferenceTranslationResult {
    std::string value;
    std::size_t referencesVisited{0};
    std::size_t referencesChanged{0};
    std::size_t referencesInvalidated{0};
    bool changed() const noexcept { return referencesChanged != 0 || referencesInvalidated != 0; }
};

// Translate A1-style references in an Excel formula after a structural edit.
// String literals are preserved verbatim. Absolute markers affect copy/fill
// semantics, not structural edits, so $A$1 is shifted exactly like A1.
ReferenceTranslationResult translateFormulaReferences(
    std::string_view formula,
    std::string_view contextSheet,
    const StructuralEdit& edit);

// Translate a reference/sqref expression (merged ranges, filters, validation,
// print areas, table refs, etc.). Supports cell ranges plus whole-row/column
// ranges and multiple areas separated by whitespace or commas.
ReferenceTranslationResult translateRangeReferences(
    std::string_view reference,
    std::string_view contextSheet,
    const StructuralEdit& edit);

// Rewrite explicit local worksheet qualifiers after a worksheet rename. String
// literals and external-workbook qualifiers are preserved. 3-D qualifiers
// such as Sheet1:Sheet3 are updated endpoint-by-endpoint.
ReferenceTranslationResult renameWorksheetReferences(
    std::string_view expression,
    std::string_view oldWorksheetName,
    std::string_view newWorksheetName);

// Replace explicit local qualifiers for a deleted worksheet with Excel's
// #REF! qualifier while preserving string literals and external workbooks.
ReferenceTranslationResult invalidateWorksheetReferences(
    std::string_view expression,
    std::string_view removedWorksheetName);

} // namespace xlpp
