#pragma once
#include <XLPP/Workbook/StructuralEdit.h>
#include <XLPP/Worksheet/Drawings/Image.h>
#include <cstddef>
#include <optional>
#include <string>
#include <string_view>

namespace xlpp::internal {

enum class StructuralAxis { Row, Column };
enum class StructuralAction { Insert, Delete };

struct StructuralEditSpec {
    StructuralAxis axis{StructuralAxis::Row};
    StructuralAction action{StructuralAction::Insert};
    std::size_t index{1};
    std::size_t amount{1};
    std::string targetSheetName;
};

struct ReferenceRewriteResult {
    std::string text;
    std::size_t referencesRewritten{0};
    std::size_t referencesInvalidated{0};
    std::size_t referencesSkippedUnsupported{0};
    bool changed() const noexcept { return referencesRewritten != 0 || referencesInvalidated != 0; }
};

// Rewrites A1 references embedded in formulas. References qualified with the
// target sheet are always considered. Unqualified references are considered
// only when ownerSheetName equals targetSheetName. String literals and
// structured-reference brackets are deliberately skipped.
ReferenceRewriteResult rewriteFormulaReferences(std::string_view formula,
                                                 std::string_view ownerSheetName,
                                                 const StructuralEditSpec& edit);

// Rewrites a worksheet sqref/reference expression (space/comma-separated A1
// areas, row ranges or column ranges). Invalidated areas are removed.
ReferenceRewriteResult rewriteReferenceList(std::string_view references,
                                             const StructuralEditSpec& edit);

// Rewrites one formula-style reference expression. This accepts an optional
// leading '=' and optional worksheet qualifier, and returns #REF! if the whole
// referenced interval is deleted.
ReferenceRewriteResult rewriteReferenceExpression(std::string_view reference,
                                                   std::string_view ownerSheetName,
                                                   const StructuralEditSpec& edit);

// Rewrites only explicit worksheet qualifiers. These helpers are used by
// workbook-level sheet rename/removal operations and deliberately leave
// external-workbook references and string literals untouched.
ReferenceRewriteResult renameWorksheetReferences(std::string_view formula,
                                                  std::string_view oldSheetName,
                                                  std::string_view newSheetName);
ReferenceRewriteResult invalidateWorksheetReferences(std::string_view formula,
                                                      std::string_view removedSheetName);

// Positional transformation used for UI/drawing anchors where deleting the
// anchored row/column clamps the marker to the edit boundary instead of
// producing #REF!.
std::size_t transformPosition(std::size_t coordinate, const StructuralEditSpec& edit) noexcept;
void transformDrawingAnchor(DrawingAnchorInfo& anchor, const StructuralEditSpec& edit) noexcept;

} // namespace xlpp::internal
