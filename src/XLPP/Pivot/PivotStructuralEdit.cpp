#include "PivotStructuralEdit.h"
#include <XLPP/Cell/CellReference.h>
#include <algorithm>
#include <cstddef>
#include <optional>
#include <string>
#include <vector>

namespace xlpp::internal {
namespace {
struct QualifiedRangeAddress {
    std::string sheetName;
    CellReference first;
    CellReference last;
};

std::string unescapeSheet(std::string value) {
    if (value.size() >= 2 && value.front() == '\'' && value.back() == '\'') {
        value = value.substr(1, value.size() - 2);
        std::size_t at = 0;
        while ((at = value.find("''", at)) != std::string::npos) {
            value.replace(at, 2, "'");
            ++at;
        }
    }
    return value;
}

std::optional<QualifiedRangeAddress> parseSource(std::string text,
                                                 std::string_view defaultSheet) {
    try {
        if (!text.empty() && text.front() == '=') text.erase(text.begin());
        std::string sheetName(defaultSheet);
        const auto bang = text.rfind('!');
        if (bang != std::string::npos) {
            sheetName = unescapeSheet(text.substr(0, bang));
            text = text.substr(bang + 1);
        }
        const auto colon = text.find(':');
        auto first = CellReference::parse(colon == std::string::npos ? text : text.substr(0, colon));
        auto last = CellReference::parse(colon == std::string::npos ? text : text.substr(colon + 1));
        if (first.row > last.row) std::swap(first.row, last.row);
        if (first.column > last.column) std::swap(first.column, last.column);
        return QualifiedRangeAddress{std::move(sheetName), first, last};
    } catch (...) {
        return std::nullopt;
    }
}

int mapFieldIndex(int value, std::size_t offset, std::size_t count, bool insert) noexcept {
    if (value < 0 || count == 0) return value;
    const int begin = static_cast<int>(offset);
    const int end = static_cast<int>(offset + count);
    const int delta = static_cast<int>(count);
    if (insert) return value >= begin ? value + delta : value;
    if (value >= end) return value - delta;
    if (value >= begin) return -1;
    return value;
}

template<class T>
void remapPrimaryFieldVector(std::vector<T>& fields,
                             std::size_t offset,
                             std::size_t count,
                             bool insert) {
    for (auto& field : fields)
        field.setFieldIndex(mapFieldIndex(field.fieldIndex(), offset, count, insert));
    if (!insert) {
        fields.erase(std::remove_if(fields.begin(), fields.end(), [](const auto& field) {
            return field.fieldIndex() < 0;
        }), fields.end());
    }
}

void repairPivotTableFieldReferences(PivotTable& pivot,
                                     std::size_t offset,
                                     std::size_t count,
                                     bool insert,
                                     StructuralEditReport& report) {
    remapPrimaryFieldVector(pivot.rowFields(), offset, count, insert);
    remapPrimaryFieldVector(pivot.columnFields(), offset, count, insert);
    remapPrimaryFieldVector(pivot.pageFields(), offset, count, insert);

    for (auto& field : pivot.dataFields()) {
        field.setFieldIndex(mapFieldIndex(field.fieldIndex(), offset, count, insert));
        field.setBaseField(mapFieldIndex(field.baseField(), offset, count, insert));
    }
    if (!insert) {
        auto& fields = pivot.dataFields();
        fields.erase(std::remove_if(fields.begin(), fields.end(), [](const auto& field) {
            return field.fieldIndex() < 0;
        }), fields.end());
    }

    for (auto& field : pivot.pageFieldSettings())
        field.setFieldIndex(mapFieldIndex(field.fieldIndex(), offset, count, insert));
    if (!insert) {
        auto& fields = pivot.pageFieldSettings();
        fields.erase(std::remove_if(fields.begin(), fields.end(), [](const auto& field) {
            return field.fieldIndex() < 0;
        }), fields.end());
    }

    for (auto& filter : pivot.filters()) {
        filter.fieldIndex = mapFieldIndex(filter.fieldIndex, offset, count, insert);
        filter.measureField = mapFieldIndex(filter.measureField, offset, count, insert);
        filter.memberPropertyField = mapFieldIndex(filter.memberPropertyField, offset, count, insert);
    }
    if (!insert) {
        auto& filters = pivot.filters();
        const auto before = filters.size();
        filters.erase(std::remove_if(filters.begin(), filters.end(), [](const auto& filter) {
            return filter.fieldIndex < 0;
        }), filters.end());
        if (filters.size() != before)
            report.diagnostics.push_back("Removed Pivot filter whose source field was deleted: " + pivot.name());
    }
}

void repairRows(PivotTable& pivot,
                const QualifiedRangeAddress& oldSource,
                const StructuralEditSpec& edit,
                StructuralEditReport& report) {
    auto& cache = pivot.cache();
    auto& records = cache.records();
    if (records.empty()) return;
    const auto expected = oldSource.last.row > oldSource.first.row
        ? oldSource.last.row - oldSource.first.row : 0;
    if (records.size() != expected) {
        report.diagnostics.push_back(
            "Pivot cache record count does not match source rows; cached rows were preserved and refresh requested: " + pivot.name());
        return;
    }

    if (edit.action == StructuralAction::Insert) {
        if (edit.index > oldSource.first.row && edit.index <= oldSource.last.row) {
            const auto offset = edit.index - oldSource.first.row - 1;
            records.insert(records.begin() + static_cast<std::ptrdiff_t>(offset),
                           edit.amount,
                           std::vector<std::string>(cache.fields().size()));
        }
        return;
    }

    const auto cutBegin = edit.index;
    const auto cutEnd = edit.index + edit.amount;
    std::size_t removeBeginRow = std::max(oldSource.first.row + 1, cutBegin);
    std::size_t removeEndRowExclusive = std::min(oldSource.last.row + 1, cutEnd);
    const bool headerDeleted = cutBegin <= oldSource.first.row && oldSource.first.row < cutEnd;
    if (headerDeleted && cutEnd <= oldSource.last.row)
        removeEndRowExclusive = std::min(oldSource.last.row + 1, cutEnd + 1);
    if (removeBeginRow >= removeEndRowExclusive) return;

    const auto beginOffset = removeBeginRow - oldSource.first.row - 1;
    const auto count = removeEndRowExclusive - removeBeginRow;
    const auto endOffset = std::min(records.size(), beginOffset + count);
    if (beginOffset < endOffset)
        records.erase(records.begin() + static_cast<std::ptrdiff_t>(beginOffset),
                      records.begin() + static_cast<std::ptrdiff_t>(endOffset));
}

void repairColumns(PivotTable& pivot,
                   const QualifiedRangeAddress& oldSource,
                   const StructuralEditSpec& edit,
                   StructuralEditReport& report) {
    auto& cache = pivot.cache();
    if (edit.action == StructuralAction::Insert) {
        if (edit.index > oldSource.first.column && edit.index <= oldSource.last.column) {
            const auto offset = edit.index - oldSource.first.column;
            cache.insertSourceFields(offset, edit.amount);
            repairPivotTableFieldReferences(pivot, offset, edit.amount, true, report);
        }
        return;
    }

    const auto cutBegin = edit.index;
    const auto cutEnd = edit.index + edit.amount;
    const auto overlapBegin = std::max(oldSource.first.column, cutBegin);
    const auto overlapEnd = std::min(oldSource.last.column + 1, cutEnd);
    if (overlapBegin >= overlapEnd) return;

    const auto offset = overlapBegin - oldSource.first.column;
    const auto requested = overlapEnd - overlapBegin;
    if (offset >= cache.fields().size()) {
        report.diagnostics.push_back(
            "Pivot source columns changed beyond the modeled cache-field width; refresh requested: " + pivot.name());
        return;
    }
    const auto safeCount = std::min(requested, cache.fields().size() - offset);
    cache.eraseSourceFields(offset, safeCount);
    repairPivotTableFieldReferences(pivot, offset, safeCount, false, report);
}
} // namespace

bool rewritePivotSourceForStructuralEdit(PivotTable& pivot,
                                         std::string_view ownerSheetName,
                                         const StructuralEditSpec& edit,
                                         StructuralEditReport& report) {
    auto& cache = pivot.cache();
    if (cache.sourceData().empty()) return false;

    const auto oldText = cache.sourceData();
    const auto oldSource = parseSource(oldText, ownerSheetName);
    const auto rewritten = rewriteFormulaReferences(oldText, ownerSheetName, edit);
    if (!rewritten.changed()) return false;

    if (oldSource && oldSource->sheetName == edit.targetSheetName) {
        // A fully deleted Pivot source is kept as a broken #REF! source while
        // preserving its cached schema/data, matching the preservation-first
        // behavior expected from an editor. This keeps the workbook savable
        // and lets the host display stale Pivot contents instead of silently
        // destroying the PivotTable.
        if (rewritten.referencesInvalidated == 0) {
            if (edit.axis == StructuralAxis::Row) repairRows(pivot, *oldSource, edit, report);
            else repairColumns(pivot, *oldSource, edit, report);
        }
        cache.setRefreshOnLoad(true);
    }

    cache.setSourceData(rewritten.text);
    report.pivotSourcesRewritten += rewritten.referencesRewritten;
    report.formulaReferencesInvalidated += rewritten.referencesInvalidated;
    if (rewritten.referencesInvalidated != 0)
        report.diagnostics.push_back(
            "Pivot cache source contains #REF! after structural deletion and requires user review: " + pivot.name());
    return true;
}

} // namespace xlpp::internal
