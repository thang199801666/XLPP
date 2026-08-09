#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Formula/ReferenceTranslator.h>
#include "WorksheetReferenceSupport.h"
#include <algorithm>
#include <cctype>
#include <string_view>
#include <utility>

namespace xlpp {
void Worksheet::insertRows(std::size_t index, std::size_t amount) {
    if (index == 0 || amount == 0) throw std::invalid_argument("Row index and amount must be positive");
    applyStructuralEdit(StructuralEdit{name_, StructuralEditKind::InsertRows, index, amount});
}

void Worksheet::deleteRows(std::size_t index, std::size_t amount) {
    if (index == 0 || amount == 0) throw std::invalid_argument("Row index and amount must be positive");
    applyStructuralEdit(StructuralEdit{name_, StructuralEditKind::DeleteRows, index, amount});
}

void Worksheet::insertColumns(std::size_t index, std::size_t amount) {
    if (index == 0 || amount == 0) throw std::invalid_argument("Column index and amount must be positive");
    applyStructuralEdit(StructuralEdit{name_, StructuralEditKind::InsertColumns, index, amount});
}

void Worksheet::deleteColumns(std::size_t index, std::size_t amount) {
    if (index == 0 || amount == 0) throw std::invalid_argument("Column index and amount must be positive");
    applyStructuralEdit(StructuralEdit{name_, StructuralEditKind::DeleteColumns, index, amount});
}

namespace {
bool structuralSameSheet(std::string_view a, std::string_view b) {
    if (a.size() != b.size()) return false;
    for (std::size_t i = 0; i < a.size(); ++i) {
        if (std::tolower(static_cast<unsigned char>(a[i])) !=
            std::tolower(static_cast<unsigned char>(b[i]))) return false;
    }
    return true;
}

void mergeStructuralReport(WorksheetStructuralEditReport& into, const WorksheetStructuralEditReport& from) {
    into.cellsMoved += from.cellsMoved;
    into.cellsRemoved += from.cellsRemoved;
    into.formulasUpdated += from.formulasUpdated;
    into.formulaMetadataUpdated += from.formulaMetadataUpdated;
    into.worksheetReferencesUpdated += from.worksheetReferencesUpdated;
    into.referencesInvalidated += from.referencesInvalidated;
    into.drawingAnchorsUpdated += from.drawingAnchorsUpdated;
    into.chartReferencesUpdated += from.chartReferencesUpdated;
    into.pivotReferencesUpdated += from.pivotReferencesUpdated;
    into.hyperlinksUpdated += from.hyperlinksUpdated;
}

bool translationContainsInvalid(const ReferenceTranslationResult& tr) {
    return tr.referencesInvalidated != 0 || tr.value.find("#REF!") != std::string::npos;
}
}

WorksheetStructuralEditReport Worksheet::applyStructuralEdit(const StructuralEdit& edit) {
    if (edit.sheetName.empty() || edit.index == 0 || edit.amount == 0)
        throw std::invalid_argument("Structural edit requires a target worksheet and positive index/amount");

    WorksheetStructuralEditReport report;
    const bool target = structuralSameSheet(name_, edit.sheetName);
    if (target) {
        const bool rowEdit = edit.kind == StructuralEditKind::InsertRows || edit.kind == StructuralEditKind::DeleteRows;
        const auto limit = rowEdit ? MaxExcelRows : MaxExcelColumns;
        if (edit.index > limit || edit.amount > limit)
            throw std::out_of_range(rowEdit ? "Structural row edit exceeds Excel worksheet bounds"
                                           : "Structural column edit exceeds Excel worksheet bounds");
        if (edit.kind == StructuralEditKind::InsertRows && maxRow() > MaxExcelRows - edit.amount)
            throw std::out_of_range("Row insertion would exceed Excel's 1,048,576-row limit");
        if (edit.kind == StructuralEditKind::InsertColumns && maxColumn() > MaxExcelColumns - edit.amount)
            throw std::out_of_range("Column insertion would exceed Excel's 16,384-column limit");
        if (edit.kind == StructuralEditKind::InsertRows)
            mergeStructuralReport(report, shiftRows(edit.index, edit.amount, true));
        else if (edit.kind == StructuralEditKind::DeleteRows)
            mergeStructuralReport(report, shiftRows(edit.index, edit.amount, false));
        else if (edit.kind == StructuralEditKind::InsertColumns)
            mergeStructuralReport(report, shiftColumns(edit.index, edit.amount, true));
        else
            mergeStructuralReport(report, shiftColumns(edit.index, edit.amount, false));
    }
    mergeStructuralReport(report, translateStructuralReferences(edit, target));
    if (target || report.formulasUpdated || report.worksheetReferencesUpdated ||
        report.chartReferencesUpdated || report.pivotReferencesUpdated || report.hyperlinksUpdated)
        dirty_ = true;
    return report;
}

WorksheetStructuralEditReport Worksheet::shiftRows(std::size_t index, std::size_t amount, bool insert) {
    WorksheetStructuralEditReport report;
    std::map<std::uint64_t, Cell> shifted;
    const auto deleteEnd = index + amount;
    for (auto& [key, source] : cells_) {
        auto row = source.row();
        const auto column = source.column();
        if (insert) {
            if (row >= index) { row += amount; ++report.cellsMoved; }
        } else {
            if (row >= index && row < deleteEnd) { ++report.cellsRemoved; continue; }
            if (row >= deleteEnd) { row -= amount; ++report.cellsMoved; }
        }
        source.setPosition(row, column);
        shifted.emplace(makeCellKey(row, column), std::move(source));
    }
    cells_ = std::move(shifted);

    std::map<std::size_t, RowDimension> shiftedRows;
    for (auto& [row, dimension] : rowDimensions_) {
        auto targetRow = row;
        if (insert) {
            if (targetRow >= index) targetRow += amount;
        } else {
            if (targetRow >= index && targetRow < deleteEnd) continue;
            if (targetRow >= deleteEnd) targetRow -= amount;
        }
        shiftedRows.emplace(targetRow, std::move(dimension));
    }
    rowDimensions_ = std::move(shiftedRows);
    dirty_ = true;
    return report;
}

WorksheetStructuralEditReport Worksheet::shiftColumns(std::size_t index, std::size_t amount, bool insert) {
    WorksheetStructuralEditReport report;
    std::map<std::uint64_t, Cell> shifted;
    const auto deleteEnd = index + amount;
    for (auto& [key, source] : cells_) {
        const auto row = source.row();
        auto column = source.column();
        if (insert) {
            if (column >= index) { column += amount; ++report.cellsMoved; }
        } else {
            if (column >= index && column < deleteEnd) { ++report.cellsRemoved; continue; }
            if (column >= deleteEnd) { column -= amount; ++report.cellsMoved; }
        }
        source.setPosition(row, column);
        shifted.emplace(makeCellKey(row, column), std::move(source));
    }
    cells_ = std::move(shifted);

    std::map<std::size_t, ColumnDimension> shiftedColumns;
    for (auto& [column, dimension] : columnDimensions_) {
        auto targetColumn = column;
        if (insert) {
            if (targetColumn >= index) targetColumn += amount;
        } else {
            if (targetColumn >= index && targetColumn < deleteEnd) continue;
            if (targetColumn >= deleteEnd) targetColumn -= amount;
        }
        shiftedColumns.emplace(targetColumn, std::move(dimension));
    }
    columnDimensions_ = std::move(shiftedColumns);
    dirty_ = true;
    return report;
}

WorksheetStructuralEditReport Worksheet::translateStructuralReferences(const StructuralEdit& edit, bool targetWorksheet) {
    WorksheetStructuralEditReport report;
    auto account = [&](const ReferenceTranslationResult& tr, std::size_t& bucket) {
        bucket += tr.referencesChanged + tr.referencesInvalidated;
        report.referencesInvalidated += tr.referencesInvalidated;
    };

    // Cell formulas, shared/array/spill metadata, and internal hyperlinks.
    for (auto& [_, cellValue] : cells_) {
        if (cellValue.hasFormula()) {
            auto tr = translateFormulaReferences(cellValue.formula(), name_, edit);
            if (tr.changed()) {
                cellValue.setFormula(std::move(tr.value));
                ++report.formulasUpdated;
                report.referencesInvalidated += tr.referencesInvalidated;
            }
        }
        auto& metadata = cellValue.formulaMetadata();
        if (!metadata.reference().empty()) {
            auto tr = translateRangeReferences(metadata.reference(), name_, edit);
            if (tr.changed()) {
                metadata.setReference(std::move(tr.value));
                ++report.formulaMetadataUpdated;
                report.referencesInvalidated += tr.referencesInvalidated;
            }
        }
        if (cellValue.hasHyperlink() && !cellValue.hyperlinkValue()->external()) {
            auto current = cellValue.hyperlinkValue()->target();
            const bool hash = !current.empty() && current.front() == '#';
            std::string_view body = hash ? std::string_view(current).substr(1) : std::string_view(current);
            auto tr = translateRangeReferences(body, name_, edit);
            if (tr.changed()) {
                if (hash) tr.value.insert(tr.value.begin(), '#');
                cellValue.hyperlink().setTarget(std::move(tr.value));
                ++report.hyperlinksUpdated;
                report.referencesInvalidated += tr.referencesInvalidated;
            }
        }
    }

    // Worksheet-owned geometry/reference collections. Unqualified references
    // only move on the target sheet; cross-sheet qualified formulas still move
    // on non-target sheets through the translator's context rules.
    if (targetWorksheet) {
        std::vector<std::string> newMerged;
        newMerged.reserve(mergedRanges_.size());
        for (const auto& range : mergedRanges_) {
            auto tr = translateRangeReferences(range, name_, edit);
            account(tr, report.worksheetReferencesUpdated);
            if (!translationContainsInvalid(tr)) newMerged.push_back(std::move(tr.value));
        }
        mergedRanges_ = std::move(newMerged);
        mergedRangesParsed_.clear();
        for (const auto& range : mergedRanges_) {
            try {
                const auto [first, last] = internal::parseWorksheetRangeAddress(range);
                mergedRangesParsed_.push_back({first.row, first.column, last.row, last.column});
            } catch (...) {}
        }

        if (freezePane_) {
            auto tr = translateRangeReferences(*freezePane_, name_, edit);
            account(tr, report.worksheetReferencesUpdated);
            if (translationContainsInvalid(tr)) freezePane_.reset();
            else freezePane_ = std::move(tr.value);
        }

        if (autoFilter_.enabled()) {
            const auto oldReference = autoFilter_.reference();
            auto tr = translateRangeReferences(oldReference, name_, edit);
            account(tr, report.worksheetReferencesUpdated);
            if (translationContainsInvalid(tr)) {
                autoFilter_.clear();
            } else {
                autoFilter_.setReference(std::move(tr.value));
                if (autoFilter_.sortStateValue()) {
                    const auto oldSort = *autoFilter_.sortStateValue();
                    auto& sort = autoFilter_.sortState();
                    if (!oldSort.reference().empty()) {
                        auto rt = translateRangeReferences(oldSort.reference(), name_, edit);
                        sort.setReference(std::move(rt.value));
                        account(rt, report.worksheetReferencesUpdated);
                    }
                    std::vector<SortCondition> conditions = oldSort.conditions();
                    sort.clear();
                    if (!oldSort.reference().empty()) {
                        auto rt = translateRangeReferences(oldSort.reference(), name_, edit);
                        sort.setReference(std::move(rt.value));
                    }
                    sort.setCaseSensitive(oldSort.caseSensitive());
                    for (const auto& condition : conditions) {
                        auto ct = translateRangeReferences(condition.reference, name_, edit);
                        if (!translationContainsInvalid(ct)) sort.addCondition(std::move(ct.value), condition.descending);
                        account(ct, report.worksheetReferencesUpdated);
                    }
                }
            }
        }
    }

    // Conditional formatting: sqref is local geometry, formulas may reference
    // the edited sheet even when the rule belongs to another worksheet.
    auto& cfEntries = conditionalFormatting_.entries();
    for (auto entryIt = cfEntries.begin(); entryIt != cfEntries.end();) {
        if (targetWorksheet) {
            auto tr = translateRangeReferences(entryIt->reference(), name_, edit);
            account(tr, report.worksheetReferencesUpdated);
            if (translationContainsInvalid(tr)) { entryIt = cfEntries.erase(entryIt); continue; }
            entryIt->setReference(std::move(tr.value));
        }
        for (auto& rule : entryIt->rules()) {
            auto formulas = rule.formulas();
            bool changed = false;
            for (auto& formula : formulas) {
                auto tr = translateFormulaReferences(formula, name_, edit);
                if (tr.changed()) { formula = std::move(tr.value); changed = true; }
                account(tr, report.worksheetReferencesUpdated);
            }
            if (changed) rule.setFormulas(std::move(formulas));
            auto translateCfvo = [&](Cfvo& cfvo) {
                if (cfvo.formula.empty()) return;
                auto tr = translateFormulaReferences(cfvo.formula, name_, edit);
                if (tr.changed()) cfvo.formula = std::move(tr.value);
                account(tr, report.worksheetReferencesUpdated);
            };
            translateCfvo(rule.getDataBar().min);
            translateCfvo(rule.getDataBar().max);
            for (auto& stop : rule.getColorScale().stops) translateCfvo(stop);
            for (auto& stop : rule.getIconSet().thresholds) translateCfvo(stop);
        }
        ++entryIt;
    }

    auto& validations = dataValidations_.items();
    for (auto it = validations.begin(); it != validations.end();) {
        if (targetWorksheet) {
            auto tr = translateRangeReferences(it->reference(), name_, edit);
            account(tr, report.worksheetReferencesUpdated);
            if (translationContainsInvalid(tr)) { it = validations.erase(it); continue; }
            it->setReference(std::move(tr.value));
        }
        if (!it->formula1().empty()) {
            auto tr = translateFormulaReferences(it->formula1(), name_, edit);
            if (tr.changed()) it->setFormula1(std::move(tr.value));
            account(tr, report.worksheetReferencesUpdated);
        }
        if (!it->formula2().empty()) {
            auto tr = translateFormulaReferences(it->formula2(), name_, edit);
            if (tr.changed()) it->setFormula2(std::move(tr.value));
            account(tr, report.worksheetReferencesUpdated);
        }
        ++it;
    }

    if (targetWorksheet) {
        for (auto it = tables_.begin(); it != tables_.end();) {
            const auto oldReference = it->reference();
            std::optional<std::pair<CellReference, CellReference>> oldBounds;
            try { oldBounds = internal::parseWorksheetRangeAddress(oldReference); } catch (...) {}
            auto tr = translateRangeReferences(oldReference, name_, edit);
            account(tr, report.worksheetReferencesUpdated);
            if (translationContainsInvalid(tr)) { it = tables_.erase(it); continue; }

            // Keep the table column model consistent with a reference whose
            // width changed. Inserting strictly inside a table creates new
            // placeholder columns; deleting overlap removes the corresponding
            // columns. Header text can later replace the generated names.
            if (oldBounds && (edit.kind == StructuralEditKind::InsertColumns ||
                              edit.kind == StructuralEditKind::DeleteColumns)) {
                const auto firstCol = oldBounds->first.column;
                const auto lastCol = oldBounds->second.column;
                auto names = std::vector<std::string>{};
                names.reserve(it->columns().size() + edit.amount);
                for (const auto& column : it->columns()) names.push_back(column.name());
                if (edit.kind == StructuralEditKind::InsertColumns &&
                    edit.index > firstCol && edit.index <= lastCol) {
                    const auto offset = std::min<std::size_t>(edit.index - firstCol, names.size());
                    for (std::size_t n = 0; n < edit.amount; ++n)
                        names.insert(names.begin() + static_cast<std::ptrdiff_t>(offset + n),
                                     "Column" + std::to_string(offset + n + 1));
                } else if (edit.kind == StructuralEditKind::DeleteColumns && !names.empty()) {
                    const auto deleteFirst = edit.index;
                    const auto deleteLast = edit.index + edit.amount - 1;
                    const auto overlapFirst = std::max(firstCol, deleteFirst);
                    const auto overlapLast = std::min(lastCol, deleteLast);
                    if (overlapFirst <= overlapLast) {
                        const auto offset = overlapFirst - firstCol;
                        const auto count = overlapLast - overlapFirst + 1;
                        if (offset < names.size()) {
                            const auto eraseCount = std::min(count, names.size() - offset);
                            names.erase(names.begin() + static_cast<std::ptrdiff_t>(offset),
                                        names.begin() + static_cast<std::ptrdiff_t>(offset + eraseCount));
                        }
                    }
                }
                if (!names.empty()) {
                    auto& columns = it->columns();
                    columns.clear();
                    for (std::size_t n = 0; n < names.size(); ++n)
                        columns.emplace_back(n + 1, std::move(names[n]));
                }
            }
            it->setReference(std::move(tr.value));
            ++it;
        }

        auto translateStoredRange = [&](std::string& value) {
            if (value.empty()) return;
            auto tr = translateRangeReferences(value, name_, edit);
            if (tr.changed()) value = std::move(tr.value);
            account(tr, report.worksheetReferencesUpdated);
        };
        translateStoredRange(printArea_);
        translateStoredRange(printTitlesRows_);
        translateStoredRange(printTitlesCols_);
    }

    // Chart data sources are dependencies, not local worksheet geometry, so
    // qualified references update from any worksheet. Imported charts record
    // selective series edits so unsupported sibling ChartML remains verbatim.
    for (std::size_t chartIndex = 0; chartIndex < charts_.size(); ++chartIndex) {
        auto& chart = charts_[chartIndex];
        for (std::size_t seriesIndex = 0; seriesIndex < chart.series_.size(); ++seriesIndex) {
            auto& series = chart.series_[seriesIndex];
            auto cats = translateFormulaReferences(series.categoriesReference(), name_, edit);
            auto vals = translateFormulaReferences(series.valuesReference(), name_, edit);
            auto title = translateFormulaReferences(series.titleReference(), name_, edit);
            bool changed = cats.changed() || vals.changed() || title.changed();
            report.referencesInvalidated += cats.referencesInvalidated + vals.referencesInvalidated + title.referencesInvalidated;
            if (changed) {
                if (chart.imported_ && !chart.stableId_.empty() &&
                    !cats.value.empty() && !vals.value.empty()) {
                    setChartSeriesReferences(chart.stableId_, seriesIndex, cats.value, vals.value);
                } else {
                    if (cats.changed()) series.setCategoriesReference(std::move(cats.value));
                    if (vals.changed()) series.setValuesReference(std::move(vals.value));
                }
                if (title.changed()) series.setTitleReference(std::move(title.value));
                ++report.chartReferencesUpdated;
            }
            auto errorBars = series.errorBars();
            bool errorChanged = false;
            for (auto& error : errorBars) {
                if (!error.plusReference.empty()) {
                    auto tr = translateFormulaReferences(error.plusReference, name_, edit);
                    if (tr.changed()) { error.plusReference = std::move(tr.value); errorChanged = true; }
                    report.referencesInvalidated += tr.referencesInvalidated;
                }
                if (!error.minusReference.empty()) {
                    auto tr = translateFormulaReferences(error.minusReference, name_, edit);
                    if (tr.changed()) { error.minusReference = std::move(tr.value); errorChanged = true; }
                    report.referencesInvalidated += tr.referencesInvalidated;
                }
            }
            if (errorChanged) {
                series.setErrorBars(std::move(errorBars));
                ++report.chartReferencesUpdated;
            }
        }
    }

    for (auto& pivot : pivotTables_) {
        if (!pivot.cache().sourceData().empty()) {
            const auto oldSource = pivot.cache().sourceData();
            auto tr = translateRangeReferences(oldSource, name_, edit);
            if (tr.changed()) {
                bool widthMayChange = false;
                try {
                    auto sourceRange = oldSource;
                    const auto bang = sourceRange.find('!');
                    if (bang != std::string::npos) sourceRange = sourceRange.substr(bang + 1);
                    const auto bounds = internal::parseWorksheetRangeAddress(sourceRange);
                    widthMayChange = (edit.kind == StructuralEditKind::InsertColumns ||
                                      edit.kind == StructuralEditKind::DeleteColumns) &&
                                     edit.index > bounds.first.column && edit.index <= bounds.second.column;
                } catch (...) {}
                pivot.cache().setSourceData(std::move(tr.value));
                // Source values are no longer authoritative after a structural
                // edit. Let the serializer rebuild records from the worksheet.
                pivot.cache().clearRecords();
                if (widthMayChange) pivot.cache().setFields({});
                ++report.pivotReferencesUpdated;
            }
            report.referencesInvalidated += tr.referencesInvalidated;
        }
        if (targetWorksheet && !pivot.location().empty()) {
            auto tr = translateRangeReferences(pivot.location(), name_, edit);
            if (tr.changed() && !translationContainsInvalid(tr)) { pivot.setLocation(std::move(tr.value)); ++report.pivotReferencesUpdated; }
            report.referencesInvalidated += tr.referencesInvalidated;
        }
    }

    if (targetWorksheet) {
        auto translateAnchor = [&](DrawingAnchorInfo& anchor) -> bool {
            if (anchor.type == DrawingAnchorType::Absolute) return false;
            const auto before = CellReference{anchor.from.row, anchor.from.column}.address();
            const auto after = CellReference{anchor.to.row, anchor.to.column}.address();
            auto tr = translateRangeReferences(before + ":" + after, name_, edit);
            if (!tr.changed() || translationContainsInvalid(tr)) return false;
            try {
                const auto [first, last] = internal::parseWorksheetRangeAddress(tr.value);
                anchor.from.row = first.row; anchor.from.column = first.column;
                anchor.to.row = last.row; anchor.to.column = last.column;
                return true;
            } catch (...) { return false; }
        };

        for (auto& image : images_) {
            if (image.imported_) {
                auto updated = image.anchorInfo_;
                if (translateAnchor(updated)) {
                    auto& imported = ensureImportedImageEdit(image);
                    imported.anchor = updated;
                    imported.moved = true;
                    image.anchorInfo_ = updated;
                    image.anchor_ = CellReference{updated.from.row, updated.from.column}.address();
                    ++report.drawingAnchorsUpdated;
                    drawingAppendDirty_ = true;
                }
            } else {
                auto tr = translateRangeReferences(image.anchor_, name_, edit);
                if (tr.changed() && !translationContainsInvalid(tr)) {
                    image.anchor_ = std::move(tr.value);
                    ++report.drawingAnchorsUpdated;
                    drawingsDirty_ = true;
                }
            }
        }
        for (auto& chart : charts_) {
            if (!chart.imported_) continue;
            auto updated = chart.anchorInfo_;
            if (translateAnchor(updated)) {
                auto& imported = ensureImportedChartEdit(chart);
                imported.anchor = updated;
                imported.moved = true;
                chart.anchorInfo_ = updated;
                ++report.drawingAnchorsUpdated;
                drawingAppendDirty_ = true;
            }
        }
    }

    if (report.pivotReferencesUpdated) pivotsDirty_ = true;
    return report;
}

WorksheetStructuralEditReport Worksheet::translateWorksheetRenameReferences(
    std::string_view oldName, std::string_view newName) {
    WorksheetStructuralEditReport report;
    auto rename = [&](std::string& value, std::size_t& bucket) {
        if (value.empty()) return;
        auto tr = renameWorksheetReferences(value, oldName, newName);
        if (tr.changed()) {
            value = std::move(tr.value);
            bucket += tr.referencesChanged;
        }
    };

    for (auto& [_, cellValue] : cells_) {
        if (cellValue.hasFormula()) {
            auto formula = cellValue.formula();
            auto tr = renameWorksheetReferences(formula, oldName, newName);
            if (tr.changed()) {
                cellValue.setFormula(std::move(tr.value));
                ++report.formulasUpdated;
            }
        }
        auto& metadata = cellValue.formulaMetadata();
        if (!metadata.reference().empty()) {
            auto tr = renameWorksheetReferences(metadata.reference(), oldName, newName);
            if (tr.changed()) {
                metadata.setReference(std::move(tr.value));
                ++report.formulaMetadataUpdated;
            }
        }
        if (cellValue.hasHyperlink() && !cellValue.hyperlinkValue()->external()) {
            auto target = cellValue.hyperlinkValue()->target();
            auto tr = renameWorksheetReferences(target, oldName, newName);
            if (tr.changed()) {
                cellValue.hyperlink().setTarget(std::move(tr.value));
                report.hyperlinksUpdated += tr.referencesChanged;
            }
        }
    }

    // Conditional-format and validation formulas may explicitly reference a
    // different worksheet even though their sqref geometry is local.
    for (auto& entry : conditionalFormatting_.entries()) {
        auto reference = entry.reference();
        rename(reference, report.worksheetReferencesUpdated);
        if (reference != entry.reference()) entry.setReference(std::move(reference));
        for (auto& rule : entry.rules()) {
            auto formulas = rule.formulas();
            bool formulasChanged = false;
            for (auto& formula : formulas) {
                const auto before = formula;
                rename(formula, report.worksheetReferencesUpdated);
                formulasChanged = formulasChanged || formula != before;
            }
            if (formulasChanged) rule.setFormulas(std::move(formulas));
            auto renameCfvo = [&](Cfvo& cfvo) {
                rename(cfvo.formula, report.worksheetReferencesUpdated);
            };
            renameCfvo(rule.getDataBar().min);
            renameCfvo(rule.getDataBar().max);
            for (auto& stop : rule.getColorScale().stops) renameCfvo(stop);
            for (auto& stop : rule.getIconSet().thresholds) renameCfvo(stop);
        }
    }
    for (auto& validation : dataValidations_.items()) {
        auto reference = validation.reference();
        rename(reference, report.worksheetReferencesUpdated);
        if (reference != validation.reference()) validation.setReference(std::move(reference));
        auto formula1 = validation.formula1();
        rename(formula1, report.worksheetReferencesUpdated);
        if (formula1 != validation.formula1()) validation.setFormula1(std::move(formula1));
        auto formula2 = validation.formula2();
        rename(formula2, report.worksheetReferencesUpdated);
        if (formula2 != validation.formula2()) validation.setFormula2(std::move(formula2));
    }

    auto filterRef = autoFilter_.reference();
    rename(filterRef, report.worksheetReferencesUpdated);
    if (autoFilter_.enabled() && filterRef != autoFilter_.reference()) autoFilter_.setReference(std::move(filterRef));
    rename(printArea_, report.worksheetReferencesUpdated);
    rename(printTitlesRows_, report.worksheetReferencesUpdated);
    rename(printTitlesCols_, report.worksheetReferencesUpdated);

    for (std::size_t chartIndex = 0; chartIndex < charts_.size(); ++chartIndex) {
        auto& chart = charts_[chartIndex];
        for (std::size_t seriesIndex = 0; seriesIndex < chart.series_.size(); ++seriesIndex) {
            auto& series = chart.series_[seriesIndex];
            auto cats = renameWorksheetReferences(series.categoriesReference(), oldName, newName);
            auto vals = renameWorksheetReferences(series.valuesReference(), oldName, newName);
            auto title = renameWorksheetReferences(series.titleReference(), oldName, newName);
            if (cats.changed() || vals.changed()) {
                if (chart.imported_ && !chart.stableId_.empty() && !cats.value.empty() && !vals.value.empty())
                    setChartSeriesReferences(chart.stableId_, seriesIndex, cats.value, vals.value);
                else {
                    if (cats.changed()) series.setCategoriesReference(std::move(cats.value));
                    if (vals.changed()) series.setValuesReference(std::move(vals.value));
                }
                ++report.chartReferencesUpdated;
            }
            if (title.changed()) {
                series.setTitleReference(std::move(title.value));
                ++report.chartReferencesUpdated;
            }
            auto errorBars = series.errorBars();
            bool errorChanged = false;
            for (auto& error : errorBars) {
                auto plus = renameWorksheetReferences(error.plusReference, oldName, newName);
                if (plus.changed()) { error.plusReference = std::move(plus.value); errorChanged = true; }
                auto minus = renameWorksheetReferences(error.minusReference, oldName, newName);
                if (minus.changed()) { error.minusReference = std::move(minus.value); errorChanged = true; }
            }
            if (errorChanged) {
                series.setErrorBars(std::move(errorBars));
                ++report.chartReferencesUpdated;
            }
        }
    }

    for (auto& pivot : pivotTables_) {
        auto source = pivot.cache().sourceData();
        auto tr = renameWorksheetReferences(source, oldName, newName);
        if (tr.changed()) {
            pivot.cache().setSourceData(std::move(tr.value));
            pivot.cache().clearRecords();
            ++report.pivotReferencesUpdated;
        }
        auto location = pivot.location();
        tr = renameWorksheetReferences(location, oldName, newName);
        if (tr.changed()) {
            pivot.setLocation(std::move(tr.value));
            ++report.pivotReferencesUpdated;
        }
    }

    if (report.formulasUpdated || report.formulaMetadataUpdated ||
        report.worksheetReferencesUpdated || report.chartReferencesUpdated ||
        report.pivotReferencesUpdated || report.hyperlinksUpdated) dirty_ = true;
    if (report.pivotReferencesUpdated) pivotsDirty_ = true;
    return report;
}

WorksheetStructuralEditReport Worksheet::invalidateWorksheetReferencesForRemoval(
    std::string_view removedName) {
    WorksheetStructuralEditReport report;

    for (auto& [_, cellValue] : cells_) {
        if (cellValue.hasFormula()) {
            auto tr = invalidateWorksheetReferences(cellValue.formula(), removedName);
            if (tr.changed()) {
                cellValue.setFormula(std::move(tr.value));
                ++report.formulasUpdated;
                report.referencesInvalidated += tr.referencesInvalidated;
            }
        }
        auto& metadata = cellValue.formulaMetadata();
        if (!metadata.reference().empty()) {
            auto tr = invalidateWorksheetReferences(metadata.reference(), removedName);
            if (tr.changed()) {
                metadata.setReference(std::move(tr.value));
                ++report.formulaMetadataUpdated;
                report.referencesInvalidated += tr.referencesInvalidated;
            }
        }
        if (cellValue.hasHyperlink() && !cellValue.hyperlinkValue()->external()) {
            auto tr = invalidateWorksheetReferences(cellValue.hyperlinkValue()->target(), removedName);
            if (tr.changed()) {
                cellValue.hyperlink().setTarget(std::move(tr.value));
                report.hyperlinksUpdated += tr.referencesInvalidated;
                report.referencesInvalidated += tr.referencesInvalidated;
            }
        }
    }

    auto invalidateFormula = [&](std::string& formula) {
        if (formula.empty()) return std::size_t{0};
        auto tr = invalidateWorksheetReferences(formula, removedName);
        const auto invalidated = tr.referencesInvalidated;
        if (tr.changed()) formula = std::move(tr.value);
        report.referencesInvalidated += invalidated;
        return invalidated;
    };
    for (auto& entry : conditionalFormatting_.entries()) {
        for (auto& rule : entry.rules()) {
            auto formulas = rule.formulas();
            bool changed = false;
            for (auto& formula : formulas) {
                const auto before = formula;
                report.worksheetReferencesUpdated += invalidateFormula(formula);
                changed = changed || formula != before;
            }
            if (changed) rule.setFormulas(std::move(formulas));
            auto invalidateCfvo = [&](Cfvo& cfvo) {
                report.worksheetReferencesUpdated += invalidateFormula(cfvo.formula);
            };
            invalidateCfvo(rule.getDataBar().min);
            invalidateCfvo(rule.getDataBar().max);
            for (auto& stop : rule.getColorScale().stops) invalidateCfvo(stop);
            for (auto& stop : rule.getIconSet().thresholds) invalidateCfvo(stop);
        }
    }
    for (auto& validation : dataValidations_.items()) {
        auto formula1 = validation.formula1();
        const auto invalid1 = invalidateFormula(formula1);
        if (invalid1) validation.setFormula1(std::move(formula1));
        report.worksheetReferencesUpdated += invalid1;
        auto formula2 = validation.formula2();
        const auto invalid2 = invalidateFormula(formula2);
        if (invalid2) validation.setFormula2(std::move(formula2));
        report.worksheetReferencesUpdated += invalid2;
    }

    for (std::size_t chartIndex = 0; chartIndex < charts_.size(); ++chartIndex) {
        auto& chart = charts_[chartIndex];
        for (std::size_t seriesIndex = 0; seriesIndex < chart.series_.size(); ++seriesIndex) {
            auto& series = chart.series_[seriesIndex];
            auto cats = invalidateWorksheetReferences(series.categoriesReference(), removedName);
            auto vals = invalidateWorksheetReferences(series.valuesReference(), removedName);
            auto title = invalidateWorksheetReferences(series.titleReference(), removedName);
            report.referencesInvalidated += cats.referencesInvalidated + vals.referencesInvalidated + title.referencesInvalidated;
            if (cats.changed() || vals.changed()) {
                if (chart.imported_ && !chart.stableId_.empty() && !cats.value.empty() && !vals.value.empty())
                    setChartSeriesReferences(chart.stableId_, seriesIndex, cats.value, vals.value);
                else {
                    if (cats.changed()) series.setCategoriesReference(std::move(cats.value));
                    if (vals.changed()) series.setValuesReference(std::move(vals.value));
                }
                ++report.chartReferencesUpdated;
            }
            if (title.changed()) {
                series.setTitleReference(std::move(title.value));
                ++report.chartReferencesUpdated;
            }
            auto errorBars = series.errorBars();
            bool errorChanged = false;
            for (auto& error : errorBars) {
                auto plus = invalidateWorksheetReferences(error.plusReference, removedName);
                report.referencesInvalidated += plus.referencesInvalidated;
                if (plus.changed()) { error.plusReference = std::move(plus.value); errorChanged = true; }
                auto minus = invalidateWorksheetReferences(error.minusReference, removedName);
                report.referencesInvalidated += minus.referencesInvalidated;
                if (minus.changed()) { error.minusReference = std::move(minus.value); errorChanged = true; }
            }
            if (errorChanged) {
                series.setErrorBars(std::move(errorBars));
                ++report.chartReferencesUpdated;
            }
        }
    }

    for (auto& pivot : pivotTables_) {
        auto tr = invalidateWorksheetReferences(pivot.cache().sourceData(), removedName);
        if (tr.changed()) {
            pivot.cache().setSourceData(std::move(tr.value));
            // Retain cached records: Excel can still display the last pivot
            // result even though its worksheet source is now broken.
            ++report.pivotReferencesUpdated;
            report.referencesInvalidated += tr.referencesInvalidated;
        }
    }

    if (report.formulasUpdated || report.formulaMetadataUpdated ||
        report.worksheetReferencesUpdated || report.chartReferencesUpdated ||
        report.pivotReferencesUpdated || report.hyperlinksUpdated) dirty_ = true;
    if (report.pivotReferencesUpdated) pivotsDirty_ = true;
    return report;
}

} // namespace xlpp
