#include <XLPP/Worksheet/Worksheet.h>
#include "XLPP/Formula/ReferenceTransformer.h"
#include "XLPP/Pivot/PivotStructuralEdit.h"
#include <algorithm>
#include <cstddef>
#include <map>
#include <optional>
#include <stdexcept>
#include <utility>
#include <vector>

namespace {
std::pair<xlpp::CellReference, xlpp::CellReference> parseRangeAddress(const std::string& address) {
    const auto colon = address.find(':');
    if (colon == std::string::npos) {
        const auto ref = xlpp::CellReference::parse(address);
        return {ref, ref};
    }
    if (address.find(':', colon + 1) != std::string::npos)
        throw std::invalid_argument("Invalid range address: " + address);
    auto first = xlpp::CellReference::parse(address.substr(0, colon));
    auto last = xlpp::CellReference::parse(address.substr(colon + 1));
    if (first.row > last.row) std::swap(first.row, last.row);
    if (first.column > last.column) std::swap(first.column, last.column);
    return {first, last};
}
}

namespace xlpp {
void Worksheet::shiftRows(std::size_t index, std::size_t amount, bool insert,
                          const StructuralEditOptions* options, StructuralEditReport* report) {
    shiftStructure(index, amount, insert, true, options, report);
}

void Worksheet::shiftColumns(std::size_t index, std::size_t amount, bool insert,
                             const StructuralEditOptions* options, StructuralEditReport* report) {
    shiftStructure(index, amount, insert, false, options, report);
}

void Worksheet::shiftStructure(std::size_t index, std::size_t amount, bool insert, bool rows,
                               const StructuralEditOptions* suppliedOptions, StructuralEditReport* report) {
    const StructuralEditOptions defaults;
    const auto& options = suppliedOptions ? *suppliedOptions : defaults;
    internal::StructuralEditSpec edit;
    edit.axis = rows ? internal::StructuralAxis::Row : internal::StructuralAxis::Column;
    edit.action = insert ? internal::StructuralAction::Insert : internal::StructuralAction::Delete;
    edit.index = index;
    edit.amount = amount;
    edit.targetSheetName = name_;

    constexpr std::size_t kMaxExcelRow = 1048576;
    constexpr std::size_t kMaxExcelColumn = 16384;
    const std::size_t coordinateLimit = rows ? kMaxExcelRow : kMaxExcelColumn;
    if (index == 0 || index > coordinateLimit || amount == 0 || amount > coordinateLimit)
        throw std::out_of_range(rows ? "Row structural edit exceeds Excel worksheet bounds"
                                     : "Column structural edit exceeds Excel worksheet bounds");
    if (insert) {
        const auto highestShiftable = coordinateLimit - amount;
        for (const auto& [_, cellValue] : cells_) {
            const auto coordinate = rows ? cellValue.row() : cellValue.column();
            if (coordinate >= index && coordinate > highestShiftable)
                throw std::out_of_range("Structural insertion would move a populated cell outside the Excel worksheet grid");
        }
        if (rows) {
            for (const auto& [coordinate, _] : rowDimensions_)
                if (coordinate >= index && coordinate > highestShiftable)
                    throw std::out_of_range("Structural insertion would move a row dimension outside the Excel worksheet grid");
        } else {
            for (const auto& [coordinate, _] : columnDimensions_)
                if (coordinate >= index && coordinate > highestShiftable)
                    throw std::out_of_range("Structural insertion would move a column dimension outside the Excel worksheet grid");
        }
    }

    if (report) {
        report->kind = rows ? (insert ? StructuralEditKind::InsertRows : StructuralEditKind::DeleteRows)
                            : (insert ? StructuralEditKind::InsertColumns : StructuralEditKind::DeleteColumns);
        report->worksheetName = name_;
        report->index = index;
        report->amount = amount;
    }

    auto mapCoordinate = [&](std::size_t coordinate) -> std::optional<std::size_t> {
        if (insert) return coordinate >= index ? coordinate + amount : coordinate;
        const auto deleteEnd = index + amount;
        if (coordinate >= index && coordinate < deleteEnd) return std::nullopt;
        return coordinate >= deleteEnd ? coordinate - amount : coordinate;
    };

    // 1. Move/delete physical cells first. Formula rewriting happens on the
    // surviving cells afterwards, so deleted formulas cannot create false
    // mutation records.
    std::map<std::uint64_t, Cell> shifted;
    for (auto& [key, source] : cells_) {
        auto row = source.row();
        auto column = source.column();
        const auto mapped = mapCoordinate(rows ? row : column);
        if (!mapped) {
            if (report) ++report->cellsDeleted;
            continue;
        }
        if (rows) row = *mapped; else column = *mapped;
        if (row != source.row() || column != source.column()) {
            source.setPosition(row, column);
            if (report) ++report->cellsMoved;
        }
        shifted.emplace(makeCellKey(row, column), std::move(source));
    }
    cells_ = std::move(shifted);
    extentsCacheValid_ = false;
    trackedCellKeys_.clear();
    for (const auto& [key, _] : cells_) trackedCellKeys_.insert(key);

    auto rewriteFormula = [&](std::string value, std::string_view ownerSheet) {
        const auto rewritten = internal::rewriteFormulaReferences(value, ownerSheet, edit);
        if (report) {
            report->formulasRewritten += rewritten.referencesRewritten;
            report->formulaReferencesInvalidated += rewritten.referencesInvalidated;
            report->referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
        }
        return rewritten.text;
    };
    auto rewriteLocalReference = [&](const std::string& value) {
        return internal::rewriteReferenceList(value, edit);
    };

    if (options.updateCellFormulas || options.updateFormulaMetadata || options.updateHyperlinks) {
        for (auto& [_, cellValue] : cells_) {
            if (options.updateCellFormulas && cellValue.hasFormula()) {
                const auto rewritten = internal::rewriteFormulaReferences(cellValue.formula(), name_, edit);
                if (report) report->referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                if (rewritten.changed()) {
                    cellValue.setFormulaTextPreservingMetadata(rewritten.text);
                    if (report) {
                        report->formulasRewritten += rewritten.referencesRewritten;
                        report->formulaReferencesInvalidated += rewritten.referencesInvalidated;
                    }
                }
            }
            const auto& formulaMetadata = static_cast<const Cell&>(cellValue).formulaMetadata();
            if (options.updateFormulaMetadata && !formulaMetadata.reference().empty()) {
                const auto rewritten = rewriteLocalReference(formulaMetadata.reference());
                if (rewritten.changed()) {
                    cellValue.formulaMetadata().setReference(rewritten.text);
                    if (report) {
                        report->formulaMetadataRewritten += rewritten.referencesRewritten;
                        report->formulaReferencesInvalidated += rewritten.referencesInvalidated;
                        report->referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                    }
                }
            }
            if (options.updateHyperlinks && cellValue.hyperlinkValue() && !cellValue.hyperlinkValue()->external()) {
                const auto& currentTarget = cellValue.hyperlinkValue()->target();
                const bool leadingHash = !currentTarget.empty() && currentTarget.front() == '#';
                const auto body = leadingHash ? std::string_view(currentTarget).substr(1) : std::string_view(currentTarget);
                const auto rewritten = internal::rewriteFormulaReferences(body, name_, edit);
                if (report) report->referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                if (rewritten.changed()) {
                    cellValue.hyperlink().setTarget(leadingHash && !rewritten.text.empty() && rewritten.text.front() != '#' ? "#" + rewritten.text : rewritten.text);
                    if (report) {
                        report->hyperlinksRewritten += rewritten.referencesRewritten;
                        report->formulaReferencesInvalidated += rewritten.referencesInvalidated;
                    }
                }
            }
        }
    }

    if (options.updateMergedCells) {
        std::vector<std::string> ranges;
        std::vector<MergedRangeCache> parsedRanges;
        ranges.reserve(mergedRanges_.size());
        parsedRanges.reserve(mergedRanges_.size());
        for (const auto& range : mergedRanges_) {
            const auto rewritten = rewriteLocalReference(range);
            if (rewritten.text.empty()) {
                if (report) ++report->worksheetRangesRemoved;
                continue;
            }
            try {
                const auto parsed = parseRangeAddress(rewritten.text);
                if (parsed.first.row == parsed.second.row && parsed.first.column == parsed.second.column) {
                    // A merge cannot collapse to one cell after deletion.
                    if (report) ++report->worksheetRangesRemoved;
                    continue;
                }
                ranges.push_back(rewritten.text);
                parsedRanges.push_back({parsed.first.row, parsed.first.column, parsed.second.row, parsed.second.column});
                if (rewritten.changed() && report) ++report->worksheetRangesRewritten;
            } catch (...) {
                ranges.push_back(range);
                const auto parsed = parseRangeAddress(range);
                parsedRanges.push_back({parsed.first.row, parsed.first.column, parsed.second.row, parsed.second.column});
                if (report) report->diagnostics.push_back("Could not rewrite merged range: " + range);
            }
        }
        mergedRanges_ = std::move(ranges);
        mergedRangesParsed_ = std::move(parsedRanges);
    }

    if (options.updateFreezePanes && freezePane_) {
        try {
            auto ref = CellReference::parse(*freezePane_);
            if (rows) ref.row = internal::transformPosition(ref.row, edit);
            else ref.column = internal::transformPosition(ref.column, edit);
            const auto updated = ref.address();
            if (updated != *freezePane_) {
                *freezePane_ = updated;
                if (report) ++report->worksheetRangesRewritten;
            }
        } catch (...) {
            if (report) report->diagnostics.push_back("Could not rewrite freeze pane: " + *freezePane_);
        }
    }

    if (options.updateFreezePanes && !sheetView_.topLeftCell().empty()) {
        try {
            auto ref = CellReference::parse(sheetView_.topLeftCell());
            if (rows) ref.row = internal::transformPosition(ref.row, edit);
            else ref.column = internal::transformPosition(ref.column, edit);
            if (const auto updated = ref.address(); updated != sheetView_.topLeftCell()) {
                sheetView_.setTopLeftCell(updated);
                if (report) ++report->worksheetRangesRewritten;
            }
        } catch (...) {
            if (report) report->diagnostics.push_back("Could not rewrite sheet-view topLeftCell: " + sheetView_.topLeftCell());
        }
    }

    if (options.updateDimensions) {
        if (rows) {
            std::map<std::size_t, RowDimension> updated;
            for (auto& [coordinate, dimension] : rowDimensions_) {
                const auto mapped = mapCoordinate(coordinate);
                if (mapped) updated.emplace(*mapped, std::move(dimension));
            }
            rowDimensions_ = std::move(updated);
        } else {
            std::map<std::size_t, ColumnDimension> updated;
            for (auto& [coordinate, dimension] : columnDimensions_) {
                const auto mapped = mapCoordinate(coordinate);
                if (mapped) updated.emplace(*mapped, std::move(dimension));
            }
            columnDimensions_ = std::move(updated);
        }
    }

    if (options.updateAutoFilter && autoFilter_.enabled()) {
        const auto oldReference = autoFilter_.reference();
        const auto oldColumns = autoFilter_.columns();
        const auto oldSort = autoFilter_.sortStateValue();
        const auto rewritten = rewriteLocalReference(oldReference);
        autoFilter_.clear();
        if (rewritten.text.empty()) {
            if (report) ++report->worksheetRangesRemoved;
        } else {
            autoFilter_.setReference(rewritten.text);
            if (rewritten.changed() && report) ++report->worksheetRangesRewritten;
            std::size_t oldStartColumn = 1, newStartColumn = 1;
            try {
                oldStartColumn = parseRangeAddress(oldReference).first.column;
                newStartColumn = parseRangeAddress(rewritten.text).first.column;
            } catch (...) {}
            for (const auto& [columnId, filter] : oldColumns) {
                std::optional<std::size_t> newId = columnId;
                if (!rows) {
                    const auto mapped = mapCoordinate(oldStartColumn + columnId);
                    if (!mapped) continue;
                    if (*mapped < newStartColumn) continue;
                    newId = *mapped - newStartColumn;
                }
                auto& target = autoFilter_.column(*newId);
                target = filter;
                target.setColumnId(*newId);
            }
            if (oldSort) {
                auto& sort = autoFilter_.sortState();
                sort.setCaseSensitive(oldSort->caseSensitive());
                if (!oldSort->reference().empty()) {
                    const auto sr = rewriteLocalReference(oldSort->reference());
                    sort.setReference(sr.text);
                    if (sr.changed() && report) ++report->worksheetRangesRewritten;
                }
                for (const auto& condition : oldSort->conditions()) {
                    const auto cr = rewriteLocalReference(condition.reference);
                    if (!cr.text.empty()) sort.addCondition(cr.text, condition.descending);
                    if (cr.changed() && report) ++report->worksheetRangesRewritten;
                }
            }
        }
    }

    if (options.updateConditionalFormatting) {
        auto& entries = conditionalFormatting_.entries();
        for (auto it = entries.begin(); it != entries.end();) {
            const auto reference = rewriteLocalReference(it->reference());
            if (reference.text.empty()) {
                it = entries.erase(it);
                if (report) ++report->worksheetRangesRemoved;
                continue;
            }
            if (reference.changed()) {
                it->setReference(reference.text);
                if (report) ++report->worksheetRangesRewritten;
            }
            for (auto& rule : it->rules()) {
                auto formulas = rule.formulas();
                bool formulasChanged = false;
                for (auto& formula : formulas) {
                    const auto rewritten = internal::rewriteFormulaReferences(formula, name_, edit);
                    if (report) report->referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                    if (rewritten.changed()) {
                        formula = rewritten.text;
                        formulasChanged = true;
                        if (report) {
                            report->formulasRewritten += rewritten.referencesRewritten;
                            report->formulaReferencesInvalidated += rewritten.referencesInvalidated;
                        }
                    }
                }
                if (formulasChanged) rule.setFormulas(std::move(formulas));
                auto rewriteCfvo = [&](Cfvo& value) {
                    if (value.type != "formula" || value.formula.empty()) return;
                    const auto rewritten = internal::rewriteFormulaReferences(value.formula, name_, edit);
                    if (report) report->referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                    if (rewritten.changed()) {
                        value.formula = rewritten.text;
                        if (report) {
                            report->formulasRewritten += rewritten.referencesRewritten;
                            report->formulaReferencesInvalidated += rewritten.referencesInvalidated;
                        }
                    }
                };
                rewriteCfvo(rule.getDataBar().min);
                rewriteCfvo(rule.getDataBar().max);
                for (auto& stop : rule.getColorScale().stops) rewriteCfvo(stop);
                for (auto& stop : rule.getIconSet().thresholds) rewriteCfvo(stop);
            }
            ++it;
        }
    }

    if (options.updateDataValidations) {
        auto& validations = dataValidations_.items();
        for (auto it = validations.begin(); it != validations.end();) {
            const auto reference = rewriteLocalReference(it->reference());
            if (reference.text.empty()) {
                it = validations.erase(it);
                if (report) ++report->worksheetRangesRemoved;
                continue;
            }
            if (reference.changed()) {
                it->setReference(reference.text);
                if (report) ++report->worksheetRangesRewritten;
            }
            if (!it->formula1().empty()) {
                const auto rewritten = internal::rewriteFormulaReferences(it->formula1(), name_, edit);
                if (report) report->referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                if (rewritten.changed()) {
                    it->setFormula1(rewritten.text);
                    if (report) {
                        report->formulasRewritten += rewritten.referencesRewritten;
                        report->formulaReferencesInvalidated += rewritten.referencesInvalidated;
                    }
                }
            }
            if (!it->formula2().empty()) {
                const auto rewritten = internal::rewriteFormulaReferences(it->formula2(), name_, edit);
                if (report) report->referencesSkippedUnsupported += rewritten.referencesSkippedUnsupported;
                if (rewritten.changed()) {
                    it->setFormula2(rewritten.text);
                    if (report) {
                        report->formulasRewritten += rewritten.referencesRewritten;
                        report->formulaReferencesInvalidated += rewritten.referencesInvalidated;
                    }
                }
            }
            ++it;
        }
    }

    if (options.updateTables) {
        for (auto it = tables_.begin(); it != tables_.end();) {
            const auto oldReference = it->reference();
            const auto rewritten = rewriteLocalReference(oldReference);
            if (rewritten.text.empty()) {
                it = tables_.erase(it);
                if (report) ++report->tablesRemoved;
                continue;
            }
            if (!rows) {
                try {
                    const auto oldBounds = parseRangeAddress(oldReference);
                    auto names = std::vector<std::string>{};
                    names.reserve(it->columns().size());
                    for (const auto& column : it->columns()) names.push_back(column.name());
                    if (insert && index > oldBounds.first.column && index <= oldBounds.second.column) {
                        const auto insertionOffset = std::min<std::size_t>(index - oldBounds.first.column, names.size());
                        for (std::size_t n = 0; n < amount; ++n)
                            names.insert(names.begin() + static_cast<std::ptrdiff_t>(insertionOffset + n),
                                         "Column" + std::to_string(insertionOffset + n + 1));
                    } else if (!insert) {
                        const auto deleteEnd = index + amount;
                        std::vector<std::string> kept;
                        for (std::size_t n = 0; n < names.size(); ++n) {
                            const auto absoluteColumn = oldBounds.first.column + n;
                            if (absoluteColumn >= index && absoluteColumn < deleteEnd) continue;
                            kept.push_back(names[n]);
                        }
                        names = std::move(kept);
                    }
                    if (!names.empty() && names.size() != it->columns().size()) {
                        it->columns().clear();
                        for (auto& name : names) it->addColumn(std::move(name));
                    }
                } catch (...) {
                    if (report) report->diagnostics.push_back("Could not update table columns for " + it->name());
                }
            }
            if (rewritten.changed()) {
                it->setReference(rewritten.text);
                if (report) ++report->tablesRewritten;
            }
            ++it;
        }
    }

    if (options.updatePrintSettings) {
        auto rewritePrint = [&](std::string& value) {
            if (value.empty()) return;
            const auto rewritten = rewriteLocalReference(value);
            if (rewritten.changed()) {
                value = rewritten.text;
                if (report) ++report->worksheetRangesRewritten;
            }
        };
        rewritePrint(printArea_);
        rewritePrint(printTitlesRows_);
        rewritePrint(printTitlesCols_);
    }

    if (options.updateDrawings) {
        for (auto& image : images_) {
            if (image.anchorInfo_.type != DrawingAnchorType::Absolute) {
                auto updated = image.anchorInfo_;
                internal::transformDrawingAnchor(updated, edit);
                if (updated.from.row != image.anchorInfo_.from.row || updated.from.column != image.anchorInfo_.from.column ||
                    updated.to.row != image.anchorInfo_.to.row || updated.to.column != image.anchorInfo_.to.column) {
                    if (image.imported_) {
                        auto& imageEdit = ensureImportedImageEdit(image);
                        imageEdit.anchor = updated;
                        imageEdit.moved = true;
                    }
                    image.anchorInfo_ = updated;
                    image.anchor_ = CellReference{updated.from.row, updated.from.column}.address();
                    if (report) ++report->drawingAnchorsRewritten;
                    drawingAppendDirty_ = true;
                }
            } else {
                try {
                    auto ref = CellReference::parse(image.anchor_);
                    if (rows) ref.row = internal::transformPosition(ref.row, edit);
                    else ref.column = internal::transformPosition(ref.column, edit);
                    image.anchor_ = ref.address();
                } catch (...) {}
            }
        }
        for (auto& chart : charts_) {
            if (chart.anchorInfo_.type == DrawingAnchorType::Absolute) continue;
            auto updated = chart.anchorInfo_;
            internal::transformDrawingAnchor(updated, edit);
            if (updated.from.row == chart.anchorInfo_.from.row && updated.from.column == chart.anchorInfo_.from.column &&
                updated.to.row == chart.anchorInfo_.to.row && updated.to.column == chart.anchorInfo_.to.column) continue;
            if (chart.imported_) {
                auto& chartEdit = ensureImportedChartEdit(chart);
                chartEdit.anchor = updated;
                chartEdit.moved = true;
                drawingAppendDirty_ = true;
            } else {
                // Generated chart drawing serialization uses the model anchor.
                drawingsDirty_ = true;
            }
            chart.anchorInfo_ = updated;
            if (report) ++report->drawingAnchorsRewritten;
        }
    }

    // Local chart references are handled here. Workbook-aware edits perform a
    // second pass over other worksheets for qualified references to this sheet.
    if (options.updateCharts) {
        for (std::size_t chartIndex = 0; chartIndex < charts_.size(); ++chartIndex) {
            auto& chart = charts_[chartIndex];
            for (std::size_t seriesIndex = 0; seriesIndex < chart.series_.size(); ++seriesIndex) {
                auto& series = chart.series_[seriesIndex];
                const auto categories = internal::rewriteFormulaReferences(series.categoriesReference(), name_, edit);
                const auto values = internal::rewriteFormulaReferences(series.valuesReference(), name_, edit);
                const auto title = internal::rewriteFormulaReferences(series.titleReference(), name_, edit);
                if (report)
                    report->referencesSkippedUnsupported += categories.referencesSkippedUnsupported +
                                                            values.referencesSkippedUnsupported +
                                                            title.referencesSkippedUnsupported;
                bool changed = categories.changed() || values.changed() || title.changed();
                if (changed) {
                    if (chart.imported_) {
                        setChartSeriesReferences(chart.stableId_, seriesIndex, categories.text, values.text);
                        // Title-reference selective patch is not yet exposed.
                        // Keep the in-memory model correct and force full chart
                        // regeneration rather than silently retaining stale XML.
                        if (title.changed()) {
                            series.setTitleReference(title.text);
                            drawingsDirty_ = true;
                            drawingAppendDirty_ = false;
                            if (report)
                                report->diagnostics.push_back("Imported chart title reference required full chart regeneration: " + chart.stableId_);
                        }
                    } else {
                        series.setCategoriesReference(categories.text);
                        series.setValuesReference(values.text);
                        series.setTitleReference(title.text);
                        drawingsDirty_ = true;
                    }
                    if (report) report->chartReferencesRewritten += categories.referencesRewritten + values.referencesRewritten + title.referencesRewritten;
                }
                auto errorBars = series.errorBars();
                bool errorChanged = false;
                for (auto& bars : errorBars) {
                    const auto plus = internal::rewriteFormulaReferences(bars.plusReference, name_, edit);
                    const auto minus = internal::rewriteFormulaReferences(bars.minusReference, name_, edit);
                    if (plus.changed()) { bars.plusReference = plus.text; errorChanged = true; }
                    if (minus.changed()) { bars.minusReference = minus.text; errorChanged = true; }
                    if (report) {
                        report->chartReferencesRewritten += plus.referencesRewritten + minus.referencesRewritten;
                        report->referencesSkippedUnsupported += plus.referencesSkippedUnsupported + minus.referencesSkippedUnsupported;
                    }
                }
                if (errorChanged) {
                    if (chart.imported_) {
                        for (const auto& bars : errorBars) setChartSeriesErrorBars(chart.stableId_, seriesIndex, bars);
                    } else {
                        series.setErrorBars(std::move(errorBars));
                        drawingsDirty_ = true;
                    }
                }
            }
        }
    }

    if (options.updatePivotSources) {
        for (auto& pivot : pivotTables_) {
            const auto location = rewriteLocalReference(pivot.location());
            if (location.changed()) {
                if (location.text.empty()) {
                    if (report) report->diagnostics.push_back("PivotTable location was fully deleted: " + pivot.name());
                } else {
                    pivot.setLocation(location.text);
                    if (report) ++report->pivotSourcesRewritten;
                    pivotsDirty_ = true;
                }
            }
            if (!pivot.cache().sourceData().empty()) {
                StructuralEditReport scratch;
                auto& activeReport = report ? *report : scratch;
                if (internal::rewritePivotSourceForStructuralEdit(pivot, name_, edit, activeReport))
                    pivotsDirty_ = true;
            }
        }
    }

    dirty_ = true;
}



} // namespace xlpp
