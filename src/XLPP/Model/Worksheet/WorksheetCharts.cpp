#include <XLPP/Worksheet/Worksheet.h>
#include "WorksheetChartValidation.h"
#include <algorithm>
#include <array>
#include <cmath>
#include <stdexcept>
#include <utility>

namespace xlpp {
using namespace internal::worksheet_chart_validation;
const Chart* Worksheet::chartByStableId(const std::string& stableId) const noexcept {
    const auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    return it == charts_.end() ? nullptr : &*it;
}

Worksheet::ImportedChartEdit& Worksheet::ensureImportedChartEdit(const Chart& chart) {
    const auto existing = std::find_if(importedChartEdits_.begin(), importedChartEdits_.end(), [&](const ImportedChartEdit& edit) {
        return edit.stableId == chart.stableId_;
    });
    if (existing != importedChartEdits_.end()) return *existing;
    ImportedChartEdit edit;
    edit.stableId = chart.stableId_;
    edit.sourceDrawingPart = chart.sourceDrawingPart_;
    edit.sourceChartPart = chart.sourceChartPart_;
    edit.sourceRelationshipId = chart.sourceRelationshipId_;
    edit.chartType = chart.type_;
    edit.primaryXAxisId = chart.primaryXAxisId_;
    edit.primaryYAxisId = chart.primaryYAxisId_;
    edit.originalAnchor = chart.anchorInfo_;
    edit.anchor = chart.anchorInfo_;
    importedChartEdits_.push_back(std::move(edit));
    return importedChartEdits_.back();
}

bool Worksheet::moveChart(const std::string& stableId, const std::string& anchor) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || it->anchorInfo_.type == DrawingAnchorType::Absolute) return false;
    const auto ref = CellReference::parse(anchor);
    const auto rowDelta = static_cast<long long>(ref.row) - static_cast<long long>(it->anchorInfo_.from.row);
    const auto columnDelta = static_cast<long long>(ref.column) - static_cast<long long>(it->anchorInfo_.from.column);
    auto updated = it->anchorInfo_;
    updated.from.row = ref.row;
    updated.from.column = ref.column;
    if (updated.type == DrawingAnchorType::TwoCell) {
        const auto newToRow = static_cast<long long>(updated.to.row) + rowDelta;
        const auto newToColumn = static_cast<long long>(updated.to.column) + columnDelta;
        if (newToRow < 1 || newToColumn < 1) return false;
        updated.to.row = static_cast<std::size_t>(newToRow);
        updated.to.column = static_cast<std::size_t>(newToColumn);
    }
    auto& edit = ensureImportedChartEdit(*it);
    edit.anchor = updated;
    edit.moved = true;
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::moveChartAbsolute(const std::string& stableId, long long xEmu, long long yEmu) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || it->anchorInfo_.type != DrawingAnchorType::Absolute || xEmu < 0 || yEmu < 0) return false;
    auto updated = it->anchorInfo_;
    updated.xEmu = xEmu;
    updated.yEmu = yEmu;
    auto& edit = ensureImportedChartEdit(*it);
    edit.anchor = updated;
    edit.moved = true;
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::resizeChart(const std::string& stableId, double widthPixels, double heightPixels) {
    if (!(widthPixels > 0.0) || !(heightPixels > 0.0)) return false;
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    auto updated = it->anchorInfo_;
    const auto oldWidth = updated.widthEmu;
    const auto oldHeight = updated.heightEmu;
    updated.widthEmu = static_cast<long long>(std::llround(widthPixels * 9525.0));
    updated.heightEmu = static_cast<long long>(std::llround(heightPixels * 9525.0));
    if (updated.type == DrawingAnchorType::TwoCell) {
        const auto resizeTerminalMarker = [](std::size_t fromIndex, long long fromOffset,
                                             std::size_t oldToIndex, long long oldToOffset,
                                             long long oldExtent, long long newExtent,
                                             std::size_t& newToIndex, long long& newToOffset) {
            if (oldExtent <= 0 || newExtent <= 0 || oldToIndex < fromIndex) return false;
            const auto span = oldToIndex - fromIndex;
            if (span == 0) {
                if (newExtent > oldExtent) return false;
                newToIndex = fromIndex;
                newToOffset = fromOffset + newExtent;
                return true;
            }
            const long double averageCellExtent =
                static_cast<long double>(oldExtent + fromOffset - oldToOffset) / static_cast<long double>(span);
            if (!(averageCellExtent > 0.0L)) return false;
            const long double terminal = static_cast<long double>(fromOffset + newExtent);
            auto cells = static_cast<long long>(std::floor(terminal / averageCellExtent));
            if (cells < 0) cells = 0;
            auto offset = static_cast<long long>(std::llround(terminal - static_cast<long double>(cells) * averageCellExtent));
            const auto roundedCell = static_cast<long long>(std::llround(averageCellExtent));
            if (roundedCell > 0 && offset >= roundedCell) { ++cells; offset = 0; }
            newToIndex = fromIndex + static_cast<std::size_t>(cells);
            newToOffset = std::max<long long>(0, offset);
            return true;
        };
        std::size_t toColumn = updated.to.column;
        std::size_t toRow = updated.to.row;
        long long toColumnOffset = updated.to.columnOffsetEmu;
        long long toRowOffset = updated.to.rowOffsetEmu;
        if (!resizeTerminalMarker(updated.from.column, updated.from.columnOffsetEmu,
                                  updated.to.column, updated.to.columnOffsetEmu, oldWidth, updated.widthEmu,
                                  toColumn, toColumnOffset) ||
            !resizeTerminalMarker(updated.from.row, updated.from.rowOffsetEmu,
                                  updated.to.row, updated.to.rowOffsetEmu, oldHeight, updated.heightEmu,
                                  toRow, toRowOffset)) return false;
        updated.to.column = toColumn;
        updated.to.columnOffsetEmu = toColumnOffset;
        updated.to.row = toRow;
        updated.to.rowOffsetEmu = toRowOffset;
    }
    auto& edit = ensureImportedChartEdit(*it);
    edit.anchor = updated;
    edit.resized = true;
    it->anchorInfo_ = updated;
    it->width_ = std::max(1, static_cast<int>(std::llround(widthPixels)));
    it->height_ = std::max(1, static_cast<int>(std::llround(heightPixels)));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartTitle(const std::string& stableId, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    edit.titleChanged = true;
    edit.title = title;
    it->title_ = std::move(title);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartStyle(const std::string& stableId, std::string style) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || style.empty()) return false;
    try {
        const auto numeric = std::stoul(style);
        if (numeric == 0 || numeric > 48) return false;
    } catch (...) { return false; }
    auto& edit = ensureImportedChartEdit(*it);
    edit.styleChanged = true;
    edit.style = style;
    it->style_ = std::move(style);
    dirty_ = true; drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartTitleRichText(const std::string& stableId, ChartRichText richText) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || !richText.present || richText.runs.empty()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    edit.titleRichTextChanged = true;
    edit.titleRichText = richText;
    it->titleRichText_ = std::move(richText);
    it->title_ = it->titleRichText_.plainText();
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartXAxisTitle(const std::string& stableId, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    if (it->primaryXAxisId_ != 0) return setChartAxisTitle(stableId, it->primaryXAxisId_, std::move(title));
    auto& edit = ensureImportedChartEdit(*it);
    edit.xAxisTitleChanged = true;
    edit.xAxisTitle = title;
    it->xAxisTitle_ = std::move(title);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartYAxisTitle(const std::string& stableId, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    if (it->primaryYAxisId_ != 0) return setChartAxisTitle(stableId, it->primaryYAxisId_, std::move(title));
    auto& edit = ensureImportedChartEdit(*it);
    edit.yAxisTitleChanged = true;
    edit.yAxisTitle = title;
    it->yAxisTitle_ = std::move(title);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartAxisTitle(const std::string& stableId, std::uint64_t axisId, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || axisId == 0) return false;
    auto axis = std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate) {
        return candidate.id == axisId;
    });
    if (axis == it->axes_.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.axisTitleEdits.begin(), edit.axisTitleEdits.end(), [&](const auto& candidate) {
        return candidate.axisId == axisId;
    });
    if (existing == edit.axisTitleEdits.end()) {
        ImportedChartEdit::AxisTitleEdit axisEdit;
        axisEdit.axisId = axisId;
        axisEdit.title = title;
        edit.axisTitleEdits.push_back(std::move(axisEdit));
    } else {
        existing->title = title;
    }
    axis->title = title;
    if (axisId == it->primaryXAxisId_) it->xAxisTitle_ = title;
    if (axisId == it->primaryYAxisId_) it->yAxisTitle_ = title;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartAxisTitleRichText(const std::string& stableId, std::uint64_t axisId, ChartRichText richText) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || axisId == 0 || !richText.present || richText.runs.empty()) return false;
    auto axis = std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate) { return candidate.id == axisId; });
    if (axis == it->axes_.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    auto existing = std::find_if(edit.axisRichTitleEdits.begin(), edit.axisRichTitleEdits.end(), [&](const auto& candidate) { return candidate.axisId == axisId; });
    if (existing == edit.axisRichTitleEdits.end()) edit.axisRichTitleEdits.push_back({axisId, richText}); else existing->richText = richText;
    axis->titleRichText = richText;
    axis->title = richText.plainText();
    if (axisId == it->primaryXAxisId_) it->xAxisTitle_ = axis->title;
    if (axisId == it->primaryYAxisId_) it->yAxisTitle_ = axis->title;
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartAxisNumberFormat(const std::string& stableId, std::uint64_t axisId, std::string formatCode, bool sourceLinked) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || axisId == 0 || formatCode.empty()) return false;
    auto axis = std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate) { return candidate.id == axisId; });
    if (axis == it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind = ImportedChartEdit::AxisFormatEdit::Kind::NumberFormat; e.axisId = axisId; e.value1 = formatCode; e.flag = sourceLinked;
    auto& edit = ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e));
    axis->numberFormat = std::move(formatCode); axis->numberFormatSourceLinked = sourceLinked;
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartAxisTicks(const std::string& stableId, std::uint64_t axisId, std::string majorTickMark, std::string minorTickMark, std::string tickLabelPosition) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || axisId == 0) return false;
    auto axis = std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate) { return candidate.id == axisId; });
    if (axis == it->axes_.end()) return false;
    const auto validTick=[](const std::string& v){ return v.empty() || v=="none" || v=="in" || v=="out" || v=="cross"; };
    const auto validPos=[](const std::string& v){ return v.empty() || v=="high" || v=="low" || v=="nextTo" || v=="none"; };
    if (!validTick(majorTickMark) || !validTick(minorTickMark) || !validPos(tickLabelPosition)) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Ticks; e.axisId=axisId; e.value1=majorTickMark; e.value2=minorTickMark; e.value3=tickLabelPosition;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e));
    axis->majorTickMark=std::move(majorTickMark); axis->minorTickMark=std::move(minorTickMark); axis->tickLabelPosition=std::move(tickLabelPosition);
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisUnits(const std::string& stableId, std::uint64_t axisId, double majorUnit, double minorUnit) {
    if (!(majorUnit > 0.0) || minorUnit < 0.0) return false;
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if (it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Units; e.axisId=axisId; e.number1=majorUnit; e.number2=minorUnit;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e));
    axis->hasMajorUnit=true; axis->majorUnit=majorUnit; axis->hasMinorUnit=minorUnit>0.0; axis->minorUnit=minorUnit;
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisScaling(const std::string& stableId, std::uint64_t axisId, ChartAxisScaling scaling) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; });
    if(axis==it->axes_.end()) return false;
    if (scaling.hasMinimum && scaling.hasMaximum && !(scaling.minimum < scaling.maximum)) return false;
    if (scaling.hasLogBase && (!(scaling.logBase >= 2.0 && scaling.logBase <= 1000.0) ||
        (scaling.hasMinimum && scaling.minimum <= 0.0) || (scaling.hasMaximum && scaling.maximum <= 0.0))) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Scaling; e.axisId=axisId; e.scaling=scaling;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->scaling=std::move(scaling);
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisCrossing(const std::string& stableId, std::uint64_t axisId, std::string crosses, std::string crossBetween) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    const auto validCross=[](const std::string& v){ return v.empty() || v=="autoZero" || v=="max" || v=="min"; };
    const auto validBetween=[](const std::string& v){ return v.empty() || v=="between" || v=="midCat"; };
    if(!validCross(crosses) || !validBetween(crossBetween)) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Crossing; e.axisId=axisId; e.value1=crosses; e.value2=crossBetween;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->crosses=std::move(crosses); axis->crossBetween=std::move(crossBetween); axis->hasCrossesAt=false; axis->crossesAt=0.0;
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisCrossesAt(const std::string& stableId, std::uint64_t axisId, double crossesAt) {
    if (!std::isfinite(crossesAt)) return false;
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::CrossesAt; e.axisId=axisId; e.number1=crossesAt;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->hasCrossesAt=true; axis->crossesAt=crossesAt; axis->crosses.clear();
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::clearChartAxisCrossesAt(const std::string& stableId, std::uint64_t axisId) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::ClearCrossesAt; e.axisId=axisId;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->hasCrossesAt=false; axis->crossesAt=0.0;
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisDisplayUnits(const std::string& stableId, std::uint64_t axisId, ChartDisplayUnits units) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0 || !units.present) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; });
    if(axis==it->axes_.end() || axis->kind!=Chart::AxisKind::Value) return false;
    static const std::array<const char*, 9> validUnits{"hundreds","thousands","tenThousands","hundredThousands","millions","tenMillions","hundredMillions","billions","trillions"};
    if (units.hasCustomUnit) {
        if (!(units.customUnit > 0.0) || !std::isfinite(units.customUnit) || !units.builtInUnit.empty()) return false;
    } else if (std::find_if(validUnits.begin(), validUnits.end(), [&](const char* value){ return units.builtInUnit==value; })==validUnits.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::DisplayUnits; e.axisId=axisId; e.displayUnits=units;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->displayUnits=std::move(units);
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::clearChartAxisDisplayUnits(const std::string& stableId, std::uint64_t axisId) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end() || axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::ClearDisplayUnits; e.axisId=axisId;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->displayUnits=ChartDisplayUnits{};
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisLineFormat(const std::string& stableId, std::uint64_t axisId, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||axisId==0||!format.present) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=ImportedChartEdit::AxisFormatEdit::Kind::Line; e.axisId=axisId; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); axis->lineFormat=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAxisGridlineFormat(const std::string& stableId, std::uint64_t axisId, bool major, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||axisId==0||!format.present) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=major?ImportedChartEdit::AxisFormatEdit::Kind::MajorGridline:ImportedChartEdit::AxisFormatEdit::Kind::MinorGridline; e.axisId=axisId; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e)); if(major){ axis->hasMajorGridlines=true; axis->majorGridlineFormat=std::move(format); } else { axis->hasMinorGridlines=true; axis->minorGridlineFormat=std::move(format); } dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartAxisGridlines(const std::string& stableId, std::uint64_t axisId, bool major) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||axisId==0) return false;
    auto axis=std::find_if(it->axes_.begin(), it->axes_.end(), [&](const Chart::Axis& candidate){ return candidate.id==axisId; }); if(axis==it->axes_.end()) return false;
    ImportedChartEdit::AxisFormatEdit e; e.kind=major?ImportedChartEdit::AxisFormatEdit::Kind::RemoveMajorGridline:ImportedChartEdit::AxisFormatEdit::Kind::RemoveMinorGridline; e.axisId=axisId;
    auto& edit=ensureImportedChartEdit(*it); edit.axisFormatEdits.push_back(std::move(e));
    if(major){ axis->hasMajorGridlines=false; axis->majorGridlineFormat=ChartLineFormat{}; } else { axis->hasMinorGridlines=false; axis->minorGridlineFormat=ChartLineFormat{}; }
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAreaLineFormat(const std::string& stableId, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false;
    ImportedChartEdit::AreaFormatEdit e; e.owner=ImportedChartEdit::AreaFormatEdit::Owner::ChartArea; e.kind=ImportedChartEdit::AreaFormatEdit::Kind::Line; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.areaFormatEdits.push_back(std::move(e)); it->chartAreaLineFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartAreaFillFormat(const std::string& stableId, ChartFillFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false;
    ImportedChartEdit::AreaFormatEdit e; e.owner=ImportedChartEdit::AreaFormatEdit::Owner::ChartArea; e.kind=ImportedChartEdit::AreaFormatEdit::Kind::Fill; e.fill=format; auto& edit=ensureImportedChartEdit(*it); edit.areaFormatEdits.push_back(std::move(e)); it->chartAreaFillFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotAreaLineFormat(const std::string& stableId, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false;
    ImportedChartEdit::AreaFormatEdit e; e.owner=ImportedChartEdit::AreaFormatEdit::Owner::PlotArea; e.kind=ImportedChartEdit::AreaFormatEdit::Kind::Line; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.areaFormatEdits.push_back(std::move(e)); it->plotAreaLineFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotAreaFillFormat(const std::string& stableId, ChartFillFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false;
    ImportedChartEdit::AreaFormatEdit e; e.owner=ImportedChartEdit::AreaFormatEdit::Owner::PlotArea; e.kind=ImportedChartEdit::AreaFormatEdit::Kind::Fill; e.fill=format; auto& edit=ensureImportedChartEdit(*it); edit.areaFormatEdits.push_back(std::move(e)); it->plotAreaFillFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotAreaLayout(const std::string& stableId, ChartManualLayout layout) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!layout.present) return false;
    ImportedChartEdit::LayoutEdit e; e.owner=ImportedChartEdit::LayoutEdit::Owner::PlotArea; e.layout=layout; auto& edit=ensureImportedChartEdit(*it); edit.layoutEdits.push_back(std::move(e)); it->plotAreaLayout_=std::move(layout); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartView3D(const std::string& stableId, ChartView3D view) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||!view.present) return false;
    if(view.hasRotationX && (view.rotationX < -90 || view.rotationX > 90)) return false;
    if(view.hasRotationY && (view.rotationY < 0 || view.rotationY > 360)) return false;
    if(view.hasHeightPercent && (view.heightPercent < 5 || view.heightPercent > 500)) return false;
    if(view.hasDepthPercent && (view.depthPercent < 20 || view.depthPercent > 2000)) return false;
    if(view.hasPerspective && (view.perspective < 0 || view.perspective > 240)) return false;
    ImportedChartEdit::View3DEdit e; e.view=view; auto& edit=ensureImportedChartEdit(*it); edit.view3DEdits.push_back(std::move(e));
    it->view3D_=std::move(view); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartFloorFormat(const std::string& stableId, ChartWallFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||!format.present) return false;
    if(format.hasThickness && format.thickness < 0) return false;
    ImportedChartEdit::WallFormatEdit e; e.owner=ImportedChartEdit::WallFormatEdit::Owner::Floor; e.format=format; auto& edit=ensureImportedChartEdit(*it); edit.wallFormatEdits.push_back(std::move(e));
    it->floorFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSideWallFormat(const std::string& stableId, ChartWallFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||!format.present) return false;
    if(format.hasThickness && format.thickness < 0) return false;
    ImportedChartEdit::WallFormatEdit e; e.owner=ImportedChartEdit::WallFormatEdit::Owner::SideWall; e.format=format; auto& edit=ensureImportedChartEdit(*it); edit.wallFormatEdits.push_back(std::move(e));
    it->sideWallFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartBackWallFormat(const std::string& stableId, ChartWallFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||!format.present) return false;
    if(format.hasThickness && format.thickness < 0) return false;
    ImportedChartEdit::WallFormatEdit e; e.owner=ImportedChartEdit::WallFormatEdit::Owner::BackWall; e.format=format; auto& edit=ensureImportedChartEdit(*it); edit.wallFormatEdits.push_back(std::move(e));
    it->backWallFormat_=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartDataTable(const std::string& stableId, ChartDataTable table) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()) return false;
    table.present=true; ImportedChartEdit::DataTableEdit e; e.table=table; auto& edit=ensureImportedChartEdit(*it); edit.dataTableEdits.push_back(std::move(e));
    it->dataTable_=std::move(table); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartDataTable(const std::string& stableId) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()) return false;
    ImportedChartEdit::DataTableEdit e; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.dataTableEdits.push_back(std::move(e));
    it->dataTable_=ChartDataTable{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotDropLines(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::DropLines; e.plotIndex=plotIndex; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasDropLines=true; it->plots_[plotIndex].dropLinesFormat=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartPlotDropLines(const std::string& stableId, std::size_t plotIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::DropLines; e.plotIndex=plotIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasDropLines=false; it->plots_[plotIndex].dropLinesFormat=ChartLineFormat{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotHighLowLines(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::HighLowLines; e.plotIndex=plotIndex; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasHighLowLines=true; it->plots_[plotIndex].highLowLinesFormat=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartPlotHighLowLines(const std::string& stableId, std::size_t plotIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::HighLowLines; e.plotIndex=plotIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasHighLowLines=false; it->plots_[plotIndex].highLowLinesFormat=ChartLineFormat{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotUpDownBars(const std::string& stableId, std::size_t plotIndex, ChartUpDownBars bars) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()||bars.gapWidth<0||bars.gapWidth>500) return false;
    bars.present=true; ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::UpDownBars; e.plotIndex=plotIndex; e.upDownBars=bars; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].upDownBars=std::move(bars); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartPlotUpDownBars(const std::string& stableId, std::size_t plotIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::PlotAuxiliaryEdit e; e.kind=ImportedChartEdit::PlotAuxiliaryEdit::Kind::UpDownBars; e.plotIndex=plotIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.plotAuxiliaryEdits.push_back(std::move(e));
    it->plots_[plotIndex].upDownBars=ChartUpDownBars{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotFirstSliceAngle(const std::string& stableId, std::size_t plotIndex, int degrees) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||plotIndex>=it->plots_.size()||degrees<0||degrees>360) return false;
    const auto type=it->plots_[plotIndex].type; if(type!=Chart::Type::Pie && type!=Chart::Type::Doughnut) return false;
    ImportedChartEdit::PlotTypeSpecificEdit e; e.kind=ImportedChartEdit::PlotTypeSpecificEdit::Kind::FirstSliceAngle; e.plotIndex=plotIndex; e.integerValue=degrees;
    auto& edit=ensureImportedChartEdit(*it); edit.plotTypeSpecificEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasFirstSliceAngle=true; it->plots_[plotIndex].firstSliceAngle=degrees; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotDoughnutHoleSize(const std::string& stableId, std::size_t plotIndex, int percent) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||plotIndex>=it->plots_.size()||percent<10||percent>90||it->plots_[plotIndex].type!=Chart::Type::Doughnut) return false;
    ImportedChartEdit::PlotTypeSpecificEdit e; e.kind=ImportedChartEdit::PlotTypeSpecificEdit::Kind::DoughnutHoleSize; e.plotIndex=plotIndex; e.integerValue=percent;
    auto& edit=ensureImportedChartEdit(*it); edit.plotTypeSpecificEdits.push_back(std::move(e));
    it->plots_[plotIndex].hasHoleSize=true; it->plots_[plotIndex].holeSize=percent; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotRadarStyle(const std::string& stableId, std::size_t plotIndex, std::string style) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||plotIndex>=it->plots_.size()||it->plots_[plotIndex].type!=Chart::Type::Radar) return false;
    if(style!="standard"&&style!="marker"&&style!="filled") return false;
    ImportedChartEdit::PlotTypeSpecificEdit e; e.kind=ImportedChartEdit::PlotTypeSpecificEdit::Kind::RadarStyle; e.plotIndex=plotIndex; e.textValue=style;
    auto& edit=ensureImportedChartEdit(*it); edit.plotTypeSpecificEdits.push_back(std::move(e));
    it->plots_[plotIndex].radarStyle=std::move(style); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotProjectedPieOptions(const std::string& stableId, std::size_t plotIndex, ChartProjectedPieOptions options) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; });
    if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    const auto type=it->plots_[plotIndex].type; if(type!=Chart::Type::PieOfPie && type!=Chart::Type::BarOfPie) return false;
    if(options.ofPieType!="pie"&&options.ofPieType!="bar") return false;
    if(options.gapWidth<0||options.gapWidth>500||options.secondPlotSize<5||options.secondPlotSize>200) return false;
    if(options.splitType!="auto"&&options.splitType!="cust"&&options.splitType!="percent"&&options.splitType!="pos"&&options.splitType!="val") return false;
    if(std::any_of(options.customSplitPoints.begin(), options.customSplitPoints.end(), [](int v){ return v<0; })) return false;
    options.present=true; options.ofPieType=(type==Chart::Type::BarOfPie?"bar":"pie");
    ImportedChartEdit::PlotTypeSpecificEdit e; e.kind=ImportedChartEdit::PlotTypeSpecificEdit::Kind::ProjectedPie; e.plotIndex=plotIndex; e.projectedPie=options;
    auto& edit=ensureImportedChartEdit(*it); edit.plotTypeSpecificEdits.push_back(std::move(e));
    it->plots_[plotIndex].projectedPie=std::move(options); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotLeaderLineFormat(const std::string& stableId, std::size_t plotIndex, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()||!format.present) return false;
    ImportedChartEdit::LeaderLineEdit e; e.plotLevel=true; e.ownerIndex=plotIndex; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.leaderLineEdits.push_back(std::move(e));
    auto& labels=it->plots_[plotIndex].dataLabels; labels.present=true; labels.showLeaderLines=true; labels.hasLeaderLines=true; labels.leaderLineFormat=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartPlotLeaderLines(const std::string& stableId, std::size_t plotIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||plotIndex>=it->plots_.size()) return false;
    ImportedChartEdit::LeaderLineEdit e; e.plotLevel=true; e.ownerIndex=plotIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.leaderLineEdits.push_back(std::move(e));
    auto& labels=it->plots_[plotIndex].dataLabels; labels.showLeaderLines=false; labels.hasLeaderLines=false; labels.leaderLineFormat=ChartLineFormat{}; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSeriesLeaderLineFormat(const std::string& stableId, std::size_t seriesIndex, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||seriesIndex>=it->series_.size()||!format.present) return false;
    ImportedChartEdit::LeaderLineEdit e; e.ownerIndex=seriesIndex; e.line=format; auto& edit=ensureImportedChartEdit(*it); edit.leaderLineEdits.push_back(std::move(e));
    auto labels=it->series_[seriesIndex].dataLabels(); labels.present=true; labels.showLeaderLines=true; labels.hasLeaderLines=true; labels.leaderLineFormat=std::move(format); it->series_[seriesIndex].setDataLabels(std::move(labels)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::removeChartSeriesLeaderLines(const std::string& stableId, std::size_t seriesIndex) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||seriesIndex>=it->series_.size()) return false;
    ImportedChartEdit::LeaderLineEdit e; e.ownerIndex=seriesIndex; e.remove=true; auto& edit=ensureImportedChartEdit(*it); edit.leaderLineEdits.push_back(std::move(e));
    auto labels=it->series_[seriesIndex].dataLabels(); labels.showLeaderLines=false; labels.hasLeaderLines=false; labels.leaderLineFormat=ChartLineFormat{}; it->series_[seriesIndex].setDataLabels(std::move(labels)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartLegend(const std::string& stableId, bool show, std::string position) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    if (show) {
        static const std::array<const char*, 5> validPositions{"l", "r", "t", "b", "tr"};
        if (std::find_if(validPositions.begin(), validPositions.end(), [&](const char* candidate) {
                return position == candidate;
            }) == validPositions.end()) return false;
    }
    auto& edit = ensureImportedChartEdit(*it);
    edit.legendChanged = true;
    edit.showLegend = show;
    edit.legendPosition = position;
    it->showLegend_ = show;
    if (show) it->legendPosition_ = std::move(position);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartLegendLayout(const std::string& stableId, ChartManualLayout layout) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!layout.present) return false;
    ImportedChartEdit::LayoutEdit e; e.owner=ImportedChartEdit::LayoutEdit::Owner::Legend; e.layout=layout; auto& edit=ensureImportedChartEdit(*it); edit.layoutEdits.push_back(std::move(e)); it->legendFormat_.present=true; it->legendFormat_.layout=std::move(layout); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartLegendOverlay(const std::string& stableId, bool overlay) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()) return false; auto& edit=ensureImportedChartEdit(*it); edit.legendOverlayChanged=true; edit.legendOverlay=overlay; it->legendFormat_.present=true; it->legendFormat_.overlay=overlay; dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartLegendLineFormat(const std::string& stableId, ChartLineFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false; auto& edit=ensureImportedChartEdit(*it); edit.legendLineFormatChanged=true; edit.legendLineFormat=format; it->legendFormat_.present=true; it->legendFormat_.line=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartLegendFillFormat(const std::string& stableId, ChartFillFormat format) {
    auto it=std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart){ return chart.imported_ && chart.stableId_==stableId; }); if(it==charts_.end()||!format.present) return false; auto& edit=ensureImportedChartEdit(*it); edit.legendFillFormatChanged=true; edit.legendFillFormat=format; it->legendFormat_.present=true; it->legendFormat_.fill=std::move(format); dirty_=true; drawingAppendDirty_=true; return true;
}

} // namespace xlpp
