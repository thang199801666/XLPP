#include <XLPP/Worksheet/Worksheet.h>
#include "WorksheetChartValidation.h"
#include <algorithm>
#include <cmath>
#include <stdexcept>
#include <utility>

namespace xlpp {
using namespace internal::worksheet_chart_validation;
bool Worksheet::setChartSeriesTitle(const std::string& stableId, std::size_t seriesIndex, std::string title) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.seriesTitleEdits.begin(), edit.seriesTitleEdits.end(), [&](const auto& candidate) {
        return candidate.seriesIndex == seriesIndex;
    });
    if (existing == edit.seriesTitleEdits.end()) {
        ImportedChartEdit::SeriesTitleEdit seriesEdit;
        seriesEdit.seriesIndex = seriesIndex;
        seriesEdit.title = title;
        edit.seriesTitleEdits.push_back(std::move(seriesEdit));
    } else {
        existing->title = title;
    }
    it->series_[seriesIndex].setTitle(std::move(title));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesReferences(const std::string& stableId, std::size_t seriesIndex,
                                          std::string categoriesReference, std::string valuesReference) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        categoriesReference.empty() || valuesReference.empty()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.seriesReferenceEdits.begin(), edit.seriesReferenceEdits.end(), [&](const auto& candidate) {
        return candidate.seriesIndex == seriesIndex;
    });
    if (existing == edit.seriesReferenceEdits.end()) {
        ImportedChartEdit::SeriesReferenceEdit seriesEdit;
        seriesEdit.seriesIndex = seriesIndex;
        seriesEdit.categoriesReference = categoriesReference;
        seriesEdit.valuesReference = valuesReference;
        edit.seriesReferenceEdits.push_back(std::move(seriesEdit));
    } else {
        existing->categoriesReference = categoriesReference;
        existing->valuesReference = valuesReference;
    }
    it->series_[seriesIndex].setCategoriesReference(std::move(categoriesReference));
    it->series_[seriesIndex].setValuesReference(std::move(valuesReference));
    it->series_[seriesIndex].setCategoriesCache(ChartSeriesCache{});
    it->series_[seriesIndex].setValuesCache(ChartSeriesCache{});
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesCategoryCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache) {
    auto it=std::find_if(charts_.begin(),charts_.end(),[&](const Chart& chart){return chart.imported_&&chart.stableId_==stableId;});
    if(it==charts_.end()||seriesIndex>=it->series_.size()||!cache.present||!validChartSeriesCache(cache)) return false;
    auto& edit=ensureImportedChartEdit(*it); ImportedChartEdit::SeriesCacheEdit e; e.kind=ImportedChartEdit::SeriesCacheEdit::Kind::Categories; e.seriesIndex=seriesIndex; e.cache=cache; edit.seriesCacheEdits.push_back(std::move(e));
    it->series_[seriesIndex].setCategoriesCache(std::move(cache)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSeriesValueCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache) {
    auto it=std::find_if(charts_.begin(),charts_.end(),[&](const Chart& chart){return chart.imported_&&chart.stableId_==stableId;});
    if(it==charts_.end()||seriesIndex>=it->series_.size()||!cache.present||!cache.numeric||!validChartSeriesCache(cache)) return false;
    auto& edit=ensureImportedChartEdit(*it); ImportedChartEdit::SeriesCacheEdit e; e.kind=ImportedChartEdit::SeriesCacheEdit::Kind::Values; e.seriesIndex=seriesIndex; e.cache=cache; edit.seriesCacheEdits.push_back(std::move(e));
    it->series_[seriesIndex].setValuesCache(std::move(cache)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartSeriesTitleCache(const std::string& stableId, std::size_t seriesIndex, ChartSeriesCache cache) {
    auto it=std::find_if(charts_.begin(),charts_.end(),[&](const Chart& chart){return chart.imported_&&chart.stableId_==stableId;});
    if(it==charts_.end()||seriesIndex>=it->series_.size()||!cache.present||cache.numeric||!validChartSeriesCache(cache)) return false;
    if(it->series_[seriesIndex].titleReference().empty()) return false;
    auto& edit=ensureImportedChartEdit(*it); ImportedChartEdit::SeriesCacheEdit e; e.kind=ImportedChartEdit::SeriesCacheEdit::Kind::Title; e.seriesIndex=seriesIndex; e.cache=cache; edit.seriesCacheEdits.push_back(std::move(e));
    it->series_[seriesIndex].setTitleCache(std::move(cache)); dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::clearChartSeriesCaches(const std::string& stableId, std::size_t seriesIndex) {
    auto it=std::find_if(charts_.begin(),charts_.end(),[&](const Chart& chart){return chart.imported_&&chart.stableId_==stableId;});
    if(it==charts_.end()||seriesIndex>=it->series_.size()) return false;
    auto& edit=ensureImportedChartEdit(*it); ImportedChartEdit::SeriesCacheEdit e; e.kind=ImportedChartEdit::SeriesCacheEdit::Kind::ClearAll; e.seriesIndex=seriesIndex; edit.seriesCacheEdits.push_back(std::move(e));
    it->series_[seriesIndex].setTitleCache(ChartSeriesCache{}); it->series_[seriesIndex].setCategoriesCache(ChartSeriesCache{}); it->series_[seriesIndex].setValuesCache(ChartSeriesCache{});
    dirty_=true; drawingAppendDirty_=true; return true;
}

bool Worksheet::setChartPlotDataLabels(const std::string& stableId, std::size_t plotIndex, Chart::DataLabels labels) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || plotIndex >= it->plots_.size()) return false;
    if (!validDataLabelPosition(labels.position)) return false;
    labels.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.plotDataLabelsEdits.begin(), edit.plotDataLabelsEdits.end(), [&](const auto& candidate) {
        return candidate.plotIndex == plotIndex;
    });
    if (existing == edit.plotDataLabelsEdits.end()) {
        ImportedChartEdit::PlotDataLabelsEdit labelsEdit;
        labelsEdit.plotIndex = plotIndex;
        labelsEdit.labels = labels;
        edit.plotDataLabelsEdits.push_back(std::move(labelsEdit));
    } else {
        existing->labels = labels;
    }
    it->plots_[plotIndex].dataLabels = std::move(labels);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesDataLabels(const std::string& stableId, std::size_t seriesIndex, Chart::DataLabels labels) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    if (!validDataLabelPosition(labels.position)) return false;
    labels.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    const auto existing = std::find_if(edit.seriesDataLabelsEdits.begin(), edit.seriesDataLabelsEdits.end(), [&](const auto& candidate) {
        return candidate.seriesIndex == seriesIndex;
    });
    if (existing == edit.seriesDataLabelsEdits.end()) {
        ImportedChartEdit::SeriesDataLabelsEdit labelsEdit;
        labelsEdit.seriesIndex = seriesIndex;
        labelsEdit.labels = labels;
        edit.seriesDataLabelsEdits.push_back(std::move(labelsEdit));
    } else {
        existing->labels = labels;
    }
    it->series_[seriesIndex].dataLabels_ = std::move(labels);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartPlotDataLabelPoint(const std::string& stableId, std::size_t plotIndex, ChartDataLabelPoint label) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || plotIndex >= it->plots_.size() || !validDataLabelPosition(label.position)) return false;
    auto& labels = it->plots_[plotIndex].dataLabels;
    labels.present = true;
    const auto existingPoint = std::find_if(labels.points.begin(), labels.points.end(), [&](const auto& point) { return point.index == label.index; });
    if (existingPoint == labels.points.end()) labels.points.push_back(label); else *existingPoint = label;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelEdit pointEdit;
    pointEdit.plotLevel = true;
    pointEdit.ownerIndex = plotIndex;
    pointEdit.label = std::move(label);
    edit.pointDataLabelEdits.push_back(std::move(pointEdit));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartPlotDataLabelPoint(const std::string& stableId, std::size_t plotIndex, std::size_t pointIndex) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || plotIndex >= it->plots_.size()) return false;
    auto& points = it->plots_[plotIndex].dataLabels.points;
    const auto existing = std::find_if(points.begin(), points.end(), [&](const auto& point) { return point.index == pointIndex; });
    if (existing == points.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelEdit pointEdit;
    pointEdit.plotLevel = true;
    pointEdit.remove = true;
    pointEdit.ownerIndex = plotIndex;
    pointEdit.label.index = pointIndex;
    edit.pointDataLabelEdits.push_back(std::move(pointEdit));
    points.erase(existing);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesDataLabelPoint(const std::string& stableId, std::size_t seriesIndex, ChartDataLabelPoint label) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validDataLabelPosition(label.position)) return false;
    auto& labels = it->series_[seriesIndex].dataLabels_;
    labels.present = true;
    const auto existingPoint = std::find_if(labels.points.begin(), labels.points.end(), [&](const auto& point) { return point.index == label.index; });
    if (existingPoint == labels.points.end()) labels.points.push_back(label); else *existingPoint = label;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelEdit pointEdit;
    pointEdit.ownerIndex = seriesIndex;
    pointEdit.label = std::move(label);
    edit.pointDataLabelEdits.push_back(std::move(pointEdit));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartSeriesDataLabelPoint(const std::string& stableId, std::size_t seriesIndex, std::size_t pointIndex) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    auto& points = it->series_[seriesIndex].dataLabels_.points;
    const auto existing = std::find_if(points.begin(), points.end(), [&](const auto& point) { return point.index == pointIndex; });
    if (existing == points.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelEdit pointEdit;
    pointEdit.remove = true;
    pointEdit.ownerIndex = seriesIndex;
    pointEdit.label.index = pointIndex;
    edit.pointDataLabelEdits.push_back(std::move(pointEdit));
    points.erase(existing);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesDataLabelPointRichText(const std::string& stableId, std::size_t seriesIndex,
                                                          std::size_t pointIndex, ChartRichText richText) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !richText.present || richText.runs.empty()) return false;
    auto& labels = it->series_[seriesIndex].dataLabels_;
    auto point = std::find_if(labels.points.begin(), labels.points.end(), [&](const auto& candidate) { return candidate.index == pointIndex; });
    if (point == labels.points.end()) {
        ChartDataLabelPoint created;
        created.index = pointIndex;
        created.richText = richText;
        labels.present = true;
        labels.points.push_back(created);
    } else point->richText = richText;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::PointDataLabelRichTextEdit richEdit;
    richEdit.seriesIndex = seriesIndex;
    richEdit.pointIndex = pointIndex;
    richEdit.richText = std::move(richText);
    edit.pointDataLabelRichTextEdits.push_back(std::move(richEdit));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesDataPointFormat(const std::string& stableId, std::size_t seriesIndex, ChartDataPointFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        !validChartLineFormat(format.line) || !validChartFillFormat(format.fill) || !validChartMarkerFormat(format.marker))
        return false;
    auto& points = it->series_[seriesIndex].dataPoints_;
    const auto existing = std::find_if(points.begin(), points.end(), [&](const auto& candidate) { return candidate.index == format.index; });
    if (existing == points.end()) points.push_back(format); else *existing = format;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::DataPointFormatEdit pointEdit;
    pointEdit.seriesIndex = seriesIndex;
    pointEdit.format = std::move(format);
    edit.dataPointFormatEdits.push_back(std::move(pointEdit));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartSeriesDataPointFormat(const std::string& stableId, std::size_t seriesIndex, std::size_t pointIndex) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    auto& points = it->series_[seriesIndex].dataPoints_;
    const auto existing = std::find_if(points.begin(), points.end(), [&](const auto& candidate) { return candidate.index == pointIndex; });
    if (existing == points.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::DataPointFormatEdit pointEdit;
    pointEdit.remove = true;
    pointEdit.seriesIndex = seriesIndex;
    pointEdit.format.index = pointIndex;
    edit.dataPointFormatEdits.push_back(std::move(pointEdit));
    points.erase(existing);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesLineFormat(const std::string& stableId, std::size_t seriesIndex, ChartLineFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validChartLineFormat(format)) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::SeriesFormatEdit formatEdit;
    formatEdit.kind = ImportedChartEdit::SeriesFormatEdit::Kind::Line;
    formatEdit.seriesIndex = seriesIndex;
    formatEdit.line = format;
    edit.seriesFormatEdits.push_back(std::move(formatEdit));
    it->series_[seriesIndex].lineFormat_ = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesFillFormat(const std::string& stableId, std::size_t seriesIndex, ChartFillFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validChartFillFormat(format)) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::SeriesFormatEdit formatEdit;
    formatEdit.kind = ImportedChartEdit::SeriesFormatEdit::Kind::Fill;
    formatEdit.seriesIndex = seriesIndex;
    formatEdit.fill = format;
    edit.seriesFormatEdits.push_back(std::move(formatEdit));
    it->series_[seriesIndex].fillFormat_ = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesMarkerFormat(const std::string& stableId, std::size_t seriesIndex, ChartMarkerFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validChartMarkerFormat(format)) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::SeriesFormatEdit formatEdit;
    formatEdit.kind = ImportedChartEdit::SeriesFormatEdit::Kind::Marker;
    formatEdit.seriesIndex = seriesIndex;
    formatEdit.marker = format;
    edit.seriesFormatEdits.push_back(std::move(formatEdit));
    it->series_[seriesIndex].markerFormat_ = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesTrendlineLineFormat(const std::string& stableId, std::size_t seriesIndex,
                                                   std::size_t trendlineIndex, ChartLineFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        trendlineIndex >= it->series_[seriesIndex].trendlines_.size() || !validChartLineFormat(format)) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::TrendlineFormatEdit formatEdit;
    formatEdit.seriesIndex = seriesIndex; formatEdit.trendlineIndex = trendlineIndex; formatEdit.line = format;
    edit.trendlineFormatEdits.push_back(std::move(formatEdit));
    it->series_[seriesIndex].trendlines_[trendlineIndex].lineFormat = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesErrorBarsLineFormat(const std::string& stableId, std::size_t seriesIndex,
                                                   ChartSeries::ErrorBarDirection direction, ChartLineFormat format) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) { return chart.imported_ && chart.stableId_ == stableId; });
    if (it == charts_.end() || seriesIndex >= it->series_.size() || !validChartLineFormat(format)) return false;
    auto& bars = it->series_[seriesIndex].errorBars_;
    const auto existing = std::find_if(bars.begin(), bars.end(), [&](const auto& candidate) { return candidate.direction == direction; });
    if (existing == bars.end()) return false;
    format.present = true;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::ErrorBarsFormatEdit formatEdit;
    formatEdit.seriesIndex = seriesIndex; formatEdit.direction = direction; formatEdit.line = format;
    edit.errorBarsFormatEdits.push_back(std::move(formatEdit));
    existing->lineFormat = std::move(format);
    dirty_ = true; drawingAppendDirty_ = true; return true;
}

bool Worksheet::setChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, std::size_t trendlineIndex,
                                         ChartSeries::Trendline trendline) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        trendlineIndex >= it->series_[seriesIndex].trendlines_.size()) return false;
    if (trendline.type == ChartSeries::TrendlineType::Polynomial && (trendline.order < 2 || trendline.order > 6)) return false;
    if (trendline.type == ChartSeries::TrendlineType::MovingAverage && (trendline.period < 2 || trendline.period > 255)) return false;
    if (!std::isfinite(trendline.forward) || !std::isfinite(trendline.backward) || trendline.forward < 0.0 || trendline.backward < 0.0) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::TrendlineEdit trendlineEdit;
    trendlineEdit.action = ImportedChartEdit::TrendlineEdit::Action::Set;
    trendlineEdit.seriesIndex = seriesIndex;
    trendlineEdit.trendlineIndex = trendlineIndex;
    trendlineEdit.trendline = trendline;
    edit.trendlineEdits.push_back(std::move(trendlineEdit));
    it->series_[seriesIndex].trendlines_[trendlineIndex] = trendline;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::addChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, ChartSeries::Trendline trendline) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    if (trendline.type == ChartSeries::TrendlineType::Polynomial && (trendline.order < 2 || trendline.order > 6)) return false;
    if (trendline.type == ChartSeries::TrendlineType::MovingAverage && (trendline.period < 2 || trendline.period > 255)) return false;
    if (!std::isfinite(trendline.forward) || !std::isfinite(trendline.backward) || trendline.forward < 0.0 || trendline.backward < 0.0) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::TrendlineEdit trendlineEdit;
    trendlineEdit.action = ImportedChartEdit::TrendlineEdit::Action::Add;
    trendlineEdit.seriesIndex = seriesIndex;
    trendlineEdit.trendlineIndex = it->series_[seriesIndex].trendlines_.size();
    trendlineEdit.trendline = trendline;
    edit.trendlineEdits.push_back(std::move(trendlineEdit));
    it->series_[seriesIndex].trendlines_.push_back(trendline);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartSeriesTrendline(const std::string& stableId, std::size_t seriesIndex, std::size_t trendlineIndex) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size() ||
        trendlineIndex >= it->series_[seriesIndex].trendlines_.size()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::TrendlineEdit trendlineEdit;
    trendlineEdit.action = ImportedChartEdit::TrendlineEdit::Action::Remove;
    trendlineEdit.seriesIndex = seriesIndex;
    trendlineEdit.trendlineIndex = trendlineIndex;
    edit.trendlineEdits.push_back(std::move(trendlineEdit));
    it->series_[seriesIndex].trendlines_.erase(it->series_[seriesIndex].trendlines_.begin() + static_cast<std::ptrdiff_t>(trendlineIndex));
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::setChartSeriesErrorBars(const std::string& stableId, std::size_t seriesIndex, ChartSeries::ErrorBars errorBars) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    if (errorBars.valueType == ChartSeries::ErrorValueType::Custom) {
        if (errorBars.plusReference.empty() || errorBars.minusReference.empty()) return false;
    } else if (!std::isfinite(errorBars.value) || errorBars.value < 0.0) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::ErrorBarsEdit barsEdit;
    barsEdit.seriesIndex = seriesIndex;
    barsEdit.errorBars = errorBars;
    barsEdit.direction = errorBars.direction;
    edit.errorBarsEdits.push_back(std::move(barsEdit));
    auto& bars = it->series_[seriesIndex].errorBars_;
    const auto existing = std::find_if(bars.begin(), bars.end(), [&](const auto& candidate) { return candidate.direction == errorBars.direction; });
    if (existing == bars.end()) bars.push_back(errorBars); else *existing = errorBars;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChartSeriesErrorBars(const std::string& stableId, std::size_t seriesIndex, ChartSeries::ErrorBarDirection direction) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end() || seriesIndex >= it->series_.size()) return false;
    auto& bars = it->series_[seriesIndex].errorBars_;
    const auto existing = std::find_if(bars.begin(), bars.end(), [&](const auto& candidate) { return candidate.direction == direction; });
    if (existing == bars.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    ImportedChartEdit::ErrorBarsEdit barsEdit;
    barsEdit.remove = true;
    barsEdit.seriesIndex = seriesIndex;
    barsEdit.direction = direction;
    edit.errorBarsEdits.push_back(std::move(barsEdit));
    bars.erase(existing);
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::removeChart(const std::string& stableId) {
    auto it = std::find_if(charts_.begin(), charts_.end(), [&](const Chart& chart) {
        return chart.imported_ && chart.stableId_ == stableId;
    });
    if (it == charts_.end()) return false;
    auto& edit = ensureImportedChartEdit(*it);
    edit.removed = true;
    const auto index = static_cast<std::size_t>(std::distance(charts_.begin(), it));
    charts_.erase(it);
    if (index < loadedChartCount_ && loadedChartCount_ > 0) --loadedChartCount_;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

} // namespace xlpp
