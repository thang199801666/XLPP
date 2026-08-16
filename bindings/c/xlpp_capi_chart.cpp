// XLPP C API — chart / image editing and chart value marshalling.
#include "xlpp_capi.h"
#include <XLPP/XLPP.h>
#include <XLPP/Chart/Chart.h>
#include <XLPP/Formula/DependencyGraph.h>
#include <cstring>
#include <filesystem>
#include <string>

#define WS2(h)   reinterpret_cast<xlpp::Worksheet*>(h)
#define IMG2(h)  reinterpret_cast<xlpp::Image*>(h)

static void copyStr(const std::string& value, char* out, int outSize) {
    if (!out || outSize <= 0) return;
    const auto n = static_cast<std::size_t>(outSize - 1);
    const auto len = value.size() < n ? value.size() : n;
    std::memcpy(out, value.data(), len);
    out[len] = '\0';
}

extern "C" {

// ============================================================
// Chart value handles (opaque heap objects used by P/Invoke)
// ============================================================

XLPP_API xlpp_chart_line_format xlpp_chart_line_value_create(void) {
    return reinterpret_cast<xlpp_chart_line_format>(new xlpp::ChartLineFormat());
}
XLPP_API void xlpp_chart_line_value_destroy(xlpp_chart_line_format v) {
    delete reinterpret_cast<xlpp::ChartLineFormat*>(v);
}
XLPP_API void xlpp_chart_line_value_set(xlpp_chart_line_format v, int present, int noFill,
                                        double width, const char* dash, const char* cap,
                                        const char* compound, const char* join) {
    auto* f = reinterpret_cast<xlpp::ChartLineFormat*>(v);
    f->present = present != 0; f->noFill = noFill != 0; f->widthPoints = width;
    f->dash = dash ? dash : ""; f->cap = cap ? cap : ""; f->compound = compound ? compound : "";
    f->join = join ? join : "";
}
XLPP_API void xlpp_chart_line_value_set_color(xlpp_chart_line_format v, int kind, const char* color) {
    auto* f = reinterpret_cast<xlpp::ChartLineFormat*>(v);
    f->color.kind = static_cast<xlpp::ChartColor::Kind>(kind);
    f->color.value = color ? color : "";
}
XLPP_API void xlpp_chart_line_value_add_color_transform(xlpp_chart_line_format v, int kind, int amount) {
    auto* f = reinterpret_cast<xlpp::ChartLineFormat*>(v);
    xlpp::ChartColorTransform t;
    t.kind = static_cast<xlpp::ChartColorTransform::Kind>(kind);
    t.value = amount;
    f->color.transforms.push_back(t);
}
XLPP_API void xlpp_chart_line_value_add_custom_dash(xlpp_chart_line_format v, double dash, double space) {
    auto* f = reinterpret_cast<xlpp::ChartLineFormat*>(v);
    f->customDash.push_back({dash, space});
}

XLPP_API xlpp_chart_fill_format xlpp_chart_fill_value_create(void) {
    return reinterpret_cast<xlpp_chart_fill_format>(new xlpp::ChartFillFormat());
}
XLPP_API void xlpp_chart_fill_value_destroy(xlpp_chart_fill_format v) {
    delete reinterpret_cast<xlpp::ChartFillFormat*>(v);
}
XLPP_API void xlpp_chart_fill_value_set(xlpp_chart_fill_format v, int present, int noFill,
                                        int kind, double angle, const char* pattern) {
    auto* f = reinterpret_cast<xlpp::ChartFillFormat*>(v);
    f->present = present != 0; f->noFill = noFill != 0;
    f->kind = static_cast<xlpp::ChartFillFormat::Kind>(kind);
    f->gradientAngleDegrees = angle; f->pattern = pattern ? pattern : "";
}
XLPP_API void xlpp_chart_fill_value_set_color(xlpp_chart_fill_format v, int role, int kind, const char* color) {
    auto* f = reinterpret_cast<xlpp::ChartFillFormat*>(v);
    xlpp::ChartColor c;
    c.kind = static_cast<xlpp::ChartColor::Kind>(kind);
    c.value = color ? color : "";
    if (role == 0) f->color = c;
    else if (role == 1) f->foregroundColor = c;
    else f->backgroundColor = c;
}
XLPP_API void xlpp_chart_fill_value_add_gradient_stop(xlpp_chart_fill_format v, int position, int kind, const char* color) {
    auto* f = reinterpret_cast<xlpp::ChartFillFormat*>(v);
    xlpp::ChartGradientStop stop;
    stop.position = position;
    stop.color.kind = static_cast<xlpp::ChartColor::Kind>(kind);
    stop.color.value = color ? color : "";
    f->gradientStops.push_back(stop);
}

XLPP_API xlpp_chart_rich_text xlpp_chart_rich_text_value_create(void) {
    return reinterpret_cast<xlpp_chart_rich_text>(new xlpp::ChartRichText());
}
XLPP_API void xlpp_chart_rich_text_value_destroy(xlpp_chart_rich_text v) {
    delete reinterpret_cast<xlpp::ChartRichText*>(v);
}
XLPP_API void xlpp_chart_rich_text_value_set_present(xlpp_chart_rich_text v, int present) {
    reinterpret_cast<xlpp::ChartRichText*>(v)->present = present != 0;
}
XLPP_API void xlpp_chart_rich_text_value_add_run(xlpp_chart_rich_text v, const char* text, int bold,
                                                 int italic, double size, const char* typeface,
                                                 int colorKind, const char* color) {
    auto* rt = reinterpret_cast<xlpp::ChartRichText*>(v);
    xlpp::ChartTextRun run;
    run.text = text ? text : "";
    run.bold = bold != 0; run.italic = italic != 0;
    run.fontSizePoints = size; run.typeface = typeface ? typeface : "";
    run.color.kind = static_cast<xlpp::ChartColor::Kind>(colorKind);
    run.color.value = color ? color : "";
    rt->runs.push_back(std::move(run));
}

XLPP_API xlpp_chart_series_cache xlpp_chart_cache_value_create(void) {
    return reinterpret_cast<xlpp_chart_series_cache>(new xlpp::ChartSeriesCache());
}
XLPP_API void xlpp_chart_cache_value_destroy(xlpp_chart_series_cache v) {
    delete reinterpret_cast<xlpp::ChartSeriesCache*>(v);
}
XLPP_API void xlpp_chart_cache_value_set(xlpp_chart_series_cache v, int present, int numeric,
                                         const char* format, unsigned long long count) {
    auto* c = reinterpret_cast<xlpp::ChartSeriesCache*>(v);
    c->present = present != 0; c->numeric = numeric != 0;
    c->formatCode = format ? format : ""; c->pointCount = static_cast<std::size_t>(count);
}
XLPP_API void xlpp_chart_cache_value_add_point(xlpp_chart_series_cache v, unsigned long long index, const char* pointValue) {
    auto* c = reinterpret_cast<xlpp::ChartSeriesCache*>(v);
    c->points.push_back({static_cast<std::size_t>(index), pointValue ? pointValue : ""});
}

// ============================================================
// Worksheet image editing
// ============================================================

XLPP_API xlpp_image xlpp_sheet_image_by_stable_id(xlpp_worksheet ws, const char* stableId) {
    try {
        for (auto& img : WS2(ws)->images()) if (img.stableId() == stableId) return reinterpret_cast<xlpp_image>(&img);
    } catch (...) {}
    return nullptr;
}
XLPP_API int xlpp_sheet_move_image(xlpp_worksheet ws, const char* id, const char* anchor) {
    try { return WS2(ws)->moveImage(id, anchor) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_move_image_absolute(xlpp_worksheet ws, const char* id, long long x, long long y) {
    try { return WS2(ws)->moveImageAbsolute(id, x, y) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_resize_image(xlpp_worksheet ws, const char* id, double w, double h) {
    try { return WS2(ws)->resizeImage(id, w, h) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_replace_image(xlpp_worksheet ws, const char* id, const char* path) {
    try { return WS2(ws)->replaceImage(id, std::filesystem::path(path)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_image(xlpp_worksheet ws, const char* id) {
    try { return WS2(ws)->removeImage(id) ? 1 : 0; } catch (...) { return 0; }
}

// ============================================================
// Worksheet chart editing
// ============================================================

XLPP_API xlpp_chart xlpp_sheet_chart_by_stable_id(xlpp_worksheet ws, const char* stableId) {
    try {
        for (std::size_t i = 0; i < WS2(ws)->chartCount(); ++i)
            if (WS2(ws)->chart(i).stableId() == stableId) return reinterpret_cast<xlpp_chart>(&WS2(ws)->chart(i));
    } catch (...) {}
    return nullptr;
}
XLPP_API int xlpp_sheet_move_chart(xlpp_worksheet ws, const char* id, const char* anchor) {
    try { return WS2(ws)->moveChart(id, anchor) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_move_chart_absolute(xlpp_worksheet ws, const char* id, long long x, long long y) {
    try { return WS2(ws)->moveChartAbsolute(id, x, y) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_resize_chart(xlpp_worksheet ws, const char* id, double w, double h) {
    try { return WS2(ws)->resizeChart(id, w, h) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart(xlpp_worksheet ws, const char* id) {
    try { return WS2(ws)->removeChart(id) ? 1 : 0; } catch (...) { return 0; }
}

XLPP_API int xlpp_sheet_set_chart_title(xlpp_worksheet ws, const char* id, const char* title) {
    try { return WS2(ws)->setChartTitle(id, title ? title : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_style(xlpp_worksheet ws, const char* id, const char* style) {
    try { return WS2(ws)->setChartStyle(id, style ? style : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_x_axis_title(xlpp_worksheet ws, const char* id, const char* title) {
    try { return WS2(ws)->setChartXAxisTitle(id, title ? title : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_y_axis_title(xlpp_worksheet ws, const char* id, const char* title) {
    try { return WS2(ws)->setChartYAxisTitle(id, title ? title : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_title(xlpp_worksheet ws, const char* id, unsigned long long axis, const char* title) {
    try { return WS2(ws)->setChartAxisTitle(id, axis, title ? title : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_number_format(xlpp_worksheet ws, const char* id, unsigned long long axis, const char* code, int linked) {
    try { return WS2(ws)->setChartAxisNumberFormat(id, axis, code ? code : "", linked != 0) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_ticks(xlpp_worksheet ws, const char* id, unsigned long long axis,
                                             const char* major, const char* minor, const char* position) {
    try { return WS2(ws)->setChartAxisTicks(id, axis, major ? major : "", minor ? minor : "", position ? position : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_units(xlpp_worksheet ws, const char* id, unsigned long long axis, double major, double minor) {
    try { return WS2(ws)->setChartAxisUnits(id, axis, major, minor) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_crossing(xlpp_worksheet ws, const char* id, unsigned long long axis, const char* crosses, const char* between) {
    try { return WS2(ws)->setChartAxisCrossing(id, axis, crosses ? crosses : "", between ? between : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_crosses_at(xlpp_worksheet ws, const char* id, unsigned long long axis, double value) {
    try { return WS2(ws)->setChartAxisCrossesAt(id, axis, value) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_clear_chart_axis_crosses_at(xlpp_worksheet ws, const char* id, unsigned long long axis) {
    try { return WS2(ws)->clearChartAxisCrossesAt(id, axis) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_display_units(xlpp_worksheet ws, const char* id, unsigned long long axis,
                                                     int present, const char* builtIn, int hasCustom, double custom,
                                                     int showLabel, xlpp_chart_rich_text label) {
    try {
        xlpp::ChartDisplayUnits units;
        units.present = present != 0; units.builtInUnit = builtIn ? builtIn : "";
        units.hasCustomUnit = hasCustom != 0; units.customUnit = custom; units.showLabel = showLabel != 0;
        if (label) units.labelRichText = *reinterpret_cast<xlpp::ChartRichText*>(label);
        return WS2(ws)->setChartAxisDisplayUnits(id, axis, units) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_clear_chart_axis_display_units(xlpp_worksheet ws, const char* id, unsigned long long axis) {
    try { return WS2(ws)->clearChartAxisDisplayUnits(id, axis) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_line_format(xlpp_worksheet ws, const char* id, unsigned long long axis, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartAxisLineFormat(id, axis, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_gridline_format(xlpp_worksheet ws, const char* id, unsigned long long axis, int major, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartAxisGridlineFormat(id, axis, major != 0, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_axis_gridlines(xlpp_worksheet ws, const char* id, unsigned long long axis, int major) {
    try { return WS2(ws)->removeChartAxisGridlines(id, axis, major != 0) ? 1 : 0; } catch (...) { return 0; }
}

XLPP_API int xlpp_sheet_set_chart_area_line_format(xlpp_worksheet ws, const char* id, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartAreaLineFormat(id, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_area_fill_format(xlpp_worksheet ws, const char* id, xlpp_chart_fill_format value) {
    try { return WS2(ws)->setChartAreaFillFormat(id, *reinterpret_cast<xlpp::ChartFillFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_area_line_format(xlpp_worksheet ws, const char* id, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartPlotAreaLineFormat(id, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_area_fill_format(xlpp_worksheet ws, const char* id, xlpp_chart_fill_format value) {
    try { return WS2(ws)->setChartPlotAreaFillFormat(id, *reinterpret_cast<xlpp::ChartFillFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_area_layout(xlpp_worksheet ws, const char* id, int present, const char* target,
                                                   const char* xm, const char* ym, const char* wm, const char* hm,
                                                   int hasX, double x, int hasY, double y, int hasW, double w, int hasH, double h) {
    try {
        xlpp::ChartManualLayout layout;
        layout.present = present != 0; layout.target = target ? target : "";
        layout.xMode = xm ? xm : ""; layout.yMode = ym ? ym : ""; layout.widthMode = wm ? wm : ""; layout.heightMode = hm ? hm : "";
        layout.hasX = hasX != 0; layout.x = x; layout.hasY = hasY != 0; layout.y = y;
        layout.hasWidth = hasW != 0; layout.width = w; layout.hasHeight = hasH != 0; layout.height = h;
        return WS2(ws)->setChartPlotAreaLayout(id, layout) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_view_3d(xlpp_worksheet ws, const char* id, int present,
                                          int hrx, int rx, int hry, int ry, int hhp, int hp,
                                          int hdp, int dp, int hra, int ra, int hpersp, int persp) {
    try {
        xlpp::ChartView3D view;
        view.present = present != 0;
        view.hasRotationX = hrx != 0; view.rotationX = rx;
        view.hasRotationY = hry != 0; view.rotationY = ry;
        view.hasHeightPercent = hhp != 0; view.heightPercent = hp;
        view.hasDepthPercent = hdp != 0; view.depthPercent = dp;
        view.hasRightAngleAxes = hra != 0; view.rightAngleAxes = ra != 0;
        view.hasPerspective = hpersp != 0; view.perspective = persp;
        return WS2(ws)->setChartView3D(id, view) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_drop_lines(xlpp_worksheet ws, const char* id, unsigned long long plot, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartPlotDropLines(id, plot, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_high_low_lines(xlpp_worksheet ws, const char* id, unsigned long long plot, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartPlotHighLowLines(id, plot, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_leader_line_format(xlpp_worksheet ws, const char* id, unsigned long long plot, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartPlotLeaderLineFormat(id, plot, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_plot_drop_lines(xlpp_worksheet ws, const char* id, unsigned long long plot) {
    try { return WS2(ws)->removeChartPlotDropLines(id, plot) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_plot_high_low_lines(xlpp_worksheet ws, const char* id, unsigned long long plot) {
    try { return WS2(ws)->removeChartPlotHighLowLines(id, plot) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_plot_up_down_bars(xlpp_worksheet ws, const char* id, unsigned long long plot) {
    try { return WS2(ws)->removeChartPlotUpDownBars(id, plot) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_plot_leader_lines(xlpp_worksheet ws, const char* id, unsigned long long plot) {
    try { return WS2(ws)->removeChartPlotLeaderLines(id, plot) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_first_slice_angle(xlpp_worksheet ws, const char* id, unsigned long long plot, int degrees) {
    try { return WS2(ws)->setChartPlotFirstSliceAngle(id, plot, degrees) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_doughnut_hole_size(xlpp_worksheet ws, const char* id, unsigned long long plot, int percent) {
    try { return WS2(ws)->setChartPlotDoughnutHoleSize(id, plot, percent) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_plot_radar_style(xlpp_worksheet ws, const char* id, unsigned long long plot, const char* style) {
    try { return WS2(ws)->setChartPlotRadarStyle(id, plot, style ? style : "") ? 1 : 0; } catch (...) { return 0; }
}

XLPP_API int xlpp_sheet_set_chart_title_rich_text(xlpp_worksheet ws, const char* id, xlpp_chart_rich_text value) {
    try { return WS2(ws)->setChartTitleRichText(id, *reinterpret_cast<xlpp::ChartRichText*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_title_rich_text(xlpp_worksheet ws, const char* id, unsigned long long axis, xlpp_chart_rich_text value) {
    try { return WS2(ws)->setChartAxisTitleRichText(id, axis, *reinterpret_cast<xlpp::ChartRichText*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_axis_scaling(xlpp_worksheet ws, const char* id, unsigned long long axis,
                                               int hasMin, double min, int hasMax, double max,
                                               int hasLog, double logBase, int reverse) {
    try {
        xlpp::ChartAxisScaling scaling;
        scaling.hasMinimum = hasMin != 0; scaling.minimum = min;
        scaling.hasMaximum = hasMax != 0; scaling.maximum = max;
        scaling.hasLogBase = hasLog != 0; scaling.logBase = logBase;
        scaling.reverseOrder = reverse != 0;
        return WS2(ws)->setChartAxisScaling(id, axis, scaling) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_legend(xlpp_worksheet ws, const char* id, int show, const char* position) {
    try { return WS2(ws)->setChartLegend(id, show != 0, position ? position : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_legend_overlay(xlpp_worksheet ws, const char* id, int overlay) {
    try { return WS2(ws)->setChartLegendOverlay(id, overlay != 0) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_legend_line_format(xlpp_worksheet ws, const char* id, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartLegendLineFormat(id, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_legend_fill_format(xlpp_worksheet ws, const char* id, xlpp_chart_fill_format value) {
    try { return WS2(ws)->setChartLegendFillFormat(id, *reinterpret_cast<xlpp::ChartFillFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_legend_layout(xlpp_worksheet ws, const char* id, int present, const char* target,
                                                const char* xm, const char* ym, const char* wm, const char* hm,
                                                int hasX, double x, int hasY, double y, int hasW, double w, int hasH, double h) {
    try {
        xlpp::ChartManualLayout layout;
        layout.present = present != 0; layout.target = target ? target : "";
        layout.xMode = xm ? xm : ""; layout.yMode = ym ? ym : ""; layout.widthMode = wm ? wm : ""; layout.heightMode = hm ? hm : "";
        layout.hasX = hasX != 0; layout.x = x; layout.hasY = hasY != 0; layout.y = y;
        layout.hasWidth = hasW != 0; layout.width = w; layout.hasHeight = hasH != 0; layout.height = h;
        return WS2(ws)->setChartLegendLayout(id, layout) ? 1 : 0;
    } catch (...) { return 0; }
}

XLPP_API int xlpp_sheet_set_chart_data_table(xlpp_worksheet ws, const char* id, int present, int hb, int vb,
                                             int outline, int legendKeys, xlpp_chart_line_format line, xlpp_chart_fill_format fill) {
    try {
        xlpp::ChartDataTable table;
        table.present = present != 0;
        table.showHorizontalBorder = hb != 0; table.showVerticalBorder = vb != 0;
        table.showOutline = outline != 0; table.showLegendKeys = legendKeys != 0;
        if (line) table.line = *reinterpret_cast<xlpp::ChartLineFormat*>(line);
        if (fill) table.fill = *reinterpret_cast<xlpp::ChartFillFormat*>(fill);
        return WS2(ws)->setChartDataTable(id, table) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_data_table(xlpp_worksheet ws, const char* id) {
    try { return WS2(ws)->removeChartDataTable(id) ? 1 : 0; } catch (...) { return 0; }
}

XLPP_API int xlpp_sheet_set_chart_series_title(xlpp_worksheet ws, const char* id, unsigned long long series, const char* title) {
    try { return WS2(ws)->setChartSeriesTitle(id, series, title ? title : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_references(xlpp_worksheet ws, const char* id, unsigned long long series,
                                                    const char* categories, const char* values) {
    try { return WS2(ws)->setChartSeriesReferences(id, series, categories ? categories : "", values ? values : "") ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_category_cache(xlpp_worksheet ws, const char* id, unsigned long long series, xlpp_chart_series_cache value) {
    try { return WS2(ws)->setChartSeriesCategoryCache(id, series, *reinterpret_cast<xlpp::ChartSeriesCache*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_value_cache(xlpp_worksheet ws, const char* id, unsigned long long series, xlpp_chart_series_cache value) {
    try { return WS2(ws)->setChartSeriesValueCache(id, series, *reinterpret_cast<xlpp::ChartSeriesCache*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_title_cache(xlpp_worksheet ws, const char* id, unsigned long long series, xlpp_chart_series_cache value) {
    try { return WS2(ws)->setChartSeriesTitleCache(id, series, *reinterpret_cast<xlpp::ChartSeriesCache*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_clear_chart_series_caches(xlpp_worksheet ws, const char* id, unsigned long long series) {
    try { return WS2(ws)->clearChartSeriesCaches(id, series) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_line_format(xlpp_worksheet ws, const char* id, unsigned long long series, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartSeriesLineFormat(id, series, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_fill_format(xlpp_worksheet ws, const char* id, unsigned long long series, xlpp_chart_fill_format value) {
    try { return WS2(ws)->setChartSeriesFillFormat(id, series, *reinterpret_cast<xlpp::ChartFillFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_leader_line_format(xlpp_worksheet ws, const char* id, unsigned long long series, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartSeriesLeaderLineFormat(id, series, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_series_leader_lines(xlpp_worksheet ws, const char* id, unsigned long long series) {
    try { return WS2(ws)->removeChartSeriesLeaderLines(id, series) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_trendline_line_format(xlpp_worksheet ws, const char* id, unsigned long long series, unsigned long long index, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartSeriesTrendlineLineFormat(id, series, index, *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_error_bars_line_format(xlpp_worksheet ws, const char* id, unsigned long long series, int direction, xlpp_chart_line_format value) {
    try { return WS2(ws)->setChartSeriesErrorBarsLineFormat(id, series, static_cast<xlpp::ChartSeries::ErrorBarDirection>(direction), *reinterpret_cast<xlpp::ChartLineFormat*>(value)) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_trendline(xlpp_worksheet ws, const char* id, unsigned long long series, unsigned long long index,
                                                   int type, int order, int period, double forward, double backward,
                                                   int equation, int r2, xlpp_chart_line_format line) {
    try {
        xlpp::ChartSeries::Trendline tr;
        tr.type = static_cast<xlpp::ChartSeries::TrendlineType>(type);
        tr.order = order; tr.period = period; tr.forward = forward; tr.backward = backward;
        tr.displayEquation = equation != 0; tr.displayRSquared = r2 != 0;
        if (line) tr.lineFormat = *reinterpret_cast<xlpp::ChartLineFormat*>(line);
        return WS2(ws)->setChartSeriesTrendline(id, series, index, tr) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_add_chart_series_trendline(xlpp_worksheet ws, const char* id, unsigned long long series,
                                                   int type, int order, int period, double forward, double backward,
                                                   int equation, int r2, xlpp_chart_line_format line) {
    try {
        xlpp::ChartSeries::Trendline tr;
        tr.type = static_cast<xlpp::ChartSeries::TrendlineType>(type);
        tr.order = order; tr.period = period; tr.forward = forward; tr.backward = backward;
        tr.displayEquation = equation != 0; tr.displayRSquared = r2 != 0;
        if (line) tr.lineFormat = *reinterpret_cast<xlpp::ChartLineFormat*>(line);
        return WS2(ws)->addChartSeriesTrendline(id, series, tr) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_set_chart_series_error_bars(xlpp_worksheet ws, const char* id, unsigned long long series,
                                                    int direction, int barType, int valueType, double value,
                                                    int noEndCap, const char* plus, const char* minus, xlpp_chart_line_format line) {
    try {
        xlpp::ChartSeries::ErrorBars eb;
        eb.direction = static_cast<xlpp::ChartSeries::ErrorBarDirection>(direction);
        eb.barType = static_cast<xlpp::ChartSeries::ErrorBarType>(barType);
        eb.valueType = static_cast<xlpp::ChartSeries::ErrorValueType>(valueType);
        eb.value = value; eb.noEndCap = noEndCap != 0;
        eb.plusReference = plus ? plus : ""; eb.minusReference = minus ? minus : "";
        if (line) eb.lineFormat = *reinterpret_cast<xlpp::ChartLineFormat*>(line);
        return WS2(ws)->setChartSeriesErrorBars(id, series, eb) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_series_trendline(xlpp_worksheet ws, const char* id, unsigned long long series, unsigned long long index) {
    try { return WS2(ws)->removeChartSeriesTrendline(id, series, index) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API int xlpp_sheet_remove_chart_series_error_bars(xlpp_worksheet ws, const char* id, unsigned long long series, int direction) {
    try { return WS2(ws)->removeChartSeriesErrorBars(id, series, static_cast<xlpp::ChartSeries::ErrorBarDirection>(direction)) ? 1 : 0; } catch (...) { return 0; }
}

// ============================================================
// Formula dependency graph
// ============================================================

XLPP_API xlpp_dependency_graph xlpp_dependency_graph_build(xlpp_workbook wb) {
    try { return reinterpret_cast<xlpp_dependency_graph>(new xlpp::FormulaDependencyGraph(xlpp::buildFormulaDependencyGraph(*reinterpret_cast<xlpp::Workbook*>(wb)))); }
    catch (...) { return nullptr; }
}
XLPP_API void xlpp_dependency_graph_destroy(xlpp_dependency_graph g) {
    delete reinterpret_cast<xlpp::FormulaDependencyGraph*>(g);
}
XLPP_API unsigned long long xlpp_dependency_graph_edge_count(xlpp_dependency_graph g) {
    return reinterpret_cast<xlpp::FormulaDependencyGraph*>(g)->edges().size();
}
XLPP_API void xlpp_dependency_graph_edge_kind(xlpp_dependency_graph g, unsigned long long index, char* out, int outSize) {
    try {
        const auto& e = reinterpret_cast<xlpp::FormulaDependencyGraph*>(g)->edges()[index];
        std::string s;
        switch (e.kind) {
            case xlpp::FormulaDependencyKind::CellOrRange: s = "cell"; break;
            case xlpp::FormulaDependencyKind::DefinedName: s = "name"; break;
            case xlpp::FormulaDependencyKind::Table: s = "table"; break;
            case xlpp::FormulaDependencyKind::ExternalReference: s = "external"; break;
            case xlpp::FormulaDependencyKind::VolatileReference: s = "volatile"; break;
        }
        copyStr(s, out, outSize);
    } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}
XLPP_API void xlpp_dependency_graph_edge_dependent(xlpp_dependency_graph g, unsigned long long index, char* out, int outSize) {
    try { copyStr(reinterpret_cast<xlpp::FormulaDependencyGraph*>(g)->edges()[index].dependentSheet + "!" + reinterpret_cast<xlpp::FormulaDependencyGraph*>(g)->edges()[index].dependentCell, out, outSize); }
    catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}
XLPP_API void xlpp_dependency_graph_edge_precedent(xlpp_dependency_graph g, unsigned long long index, char* out, int outSize) {
    try {
        const auto& e = reinterpret_cast<xlpp::FormulaDependencyGraph*>(g)->edges()[index];
        copyStr(e.precedentSheet + "!" + e.precedentReference, out, outSize);
    } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}
XLPP_API void xlpp_dependency_graph_edge_symbol(xlpp_dependency_graph g, unsigned long long index, char* out, int outSize) {
    try { copyStr(reinterpret_cast<xlpp::FormulaDependencyGraph*>(g)->edges()[index].symbol, out, outSize); }
    catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}
XLPP_API int xlpp_dependency_graph_depends_on(xlpp_dependency_graph g, const char* ds, const char* dc, const char* ps, const char* pc) {
    try { return reinterpret_cast<xlpp::FormulaDependencyGraph*>(g)->dependsOn(ds, dc, ps, pc) ? 1 : 0; }
    catch (...) { return 0; }
}

} // extern "C"
