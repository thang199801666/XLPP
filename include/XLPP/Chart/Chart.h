#pragma once
#include <XLPP/Core/StableVector.h>
#include <string>
#include <vector>
#include <memory>
#include <cstdint>
#include <algorithm>
#include <utility>
#include <cmath>
#include <iomanip>
#include <sstream>
#include <set>
#include <XLPP/Worksheet/Drawings/Image.h>

namespace xlpp {

struct ChartColorTransform {
    enum class Kind { Alpha, AlphaMod, AlphaOff, Tint, Shade, LumMod, LumOff, SatMod, SatOff };
    Kind kind{Kind::Alpha};
    int value{100000};
};

struct ChartColor {
    enum class Kind { None, SRgb, Scheme, System, Preset, Unknown };
    Kind kind{Kind::None};
    std::string value;
    std::vector<ChartColorTransform> transforms;
    bool present() const noexcept { return kind != Kind::None && !value.empty(); }
};

struct ChartCustomDashStop {
    double dash{0.0};
    double space{0.0};
};

struct ChartLineFormat {
    bool present{false};
    bool noFill{false};
    ChartColor color{};
    double widthPoints{0.0};
    std::string dash;
    std::string cap;
    std::string compound;
    std::string join;
    std::vector<ChartCustomDashStop> customDash;
};

struct ChartGradientStop {
    int position{0};
    ChartColor color{};
};

struct ChartFillFormat {
    enum class Kind { None, NoFill, Solid, Gradient, Pattern };
    bool present{false};
    bool noFill{false};
    ChartColor color{};
    Kind kind{Kind::None};
    std::vector<ChartGradientStop> gradientStops;
    double gradientAngleDegrees{0.0};
    std::string pattern;
    ChartColor foregroundColor{};
    ChartColor backgroundColor{};
};

struct ChartTextRun {
    std::string text;
    bool bold{false};
    bool italic{false};
    double fontSizePoints{0.0};
    std::string typeface;
    ChartColor color{};
};

struct ChartTextStyle {
    bool present{false};
    bool bold{false};
    bool italic{false};
    double fontSizePoints{0.0};
    std::string typeface;
    ChartColor color{};
};

struct ChartRichText {
    bool present{false};
    std::vector<ChartTextRun> runs;
    std::string plainText() const {
        std::string result;
        for (const auto& run : runs) result += run.text;
        return result;
    }
};


struct ChartCachePoint {
    std::size_t index{0};
    std::string value;
};

struct ChartSeriesCache {
    bool present{false};
    bool numeric{false};
    std::string formatCode;
    std::size_t pointCount{0};
    std::vector<ChartCachePoint> points;

    std::size_t effectivePointCount() const noexcept {
        std::size_t count = pointCount;
        for (const auto& point : points) count = std::max(count, point.index + 1);
        return count;
    }
    bool hasDuplicateIndexes() const {
        std::set<std::size_t> seen;
        for (const auto& point : points) if (!seen.insert(point.index).second) return true;
        return false;
    }
    bool ordered() const noexcept {
        for (std::size_t i = 1; i < points.size(); ++i) if (points[i - 1].index >= points[i].index) return false;
        return true;
    }
    bool sparse() const noexcept {
        if (!present) return false;
        const auto count = effectivePointCount();
        if (points.size() < count) return true;
        for (std::size_t i = 0; i < points.size(); ++i) if (points[i].index != i) return true;
        return false;
    }
    bool valid(bool allowSparse = true) const {
        if (!present) return true;
        if (hasDuplicateIndexes()) return false;
        if (pointCount != 0 && pointCount < effectivePointCount()) return false;
        if (!allowSparse && sparse()) return false;
        return true;
    }
};

struct ChartResolvedColor {
    bool present{false};
    int red{0};
    int green{0};
    int blue{0};
    double alpha{1.0};
    std::string srgb() const {
        if (!present) return {};
        std::ostringstream out;
        out << std::uppercase << std::hex << std::setfill('0')
            << std::setw(2) << std::clamp(red, 0, 255)
            << std::setw(2) << std::clamp(green, 0, 255)
            << std::setw(2) << std::clamp(blue, 0, 255);
        return out.str();
    }
};

struct ChartThemeFontScheme {
    bool present{false};
    std::string name;
    std::string majorLatinTypeface;
    std::string minorLatinTypeface;
};

struct ChartThemeEffectScheme {
    bool present{false};
    std::string name;
    std::size_t fillStyleCount{0};
    std::size_t lineStyleCount{0};
    std::size_t effectStyleCount{0};
    std::size_t backgroundFillStyleCount{0};
};

struct ChartThemeColor {
    std::string name;
    std::string srgb;
};

struct ChartThemePalette {
    bool present{false};
    std::vector<ChartThemeColor> colors;
    ChartThemeFontScheme fontScheme{};
    ChartThemeEffectScheme effectScheme{};
    std::string baseColor(const std::string& schemeName) const {
        const auto it = std::find_if(colors.begin(), colors.end(), [&](const auto& item) { return item.name == schemeName; });
        return it == colors.end() ? std::string{} : it->srgb;
    }
    std::string resolveBase(const ChartColor& color) const {
        if (color.kind == ChartColor::Kind::SRgb) return color.value;
        if (color.kind == ChartColor::Kind::Scheme) return baseColor(color.value);
        return {};
    }
    ChartResolvedColor resolve(const ChartColor& color) const {
        auto hex = resolveBase(color);
        if (hex.size() != 6) return {};
        auto hexByte = [&](std::size_t offset) -> int {
            try { return std::stoi(hex.substr(offset, 2), nullptr, 16); } catch (...) { return -1; }
        };
        int r = hexByte(0), g = hexByte(2), b = hexByte(4);
        if (r < 0 || g < 0 || b < 0) return {};
        double rd = r / 255.0, gd = g / 255.0, bd = b / 255.0, alpha = 1.0;
        auto clamp01 = [](double v) { return std::clamp(v, 0.0, 1.0); };
        auto rgbToHsl = [&](double& h, double& sat, double& lum) {
            const double mx = std::max({rd, gd, bd}), mn = std::min({rd, gd, bd});
            const double d = mx - mn; lum = (mx + mn) / 2.0;
            sat = d == 0.0 ? 0.0 : d / (1.0 - std::abs(2.0 * lum - 1.0));
            if (d == 0.0) h = 0.0;
            else if (mx == rd) h = std::fmod((gd - bd) / d, 6.0) / 6.0;
            else if (mx == gd) h = (((bd - rd) / d) + 2.0) / 6.0;
            else h = (((rd - gd) / d) + 4.0) / 6.0;
            if (h < 0.0) h += 1.0;
        };
        auto hslToRgb = [&](double h, double sat, double lum) {
            sat = clamp01(sat); lum = clamp01(lum);
            const double c = (1.0 - std::abs(2.0 * lum - 1.0)) * sat;
            const double hp = h * 6.0; const double x = c * (1.0 - std::abs(std::fmod(hp, 2.0) - 1.0));
            double rr=0, gg=0, bb=0;
            if (hp < 1) { rr=c; gg=x; } else if (hp < 2) { rr=x; gg=c; } else if (hp < 3) { gg=c; bb=x; }
            else if (hp < 4) { gg=x; bb=c; } else if (hp < 5) { rr=x; bb=c; } else { rr=c; bb=x; }
            const double m = lum - c / 2.0; rd=clamp01(rr+m); gd=clamp01(gg+m); bd=clamp01(bb+m);
        };
        for (const auto& transform : color.transforms) {
            const double v = transform.value / 100000.0;
            using K = ChartColorTransform::Kind;
            if (transform.kind == K::Alpha) alpha = clamp01(v);
            else if (transform.kind == K::AlphaMod) alpha = clamp01(alpha * v);
            else if (transform.kind == K::AlphaOff) alpha = clamp01(alpha + v);
            else if (transform.kind == K::Tint) { rd=clamp01(rd+(1-rd)*v); gd=clamp01(gd+(1-gd)*v); bd=clamp01(bd+(1-bd)*v); }
            else if (transform.kind == K::Shade) { rd=clamp01(rd*v); gd=clamp01(gd*v); bd=clamp01(bd*v); }
            else {
                double h=0, sat=0, lum=0; rgbToHsl(h,sat,lum);
                if (transform.kind == K::LumMod) lum *= v;
                else if (transform.kind == K::LumOff) lum += v;
                else if (transform.kind == K::SatMod) sat *= v;
                else if (transform.kind == K::SatOff) sat += v;
                hslToRgb(h,sat,lum);
            }
        }
        ChartResolvedColor result; result.present=true; result.red=static_cast<int>(std::lround(rd*255.0));
        result.green=static_cast<int>(std::lround(gd*255.0)); result.blue=static_cast<int>(std::lround(bd*255.0)); result.alpha=alpha; return result;
    }
    std::string resolveFinalRgb(const ChartColor& color) const { return resolve(color).srgb(); }
};

struct ChartStyleResources {
    bool chartStylePresent{false};
    bool colorStylePresent{false};
    std::string chartStylePart;
    std::string colorStylePart;
};

struct ChartMarkerFormat {
    bool present{false};
    std::string symbol;
    int size{0};
    ChartFillFormat fill{};
    ChartLineFormat line{};
};

struct ChartDataLabelPoint {
    std::size_t index{0};
    bool deleted{false};
    bool showLegendKey{false};
    bool showValue{false};
    bool showCategoryName{false};
    bool showSeriesName{false};
    bool showPercent{false};
    bool showBubbleSize{false};
    bool showLeaderLines{false};
    std::string position;
    std::string separator;
    ChartRichText richText{};
};

struct ChartDataLabels {
    bool present{false};
    bool showLegendKey{false};
    bool showValue{false};
    bool showCategoryName{false};
    bool showSeriesName{false};
    bool showPercent{false};
    bool showBubbleSize{false};
    bool showLeaderLines{false};
    bool hasLeaderLines{false};
    ChartLineFormat leaderLineFormat{};
    std::string position;
    std::string separator;
    std::vector<ChartDataLabelPoint> points;
};

struct ChartDataTable {
    bool present{false};
    bool showHorizontalBorder{false};
    bool showVerticalBorder{false};
    bool showOutline{false};
    bool showLegendKeys{false};
    ChartFillFormat fill{};
    ChartLineFormat line{};
    ChartTextStyle textStyle{};
};

struct ChartUpDownBars {
    bool present{false};
    int gapWidth{150};
    ChartFillFormat upFill{};
    ChartLineFormat upLine{};
    ChartFillFormat downFill{};
    ChartLineFormat downLine{};
};

struct ChartProjectedPieOptions {
    bool present{false};
    std::string ofPieType{"pie"}; // pie or bar
    int gapWidth{150};
    std::string splitType{"auto"}; // auto, cust, percent, pos, val
    bool hasSplitPosition{false};
    double splitPosition{0.0};
    std::vector<int> customSplitPoints;
    int secondPlotSize{75};
    bool hasSeriesLines{false};
    ChartLineFormat seriesLinesFormat{};
};

struct ChartManualLayout {
    bool present{false};
    std::string target;
    std::string xMode;
    std::string yMode;
    std::string widthMode;
    std::string heightMode;
    bool hasX{false};
    bool hasY{false};
    bool hasWidth{false};
    bool hasHeight{false};
    double x{0.0};
    double y{0.0};
    double width{0.0};
    double height{0.0};
};

struct ChartLegendFormat {
    bool present{false};
    bool overlay{false};
    ChartManualLayout layout{};
    ChartFillFormat fill{};
    ChartLineFormat line{};
};

struct ChartAxisScaling {
    bool hasMinimum{false};
    bool hasMaximum{false};
    bool hasLogBase{false};
    double minimum{0.0};
    double maximum{0.0};
    double logBase{0.0};
    bool reverseOrder{false};
};

struct ChartDisplayUnits {
    bool present{false};
    std::string builtInUnit;
    bool hasCustomUnit{false};
    double customUnit{0.0};
    bool showLabel{false};
    ChartRichText labelRichText{};
};

struct ChartDataPointFormat {
    std::size_t index{0};
    ChartFillFormat fill{};
    ChartLineFormat line{};
    ChartMarkerFormat marker{};
};

struct ChartView3D {
    bool present{false};
    bool hasRotationX{false};
    bool hasRotationY{false};
    bool hasHeightPercent{false};
    bool hasDepthPercent{false};
    bool hasRightAngleAxes{false};
    bool hasPerspective{false};
    int rotationX{0};
    int rotationY{0};
    int heightPercent{100};
    int depthPercent{100};
    bool rightAngleAxes{true};
    int perspective{30};
};

struct ChartWallFormat {
    bool present{false};
    bool hasThickness{false};
    int thickness{0};
    ChartFillFormat fill{};
    ChartLineFormat line{};
};

class ChartSeries {
    friend class Worksheet;
public:
    enum class TrendlineType { Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage };
    enum class ErrorBarDirection { X, Y };
    enum class ErrorBarType { Both, Plus, Minus };
    enum class ErrorValueType { FixedValue, Percentage, StandardDeviation, StandardError, Custom };

    struct Trendline {
        TrendlineType type{TrendlineType::Linear};
        int order{2};
        int period{2};
        double forward{0.0};
        double backward{0.0};
        bool displayEquation{false};
        bool displayRSquared{false};
        ChartLineFormat lineFormat{};
    };

    struct ErrorBars {
        ErrorBarDirection direction{ErrorBarDirection::Y};
        ErrorBarType barType{ErrorBarType::Both};
        ErrorValueType valueType{ErrorValueType::FixedValue};
        double value{0.0};
        bool noEndCap{false};
        std::string plusReference;
        std::string minusReference;
        ChartLineFormat lineFormat{};
    };
    ChartSeries() = default;
    explicit ChartSeries(std::string title) : title_(std::move(title)) {}

    const std::string& title() const noexcept { return title_; }
    void setTitle(std::string v) { title_ = std::move(v); }
    const std::string& valuesReference() const noexcept { return valuesReference_; }
    void setValuesReference(std::string v) { valuesReference_ = std::move(v); }
    const std::string& categoriesReference() const noexcept { return categoriesReference_; }
    void setCategoriesReference(std::string v) { categoriesReference_ = std::move(v); }
    const std::string& titleReference() const noexcept { return titleReference_; }
    void setTitleReference(std::string v) { titleReference_ = std::move(v); }
    const ChartSeriesCache& titleCache() const noexcept { return titleCache_; }
    const ChartSeriesCache& categoriesCache() const noexcept { return categoriesCache_; }
    const ChartSeriesCache& valuesCache() const noexcept { return valuesCache_; }
    const ChartSeriesCache& bubbleSizeCache() const noexcept { return bubbleSizeCache_; }
    void setTitleCache(ChartSeriesCache v) { titleCache_ = std::move(v); }
    void setCategoriesCache(ChartSeriesCache v) { categoriesCache_ = std::move(v); }
    void setValuesCache(ChartSeriesCache v) { valuesCache_ = std::move(v); }
    void setBubbleSizeCache(ChartSeriesCache v) { bubbleSizeCache_ = std::move(v); }
    const std::string& bubbleSizeReference() const noexcept { return bubbleSizeReference_; }
    void setBubbleSizeReference(std::string v) { bubbleSizeReference_ = std::move(v); }
    bool smooth() const noexcept { return smooth_; }
    bool hasSmooth() const noexcept { return hasSmooth_; }
    void setSmooth(bool v) noexcept { smooth_ = v; hasSmooth_ = true; }
    void clearSmooth() noexcept { hasSmooth_ = false; smooth_ = false; }
    const std::vector<Trendline>& trendlines() const noexcept { return trendlines_; }
    const std::vector<ErrorBars>& errorBars() const noexcept { return errorBars_; }
    const ChartDataLabels& dataLabels() const noexcept { return dataLabels_; }
    const ChartLineFormat& lineFormat() const noexcept { return lineFormat_; }
    const ChartFillFormat& fillFormat() const noexcept { return fillFormat_; }
    const ChartMarkerFormat& markerFormat() const noexcept { return markerFormat_; }
    const std::vector<ChartDataPointFormat>& dataPoints() const noexcept { return dataPoints_; }
    const ChartDataPointFormat* dataPoint(std::size_t pointIndex) const noexcept {
        const auto it = std::find_if(dataPoints_.begin(), dataPoints_.end(), [&](const auto& point) { return point.index == pointIndex; });
        return it == dataPoints_.end() ? nullptr : &*it;
    }
    void setTrendlines(std::vector<Trendline> value) { trendlines_ = std::move(value); }
    void setErrorBars(std::vector<ErrorBars> value) { errorBars_ = std::move(value); }
    void setDataLabels(ChartDataLabels value) { dataLabels_ = std::move(value); }
    void setLineFormat(ChartLineFormat value) { lineFormat_ = std::move(value); }
    void setFillFormat(ChartFillFormat value) { fillFormat_ = std::move(value); }
    void setMarkerFormat(ChartMarkerFormat value) { markerFormat_ = std::move(value); }
    void setDataPoints(std::vector<ChartDataPointFormat> value) { dataPoints_ = std::move(value); }

    // Reference helpers: "=SheetName!$B$2:$B$10"
    void reference(std::string sheetName, std::string rangeRef) {
        valuesReference_ = "='" + sheetName + "'!" + rangeRef;
    }
    void categories(std::string sheetName, std::string rangeRef) {
        categoriesReference_ = "='" + sheetName + "'!" + rangeRef;
    }

private:
    std::string title_, titleReference_, valuesReference_, categoriesReference_, bubbleSizeReference_;
    ChartSeriesCache titleCache_{}, categoriesCache_{}, valuesCache_{}, bubbleSizeCache_{};
    bool smooth_{false};
    bool hasSmooth_{false};
    std::vector<Trendline> trendlines_;
    std::vector<ErrorBars> errorBars_;
    ChartDataLabels dataLabels_{};
    ChartLineFormat lineFormat_{};
    ChartFillFormat fillFormat_{};
    ChartMarkerFormat markerFormat_{};
    std::vector<ChartDataPointFormat> dataPoints_;
};

class Chart {
    friend class Worksheet;
public:
    // Numeric values 0..16 are kept stable for the public C ABI and existing bindings.
    enum class Type {
        Bar = 0, Line = 1, Pie = 2, Scatter = 3, Doughnut = 4, Radar = 5, Area = 6, Bubble = 7, Stock = 8,
        Bar3D = 9, Line3D = 10, Area3D = 11, Pie3D = 12, Surface = 13, Surface3D = 14,
        PieOfPie = 15, BarOfPie = 16,
        HorizontalBar = 17,
        Histogram = 18, Pareto = 19, BoxWhisker = 20, Waterfall = 21, Funnel = 22,
        Treemap = 23, Sunburst = 24, FilledMap = 25
    };
    enum class Grouping { Standard, Stacked, PercentStacked, Clustered };
    enum class AxisKind { Category, Value, Date, Series };
    enum class BarDirection { Column, Bar };
    enum class ScatterStyle { None, Line, LineMarker, Marker, Smooth, SmoothMarker };
    enum class BubbleSizeRepresents { Area, Width };

    struct Axis {
        AxisKind kind{AxisKind::Value};
        std::uint64_t id{0};
        std::uint64_t crossAxisId{0};
        std::string position;
        std::string title;
        ChartRichText titleRichText{};
        bool secondary{false};
        std::string numberFormat;
        bool numberFormatSourceLinked{true};
        std::string majorTickMark;
        std::string minorTickMark;
        std::string tickLabelPosition;
        bool hasMajorUnit{false};
        bool hasMinorUnit{false};
        double majorUnit{0.0};
        double minorUnit{0.0};
        std::string crosses;
        std::string crossBetween;
        bool hasCrossesAt{false};
        double crossesAt{0.0};
        ChartAxisScaling scaling{};
        ChartDisplayUnits displayUnits{};
        bool hasMajorGridlines{false};
        bool hasMinorGridlines{false};
        ChartLineFormat lineFormat{};
        ChartLineFormat majorGridlineFormat{};
        ChartLineFormat minorGridlineFormat{};
    };

    using DataLabels = ChartDataLabels;

    struct Plot {
        Type type{Type::Bar};
        Grouping grouping{Grouping::Standard};
        std::vector<std::uint64_t> axisIds;
        std::size_t firstSeries{0};
        std::size_t seriesCount{0};
        bool usesSecondaryAxes{false};
        DataLabels dataLabels{};
        bool hasDropLines{false};
        ChartLineFormat dropLinesFormat{};
        bool hasHighLowLines{false};
        ChartLineFormat highLowLinesFormat{};
        ChartUpDownBars upDownBars{};
        bool hasGapDepth{false};
        int gapDepth{150};
        bool hasWireframe{false};
        bool wireframe{false};
        std::string shape;
        bool hasFirstSliceAngle{false};
        int firstSliceAngle{0};
        bool hasHoleSize{false};
        int holeSize{10};
        std::string radarStyle;
        ChartProjectedPieOptions projectedPie{};
        BarDirection barDirection{BarDirection::Column};
        ScatterStyle scatterStyle{ScatterStyle::Marker};
        bool hasBubbleScale{false};
        int bubbleScale{100};
        bool showNegativeBubbles{false};
        BubbleSizeRepresents bubbleSizeRepresents{BubbleSizeRepresents::Area};
        bool bubble3D{false};
        // ChartEx options.  These are ignored by legacy ChartML types.
        double histogramBinWidth{0.0};
        int histogramBinCount{0};
        bool histogramAutomaticBins{true};
        bool histogramHasUnderflow{false};
        double histogramUnderflow{0.0};
        bool histogramHasOverflow{false};
        double histogramOverflow{0.0};
        bool boxWhiskerShowInnerPoints{true};
        bool boxWhiskerShowOutlierPoints{true};
        bool boxWhiskerShowMeanLine{false};
        bool boxWhiskerShowMeanMarker{true};
        bool boxWhiskerQuartileInclusive{false};
        bool waterfallShowConnectorLines{true};
        std::string mapProjection{"automatic"};
        std::string mapArea{"automatic"};
        std::string mapLabels{"bestFitOnly"};
    };

    Chart(Type type = Type::Bar) : type_(type) {}

    Type type() const noexcept { return type_; }
    Grouping grouping() const noexcept { return grouping_; }
    void setGrouping(Grouping v) noexcept { grouping_ = v; }

    const std::string& title() const noexcept { return title_; }
    void setTitle(std::string v) { title_ = std::move(v); }
    const ChartRichText& titleRichText() const noexcept { return titleRichText_; }
    void setTitleRichText(ChartRichText v) { titleRichText_ = std::move(v); if (titleRichText_.present) title_ = titleRichText_.plainText(); }
    const std::string& xAxisTitle() const noexcept { return xAxisTitle_; }
    void setXAxisTitle(std::string v) { xAxisTitle_ = std::move(v); }
    const std::string& yAxisTitle() const noexcept { return yAxisTitle_; }
    void setYAxisTitle(std::string v) { yAxisTitle_ = std::move(v); }

    const std::string& style() const noexcept { return style_; }
    void setStyle(std::string v) { style_ = std::move(v); }
    const ChartThemePalette& themePalette() const noexcept { return themePalette_; }
    void setThemePalette(ChartThemePalette v) { themePalette_ = std::move(v); }
    const ChartStyleResources& styleResources() const noexcept { return styleResources_; }
    void setStyleResources(ChartStyleResources v) { styleResources_ = std::move(v); }
    std::string resolveThemeBaseColor(const ChartColor& color) const { return themePalette_.resolveBase(color); }
    ChartResolvedColor resolveThemeColor(const ChartColor& color) const { return themePalette_.resolve(color); }
    std::string resolveThemeFinalRgb(const ChartColor& color) const { return themePalette_.resolveFinalRgb(color); }

    int width() const noexcept { return width_; } void setWidth(int v) noexcept { width_ = v; }
    int height() const noexcept { return height_; } void setHeight(int v) noexcept { height_ = v; }

    ChartSeries& addSeries(ChartSeries series) { series_.push_back(std::move(series)); return series_.back(); }
    const StableVector<ChartSeries>& series() const noexcept { return series_; }
    StableVector<ChartSeries>& series() noexcept { return series_; }

    bool showLegend() const noexcept { return showLegend_; }
    void setShowLegend(bool v) noexcept { showLegend_ = v; }
    const std::string& legendPosition() const noexcept { return legendPosition_; }
    void setLegendPosition(std::string v) { legendPosition_ = std::move(v); }
    const ChartLegendFormat& legendFormat() const noexcept { return legendFormat_; }
    void setLegendFormat(ChartLegendFormat v) { legendFormat_ = std::move(v); }
    const ChartManualLayout& plotAreaLayout() const noexcept { return plotAreaLayout_; }
    void setPlotAreaLayout(ChartManualLayout v) { plotAreaLayout_ = std::move(v); }
    const ChartFillFormat& chartAreaFillFormat() const noexcept { return chartAreaFillFormat_; }
    const ChartLineFormat& chartAreaLineFormat() const noexcept { return chartAreaLineFormat_; }
    const ChartFillFormat& plotAreaFillFormat() const noexcept { return plotAreaFillFormat_; }
    const ChartLineFormat& plotAreaLineFormat() const noexcept { return plotAreaLineFormat_; }
    void setChartAreaFillFormat(ChartFillFormat v) { chartAreaFillFormat_ = std::move(v); }
    void setChartAreaLineFormat(ChartLineFormat v) { chartAreaLineFormat_ = std::move(v); }
    void setPlotAreaFillFormat(ChartFillFormat v) { plotAreaFillFormat_ = std::move(v); }
    void setPlotAreaLineFormat(ChartLineFormat v) { plotAreaLineFormat_ = std::move(v); }
    const ChartDataTable& dataTable() const noexcept { return dataTable_; }
    void setDataTable(ChartDataTable v) { dataTable_ = std::move(v); }
    const ChartView3D& view3D() const noexcept { return view3D_; }
    void setView3D(ChartView3D v) { view3D_ = std::move(v); }
    const ChartWallFormat& floorFormat() const noexcept { return floorFormat_; }
    const ChartWallFormat& sideWallFormat() const noexcept { return sideWallFormat_; }
    const ChartWallFormat& backWallFormat() const noexcept { return backWallFormat_; }
    void setFloorFormat(ChartWallFormat v) { floorFormat_ = std::move(v); }
    void setSideWallFormat(ChartWallFormat v) { sideWallFormat_ = std::move(v); }
    void setBackWallFormat(ChartWallFormat v) { backWallFormat_ = std::move(v); }

    // Read-only structural metadata for imported charts. A combined chart has
    // more than one plot (for example bar + line). Axis IDs are the native
    // OOXML c:axId values and remain stable across selective edits.
    const std::vector<Axis>& axes() const noexcept { return axes_; }
    const std::vector<Plot>& plots() const noexcept { return plots_; }
    std::vector<Plot>& plots() noexcept { return plots_; }
    Plot& addPlot(Type type, std::size_t firstSeries, std::size_t seriesCount, bool secondaryAxes = false) {
        Plot plot;
        plot.type = type;
        plot.firstSeries = firstSeries;
        plot.seriesCount = seriesCount;
        plot.usesSecondaryAxes = secondaryAxes;
        if (type == Type::HorizontalBar) plot.barDirection = BarDirection::Bar;
        plots_.push_back(std::move(plot));
        return plots_.back();
    }
    Plot& primaryPlot() {
        if (plots_.empty()) {
            Plot plot;
            plot.type = type_;
            plot.grouping = grouping_;
            if (type_ == Type::HorizontalBar) plot.barDirection = BarDirection::Bar;
            plots_.push_back(std::move(plot));
        }
        return plots_.front();
    }
    const Plot* primaryPlotOrNull() const noexcept { return plots_.empty() ? nullptr : &plots_.front(); }
    bool combined() const noexcept { return plots_.size() > 1; }
    bool modern() const noexcept { return isModernType(type_); }
    static bool isModernType(Type type) noexcept {
        return type == Type::Histogram || type == Type::Pareto || type == Type::BoxWhisker ||
               type == Type::Waterfall || type == Type::Funnel || type == Type::Treemap ||
               type == Type::Sunburst || type == Type::FilledMap;
    }
    bool hasSecondaryAxes() const noexcept {
        return std::any_of(axes_.begin(), axes_.end(), [](const Axis& axis) { return axis.secondary; });
    }
    std::uint64_t primaryXAxisId() const noexcept { return primaryXAxisId_; }
    std::uint64_t primaryYAxisId() const noexcept { return primaryYAxisId_; }
    const Axis* axisById(std::uint64_t axisId) const noexcept {
        const auto it = std::find_if(axes_.begin(), axes_.end(), [&](const Axis& axis) { return axis.id == axisId; });
        return it == axes_.end() ? nullptr : &*it;
    }

    // Package-origin metadata is populated for charts discovered in an existing
    // worksheet drawing.  It allows safe selective edits while unsupported
    // chart XML remains preserved verbatim in the package.
    const DrawingAnchorInfo& anchorInfo() const noexcept { return anchorInfo_; }
    const std::string& stableId() const noexcept { return stableId_; }
    const std::string& sourceDrawingPart() const noexcept { return sourceDrawingPart_; }
    const std::string& sourceChartPart() const noexcept { return sourceChartPart_; }
    const std::string& sourceRelationshipId() const noexcept { return sourceRelationshipId_; }
    const std::string& drawingObjectName() const noexcept { return drawingObjectName_; }
    bool imported() const noexcept { return imported_; }

    // Loader-facing setters. Normal callers generally use the inspection
    // accessors above and Worksheet's stable-ID mutation APIs.
    void setAnchorInfo(DrawingAnchorInfo value) noexcept { anchorInfo_ = std::move(value); }
    void setStableId(std::string value) { stableId_ = std::move(value); }
    void setSourceDrawingPart(std::string value) { sourceDrawingPart_ = std::move(value); }
    void setSourceChartPart(std::string value) { sourceChartPart_ = std::move(value); }
    void setSourceRelationshipId(std::string value) { sourceRelationshipId_ = std::move(value); }
    void setDrawingObjectName(std::string value) { drawingObjectName_ = std::move(value); }
    void setImported(bool value) noexcept { imported_ = value; }
    void setAxes(std::vector<Axis> value) { axes_ = std::move(value); }
    void setPlots(std::vector<Plot> value) { plots_ = std::move(value); }
    void setPrimaryAxisIds(std::uint64_t xAxisId, std::uint64_t yAxisId) noexcept {
        primaryXAxisId_ = xAxisId; primaryYAxisId_ = yAxisId;
    }

    static std::string typeName(Type type, Grouping grouping = Grouping::Standard);

private:
    Type type_{Type::Bar};
    Grouping grouping_{Grouping::Standard};
    std::string title_, xAxisTitle_, yAxisTitle_;
    ChartRichText titleRichText_{};
    std::string style_{"2"};
    ChartThemePalette themePalette_{};
    ChartStyleResources styleResources_{};
    int width_{600}, height_{400};
    StableVector<ChartSeries> series_;
    bool showLegend_{true};
    std::string legendPosition_{"r"};
    ChartLegendFormat legendFormat_{};
    ChartManualLayout plotAreaLayout_{};
    ChartFillFormat chartAreaFillFormat_{};
    ChartLineFormat chartAreaLineFormat_{};
    ChartFillFormat plotAreaFillFormat_{};
    ChartLineFormat plotAreaLineFormat_{};
    ChartDataTable dataTable_{};
    ChartView3D view3D_{};
    ChartWallFormat floorFormat_{};
    ChartWallFormat sideWallFormat_{};
    ChartWallFormat backWallFormat_{};
    DrawingAnchorInfo anchorInfo_{};
    std::string stableId_;
    std::string sourceDrawingPart_;
    std::string sourceChartPart_;
    std::string sourceRelationshipId_;
    std::string drawingObjectName_{"Chart"};
    bool imported_{false};
    std::vector<Axis> axes_;
    std::vector<Plot> plots_;
    std::uint64_t primaryXAxisId_{0};
    std::uint64_t primaryYAxisId_{0};
};

inline std::string Chart::typeName(Chart::Type type, Chart::Grouping grouping) {
    switch (type) {
    case Type::Bar:        return grouping == Grouping::Stacked ? "barStacked" : (grouping == Grouping::PercentStacked ? "barPercentStacked" : "barChart");
    case Type::Line:       return grouping == Grouping::Stacked ? "lineStacked" : "lineChart";
    case Type::Pie:        return "pieChart";
    case Type::Scatter:    return "scatterChart";
    case Type::Doughnut:   return "doughnutChart";
    case Type::Radar:      return "radarChart";
    case Type::Area:       return grouping == Grouping::Stacked ? "areaStacked" : "areaChart";
    case Type::Bubble:     return "bubbleChart";
    case Type::Stock:      return "stockChart";
    case Type::PieOfPie:   return "ofPieChart";
    case Type::BarOfPie:   return "ofPieChart";
    case Type::HorizontalBar: return "barChart";
    case Type::Bar3D:      return "bar3DChart";
    case Type::Line3D:     return "line3DChart";
    case Type::Area3D:     return "area3DChart";
    case Type::Pie3D:      return "pie3DChart";
    case Type::Surface:    return "surfaceChart";
    case Type::Surface3D:  return "surface3DChart";
    case Type::Histogram:  return "histogram";
    case Type::Pareto:     return "pareto";
    case Type::BoxWhisker: return "boxWhisker";
    case Type::Waterfall:  return "waterfall";
    case Type::Funnel:     return "funnel";
    case Type::Treemap:    return "treemap";
    case Type::Sunburst:   return "sunburst";
    case Type::FilledMap:  return "regionMap";
    }
    return "barChart";
}

}
