#pragma once
#include <string>
#include <vector>
#include <memory>

namespace xlpp {

class ChartSeries {
public:
    ChartSeries() = default;
    explicit ChartSeries(std::string title) : title_(std::move(title)) {}

    const std::string& title() const noexcept { return title_; }
    void setTitle(std::string v) { title_ = std::move(v); }
    const std::string& valuesReference() const noexcept { return valuesReference_; }
    void setValuesReference(std::string v) { valuesReference_ = std::move(v); }
    const std::string& categoriesReference() const noexcept { return categoriesReference_; }
    void setCategoriesReference(std::string v) { categoriesReference_ = std::move(v); }

    // Reference helpers: "=SheetName!$B$2:$B$10"
    void reference(std::string sheetName, std::string rangeRef) {
        valuesReference_ = "='" + sheetName + "'!" + rangeRef;
    }
    void categories(std::string sheetName, std::string rangeRef) {
        categoriesReference_ = "='" + sheetName + "'!" + rangeRef;
    }

private:
    std::string title_, valuesReference_, categoriesReference_;
};

class Chart {
public:
    enum class Type { Bar, Line, Pie, Scatter, Doughnut, Radar, Area, Bubble };
    enum class Grouping { Standard, Stacked, PercentStacked, Clustered };

    Chart(Type type = Type::Bar) : type_(type) {}

    Type type() const noexcept { return type_; }
    Grouping grouping() const noexcept { return grouping_; }
    void setGrouping(Grouping v) noexcept { grouping_ = v; }

    const std::string& title() const noexcept { return title_; }
    void setTitle(std::string v) { title_ = std::move(v); }
    const std::string& xAxisTitle() const noexcept { return xAxisTitle_; }
    void setXAxisTitle(std::string v) { xAxisTitle_ = std::move(v); }
    const std::string& yAxisTitle() const noexcept { return yAxisTitle_; }
    void setYAxisTitle(std::string v) { yAxisTitle_ = std::move(v); }

    const std::string& style() const noexcept { return style_; }
    void setStyle(std::string v) { style_ = std::move(v); }

    int width() const noexcept { return width_; } void setWidth(int v) noexcept { width_ = v; }
    int height() const noexcept { return height_; } void setHeight(int v) noexcept { height_ = v; }

    ChartSeries& addSeries(ChartSeries series) { series_.push_back(std::move(series)); return series_.back(); }
    const std::vector<ChartSeries>& series() const noexcept { return series_; }
    std::vector<ChartSeries>& series() noexcept { return series_; }

    bool showLegend() const noexcept { return showLegend_; }
    void setShowLegend(bool v) noexcept { showLegend_ = v; }
    const std::string& legendPosition() const noexcept { return legendPosition_; }
    void setLegendPosition(std::string v) { legendPosition_ = std::move(v); }

    static std::string typeName(Type type, Grouping grouping = Grouping::Standard);

private:
    Type type_{Type::Bar};
    Grouping grouping_{Grouping::Standard};
    std::string title_, xAxisTitle_, yAxisTitle_;
    std::string style_{"2"};
    int width_{600}, height_{400};
    std::vector<ChartSeries> series_;
    bool showLegend_{true};
    std::string legendPosition_{"r"};
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
    }
    return "barChart";
}

}
