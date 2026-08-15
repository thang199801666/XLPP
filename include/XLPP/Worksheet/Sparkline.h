#pragma once

#include <string>
#include <vector>

namespace xlpp {

// One inline sparkline: a source data range mapped onto a target cell.
struct Sparkline {
    // Source data reference, e.g. "'Data'!A2:A10".
    std::string reference;
    // Target anchor cell, e.g. "B2".
    std::string location;
};

// Grouping container for a set of sparklines that share display settings.
// XL++ models the common x14:sparklineGroup attributes; untouched XML is
// preserved through sparklinesRawXml when the group is not edited.
struct SparklineGroup {
    // Group type: "line", "column" or "stacked".
    std::string type{"line"};
    // Line style: "smooth", "straight" or empty for default.
    std::string lineStyle;
    // Colors (ARGB) for markers/negative/axis.
    std::string markersColor;
    std::string negativeColor;
    std::string axisColor;
    // Display flags.
    bool displayHidden{false};
    bool displayXAxis{false};
    bool displayMarkers{false};
    bool high{false};
    bool low{false};
    bool first{false};
    bool last{false};
    bool negative{false};
    bool colorSeries{false};
    bool colorAxis{false};
    bool colorMarkers{false};
    bool colorFirst{false};
    bool colorLast{false};
    bool colorHigh{false};
    bool colorLow{false};
    // Right-to-left direction.
    bool rightToLeft{false};
    // Date axis reference (optional).
    std::string dateAxis;
    std::vector<Sparkline> sparklines;
    // Raw x14:sparklineGroup subtree for byte-preserving carry-through.
    std::string rawXml;
    bool empty() const noexcept {
        return sparklines.empty() && rawXml.empty();
    }
};

} // namespace xlpp
