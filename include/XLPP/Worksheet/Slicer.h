#pragma once
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

// One selectable value in a slicer cache.
struct SlicerItem {
    std::string value;
    bool selected{true};
};

// Slicer (filter control) bound to a pivot-table field. A slicer cache
// (xl/slicerCaches/slicerCacheN.xml) holds the unique field values while the
// per-worksheet slicer part (xl/slicers/slicerN.xml) holds the visual layout.
struct Slicer {
    std::string name;          // unique slicer name, e.g. "Slicer_City"
    std::string caption;       // visible caption
    std::string worksheetName; // worksheet that hosts the slicer control
    std::string sourceName;    // pivot field name that drives the filter
    std::string pivotTableName;// owning pivot table
    std::string style{"SlicerStyleLight2"};
    int columnCount{1};
    std::vector<SlicerItem> items;
};

} // namespace xlpp
