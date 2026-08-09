#pragma once
#include <cstddef>
#include <optional>

namespace xlpp {
struct RowDimension {
    std::optional<double> height;
    bool hidden{false};
    int outlineLevel{0};
    bool collapsed{false};
};

struct ColumnDimension {
    std::optional<double> width;
    bool hidden{false};
    bool bestFit{false};
    int outlineLevel{0};
    bool collapsed{false};
};

struct WorksheetExtents {
    std::size_t minRow{0};
    std::size_t minColumn{0};
    std::size_t maxRow{0};
    std::size_t maxColumn{0};
};
}
