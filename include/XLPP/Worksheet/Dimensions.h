#pragma once
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
}
