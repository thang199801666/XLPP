# XL++ API Reference

XL++ is a C++20 library for reading and writing Excel `.xlsx` workbooks.

## Quick Example

```cpp
#include <XLPP/XLPP.h>

int main() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Report");
    sheet.cell("A1").setValue("Hello");
    sheet.cell("B1").setValue(42.0);
    workbook.save("report.xlsx");
}
```

## Public API

- `xlpp::Workbook` manages workbook files and worksheets.
- `xlpp::Worksheet` provides cells, ranges, formatting, tables, and layout.
- `xlpp::Cell` stores values, formulas, styles, comments, and hyperlinks.
- Streaming APIs support large workbooks with bounded memory usage.

The public headers are located in `include/XLPP`.

## Python

See the complete [Python guide](python.md) for installation, workbook usage,
formatting, tables, charts, NumPy helpers, and troubleshooting.
