# XL++ Refinement Roadmap
## Post P0Z-H / v1.11.0

### 1. Mục tiêu chiến lược

Mục tiêu của giai đoạn tiếp theo không còn là tăng số lượng API hoặc tăng LOC.

XL++ cần tiến từ:

**“C++ Excel library có feature coverage rất rộng”**

sang:

**“Production-grade Excel engine có khả năng chỉnh sửa workbook thực tế với độ tin cậy cao, preservation tốt và compatibility được xác nhận trực tiếp bằng Microsoft Excel.”**

Các mục tiêu chính:

1. Tăng độ tin cậy khi round-trip workbook được tạo bởi Excel.
2. Không làm mất các feature mà XL++ chưa hiểu khi chỉnh sửa phần khác của workbook.
3. Hoàn thiện Pivot, VBA và Drawing ecosystems.
4. Nâng Formula engine từ mức mạnh lên mức có kiến trúc calculation engine lâu dài.
5. Đưa Python/C#/C ABI thành SDK thực sự.
6. Xây dựng interoperability test infrastructure với Excel Desktop.
7. Loại bỏ technical debt và tài liệu lịch sử không còn phản ánh code hiện tại.

---

# Phase 21 — Documentation & Capability Baseline

## Mục tiêu

Tạo một baseline duy nhất phản ánh chính xác trạng thái source hiện tại trước khi mở rộng tiếp.

## 21.1 Current capability specification

Tạo:

`docs/CURRENT_CAPABILITIES.md`

Phân loại từng feature thành:

- Fully supported
- Supported with limitations
- Preservation-only
- Experimental
- Unsupported

Các nhóm:

- Workbook
- Worksheet
- Cells
- Styles
- Tables
- Formulas
- Charts
- Pivot
- Images
- DrawingML
- VBA
- Encryption
- Streaming
- External links
- Comments
- Validation
- Bindings

Không được dùng các % coverage chung chung nếu không có metric cụ thể.

## 21.2 Retire roadmap lịch sử

Di chuyển các roadmap cũ vào:

`docs/history/`

Ví dụ:

`MISSING_FEATURES_AND_DEVELOPMENT_ROADMAP.md`

không còn được coi là nguồn trạng thái hiện tại.

## 21.3 Feature capability manifest

Tạo machine-readable:

`docs/capabilities.json`

Ví dụ:

```json
{
  "pivot": {
    "basic": "supported",
    "grouping": "supported",
    "pivotChart": "unsupported",
    "slicer": "unsupported",
    "olap": "preserve-only"
  }
}
```

Dùng manifest này cho:

- documentation;
- tests;
- bindings;
- release notes.

## Exit criteria

- Không còn tài liệu active nào mâu thuẫn với source.
- Capability matrix được test trong CI.
- Version và feature manifest đồng bộ.

---

# Phase 22 — Microsoft Excel Interoperability Harness

Đây là phase ưu tiên cao nhất.

## 22.1 Excel Desktop automation runner

Tạo Windows-only interoperability runner:

`tests/excel_interop/`

Sử dụng Excel COM Automation.

Workflow:

```text
XL++ generate/edit
        ↓
Excel.Application.Open
        ↓
detect repair/recovery
        ↓
Save
        ↓
Close
        ↓
Reopen
        ↓
inspect workbook
```

## 22.2 Detect Excel repair

Phát hiện:

- file repaired;
- removed records;
- repaired records;
- unreadable content;
- XML corruption;
- chart repair;
- pivot repair;
- VBA corruption.

Thu thập recovery logs nếu Excel tạo.

## 22.3 Workbook semantic inspection

COM runner kiểm tra:

- worksheet count;
- formulas;
- tables;
- chart type;
- pivot existence;
- VBA module count;
- names;
- shapes;
- external links.

## 22.4 Excel-created corpus

Tạo corpus workbook do Excel Desktop sinh:

```text
tests/corpus/excel/
    charts/
    pivot/
    formulas/
    styles/
    vba/
    drawings/
    external-links/
    power-query/
    slicers/
    enterprise/
```

Mỗi workbook có manifest mô tả expected behavior.

## Exit criteria

Ít nhất:

- 100 Excel-created workbooks.
- 0 unexpected repair dialogs.
- 0 unexpected removed records.
- CI report cho từng workbook.

---

# Phase 23 — Preservation Engine 2.0

Mục tiêu: chỉnh sửa một cell không được phá feature không liên quan.

## 23.1 Unknown relationship preservation

Audit toàn bộ package relationships.

Mọi relationship chưa hiểu phải được:

- giữ nguyên target;
- giữ nguyên relationship type;
- giữ nguyên part;
- không đổi ID không cần thiết.

## 23.2 Unknown XML extension preservation

Preserve:

- `extLst`
- namespace extensions
- Office 2010+
- Office 2013+
- Office 2016+
- Microsoft 365 extension nodes.

Đặc biệt:

- charts;
- pivot;
- conditional formatting;
- data validation;
- drawings.

## 23.3 Content type preservation

Không regenerate `[Content_Types].xml` theo cách làm mất unknown overrides/defaults.

## 23.4 Stable part topology

Không renumber part nếu không cần.

Ví dụ:

```text
chart7.xml
pivotCacheDefinition4.xml
drawing12.xml
```

nên giữ identity khi chỉnh sửa unrelated content.

## 23.5 Preservation regression

Các test dạng:

```text
Excel workbook
→ checksum unknown parts
→ XL++ edit A1
→ save
→ compare preserved parts
```

## Exit criteria

Unrelated edits giữ nguyên byte-for-byte càng nhiều unknown parts càng tốt.

---

# Phase 24 — Pivot Engine 2.0

Pivot trở thành subsystem lớn riêng.

## 24.1 Calculated Fields

Hỗ trợ:

- create;
- remove;
- formula;
- caption;
- number format;
- round-trip.

## 24.2 Calculated Items

Hỗ trợ calculated item trong PivotField.

Phải validate các trường hợp Excel cấm.

## 24.3 Advanced field behaviors

Hoàn thiện:

- AutoShow;
- Top N;
- manual item ordering;
- custom lists;
- multiple page fields;
- show items with no data;
- missing item retention;
- retain deleted items;
- subtotal ordering.

## 24.4 Pivot formatting

Hỗ trợ:

- pivot formats;
- field-level format;
- subtotal format;
- grand-total format;
- label/data conditional formats.

## 24.5 PivotChart

Tạo model riêng:

`PivotChart`

Không giả lập bằng normal chart.

Hỗ trợ:

- chart ↔ pivot table relationship;
- field buttons;
- category/value mapping;
- cache linkage.

## 24.6 Slicers

Triển khai:

- slicer cache;
- slicer;
- worksheet shape;
- pivot connection;
- item selection.

## 24.7 Timelines

Sau slicer:

- timeline cache;
- timeline state;
- Pivot connection.

## Exit criteria

Pivot non-OLAP đạt mức gần Excel Desktop cho use case thông thường.

---

# Phase 25 — VBA Engine 2.0

## 25.1 References

Parse/write VBA references:

- registered references;
- project references;
- control references;
- original typelib metadata.

Public API:

```cpp
project.references()
project.addReference(...)
project.removeReference(...)
```

## 25.2 UserForm parser

Đọc UserForm:

- designer stream;
- controls;
- properties;
- module code.

Không rebuild ngay ở bước đầu.

Mode đầu tiên:

**read + preserve + inspect**.

## 25.3 FRX support

Parse và preserve binary resource:

- images;
- icons;
- list data;
- control resources.

## 25.4 UserForm authoring

Sau parser ổn định:

- Frame
- Label
- TextBox
- CommandButton
- CheckBox
- ComboBox
- ListBox

Không cần implement toàn bộ MSForms ngay.

## 25.5 VBA protection

Đầu tiên:

- detect locked project;
- preserve password/protection records.

Sau đó mới cân nhắc editing.

## 25.6 Digital signature

Tách subsystem:

`VbaSignature`

Mục tiêu đầu tiên:

- detect;
- inspect certificate metadata;
- preserve.

Sau đó:

- remove explicitly;
- re-sign khi có certificate.

Không âm thầm phá signature.

## 25.7 Codepage/Unicode

Corpus nhiều locale:

- Vietnamese;
- Japanese;
- Chinese;
- Korean;
- Cyrillic;
- Western European.

## Exit criteria

XL++ có thể inspect và preserve phần lớn project VBA thực tế, bao gồm UserForms.

---

# Phase 26 — Chart Fidelity 2.0

Chart type coverage hiện đã đủ rộng.

Phase này tập trung vào fidelity.

## 26.1 Chart imported model

Reader phải giữ:

- unknown chart nodes;
- custom layout;
- formatting;
- extension lists;
- embedded theme data.

## 26.2 Rich formatting

Bổ sung:

- gradient fill;
- pattern fill;
- custom line dash;
- marker formatting;
- text formatting;
- rich title;
- rich axis title;
- data-label formatting.

## 26.3 Axis completeness

Hoàn thiện:

- log axis;
- date axis;
- crossing options;
- custom units;
- display units;
- reverse order;
- secondary category axis;
- multi-level categories.

## 26.4 Trendlines

Đầy đủ:

- linear;
- exponential;
- logarithmic;
- polynomial;
- power;
- moving average;
- equation display;
- R²;
- forecast.

## 26.5 Error bars

Hoàn thiện:

- fixed;
- percentage;
- standard deviation;
- standard error;
- custom positive/negative.

## 26.6 Chart template fidelity

Không nhất thiết hỗ trợ `.crtx` ngay, nhưng internal model phải đủ để map các style phổ biến.

## Exit criteria

Imported Excel chart → edit title/series → save phải gần như giữ nguyên visual appearance.

---

# Phase 27 — DrawingML & Shapes

Đây hiện là một trong những gap lớn.

## 27.1 Shape model

Tạo:

```cpp
Shape
TextBox
Connector
GroupShape
```

## 27.2 Preset geometry

Hỗ trợ common shapes:

- rectangle;
- round rectangle;
- ellipse;
- line;
- arrow;
- triangle;
- diamond;
- callout.

## 27.3 TextBox

Hỗ trợ:

- rich text;
- paragraphs;
- alignment;
- margins;
- rotation;
- vertical alignment.

## 27.4 Shape formatting

- fill;
- gradient;
- outline;
- transparency;
- shadow;
- rotation.

## 27.5 Group shapes

Preserve coordinate transforms.

## 27.6 SmartArt

Không cố semantic authoring ngay.

Target:

**perfect preservation first**.

## Exit criteria

Common business workbook chứa shapes/textboxes có thể được chỉnh sửa mà không bị mất drawing.

---

# Phase 28 — Formula Engine 2.0

Không chạy theo số lượng function vô hạn.

Tập trung vào kiến trúc.

## 28.1 Function registry

Chuẩn hóa:

```cpp
FunctionRegistry
FunctionDescriptor
FunctionArguments
EvaluationContext
```

Tránh switch/dispatch ngày càng lớn.

## 28.2 Value system

Chuẩn hóa:

```text
Blank
Number
String
Boolean
Error
Array
Reference
Lambda
```

## 28.3 Reference semantics

Phân biệt rõ:

- scalar;
- array;
- reference;
- range;
- implicit intersection.

## 28.4 Dynamic arrays

Hoàn thiện:

- spill dependency;
- blocked spill;
- resize spill;
- chained dynamic arrays.

## 28.5 LET/LAMBDA

Tiếp tục từ LET tới:

- LAMBDA;
- named LAMBDA;
- MAP;
- REDUCE;
- SCAN;
- BYROW;
- BYCOL.

## 28.6 Calculation modes

- automatic;
- manual;
- automatic except data tables.

## 28.7 Dirty graph

Chỉ evaluate dependency bị thay đổi.

## 28.8 Multithreaded calculation

Sau dependency graph ổn định:

- DAG scheduling;
- worker pool;
- deterministic results.

## 28.9 Formula compatibility corpus

Corpus từ Excel với expected cached value.

## Exit criteria

Calculation engine có architecture scale được thay vì đơn thuần tăng function count.

---

# Phase 29 — Common Excel Feature Completion

Đây là nhóm feature nhỏ nhưng ảnh hưởng lớn tới usability.

## 29.1 AutoFilter

Hoàn thiện:

- custom filters;
- dynamic filters;
- color filters;
- icon filters;
- Top10;
- date groups.

## 29.2 Conditional Formatting

Hoàn thiện:

- color scales;
- data bars;
- icon sets;
- formulas;
- priority;
- stopIfTrue;
- Office extensions.

## 29.3 Data Validation

- list;
- custom;
- date/time;
- decimal;
- whole;
- text length;
- formula1/formula2;
- prompt/error messages.

## 29.4 Comments

Legacy comments hoàn thiện.

Sau đó:

**Threaded comments / modern comments**.

## 29.5 Workbook views

- active sheet;
- selected tabs;
- first visible sheet;
- window position;
- zoom;
- split/freeze;
- workbook window state.

## 29.6 Print/page layout

- page breaks;
- repeating rows/columns;
- print area;
- print titles;
- header/footer;
- margins;
- paper size;
- orientation.

## Exit criteria

Common business workbook không còn gặp những thiếu hụt nhỏ nhưng khó chịu.

---

# Phase 30 — External Data & Connections

## 30.1 External workbook links

Hoàn thiện:

- externalBook;
- sheet names;
- defined names;
- cached values.

## 30.2 Connections

Parse/preserve:

- workbook connections;
- ODBC;
- OLE DB;
- text/web connection metadata.

Mode đầu tiên:

**read + preserve**.

## 30.3 Query tables

Hỗ trợ relationship giữa:

```text
worksheet
↔ queryTable
↔ connection
```

## 30.4 Power Query

Không attempt full M engine.

Target đầu tiên:

- detect;
- inspect metadata;
- preserve;
- expose raw package parts.

## Exit criteria

XL++ có thể safely edit workbook chứa enterprise data connections.

---

# Phase 31 — Data Model / OLAP Research Track

Không gộp trực tiếp vào core roadmap.

Tạo research subsystem riêng.

## 31.1 Detect Data Model

Nhận diện:

- model relationships;
- embedded model parts;
- pivot caches sử dụng OLAP.

## 31.2 Preserve perfectly

Đầu tiên không chỉnh sửa.

Goal:

```text
edit normal worksheet
→ Data Model remains valid
```

## 31.3 Metadata inspection

Expose:

- tables;
- columns;
- relationships;
- measures nếu đọc được.

## 31.4 Decision gate

Chỉ triển khai authoring Data Model nếu:

- format hiểu đủ;
- có corpus;
- business value hợp lý.

Không để phase này kéo chậm core library.

---

# Phase 32 — API & SDK Refinement

## 32.1 C++ API consistency

Audit naming:

```text
getX()
setX()
x()
hasX()
removeX()
```

Chọn một convention nhất quán.

## 32.2 Error model

Chuẩn hóa:

```cpp
ErrorCode
Exception
ValidationIssue
Warning
```

Không throw các exception generic không cần thiết.

## 32.3 Handles

Đánh giá tất cả public handles:

- Workbook
- Worksheet
- Cell
- Range
- Table
- Chart
- Pivot

Đảm bảo semantics sau:

- copy;
- move;
- deletion;
- workbook destruction.

## 32.4 API deprecation policy

Từ đây trở đi:

- API breaking → major version;
- old API có `[[deprecated]]`;
- migration note.

## 32.5 C ABI

Đóng ABI policy chính thức.

Không sửa size struct cũ.

Dùng:

- versioned struct;
- additive functions;
- capability query.

## Exit criteria

Public API đủ ổn định để bên thứ ba phụ thuộc lâu dài.

---

# Phase 33 — Python SDK

## 33.1 Binary wheel

Build wheel cho:

### Windows
- x64

### Linux
- manylinux x86_64

### macOS
- arm64
- x86_64/universal nếu hợp lý.

## 33.2 Python versions

Target:

- 3.10
- 3.11
- 3.12
- 3.13

## 33.3 Pythonic layer

Không expose C++ API một cách máy móc.

Ví dụ:

```python
with xlpp.Workbook.open(path) as wb:
    ws = wb["Sheet1"]
```

## 33.4 Exceptions

Map native errors sang exception Python rõ ràng.

## 33.5 PyPI workflow

Thêm workflow thực sự:

`.github/workflows/pypi-publish.yml`

và kiểm tra tài liệu không claim artifact chưa tồn tại.

## Exit criteria

```bash
pip install xlpp
```

hoạt động trên platform target.

---

# Phase 34 — .NET SDK

## 34.1 NuGet packaging

Tạo:

```text
XLPP.Native
XLPP
```

hoặc package thống nhất nếu distribution hợp lý.

## 34.2 RID binaries

- win-x64
- linux-x64
- osx-arm64

## 34.3 SafeHandle

Mọi native handle phải được wrap bằng `SafeHandle`.

## 34.4 Span/Memory

Cho bulk APIs:

```csharp
Span<double>
ReadOnlySpan<byte>
Memory<T>
```

giảm marshal overhead.

## 34.5 Async

Chỉ thêm async cho operation thật sự I/O-heavy:

```csharp
SaveAsync
OpenAsync
```

Không fake async bằng `Task.Run` cho mọi API.

## Exit criteria

NuGet có thể dùng như SDK bình thường mà người dùng không cần tự copy native DLL.

---

# Phase 35 — Performance Engineering

Chỉ optimization dựa trên benchmark.

## 35.1 Benchmark suite

Tạo:

`benchmarks/`

Scenario:

- open workbook;
- save workbook;
- 1M cell write;
- 1M streaming write;
- formula calculation;
- styles;
- chart generation;
- Pivot generation;
- encrypted save.

## 35.2 Competitor baseline

Khi phù hợp so với:

- OpenPyXL;
- XlsxWriter;
- ClosedXML;
- EPPlus;
- Open XML SDK.

Không cần thắng mọi benchmark.

## 35.3 Memory benchmark

Theo dõi:

- peak RSS;
- allocations;
- ZIP buffers;
- XML DOM footprint.

## 35.4 Performance regression gate

Không cho PR vô tình làm operation quan trọng chậm >10–15%.

## Exit criteria

Performance trở thành measurable contract.

---

# Phase 36 — Reliability CI 2.0

## 36.1 Strict warning CI

Mỗi PR chạy:

- GCC
- Clang
- MSVC `/W4`

## 36.2 Sanitizers

Nightly:

- ASan;
- UBSan.

Nếu khả thi:

- TSan cho các subsystem threaded.

## 36.3 Fuzzing

Continuous corpus fuzzing:

- ZIP;
- XML;
- workbook loader;
- formula parser;
- CFB;
- encryption;
- VBA.

## 36.4 Malformed workbook corpus

Tạo fixtures:

- truncated ZIP;
- bad relationships;
- cyclic relationships;
- malformed dimensions;
- invalid shared strings;
- broken style indexes;
- corrupted VBA CFB.

## Exit criteria

Parser phải fail safely, không crash và không OOM bất thường.

---

# Phase 37 — Large Enterprise Corpus

Đây là phase quyết định khả năng production thực sự.

## 37.1 Corpus classification

Thu thập workbook anonymized theo nhóm:

```text
small
medium
large
formula-heavy
chart-heavy
pivot-heavy
vba-heavy
enterprise
weird
```

## 37.2 Round-trip matrix

Mỗi workbook chạy:

```text
Open
→ no-op Save
→ Excel validate

Edit A1
→ Save
→ Excel validate

Add Sheet
→ Save
→ Excel validate

Copy Sheet
→ Save
→ Excel validate
```

## 37.3 Preservation score

Tính metric:

```text
known semantic preservation
unknown part preservation
Excel repair count
package delta
```

## Exit criteria

Ít nhất 500 workbook thực tế trước khi claim broad production compatibility.

---

# Phase 38 — Security Hardening

## 38.1 Parser resource limits

Global configurable limits:

```cpp
ParseLimits
```

bao gồm:

- max ZIP entries;
- max inflated bytes;
- max XML depth;
- max relationship count;
- max shared strings;
- max cells;
- max VBA stream size.

## 38.2 ZIP bomb defense

Ratio + total output limits.

## 38.3 XML hardening

Không external entities.

Không DTD expansion.

## 38.4 Encryption security

Audit:

- password handling;
- zeroization;
- random generation;
- key lifetime.

## 38.5 Security policy

Tạo:

`SECURITY.md`

## Exit criteria

XL++ đủ an toàn để xử lý workbook không tin cậy ở server.

---

# Phase 39 — Release Engineering

## 39.1 Semantic versioning

Chính thức:

```text
MAJOR.MINOR.PATCH
```

## 39.2 CHANGELOG

Tạo:

`CHANGELOG.md`

theo release.

## 39.3 Migration guides

Khi breaking API:

`docs/migrations/`

## 39.4 Reproducible build

Cố gắng deterministic archive/package.

## 39.5 Signed binaries

Nếu public distribution:

- Windows code signing;
- package checksum;
- GitHub provenance.

## Exit criteria

Một release có thể được tái tạo, kiểm tra và distribute chuyên nghiệp.

---

# Phase 40 — XL++ 2.0 Readiness Gate

Không release 2.0 chỉ vì đã tới Phase 40.

2.0 chỉ được phép khi đạt các gate sau.

## Core

- stable C++ public API;
- stable C ABI;
- no known critical structural corruption;
- durable save;
- transaction semantics.

## Compatibility

- >= 500 workbook enterprise corpus;
- Excel Desktop automation;
- zero known systematic repair issue;
- preservation score cao.

## Tests

- >200 logical suites;
- normal CI;
- strict;
- ASan;
- UBSan;
- fuzz;
- Excel interop.

Không đặt KPI dựa trên số check đơn thuần.

## Bindings

Python:

```bash
pip install xlpp
```

C#:

```bash
dotnet add package XLPP
```

hoạt động từ package chính thức.

## Documentation

Có:

- Getting Started
- API Reference
- Feature Matrix
- Examples
- Migration Guide
- Security Policy
- ABI Policy

---

# Thứ tự ưu tiên thực tế

Không nên chạy Phase 21 → 40 máy móc.

### Priority P0 — làm ngay

1. Phase 21 — capability/documentation cleanup.
2. Phase 22 — Excel Desktop interoperability harness.
3. Phase 23 — Preservation Engine 2.0.
4. Phase 36 — strict/sanitizer/fuzz CI.
5. Phase 37 — enterprise corpus.

Đây là nhóm giúp tăng **độ tin cậy** nhanh nhất.

### Priority P1 — feature refinement

6. Phase 24 — Pivot Engine 2.0.
7. Phase 25 — VBA Engine 2.0.
8. Phase 26 — Chart Fidelity.
9. Phase 27 — DrawingML.
10. Phase 29 — common Excel feature completion.

### Priority P2 — engine/API

11. Phase 28 — Formula Engine 2.0.
12. Phase 32 — C++/C ABI stabilization.
13. Phase 33 — Python SDK.
14. Phase 34 — .NET SDK.
15. Phase 35 — performance engineering.

### Priority P3 — specialized enterprise features

16. Phase 30 — connections.
17. Phase 31 — Data Model/OLAP research.
18. Phase 38 — extended security.
19. Phase 39 — release engineering.
20. Phase 40 — XL++ 2.0 readiness.

---

# Những việc KHÔNG nên ưu tiên lúc này

Không nên:

- tăng LOC để trông “đủ lớn”;
- thêm hàng trăm formula function trước interoperability;
- implement Data Model trước khi preservation engine đủ chắc;
- tự sinh SmartArt ngay;
- tự viết full ActiveX engine ngay;
- tự viết VBA p-code compiler;
- cố deserialize tất cả unknown XML rồi regenerate;
- phá C ABI để API đẹp hơn;
- đánh dấu feature “complete” chỉ vì serializer tạo được XML.

Từ giai đoạn này, tiêu chí quan trọng phải là:

**Excel opens it without repair + round-trip preserves unrelated content + API remains stable.**

---

# Target sau roadmap

Nếu hoàn thành P0 + phần lớn P1:

### XL++ 1.x mature target

- Core Excel editing: **9.5/10**
- Chart: **9.3/10**
- Pivot non-OLAP: **9.0/10**
- VBA: **8.5–9.0/10**
- Preservation: **9.3/10**
- Formula: **8.5–9.0/10**
- Streaming: **9.5/10**
- Encryption: **9.3/10**
- Bindings: **9.0/10**
- Enterprise interoperability: **8.5–9.0/10**

Quan trọng hơn con số:

**XL++ lúc đó có thể được xem là một production-grade native Excel engine, thay vì một thư viện chỉ có feature coverage rộng.**

---

# Milestone đề xuất tiếp theo

Milestone kế tiếp nên là:

## P0Z-I — Excel Interoperability & Preservation Foundation
### Version đề xuất: v1.12.0

Scope:

1. `CURRENT_CAPABILITIES.md`
2. retire stale roadmap
3. Excel Desktop COM test harness
4. Excel recovery-log detection
5. Excel-created corpus structure
6. Preservation Engine audit
7. unknown relationship preservation
8. unknown `extLst` preservation
9. stable package part identities
10. package-delta regression framework
11. strict-warning CI
12. sanitizer nightly CI
13. fuzz CI
14. compatibility report generation

Milestone tiếp theo sau đó:

## P0Z-J — Pivot & VBA Ecosystem
### v1.13.0

- calculated Pivot fields/items;
- PivotChart;
- slicer foundation;
- VBA references;
- UserForm parsing;
- FRX preservation.

Sau P0Z-J:

## P0Z-K — Chart/Drawing Fidelity & Common Excel Completion
### v1.14.0

Sau đó mới nên mở một nhánh:

## P1A — Formula Engine 2.0 & SDK Stabilization

Cách chia này sẽ giúp mỗi release có một mục tiêu kỹ thuật rõ ràng, dễ regression và dễ xác định chính xác XL++ đang tiến bộ ở đâu.