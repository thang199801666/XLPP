# XL++ Core Architecture

XL++ exposes one public C++20 library while its implementation is divided into semantic model, services, OOXML codecs and package primitives. `Workbook` is a facade/orchestration surface; it is not the storage location for every parser or serializer.

## Dependency direction

```text
Public API / Workbook facade
            |
   +--------+---------+----------------+
   |                  |                |
 Formula         Dependencies       Validation
   |                  |                |
   +------------------+----------------+
                      |
                Document Model
                      |
       +--------------+---------------+
       |              |               |
     Charts          Pivot          Drawings
       |              |               |
       +--------------+---------------+
                      |
                OOXML adapters
                      |
                Preservation
                      |
                     OPC
                      |
                     ZIP
                      |
                  Platform/Core
```

Encryption wraps/unwraps the serialized OOXML package at the file boundary. Streaming uses optimized read/write paths while sharing the same package/XML safety invariants.

## Implementation tree

```text
src/XLPP/
├── Core/                         low-level utilities/threading
├── Platform/                     mapped files/platform facilities
├── Package/
│   ├── Zip/                      ZIP/ZIP64 reader/writer
│   ├── Opc/                      OPC relationships/package graph
│   └── Xml/                      scanner/pull reader/XML utilities
├── Model/
│   ├── Workbook/                 workbook semantic model
│   └── Worksheet/                worksheet/range semantic model
├── Formula/                      calculation/reference dependencies
├── Dependencies/                 structural/topology rewrite
├── IO/                           load/save facade, transactions, binary IO
├── Preservation/                 retained-part and XML merge policy
├── OOXML/
│   ├── Common/                   namespaces, relationships, rich text, drawing XML support
│   ├── Workbook/                 WorkbookReader / WorkbookWriter / reference support
│   ├── Worksheet/                worksheet reader/writer/features/batch serialization
│   ├── Charts/                   ChartReader / ChartSerializer / ImportedChartXmlEditor / package patcher
│   ├── Drawings/                 drawing reader/writer/package support/preserved editor
│   ├── Pivot/                    PivotTable/cache semantic codec
│   ├── Styles/                   style/differential-style codec
│   ├── Tables/                   table writer
│   └── Comments/                 comments/VML codec
├── Charts/                       chart-domain services (cache synchronization)
├── Streaming/                    streaming worksheet/package paths
├── Encryption/                   Agile/Standard Office encryption + CFB
├── VBA/                          VBA project/binary lifecycle
└── Validation/                   workbook validation
```

## CMake modules

The implementation compiles through 35 internal object targets before aggregation into `xlpp_static` / `XLPP::xlpp`.

Core/service modules:

- `xlpp_model_workbook_obj`
- `xlpp_model_worksheet_obj`
- `xlpp_formula_engine_obj`
- `xlpp_formula_parser_obj`
- `xlpp_formula_functions_obj`
- `xlpp_formula_dependency_obj`
- `xlpp_formula_reference_obj`
- `xlpp_dependencies_obj`
- `xlpp_package_zip_obj`
- `xlpp_package_opc_obj`
- `xlpp_package_xml_obj`
- `xlpp_io_obj`
- `xlpp_preservation_obj`
- `xlpp_charts_obj`
- `xlpp_streaming_obj`
- `xlpp_encryption_obj`
- `xlpp_vba_obj`
- `xlpp_validation_obj`

OOXML subdomain modules:

- `xlpp_ooxml_common_obj`
- `xlpp_ooxml_workbook_obj`
- `xlpp_ooxml_worksheet_obj`
- `xlpp_ooxml_chart_support_obj`
- `xlpp_ooxml_chart_reader_obj`
- `xlpp_ooxml_chart_serializer_obj`
- `xlpp_ooxml_chart_xml_editor_obj`
- `xlpp_ooxml_chart_package_patcher_obj`
- `xlpp_ooxml_drawings_obj`
- `xlpp_ooxml_pivot_obj`
- `xlpp_ooxml_styles_obj`
- `xlpp_ooxml_tables_obj`
- `xlpp_ooxml_comments_obj`

Object libraries are private build boundaries; consumers still link one XL++ target.

## Public/private boundary

`include/XLPP/` is the supported public surface. Headers under `src/XLPP/` are implementation details and may move between releases. Internal includes use source-root-relative paths such as `"Package/Zip/ZipArchive.h"`.

## Workbook facade

Path/stream public entry points live in `IO/WorkbookIO.cpp`; atomic/temp-file mechanics live in `IO/FileTransaction.*`; raw binary file primitives live in `IO/BinaryFile.*`.

There is no catch-all `WorkbookCodec.cpp` in P0Z-B. OOXML orchestration is split into `WorkbookReader.cpp` and `WorkbookWriter.cpp`, with worksheet/chart/drawing/pivot/style/table/comment responsibilities delegated to their own adapters. This keeps Workbook-level code focused on package orchestration rather than feature-specific XML parsing.

## Preservation-first rule

When XL++ understands an object semantically, its model/OOXML adapter owns the generated XML for that semantic region. Unknown package parts, relationships, extension lists and unmodeled siblings are preserved whenever possible. Imported-object edits should patch the narrowest owning subtree/part instead of regenerating unrelated OOXML.

Generic XML merge helpers belong to `Preservation/`; semantic Chart/Image/Pivot changes remain inside their feature codecs.

## Tests

Native tests are organized under `tests/unit/`:

```text
tests/unit/
├── TestMain.cpp
├── TestFramework.h
├── Model/ModelWorkbookTests.cpp
├── Formula/FormulaDependencyTests.cpp
└── Regression/
    ├── RegressionRegistry.cpp
    ├── PackageStreamingTests.cpp
    ├── WorkbookTests.cpp
    ├── DrawingsTests.cpp
    ├── ChartsTests.cpp
    ├── PivotTests.cpp
    └── CompatibilityTests.cpp
```

The former ~9.7k-line test monolith and the later ~7.6k-line regression TU are both gone. Regression coverage is split by Package/Streaming, Workbook, Drawings, Charts, Pivot and Compatibility domains while a small registry preserves suite order.

## Next decomposition targets

1. Split the pure `ImportedChartXmlEditor.cpp` further only when chart-level and series-level shared helpers have a clean internal owner.
2. Decompose the formula evaluator and relationship-graph implementation when those boundaries reduce rebuild cost without duplicating state.
3. Extend architecture-boundary checks as private interfaces stabilize.
4. Keep Python/C# parity synchronized through `BindingParityTests` and the dedicated binary binding CI job.


## P0Z-D formula and binding boundary

Formula parsing and the Excel function catalog now depend on a private `FormulaEvaluationContext` rather than the concrete calculation engine. This prevents parser/function changes from recompiling or reaching into workbook calculation state. The five Formula object modules are aggregated like the other private domains.

High-level bindings are treated as maintained surfaces rather than downstream samples. Required C/Python/C# symbols are checked by `cmake/check_binding_parity.cmake`, while the GitHub binding job performs real Python and .NET builds/tests with their SDK dependencies.
