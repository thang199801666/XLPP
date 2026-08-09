# Streaming and performance

Use the normal `Workbook` model when you need random access, styles, charts,
pivots, formulas, or structural edits. Use streaming when data volume is large
and access is naturally sequential.

## Streaming writer

```python
writer = xlpp.StreamingWorkbookWriter("large.xlsx")
sheet = writer.add_worksheet("Data")

sheet.append(["Row", "Label", "Value"])
for index in range(1_000_000):
    sheet.append([index, f"row-{index}", index * 0.5])

writer.close()
```

Always call `close()` to finish ZIP/package output. Check `writer.closed` when
ownership is shared across application layers.

## Shared-string modes

```python
writer = xlpp.StreamingWorkbookWriter(
    "large.xlsx",
    xlpp.SharedStringMode.BOUNDED_LRU,
    4096,
)
```

| Mode | Behavior | Use when |
| --- | --- | --- |
| `DISABLED` | Writes inline strings | Strings are mostly unique or memory is critical |
| `HASH` | Deduplicates all shared strings | Repetition is high and the string set fits memory |
| `BOUNDED_LRU` | Deduplicates through a bounded cache | You want predictable memory with local repetition |

Configure compression and worker count before closing:

```python
writer.set_compression_level(xlpp.CompressionLevel.FASTEST)
writer.set_compression_strategy(xlpp.CompressionStrategy.DEFAULT)
writer.set_parallel_workers(4)
```

## Streaming reader

```python
options = xlpp.StreamingReaderOptions()
options.max_file_bytes = 2 * 1024 * 1024 * 1024
options.max_entry_bytes = 512 * 1024 * 1024
options.max_total_bytes = 4 * 1024 * 1024 * 1024
options.max_entries = 50_000
options.validate_cell_references = True

reader = xlpp.StreamingWorkbookReader("large.xlsx", options)
print(reader.worksheet_names())

for row_number, cells in reader.worksheet("Data"):
    values = [cell.value for cell in cells]
    print(row_number, values)
```

The convenience iterator materializes the selected worksheet into Python list
objects. For early termination and lower Python-side materialization, use a
callback:

```python
def consume(row_number, cells):
    print(row_number, [cell.value for cell in cells])
    return row_number < 1000

reader.for_each_row("Data", consume)
```

## NumPy arrays

```python
import numpy as np

values = np.arange(50_000, dtype=float).reshape(5_000, 10)
sheet.write_array(values, row=1, col=1)

result = sheet.to_array(
    min_row=1,
    min_col=1,
    max_row=5_000,
    max_col=10,
)
```

`write_array()` accepts 2D numeric arrays and supports `transpose=True`.
`to_array()` returns numeric data; missing and non-numeric cells become `0.0`.

## pandas

```python
sheet.from_records(
    df.values.tolist(),
    columns=list(df.columns),
)

records = sheet.to_records()
df2 = pd.DataFrame(records[1:], columns=records[0])
```

Records preserve mixed Python value types better than numeric arrays.

## Performance checklist

### Writing

- Prefer `append()`, `from_records()`, or `write_array()` over per-cell calls.
- Use streaming for append-only large exports.
- Choose `FASTEST` compression when throughput matters more than file size.
- Avoid creating many near-duplicate styles.
- Avoid formatting large empty ranges.
- Reuse strings through an appropriate shared-string mode.

### Reading

- Use streaming when rows can be consumed sequentially.
- Set realistic ZIP/resource limits for untrusted files.
- Read only the worksheet and range required by the operation.
- Avoid converting numeric blocks to Python objects when NumPy is sufficient.

### Saving an existing workbook

- Disable formula calculation and chart-cache synchronization unless needed.
- Use `changed_references_only` for cache synchronization.
- Preserve imported enterprise parts rather than rebuilding unrelated content.
- Benchmark with production-shaped workbooks, not only dense numeric grids.

## Choosing an API

| Requirement | Recommended API |
| --- | --- |
| Random cell edits and formatting | `Workbook` / `Worksheet` |
| Charts, pivots, formulas, structural edits | `Workbook` model |
| Million-row sequential export | `StreamingWorkbookWriter` |
| Sequential row ingestion | `StreamingWorkbookReader.for_each_row()` |
| Dense numeric matrix | `write_array()` / `to_array()` |
| Mixed tabular data | `from_records()` / `to_records()` |
| HTTP/database binary payload | `save_bytes()` / `load_bytes()` |

See [Performance](../performance.md) for repository benchmarks and methodology.

