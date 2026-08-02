# Performance

## Benchmarks

All benchmarks run on Windows 11, AMD Ryzen 5 5500U, MSVC 19.51, Release build.

### XML Scanner

| Rows × Cells | XLPP (SIMD) | Baseline | Speedup |
|-------------|-------------|----------|---------|
| 5,000 × 4 | **10.5 ms** | 63 ms | **6.0×** |

### Streaming read/write

| Rows | Write | Read | File Size |
|------|-------|------|-----------|
| 5,000 × 3 types | **129 ms** | **58 ms** | 87 KB |

### Comparison (estimated — based on published benchmarks)

| Library | 250K rows × 15 cols write | 250K rows read | Threads |
|---------|--------------------------|---------------|---------|
| **XLPP** | ~2.5s | ~1s (mmap) | 4 |
| openpyxl | ~45s | ~30s | 1 |
| xlsxwriter | ~3s | — (write only) | 1 |
| ClosedXML | ~8s | ~16s | 1 |

## Optimization techniques

| Technique | Where | Impact |
|-----------|-------|--------|
| SIMD (SSE2) | XML scanning | 6× speedup |
| uint64_t cell keys | Cell storage | No sort needed |
| ThreadPool | Save + ZIP | Scales with cores |
| mmap I/O | ZIP reading | Zero-copy |
| Default style skip | Save path | ~90% hash savings |
| Column merge | Output XML | Smaller files |
| Row chunking | Per-row save | Large single sheets |
| Style dedup | styles.xml | Unique styles only |
| SST dedup | shared strings | Unique strings only |

## Tuning save performance

```cpp
SaveOptions opt;

// Compression: trade size vs speed
opt.compressionLevel = CompressionLevel::Fastest;  // faster, larger
opt.compressionLevel = CompressionLevel::Best;     // slower, smaller

// Parallelism
opt.parallelWorkers = std::thread::hardware_concurrency();
opt.parallelSheets = true;   // per-sheet parallelism
opt.parallelRows = true;     // per-row inside single sheet
```
