# Performance

## Benchmarks

All benchmarks run on Windows 11, AMD Ryzen 5 5500U, MSVC 19.51, Release build.

### C++ Native (end-to-end)

| Rows × Cols | Cells | Build | Save | Load | File Size |
|------------|-------|-------|------|------|-----------|
| 1K × 10 | 10K | **3.6 ms** | **12.9 ms** | **20.5 ms** | 43 KB |
| 10K × 15 | 150K | **61 ms** | **305 ms** | **153 ms** | 480 KB |
| 50K × 10 | 500K | **203 ms** | **1,125 ms** | **531 ms** | 1.9 MB |
| 100K × 10 | 1M | **411 ms** | **2,223 ms** | **1,064 ms** | 3.8 MB |

**Throughput**: 450K cells/sec write, 940K cells/sec read at 1M cells.

### Python Binding (vs openpyxl vs xlsxwriter)

| Rows × Cols | XLPP write | openpyxl write | xlsxwriter write | XLPP read | openpyxl read |
|------------|------------|----------------|-------------------|-----------|---------------|
| 500 × 5 | **8.5 ms** | 42.5 ms | 33.4 ms | **15.0 ms** | 32.5 ms |
| 1K × 10 | **36.1 ms** | 81.3 ms | 73.9 ms | **28.7 ms** | 70.7 ms |
| 5K × 10 | **212 ms** | 333 ms | 244 ms | **102 ms** | 263 ms |
| 10K × 15 | **654 ms** | 1225 ms | 877 ms | **399 ms** | 910 ms |

**Speedup vs openpyxl**: 1.9× write, 2.3× read (10K rows).

### File Size

| Rows × Cols | XLPP | openpyxl | xlsxwriter |
|------------|------|----------|------------|
| 10K × 15 | 480 KB | 2.61 MB | 2.61 MB |

### C++ XML Scanner (internal)

| Rows × Cells | XLPP (SIMD) | Baseline | Speedup |
|-------------|-------------|----------|---------|
| 5,000 × 4 | **10.5 ms** | 63 ms | **6.0×** |

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
