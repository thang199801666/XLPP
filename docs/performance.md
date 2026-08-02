# Performance

## Benchmarks

All benchmarks run on Windows 11, AMD Ryzen 5 5500U, Python 3.14, Release builds.

### Write + Read (Python binding)

| Rows × Cols | XLPP write | openpyxl write | xlsxwriter write | XLPP read | openpyxl read |
|------------|------------|----------------|-------------------|-----------|---------------|
| 500 × 5 | **8.5 ms** | 42.5 ms | 33.4 ms | **15.0 ms** | 32.5 ms |
| 1K × 10 | **36.1 ms** | 81.3 ms | 73.9 ms | **28.7 ms** | 70.7 ms |
| 5K × 10 | **212 ms** | 333 ms | 244 ms | **102 ms** | 263 ms |
| 10K × 15 | **654 ms** | 1225 ms | 877 ms | **399 ms** | 910 ms |

**Speedup**: XLPP writes **1.9× faster** and reads **2.3× faster** than openpyxl (10K rows).

### File Size

| Rows × Cols | XLPP | openpyxl | xlsxwriter |
|------------|------|----------|------------|
| 10K × 15 | 2.66 MB | 2.61 MB | 2.61 MB |

All three libraries produce similar file sizes.

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
