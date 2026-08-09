# XL++ Documentation

Welcome to the XL++ documentation. XL++ is a high-performance C++20 library for reading and writing Excel `.xlsx` files.

## Get started

- [Quick Start — C++](quickstart-cpp.md)
- [Quick Start — Python](quickstart-python.md)
- [Quick Start — C#](quickstart-csharp.md)
- [API Reference](api-reference.md)
- [Architecture](architecture.md)
- [Performance](performance.md)
- [Current Capabilities](CURRENT_CAPABILITIES.md)
- [P0Z-I Phase 28–37 Refinement](P0ZI_PHASE28_TO_37_REFINEMENT.md)
- [Building from Source](building.md)

## Overview

XL++ is designed to be:

- **Fast** — SIMD-accelerated XML parsing (~6× speedup), multi-threaded save, mmap I/O
- **Minimal** — Only one dependency: zlib (bundled)
- **Compatible** — openpyxl-inspired API for C++, Python, and C# bindings
- **Robust** — strict-warning/sanitizer gates, Clang/libFuzzer harness, enterprise preservation corpus, strict OOXML support and ZIP64 for >4GB files
- **Preservation-aware** — broad read/write/edit coverage with explicit preservation-first boundaries for unsupported enterprise/extension features
