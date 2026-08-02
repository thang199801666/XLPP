# XL++ Documentation

Welcome to the XL++ documentation. XL++ is a high-performance C++20 library for reading and writing Excel `.xlsx` files.

## Get started

- [Quick Start — C++](quickstart-cpp.md)
- [Quick Start — Python](quickstart-python.md)
- [Quick Start — C#](quickstart-csharp.md)
- [API Reference](api-reference.md)
- [Architecture](architecture.md)
- [Performance](performance.md)
- [Building from Source](building.md)

## Overview

XL++ is designed to be:

- **Fast** — SIMD-accelerated XML parsing (~6× speedup), multi-threaded save, mmap I/O
- **Minimal** — Only one dependency: zlib (bundled)
- **Compatible** — openpyxl-inspired API for C++, Python, and C# bindings
- **Robust** — 200-round fuzz testing, strict OOXML support, ZIP64 for >4GB files
- **Complete** — Read, write, and round-trip with full style/formula/image/chart fidelity
