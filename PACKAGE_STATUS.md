# XL++ Full Project Package — P0U

This directory is the complete XL++ source tree assembled from the supplied baseline package and all preservation/chart development patches through P0U.

## Included development batches

- P0 — Preservation Core
- P0B — Object Graph Validation
- P0C — Pivot Preservation
- P0D — Drawing Preservation Foundation
- P0E — Safe Imported Image Mutation
- P0F — Multi-Drawing & Unknown DrawingML Preservation
- P0G — Extended Owner Graph & Corruption Hardening
- P0H — Chart Preservation & Selective Edit Foundation
- P0I — Deeper Selective Chart Editing & Lifecycle
- P0J — Chart Axis & Series Structure Foundation
- P0K — Chart Labels, Trendlines & Error Bars Foundation
- P0L — Per-point Data Labels, Custom Error Bars & Rich Chart Formatting Foundation
- P0M — Data-point Styling, Rich Text & Advanced Chart Formatting Foundation
- P0N — Chart Layout, Legend/Axis Formatting & Rich Text Expansion
- P0O — Axis Scaling, Display Units & Chart/Plot Area Formatting
- P0P — Chart Auxiliary Objects Foundation
- P0Q — Stock Chart Structure, Generation & Data-table Text
- P0R — 3D/Surface/Projected Chart Preservation Foundation
- P0S — Projected Pie, Doughnut & Radar Expansion
- P0T — Chart Style, Theme & Series Cache Foundation
- P0U — Cache Synchronization & Theme Transform Engine

## Verification of this assembled tree

Clean out-of-tree CMake verification performed on 2026-08-07:

- Configure: PASS
- Full build: PASS
- Standalone public-header checks: PASS
- C API target: PASS
- Package validator target: PASS
- Unit tests: 154/154 suites PASS
- Checks: 2,708/2,708 PASS
- Package consumers (`find_package` + `add_subdirectory`): PASS

The source package intentionally excludes generated build directories and stale compiled binaries (`build`, `bin`, `obj`, Visual Studio intermediate files, and prebuilt C API DLLs). Rebuild binaries from source for the target platform.

Project version file remains `1.1.2`; packaging this development state does not change the project's versioning policy.
