# Enterprise workbook corpus

This directory stores the manifest, not private customer workbooks. The runner supports the four mandatory round-trip probes: no-op save, edit A1, add sheet, and copy sheet. Production/private corpora can point the same schema at an external folder in CI.

A corpus item records its class (`formula-heavy`, `chart-heavy`, `pivot-heavy`, `vba-heavy`, `enterprise`, etc.) and the operations expected to remain valid. `tools/enterprise_corpus_runner.py` also records package-part SHA-256 deltas so preservation regressions are visible rather than hidden behind a simple open/save pass.
