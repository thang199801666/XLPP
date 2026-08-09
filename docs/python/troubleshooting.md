# Troubleshooting

## `ModuleNotFoundError: No module named 'xlpp'`

Check the interpreter and pip belong to the same environment:

```bash
python -c "import sys; print(sys.executable)"
python -m pip show xlpp
```

Install with `python -m pip`, not an unqualified `pip` from another environment.

## Import uses the wrong build

```python
import xlpp
print(xlpp.__file__)
print(xlpp.__version__)
```

Remove stale in-place `.pyd`/`.so` files or old editable installs from the
working directory and environment.

## pip starts compiling C++

No wheel matches the interpreter/platform. Confirm:

- CPython version is in the published wheel matrix.
- Python architecture is 64-bit.
- Platform architecture matches the wheel.
- pip is current enough to understand wheel tags.

```bash
python -m pip install --upgrade pip
python -m pip install -v xlpp
```

## `Python.h` or Python library not found

Install Python development files and ensure the compiler architecture matches
Python. On Windows, run from an x64 Visual Studio developer environment. On
Linux, install the development package supplied by the distribution.

## MSVC `C1128`, section limit, or object too large

The binding is a large pybind11 translation unit. The source build enables
`/bigobj`, but custom builds must retain it. Use a 64-bit compiler process and
avoid highly parallel builds on low-memory machines.

## Workbook will not open

1. Capture the full exception.
2. Verify the input is OOXML `.xlsx`/`.xlsm`, not legacy binary `.xls`.
3. Inspect encryption with `inspect_office_encryption()`.
4. Set the password in `LoadOptions` when encrypted.
5. Use lenient loading only for intentional recovery.
6. Run the package validator for relationship/content-type diagnostics.

## Password is rejected

```python
info = xlpp.inspect_office_encryption("input.xlsx")
print(info.encrypted, info.supported, info.mode)
```

Confirm the encryption profile is supported and that password text reaches
the process unchanged. Avoid shell quoting for secrets.

## Formula cached value is stale

Writing a formula does not guarantee all consumers recalculate immediately.

```python
report = book.calculate_formulas()
print(report.unsupported_formulas, report.evaluation_errors)
```

For Excel-side calculation, set workbook calculation properties or enable
full calculation on load. Inspect unsupported functions before trusting native
cached values.

## Chart shows old data

```python
options = xlpp.ChartCacheSyncOptions()
options.changed_references_only = False
report = book.synchronize_chart_caches(options)
print(report.caches_updated, report.warnings)
```

Verify each series reference points to the intended worksheet and range.

## Pivot source width mismatch

The number of pivot cache fields must equal the number of columns in
`pivot.cache.source_data`. Add or remove fields or correct the range.

## Excel repairs the output file

- Save to a new file and keep the original.
- Run `book.validate()` before save.
- Run `xlpp-package-validator before.xlsx --compare after.xlsx`.
- Inspect Excel's repair log.
- Reduce the edit to the smallest reproducer.
- Preserve advanced chart/pivot/enterprise parts rather than rebuilding them.

## VBA disappeared

- Confirm `book.has_vba_project` immediately after load.
- Save as `.xlsm`.
- Do not call `remove_vba_project()`.
- Check whether a template conversion workflow removed macro content.

## VBA signature is invalid

Any module/project modification can invalidate signatures. XL++ does not
re-sign VBA projects. Use an external signing process after editing.

## Streaming reader uses too much Python memory

The worksheet convenience iterator constructs a Python list. Use
`for_each_row()` for callback-driven processing and early termination.

## NumPy helper is unavailable

```bash
python -m pip install numpy
```

Confirm NumPy is installed in the same interpreter that imports XL++.

## openpyxl and XL++ produce different values

Common causes:

- Excel stores numbers as doubles.
- Formula cached values differ from formula text.
- Date epoch is 1900 versus 1904.
- Number formats change interpretation.
- openpyxl `data_only=True` reads cached values.
- XL++ native calculation supports only part of Excel's full behavior.

Compare raw values, formulas, cached values, date epoch, and number formats
separately.

## Report a reproducible issue

Include:

- XL++ and Python versions
- Operating system and architecture
- Wheel filename or source compiler
- Minimal code reproducer
- Minimal sanitized workbook when possible
- Full exception/build output
- Excel repair log or package validator output
- Whether the issue reproduces after a no-op load/save round trip
