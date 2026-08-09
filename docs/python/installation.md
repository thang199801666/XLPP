# Install and compatibility

## Install a wheel

```bash
python -m pip install --upgrade xlpp
```

Verify the import and version:

```python
import xlpp

print(xlpp.__version__)
```

## Published wheel matrix

The current GitHub release workflow builds:

| Platform | Architecture | CPython |
| --- | --- | --- |
| Windows | AMD64 | 3.10, 3.11, 3.12, 3.13 |
| Linux manylinux | x86_64 | 3.10, 3.11, 3.12, 3.13 |
| macOS 10.15+ | Apple Silicon | 3.10, 3.11, 3.12, 3.13 |

The source package declares Python 3.8 or newer. A Python/platform combination
outside the wheel matrix is installed from source and therefore needs a C++20
compiler and Python development files.

!!! tip "Confirm that pip selected a wheel"

    Run `python -m pip install -v xlpp`. A wheel filename contains tags such as
    `cp312-win_amd64` or `cp313-manylinux_x86_64`. If pip begins compiling C++
    sources, no compatible wheel was found.

## Virtual environment

=== "Windows"

    ```powershell
    py -3.12 -m venv .venv
    .venv\Scripts\Activate.ps1
    python -m pip install --upgrade pip xlpp
    ```

=== "Linux / macOS"

    ```bash
    python3.12 -m venv .venv
    source .venv/bin/activate
    python -m pip install --upgrade pip xlpp
    ```

## Optional packages

XL++ itself does not require NumPy, pandas, Pillow, or openpyxl.

```bash
python -m pip install numpy pandas
```

- NumPy enables `write_array()` and `to_array()`.
- pandas integration uses `from_records()` and `to_records()`.
- openpyxl is useful for interoperability tests during migration.

## Build from source

From the repository root:

```bash
python -m pip install --upgrade pip setuptools wheel pybind11
python -m pip install .
```

The build compiles the Python binding, XL++ native sources, and bundled zlib
into one extension module.

### Windows requirements

- 64-bit CPython
- Visual Studio with Desktop development with C++
- MSVC C++20 support
- A matching x64 developer environment

### Linux/macOS requirements

- A C++20 compiler
- Python development headers
- Standard build tools
- Enough memory for the large pybind11 translation unit

## Test a source checkout

```bash
python -m pip install pytest numpy openpyxl
python setup.py build_ext --inplace
python -m pytest bindings/python/tests -q
```

See [Troubleshooting](troubleshooting.md) when pip compiles unexpectedly,
cannot find Python headers, or imports a stale local module.

