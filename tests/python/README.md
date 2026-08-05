# Python Examples and Tests

Install the optional comparison dependencies:

```powershell
python -m pip install openpyxl numpy pandas pytest
```

Run the normal feature tests:

```powershell
python -m pytest tests/python -v
```

Run the OpenPyXL comparison benchmark:

```powershell
python tests/python/benchmark_openpyxl.py --rows 10000 --columns 10
```

Run the one-million-row example separately:

```powershell
python tests/python/large_file_1m.py --rows 1000000 --output large-xlpp.xlsx
```

The one-million-row test is intentionally not part of the default test suite because it is a long-running, disk-intensive workload.
