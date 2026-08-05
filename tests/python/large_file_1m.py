"""Generate and read a large workbook without loading it into OpenPyXL."""

import argparse
import time
from pathlib import Path

import xlpp


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=1_000_000)
    parser.add_argument("--columns", type=int, default=4)
    parser.add_argument("--output", default="large-xlpp.xlsx")
    args = parser.parse_args()

    output = Path(args.output)
    book = xlpp.Workbook()
    sheet = book.add_worksheet("MillionRows")
    sheet.append(["row", "value", "square", "text"])
    start = time.perf_counter()
    for row in range(1, args.rows + 1):
        sheet.append([row, row * 0.5, row * row, f"row-{row}"])
        if row % 100_000 == 0:
            print(f"appended {row:,} rows")
    write_start = time.perf_counter()
    book.save(str(output))
    write_seconds = time.perf_counter() - write_start

    loaded = xlpp.Workbook()
    read_start = time.perf_counter()
    loaded.load(str(output))
    read_seconds = time.perf_counter() - read_start
    result = loaded["MillionRows"]
    print(f"rows={result.max_row:,} columns={result.max_column}")
    print(f"generate={time.perf_counter() - start:.3f}s write={write_seconds:.3f}s read={read_seconds:.3f}s")
    print(f"file={output} bytes={output.stat().st_size:,}")


if __name__ == "__main__":
    main()
