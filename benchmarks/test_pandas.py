import sys, os, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "bindings", "python"))
import xlpp
import numpy as np
import pandas as pd

print("=== NumPy integration ===")

# Bulk write 2D array
arr = np.random.rand(1000, 10)
wb = xlpp.Workbook()
ws = wb.add_worksheet("Numpy")

t0 = time.perf_counter()
ws.write_array(arr, 1, 1)
t1 = time.perf_counter()
print(f"write_array 1000x10: {int((t1-t0)*1000)} ms")

# Read back as array
t2 = time.perf_counter()
out = ws.to_array(1, 1, 1000, 10)
t3 = time.perf_counter()
print(f"to_array: {int((t3-t2)*1000)} ms, shape={out.shape}")
assert np.allclose(out, arr), "Array round-trip failed!"
print("  Round-trip OK")

print("\n=== from_records / to_records ===")
ws2 = wb.add_worksheet("Records")
ws2.from_records([
    ["Alice", 30, 9.5],
    ["Bob", 25, 8.0],
    ["Carol", 35, 7.5],
], columns=["Name", "Age", "Score"])
records = ws2.to_records()
for row in records:
    print(" ", row)

print("\n=== Pandas integration ===")
df = pd.DataFrame({
    "Name": ["Alice", "Bob", "Carol"],
    "Age": [30, 25, 35],
    "Score": [9.5, 8.0, 7.5],
})
print(df)

# DataFrame -> records -> worksheet
ws3 = wb.add_worksheet("DataFrame")
ws3.from_records(df.values.tolist(), columns=list(df.columns))
records3 = ws3.to_records()
df2 = pd.DataFrame(records3[1:], columns=records3[0])
print("\nRound-tripped DataFrame:")
print(df2)
# Excel stores all numbers as double, so int -> float is expected
df_expected = df.copy()
df_expected["Age"] = df_expected["Age"].astype(float)
assert df2.equals(df_expected), f"DataFrame round-trip failed: {df2.dtypes} vs {df_expected.dtypes}"
print("\nPandas OK!")

# Save + reload
wb.save(os.path.join(os.path.dirname(__file__), "pandas_test.xlsx"))
wb2 = xlpp.Workbook()
wb2.load(os.path.join(os.path.dirname(__file__), "pandas_test.xlsx"))
print("\nReloaded sheets:", wb2.sheet_names)
print("All OK!")
os.remove(os.path.join(os.path.dirname(__file__), "pandas_test.xlsx"))
