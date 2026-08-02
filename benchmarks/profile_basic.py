import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "bindings", "python"))
import xlpp, time, random, tempfile

N = 5000
print(f"Profile: {N} rows x 10 cols")

data = []
for r in range(N):
    row = []
    for c in range(10):
        if c == 0: row.append(f"Item-{r}")
        elif c == 1: row.append(r)
        elif c == 2: row.append(round(random.uniform(0,100000),2))
        else: row.append("text")
    data.append(row)

t0 = time.perf_counter()
wb = xlpp.Workbook()
ws = wb.add_worksheet("Data")
for row in data:
    ws.append(row)
t1 = time.perf_counter()
print(f"Append: {int((t1-t0)*1000)} ms")

ws["A1"].font().bold = True
fp = os.path.join(tempfile.gettempdir(), "xlpp_profile.xlsx")
t2 = time.perf_counter()
wb.save(fp)
t3 = time.perf_counter()
print(f"Save:   {int((t3-t2)*1000)} ms")
print(f"Total:  {int((t3-t0)*1000)} ms")
os.remove(fp)
