using System.Diagnostics;
using ClosedXML.Excel;
using XLPP;

const int rows = 10_000;
const int columns = 15;
var temp = Path.GetTempPath();

// Warm up both serializers before measuring the real workloads.
var warmupXlppPath = Path.Combine(temp, "xlpp-csharp-warmup.xlsx");
using (var workbook = new Workbook())
{
    var sheet = workbook.AddWorksheet("Warmup");
    sheet["A1"].Value = "warmup";
    workbook.Save(warmupXlppPath);
}
File.Delete(warmupXlppPath);
var warmupClosedPath = Path.Combine(temp, "closedxml-warmup.xlsx");
using (var workbook = new XLWorkbook())
{
    var sheet = workbook.Worksheets.Add("Warmup");
    sheet.Cell("A1").Value = "warmup";
    workbook.SaveAs(warmupClosedPath);
}
File.Delete(warmupClosedPath);

static string Value(int row, int column) => column switch
{
    0 => $"Item-{row}",
    1 => (row + 1).ToString(),
    2 => ((row * 7919 % 10_000_000) / 100.0).ToString("F2"),
    _ => "Lorem ipsum dolor sit amet"
};

static void Print(string library, string operation, Stopwatch timer, string path)
{
    timer.Stop();
    Console.WriteLine($"BENCHMARK,csharp,{library},{operation},{timer.Elapsed.TotalMilliseconds:F2},{new FileInfo(path).Length}");
}

static void PrintRead(string library, string operation, Stopwatch timer)
{
    timer.Stop();
    Console.WriteLine($"BENCHMARK,csharp,{library},{operation},{timer.Elapsed.TotalMilliseconds:F2},0");
}

static void RunXlppScenario(string scenario, string path)
{
    using (var workbook = new Workbook())
    {
        var data = workbook.AddWorksheet("Data");
        var lookup = workbook.AddWorksheet("Lookup");
        var timer = Stopwatch.StartNew();
        for (var row = 1; row <= 500; row++)
        {
            lookup.Cell((ulong)row, 1).Value = $"Item-{row - 1}";
            lookup.Cell((ulong)row, 2).Value = row * 1.25;
        }
        for (var row = 0; row < rows; row++)
        {
            for (var column = 0; column < columns; column++)
            {
                if (scenario == "formula" && column == columns - 1)
                    data.Cell((ulong)(row + 1), (ulong)(column + 1)).Formula =
                        $"=VLOOKUP(A{row + 1},Lookup!$A$1:$B$500,2,FALSE)";
                else
                    data.Cell((ulong)(row + 1), (ulong)(column + 1)).Value = Value(row, column);
            }
        }
        data["A1"].Font.SetBold(true);
        data.FreezePanes("A2");
        workbook.Save(path);
        Print("XLPP", $"{scenario}_write", timer, path);
    }
    var readTimer = Stopwatch.StartNew();
    using (var workbook = new Workbook())
    {
        workbook.Load(path);
        _ = workbook["Data"].Cell((ulong)rows, (ulong)columns).Formula;
    }
    PrintRead("XLPP", $"{scenario}_read", readTimer);
}

static void RunClosedXmlScenario(string scenario, string path)
{
    var timer = Stopwatch.StartNew();
    using (var workbook = new XLWorkbook())
    {
        var data = workbook.Worksheets.Add("Data");
        var lookup = workbook.Worksheets.Add("Lookup");
        for (var row = 1; row <= 500; row++)
        {
            lookup.Cell(row, 1).Value = $"Item-{row - 1}";
            lookup.Cell(row, 2).Value = row * 1.25;
        }
        for (var row = 0; row < rows; row++)
        {
            for (var column = 0; column < columns; column++)
            {
                var cell = data.Cell(row + 1, column + 1);
                if (scenario == "formula" && column == columns - 1)
                    cell.FormulaA1 = $"VLOOKUP(A{row + 1},Lookup!$A$1:$B$500,2,FALSE)";
                else
                    cell.Value = Value(row, column);
            }
        }
        data.Cell("A1").Style.Font.Bold = true;
        data.SheetView.FreezeRows(1);
        workbook.SaveAs(path);
    }
    Print("ClosedXML", $"{scenario}_write", timer, path);
    var readTimer = Stopwatch.StartNew();
    using (var workbook = new XLWorkbook(path))
        _ = workbook.Worksheet("Data").Cell(rows, columns).FormulaA1;
    PrintRead("ClosedXML", $"{scenario}_read", readTimer);
}

var xlppPath = Path.Combine(temp, "xlpp-csharp-benchmark.xlsx");
var closedPath = Path.Combine(temp, "closedxml-benchmark.xlsx");

using (var workbook = new Workbook())
{
    var sheet = workbook.AddWorksheet("Data");
    var timer = Stopwatch.StartNew();
    for (var row = 0; row < rows; row++)
        for (var column = 0; column < columns; column++)
            sheet.Cell((ulong)(row + 1), (ulong)(column + 1)).Value = Value(row, column);
    sheet["A1"].Font.SetBold(true);
    sheet.FreezePanes("A2");
    workbook.Save(xlppPath);
    Print("XLPP", "write", timer, xlppPath);
}

var xlppReadTimer = Stopwatch.StartNew();
using (var workbook = new Workbook())
{
    workbook.Load(xlppPath);
    _ = workbook["Data"].MaxRow;
}
PrintRead("XLPP", "read", xlppReadTimer);

using (var workbook = new XLWorkbook())
{
    var sheet = workbook.Worksheets.Add("Data");
    var timer = Stopwatch.StartNew();
    for (var row = 0; row < rows; row++)
        for (var column = 0; column < columns; column++)
            sheet.Cell(row + 1, column + 1).Value = Value(row, column);
    sheet.Cell("A1").Style.Font.Bold = true;
    sheet.SheetView.FreezeRows(1);
    workbook.SaveAs(closedPath);
    Print("ClosedXML", "write", timer, closedPath);
}

var closedReadTimer = Stopwatch.StartNew();
using (var workbook = new XLWorkbook(closedPath))
    _ = workbook.Worksheet("Data").LastRowUsed()?.RowNumber();
PrintRead("ClosedXML", "read", closedReadTimer);

File.Delete(xlppPath);
File.Delete(closedPath);

foreach (var scenario in new[] { "lookup", "formula" })
{
    var scenarioXlppPath = Path.Combine(temp, $"xlpp-csharp-{scenario}.xlsx");
    var scenarioClosedPath = Path.Combine(temp, $"closedxml-{scenario}.xlsx");
    RunXlppScenario(scenario, scenarioXlppPath);
    RunClosedXmlScenario(scenario, scenarioClosedPath);
    File.Delete(scenarioXlppPath);
    File.Delete(scenarioClosedPath);
}
