using System.Diagnostics;
using ClosedXML.Excel;
using XLPP;

const int rows = 10_000;
const int columns = 10;
var temp = Path.GetTempPath();

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

File.Delete(xlppPath);
File.Delete(closedPath);
