using XLPP;

using var workbook = new Workbook();
var sheet = workbook.AddWorksheet("CSharp");
if (sheet.Name != "CSharp")
    throw new InvalidOperationException("Worksheet name binding failed.");
sheet["A1"].Value = "XL++";
sheet["B1"].Value = 42;
sheet["C1"].Formula = "=B1*2";
sheet["A1"].Font.SetBold(true);
sheet["A1"].Font.SetUnderline(true);
sheet["A1"].NumberFormat = "@";
if (sheet["A1"].NumberFormat != "@")
    throw new InvalidOperationException("Cell number format binding failed.");
if (sheet["A1"].Address != "A1")
    throw new InvalidOperationException("Cell address binding failed.");
if (sheet.Dimensions != "A1:C1")
    throw new InvalidOperationException($"Worksheet dimensions binding failed: {sheet.Dimensions}");
workbook.Properties.Title = "C# binding test";
sheet.MergeCells("A2:C2");
sheet.FreezePanes("A2");
if (sheet.FrozenPane != "A2" || sheet.MergedRanges.Count != 1)
    throw new InvalidOperationException("Worksheet metadata binding failed.");
if (!sheet.IsMerged("A2"))
    throw new InvalidOperationException("Merged-cell query binding failed.");
sheet.ClearFreezePanes();
if (sheet.FrozenPane != null)
    throw new InvalidOperationException("Clear freeze-pane binding failed.");
var table = sheet.AddTable("SmokeTable", "A1:C1");
table.AddColumn("Value");
if (table.Name != "SmokeTable" || table.ColumnCount != 1)
    throw new InvalidOperationException("Table binding failed.");
table.DisplayName = "SmokeDisplay";
table.StyleName = "TableStyleMedium4";
table.ShowRowStripes = false;
var validation = sheet.AddListValidation("A1:A3", "\"Open,Closed\"");
validation.AllowBlank = true;
validation.SetPrompt("Status", "Choose a status");
validation.SetError("Invalid", "Use an allowed status");
var rule = sheet.AddFormulaRule("A1:A3", "A1>0");
rule.Priority = 1;
rule.StopIfTrue = true;
rule.SetFontColor("FFFF0000");
sheet.AddDataBarRule("B1:B3");

// Stable native child handles must remain usable when their owning collection grows.
var pinnedStyle = workbook.AddNamedStyle("PinnedStyle");
for (var i = 0; i < 128; ++i)
    workbook.AddNamedStyle($"Style{i}");
pinnedStyle.Name = "PinnedStyleUpdated";
if (pinnedStyle.Name != "PinnedStyleUpdated")
    throw new InvalidOperationException("NamedStyle handle became stale after collection growth.");

var pinnedName = workbook.AddDefinedName("PinnedName", "CSharp!$A$1");
for (var i = 0; i < 128; ++i)
    workbook.AddDefinedName($"Name{i}", "CSharp!$A$1");
pinnedName.Value = "CSharp!$B$1";
if (pinnedName.Value != "CSharp!$B$1")
    throw new InvalidOperationException("DefinedName handle became stale after collection growth.");

var pinnedColumn = table.AddColumn("PinnedColumn");
for (var i = 0; i < 96; ++i)
    table.AddColumn($"Column{i}");
pinnedColumn.Name = "PinnedColumnUpdated";
if (pinnedColumn.Name != "PinnedColumnUpdated")
    throw new InvalidOperationException("TableColumn handle became stale after collection growth.");
using (var chart = new Chart(ChartType.Bar))
{
    chart.Title = "Sales";
    chart.XAxisTitle = "Quarter";
    chart.YAxisTitle = "Units";
    chart.Style = "10";
    chart.Grouping = 1;
    chart.SetSize(800, 500);
    chart.SetLegend(true, "b");
    var series = chart.AddSeries("Units");
    series.ValuesReference = "'CSharp'!$B$1:$B$1";
    series.CategoriesReference = "'CSharp'!$A$1:$A$1";
    sheet.AddChart(chart);
}

var output = Path.Combine(Environment.CurrentDirectory, "csharp-test.xlsx");
if (!workbook.Save(output))
    throw new InvalidOperationException("XL++ failed to save the workbook.");

using var reloaded = new Workbook();
if (!reloaded.Load(output))
    throw new InvalidOperationException($"XL++ failed to reload the workbook: {Workbook.LastError}");
var reloadedSheet = reloaded["CSharp"];
if (reloaded.Properties.Title != "C# binding test" || reloadedSheet.Dimensions != "A1:C1")
    throw new InvalidOperationException("C# binding round-trip metadata failed.");

Console.WriteLine($"C# binding OK: {output}");
