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
if (sheet["A1"].Address != "A1")
    throw new InvalidOperationException("Cell address binding failed.");
if (sheet.Dimensions != "A1:C1")
    throw new InvalidOperationException($"Worksheet dimensions binding failed: {sheet.Dimensions}");
workbook.Properties.Title = "C# binding test";
sheet.MergeCells("A2:C2");
sheet.FreezePanes("A2");
if (sheet.FrozenPane != "A2" || sheet.MergedRanges.Count != 1)
    throw new InvalidOperationException("Worksheet metadata binding failed.");
var table = sheet.AddTable("SmokeTable", "A1:C1");
table.AddColumn("Value");
if (table.Name != "SmokeTable" || table.ColumnCount != 1)
    throw new InvalidOperationException("Table binding failed.");
var validation = sheet.AddListValidation("A1:A3", "\"Open,Closed\"");
validation.AllowBlank = true;
validation.SetPrompt("Status", "Choose a status");
validation.SetError("Invalid", "Use an allowed status");
var rule = sheet.AddFormulaRule("A1:A3", "A1>0");
rule.Priority = 1;
rule.StopIfTrue = true;
rule.SetFontColor("FFFF0000");
sheet.AddDataBarRule("B1:B3");

var output = Path.Combine(Environment.CurrentDirectory, "csharp-test.xlsx");
if (!workbook.Save(output))
    throw new InvalidOperationException("XL++ failed to save the workbook.");

Console.WriteLine($"C# binding OK: {output}");
