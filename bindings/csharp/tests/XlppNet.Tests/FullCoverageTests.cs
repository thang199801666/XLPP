using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using XLPP;

namespace XlppNet.Tests
{
    public class FullCoverageTests
    {
        private static string TempPath(string name) =>
            Path.Combine(Path.GetTempPath(), $"xlpp-{name}-{Guid.NewGuid():N}.xlsx");

        [Fact]
        public void Workbook_CopySheet_And_Date1904()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Src");
            ws["A1"].Value = "hello";
            var copy = wb.CopyWorksheet(ws, "Dst");
            Assert.Equal("Dst", copy.Name);
            Assert.Equal(2, wb.SheetCount);
            Assert.Equal(0, wb.IndexOf(ws));
            wb.Date1904 = true;
            Assert.True(wb.Date1904);
            wb.Clear();
            Assert.Equal(0, wb.SheetCount);
        }

        [Fact]
        public void Workbook_Collections_And_PreservedRelationships()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("First");
            wb.AddWorksheet("Second");
            Assert.Equal(new[] { "First", "Second" }, wb.Worksheets.Select(x => x.Name));
            Assert.Equal(0, wb.Index(wb.Worksheets[0]));
            Assert.Empty(wb.NamedStyles);
            Assert.Empty(wb.DefinedNames);
            Assert.Empty(wb.PreservedRelationships);
        }

        [Fact]
        public void Worksheet_Collections_StaySynchronized()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.AddTable("Table1", "A1:B2");
            ws.AddChart(ChartType.Line);
            ws.MergeCells("A1:B1");
            Assert.Single(ws.Tables);
            Assert.Single(ws.Charts);
            Assert.Empty(ws.Images);
            Assert.Empty(ws.PivotTables);
            Assert.Contains("A1:B1", ws.MergedRanges);
        }

        [Fact]
        public void Workbook_NamedStyle_And_DefinedName()
        {
            using var wb = new Workbook();
            var ns = wb.AddNamedStyle("MyStyle");
            ns.Style.Font.SetBold(true);
            ns.Style.Font.SetSize(14);
            Assert.Equal("MyStyle", wb.GetNamedStyle("MyStyle")!.Name);
            Assert.Equal(1, wb.NamedStyleCount);

            var dn = wb.AddDefinedName("MyRange", "Sheet1!$A$1:$A$10");
            Assert.Equal("MyRange", dn.Name);
            dn.Hidden = true;
            Assert.True(dn.Hidden);
            Assert.Equal(1, wb.DefinedNameCount);
            var got = wb.GetDefinedName("MyRange");
            Assert.Equal("Sheet1!$A$1:$A$10", got!.Value);
        }

        [Fact]
        public void Properties_AllFields_RoundTrip()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("S");
            wb.Properties.Title = "Title";
            wb.Properties.Subject = "Subject";
            wb.Properties.Creator = "Creator";
            wb.Properties.Description = "Desc";
            wb.Properties.Keywords = "k1,k2";
            wb.Properties.Category = "Cat";
            wb.Properties.LastModifiedBy = "User";
            var path = TempPath("props");
            Assert.True(wb.Save(path));
            using var wb2 = new Workbook();
            Assert.True(wb2.Load(path));
            Assert.Equal("Title", wb2.Properties.Title);
            Assert.Equal("Subject", wb2.Properties.Subject);
            Assert.Equal("Desc", wb2.Properties.Description);
            Assert.Equal("k1,k2", wb2.Properties.Keywords);
            Assert.Equal("Cat", wb2.Properties.Category);
            Assert.Equal("User", wb2.Properties.LastModifiedBy);
            File.Delete(path);
        }

        [Fact]
        public void Cell_Date_Error_And_NumberFormat()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("S");
            ws["A1"].Value = new DateTime(2024, 3, 15);
            Assert.True(ws["A1"].IsDate);
            var d = ws["A1"].Date;
            Assert.Equal(2024, d!.Value.Year);
            Assert.Equal(3, d.Value.Month);
            Assert.Equal(15, d.Value.Day);

            ws["B1"].SetError(CellError.DivisionByZero);
            Assert.True(ws["B1"].IsError);
            Assert.Equal(CellError.DivisionByZero, ws["B1"].Error);

            ws["C1"].NumberFormat = "0.00";
            Assert.Equal("0.00", ws["C1"].NumberFormat);
            ws["C1"].Value = 3.14159;
            Assert.Equal(3.14159, ws["C1"].Value);
        }

        [Fact]
        public void Cell_Formulas_Advanced()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("S");
            ws["A1"].SetArrayFormula("SUM(A2:A10)", "A1");
            Assert.True(ws["A1"].HasFormula);
            ws["A2"].SetDynamicArrayFormula("_xlfn.SORT(A1:A10)", "A2");
            Assert.Contains("SORT", ws["A2"].Formula!);
            ws["A3"].SetSharedFormula("A1+B1", 0, "A3");
            Assert.Equal("A1+B1", ws["A3"].Formula);
            ws["A3"].ClearFormula();
            Assert.False(ws["A3"].HasFormula);
        }

        [Fact]
        public void Cell_Hyperlink_And_Comment()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("S");
            var c = ws["A1"];
            c.SetHyperlink("https://example.com", "Example", "Tip", true);
            Assert.True(c.HasHyperlink);
            Assert.Equal("https://example.com", c.Hyperlink.Target);
            c.ClearHyperlink();
            Assert.False(c.HasHyperlink);

            c.SetComment("Note", "Author");
            Assert.True(c.HasComment);
            Assert.Equal("Note", c.Comment.Text);
            Assert.Equal("Author", c.Comment.Author);
            c.ClearComment();
            Assert.False(c.HasComment);
        }

        [Fact]
        public void Worksheet_Range_Dimensions_And_Merged()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("S");
            ws["A1"].Value = 1.0;
            ws["B1"].Value = 2.0;
            ws["C1"].Value = 3.0;
            var r = ws.Range("A1:C1");
            Assert.Equal("A1:C1", r.Address);
            Assert.Equal(1UL, r.RowCount);
            Assert.Equal(3UL, r.ColumnCount);
            Assert.Equal(3, r.Values.Length);
            ws.MergeCells("A1:C1");
            Assert.True(ws.IsMerged("A1"));
            Assert.Equal(1, ws.MergedCount);
            Assert.Equal("A1:C1", ws.MergedAt(0));
            ws.UnmergeCells("A1:C1");
            Assert.False(ws.IsMerged("A1"));

            var rd = ws.RowDimension(1);
            rd.Height = 30;
            Assert.Equal(30, rd.Height);
            var cd = ws.ColumnDimension(1);
            cd.Width = 20;
            Assert.Equal(20, cd.Width);
        }

        [Fact]
        public void Worksheet_Table_Image_Chart_Pivot()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Data");
            ws.AppendRow("Q1", "10");
            var table = ws.AddTable("Sales", "A1:B1");
            Assert.Equal("Sales", table.Name);
            table.DisplayName = "SalesDisp";
            table.ShowHeaderRow = false;
            table.ShowTotalsRow = true;
            table.StyleInfo.Name = "TableStyleMedium4";
            table.StyleInfo.ShowRowStripes = false;
            table.AddColumn("Value");
            Assert.Equal(1, table.ColumnCount);
            Assert.False(table.ShowHeaderRow);
            Assert.True(table.ShowTotalsRow);

            var chart = ws.AddChart(ChartType.Bar);
            chart.Title = "Sales Chart";
            chart.Width = 800;
            chart.Height = 500;
            chart.Grouping = ChartGrouping.Stacked;
            chart.SetLegendPosition("b");
            var series = chart.AddSeries("Units");
            series.ValuesReference = "'Data'!$B$1:$B$1";
            series.CategoriesReference = "'Data'!$A$1:$A$1";
            Assert.Equal(1, chart.SeriesCount);
            Assert.Equal(1, ws.ChartCount);

            var pivot = ws.AddPivotTable("Pivot1", "D1");
            pivot.AddRowField("Region");
            pivot.AddDataField();
            Assert.Equal(1, ws.PivotCount);
        }

        [Fact]
        public void Worksheet_AutoFilter_PageSetup_Protection()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("S");
            ws.AutoFilter.Reference = "A1:B2";
            Assert.True(ws.AutoFilter.Enabled);
            var col = ws.AutoFilter.Column(0);
            col.AddValue("foo");
            Assert.Equal(1, col.ValueCount);

            ws.PageSetup.Orientation = PageOrientation.Landscape;
            ws.PageSetup.Scale = 90;
            Assert.Equal(PageOrientation.Landscape, ws.PageSetup.Orientation);
            ws.PageMargins.Left = 0.5;
            Assert.Equal(0.5, ws.PageMargins.Left);
            ws.PrintOptions.GridLines = true;
            Assert.True(ws.PrintOptions.GridLines);
            ws.HeaderFooter.OddHeader = "&CHead";
            ws.HeaderFooter.DifferentOddEven = true;
            Assert.True(ws.HeaderFooter.DifferentOddEven);

            ws.Protection.Enabled = true;
            ws.Protection.InsertRows = false;
            Assert.True(ws.Protection.Enabled);

            var path = TempPath("af");
            Assert.True(wb.Save(path));
            File.Delete(path);
        }

        [Fact]
        public void ConditionalFormatting_And_DataValidation()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("S");
            var cf = ws.ConditionalFormatting;
            var entry = cf.AddEntry("A1:A10");
            var rule = entry.AddRule(ConditionalRuleType.CellIs);
            rule.AddFormula("5");
            rule.Operator = ConditionalOperator.GreaterThan;
            Assert.Equal(1, cf.EntryCount);
            Assert.Equal(1, entry.RuleCount);
            Assert.Equal(ConditionalRuleType.CellIs, rule.Type);

            var dv = ws.DataValidations.Add(DataValidationType.Whole, "B1:B10");
            dv.Operator = DataValidationOperator.Between;
            dv.Formula1 = "1";
            dv.Formula2 = "100";
            Assert.Equal(1, ws.DataValidations.Count);
            Assert.Equal(DataValidationType.Whole, dv.Type);
        }

        [Fact]
        public void Streaming_Reader_Writer_RoundTrip()
        {
            var path = TempPath("stream");
            using (var writer = new StreamingWorkbookWriter(path))
            {
                var sheet = writer.AddWorksheet("Big");
                for (int i = 0; i < 100; i++)
                    sheet.AppendRow(new[] { $"row{i}", i.ToString() });
                Assert.Equal(100UL, writer.RowCount(sheet));
                writer.Close();
            }

            using (var reader = new StreamingWorkbookReader(path))
            {
                Assert.Equal(1, reader.SheetCount);
                Assert.Single(reader.SheetNames);
                var rows = reader.ReadSheet(0);
                Assert.Equal(100, rows.Count);
                Assert.Equal(2, rows[0].Cells.Count);
                Assert.Equal("row0", rows[0].Cells[0].Value);
            }
            File.Delete(path);
        }

        [Fact]
        public void Worksheet_SheetView_And_PrintTitles()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("S");
            ws.SheetView.ZoomScale = 120;
            ws.SheetView.ShowGridLines = false;
            Assert.Equal(120, ws.SheetView.ZoomScale);
            Assert.False(ws.SheetView.ShowGridLines);
            ws.PrintArea = "A1:D20";
            Assert.Equal("A1:D20", ws.PrintArea);
            ws.PrintTitlesRows = "$1:$1";
            Assert.Equal("$1:$1", ws.PrintTitlesRows);
        }

        [Fact]
        public void Workbook_CalcAndCustomProperties()
        {
            using var wb = new Workbook();
            wb.CalcProperties.FullCalcOnLoad = true;
            Assert.True(wb.CalcProperties.FullCalcOnLoad);
            var cp = wb.CustomProperties;
            var prop = cp.Add("myKey", "myValue");
            Assert.Equal("myKey", prop.Name);
            Assert.Equal("myValue", prop.Value);
            Assert.Equal(1, cp.Count);
        }
    }
}
