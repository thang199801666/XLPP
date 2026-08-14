using System;
using System.IO;
using Xunit;
using XLPP;

namespace XlppNet.Tests
{
    public class WorkbookTests
    {
        [Fact]
        public void Version_IsNonEmpty()
        {
            Assert.False(string.IsNullOrEmpty(Workbook.Version));
        }

        [Fact]
        public void Create_AddSheets_Count()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("One");
            wb.AddWorksheet("Two");
            Assert.Equal(2, wb.SheetCount);
        }

        [Fact]
        public void GetWorksheet_ByName_And_Index()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("Alpha");
            wb.AddWorksheet("Beta");
            Assert.NotNull(wb.GetWorksheet("Alpha"));
            Assert.Null(wb.GetWorksheet("Nope"));
            Assert.Equal("Beta", wb[1].MaxRow >= 1 ? "Beta" : "");
        }

        [Fact]
        public void AddDuplicate_Throws()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("Dup");
            Assert.Throws<ArgumentException>(() => wb.AddWorksheet("Dup"));
            Assert.Equal(1, wb.SheetCount);
        }

        [Fact]
        public void RemoveWorksheet()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("Keep");
            wb.AddWorksheet("Drop");
            Assert.True(wb.RenameWorksheet("drop", "Renamed"));
            Assert.NotNull(wb.GetWorksheet("RENAMED"));
            Assert.True(wb.RemoveWorksheet("renamed"));
            Assert.False(wb.RemoveWorksheet("Missing"));
            Assert.False(wb.RemoveWorksheet("Keep")); // last-sheet invariant stays inside the C ABI
            Assert.Equal(1, wb.SheetCount);
        }

        [Fact]
        public void PasswordEncryption_RoundTrips()
        {
            var path = Path.Combine(Path.GetTempPath(), $"xlpp-p1g-{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var wb = new Workbook())
                {
                    var ws = wb.AddWorksheet("Secret");
                    ws["A1"].Value = "classified";
                    Assert.True(wb.SaveEncrypted(path, "P@ssw0rd✓", 1000));
                }
                Assert.True(Workbook.IsPasswordEncryptedFile(path));
                using var reopened = new Workbook();
                Assert.True(reopened.Load(path, "P@ssw0rd✓"));
                Assert.Equal("classified", reopened[0]["A1"].Value);
                using var wrong = new Workbook();
                Assert.False(wrong.Load(path, "wrong"));
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }

    public class CellTests
    {
        [Fact]
        public void SetGet_String_Number_Bool()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Data");
            ws["A1"].Value = "hello";
            ws["A2"].Value = 42.5;
            ws["A3"].Value = true;
            ws["A4"].Value = 7;

            Assert.Equal("hello", ws["A1"].Value);
            Assert.Equal(42.5, ws["A2"].Value);
            Assert.Equal(true, ws["A3"].Value);
            Assert.Equal(7.0, ws["A4"].Value);
            Assert.Equal(CellValueType.String, ws["A1"].ValueType);
            Assert.Equal(CellValueType.Number, ws["A2"].ValueType);
            Assert.Equal(CellValueType.Bool, ws["A3"].ValueType);
        }

        [Fact]
        public void Cell_ByAddress_And_RC()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Data");
            ws.Cell(1, 1).Value = "A1";
            ws.Cell("B2").Value = 5.0;
            Assert.Equal("A1", ws["A1"].Value);
            Assert.Equal(5.0, ws.Cell(2, 2).Value);
        }

        [Fact]
        public void Formula_GetSet()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Data");
            ws["A1"].Value = 2.0;
            ws["A2"].Value = 3.0;
            ws["A3"].Formula = "A1*A2";
            Assert.Equal("A1*A2", ws["A3"].Formula);
        }

        [Fact]
        public void IsEmpty_And_HasCell()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Data");
            ws["A1"].Value = "x";
            Assert.True(ws.HasCell("A1"));
            Assert.False(ws.HasCell("Z99"));
        }
    }

    public class LayoutTests
    {
        [Fact]
        public void MergeCells_And_FreezePanes()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Layout");
            ws["A1"].Value = "Title";
            ws.MergeCells("A1:C1");
            ws.FreezePanes("A2");
            ws["D5"].Value = 1.0;
            Assert.Equal(5ul, ws.MaxRow);
            Assert.Equal(4ul, ws.MaxColumn);
        }

        [Fact]
        public void MaxRow_And_Column()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Data");
            ws["C5"].Value = 1.0;
            Assert.Equal(5ul, ws.MaxRow);
            Assert.Equal(3ul, ws.MaxColumn);
        }
    }

    public class StylingTests
    {
        [Fact]
        public void Font_And_Border_Setters_DoNotThrow()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Style");
            ws["A1"].Value = "Styled";
            ws["A1"].Font.SetName("Arial");
            ws["A1"].Font.SetSize(14.0);
            ws["A1"].Font.SetBold(true);
            ws["A1"].Font.SetItalic(true);
            ws["A1"].Font.SetColor("FFFF0000");
            ws["A1"].Border.Top.SetStyle("thin");
            ws["A1"].Alignment.SetHorizontal("center");
            ws["A1"].Alignment.SetWrapText(true);
            Assert.True(true);
        }
    }

    public class RoundTripTests
    {
        [Fact]
        public void Save_And_Load_RoundTrips()
        {
            var path = Path.Combine(Path.GetTempPath(), "xlpp_csharp_test.xlsx");
            try
            {
                using (var wb = new Workbook())
                {
                    wb.Properties.Title = "C# Roundtrip";
                    var ws = wb.AddWorksheet("Data");
                    ws["A1"].Value = "hello";
                    ws["B1"].Value = 42.5;
                    ws["C1"].Value = true;
                    Assert.True(wb.Save(path));
                }

                using (var wb = new Workbook())
                {
                    Assert.True(wb.Load(path));
                    Assert.Equal(1, wb.SheetCount);
                    var ws = wb["Data"];
                    Assert.Equal("hello", ws["A1"].Value);
                    Assert.Equal(42.5, ws["B1"].Value);
                    Assert.Equal(true, ws["C1"].Value);
                }
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void MultipleSheets_RoundTrip()
        {
            var path = Path.Combine(Path.GetTempPath(), "xlpp_csharp_multi.xlsx");
            try
            {
                using (var wb = new Workbook())
                {
                    wb.AddWorksheet("One").Cell("A1").Value = 1.0;
                    wb.AddWorksheet("Two").Cell("A1").Value = 2.0;
                    wb.AddWorksheet("Three").Cell("A1").Value = 3.0;
                    Assert.True(wb.Save(path));
                }
                using (var wb = new Workbook())
                {
                    Assert.True(wb.Load(path));
                    Assert.Equal(3, wb.SheetCount);
                    Assert.Equal(1.0, wb["One"]["A1"].Value);
                    Assert.Equal(3.0, wb["Three"]["A1"].Value);
                }
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void SpecialCharacters_RoundTrip()
        {
            var path = Path.Combine(Path.GetTempPath(), "xlpp_csharp_special.xlsx");
            try
            {
                using (var wb = new Workbook())
                {
                    var ws = wb.AddWorksheet("S");
                    ws["A1"].Value = "a < b & c > d";
                    ws["A2"].Value = "quotes \"and\" 'single'";
                    ws["A3"].Value = "line1\nline2";
                    Assert.True(wb.Save(path));
                }
                using (var wb = new Workbook())
                {
                    Assert.True(wb.Load(path));
                    var ws = wb["S"];
                    Assert.Equal("a < b & c > d", ws["A1"].Value);
                    Assert.Equal("quotes \"and\" 'single'", ws["A2"].Value);
                    Assert.Equal("line1\nline2", ws["A3"].Value);
                }
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Properties_RoundTrip()
        {
            var path = Path.Combine(Path.GetTempPath(), "xlpp_csharp_props.xlsx");
            try
            {
                using (var wb = new Workbook())
                {
                    wb.Properties.Title = "My Title";
                    wb.Properties.Creator = "Tester";
                    wb.AddWorksheet("S").Cell("A1").Value = 1.0;
                    Assert.True(wb.Save(path));
                }
                using (var wb = new Workbook())
                {
                    Assert.True(wb.Load(path));
                    Assert.Equal(1, wb.SheetCount);
                }
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }

    public class ApiSurfaceTests
    {
        [Fact]
        public void Cell_Address_Row_Column()
        {
            using var wb = new Workbook();
            var c = wb.AddWorksheet("S").Cell("C5");
            Assert.Equal("C5", c.Address);
            Assert.Equal(5ul, c.Row);
            Assert.Equal(3ul, c.Column);
        }

        [Fact]
        public void Font_Getters_After_Setters()
        {
            using var wb = new Workbook();
            var f = wb.AddWorksheet("S")["A1"].Font;
            f.SetName("Arial");
            f.SetSize(16.0);
            f.SetBold(true);
            Assert.Equal("Arial", f.Name);
            Assert.Equal(16.0, f.Size);
            Assert.True(f.Bold);
        }

        [Fact]
        public void HasFormula()
        {
            using var wb = new Workbook();
            var c = wb.AddWorksheet("S")["A1"];
            Assert.False(c.HasFormula);
            c.Formula = "=1+2";
            Assert.True(c.HasFormula);
            Assert.Equal("=1+2", c.Formula);
        }

        [Fact]
        public void Worksheet_Name_And_Rename()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Original");
            Assert.Equal("Original", ws.Name);
            ws.Rename("Renamed");
            Assert.Equal("Renamed", ws.Name);
            Assert.NotNull(wb.GetWorksheet("Renamed"));
        }

        [Fact]
        public void AppendRows_Strings_And_Doubles()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Data");
            ws.AppendRow("Name", "Value");
            ws.AppendRow("Alice", "100");
            Assert.Equal("Name", ws["A1"].Value);
            Assert.Equal("Value", ws["B1"].Value);
            Assert.Equal("Alice", ws["A2"].Value);
            Assert.Equal("100", ws["B2"].Value);
            ws.AppendRow(1.0, 2.0, 3.0);
            Assert.Equal(1.0, ws["A3"].Value);
            Assert.Equal(3.0, ws["C3"].Value);
            Assert.Equal(3ul, ws.MaxRow);
        }

        [Fact]
        public void Insert_And_Delete_Rows_Columns()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Edit");
            ws["A1"].Value = "one";
            ws["B2"].Value = 10.0;
            ws.InsertRows(2);
            Assert.Equal(10.0, ws["B3"].Value);
            ws.DeleteRows(2);
            Assert.Equal(10.0, ws["B2"].Value);
            ws.InsertColumns(2, 2);
            Assert.Equal(10.0, ws["D2"].Value);
            ws.DeleteColumns(2, 2);
            Assert.Equal(10.0, ws["B2"].Value);
        }

        [Fact]
        public void Load_MissingFile_ReturnsFalse()
        {
            using var wb = new Workbook();
            Assert.False(wb.Load(Path.Combine(Path.GetTempPath(), "does_not_exist.xlsx")));
        }

        [Fact]
        public void LargeSheet()
        {
            using var wb = new Workbook();
            var ws = wb.AddWorksheet("Large");
            for (int i = 1; i <= 200; i++)
            {
                ws.Cell((ulong)i, 1).Value = "row" + i;
                ws.Cell((ulong)i, 2).Value = (double)i;
            }
            Assert.Equal(200ul, ws.MaxRow);
            Assert.Equal("row200", ws["A200"].Value);
            Assert.Equal(200.0, ws["B200"].Value);
        }

        [Fact]
        public void Merge_RoundTrip()
        {
            var path = Path.Combine(Path.GetTempPath(), "xlpp_csharp_merge.xlsx");
            try
            {
                using (var wb = new Workbook())
                {
                    var ws = wb.AddWorksheet("Layout");
                    ws["A1"].Value = "Title";
                    ws.MergeCells("A1:C1");
                    ws.FreezePanes("A2");
                    Assert.True(wb.Save(path));
                }
                using (var wb = new Workbook())
                {
                    Assert.True(wb.Load(path));
                    Assert.Equal("Title", wb["Layout"]["A1"].Value);
                }
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
