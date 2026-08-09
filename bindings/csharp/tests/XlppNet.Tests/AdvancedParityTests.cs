using System;
using Xunit;
using XLPP;

namespace XlppNet.Tests
{
    public class AdvancedParityTests
    {
        [Fact]
        public void Formula_ExternalResolver_And_DependencyGraph()
        {
            using var wb = new Workbook();
            var calc = wb.AddWorksheet("Calc");
            calc["A1"].Formula = "='[External.xlsx]Data'!A1+1";
            var options = new CalculationOptions
            {
                ExternalReferenceResolver = (book, sheet, address) =>
                    book == "External.xlsx" && sheet == "Data" && address == "A1"
                        ? ExternalReferenceValue.Number(41.0)
                        : null
            };
            var report = wb.CalculateFormulas(options);
            Assert.True(report.Success);
            Assert.Equal(42.0, calc["A1"].Value);
            Assert.True(report.ExternalReferencesResolved >= 1);

            var consumer = wb.AddWorksheet("Consumer");
            consumer["A1"].Formula = "=Calc!A1";
            using var graph = wb.DependencyGraph();
            Assert.True(graph.DependsOn("Consumer", "A1", "Calc", "A1"));
            Assert.NotEmpty(graph.PrecedentsOf("Consumer", "A1"));
            Assert.NotEmpty(graph.DependentsOf("Calc", "A1"));
        }

        [Fact]
        public void MemoryIo_LoadCallbacks_And_SafetyOptions()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("Data")["A1"].Value = "roundtrip";
            var payload = wb.SaveBytes(new SaveOptions { AtomicWrite = true, ValidateBeforeSave = true });
            Assert.NotEmpty(payload);

            ulong progressCalls = 0;
            using var loaded = new Workbook();
            Assert.True(loaded.Load(payload, new LoadOptions
            {
                Cancel = () => false,
                Progress = (_, _) => ++progressCalls
            }));
            Assert.Equal("roundtrip", loaded.GetWorksheet("Data")!["A1"].Value);
            Assert.True(progressCalls > 0);
        }

        [Fact]
        public void Structural_Rename_Validation_ChartTracking_And_Vba()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("Data")["A1"].Value = 5.0;
            var calc = wb.AddWorksheet("Calc");
            calc["A1"].Formula = "=Data!A1";

            var rename = wb.RenameWorksheet("Data", "Input");
            Assert.True(rename.Success);
            Assert.Contains("Input", calc["A1"].Formula!);

            var structural = wb.ApplyStructuralEdit("Input", StructuralEditKind.InsertRows, 1, 1,
                new StructuralEditOptions { Transactional = true });
            Assert.True(structural.Success);
            Assert.True(wb.Validate().Ok);

            Assert.True(wb.SynchronizeChartCaches().Success);
            wb.ResetChartCacheDependencyTracking();
            Assert.Equal(0UL, wb.TrackedChartCacheDependencyCount);

            var input = wb.GetWorksheet("Input")!;
            input.VbaCodeName = "InputSheet";
            wb.VbaProjectProperties = new VbaProjectProperties("ParityMacros", "C# VBA parity", "parity.chm", 9, "CSharpBinding = 1");
            Assert.True(wb.SetVbaDocumentModuleText("InputSheet", "Private Sub Worksheet_Activate()\r\nEnd Sub\r\n"));
            Assert.True(wb.SetVbaClassModuleText("ParityClass", "Public Function Value() As Long\r\nValue = 1\r\nEnd Function\r\n", true, true));
            Assert.True(wb.SetVbaModuleText("Module1", "Sub Hello()\r\nEnd Sub\r\n"));
            Assert.True(wb.HasVbaProject);
            Assert.True(wb.VbaSourceEditable);
            Assert.Equal("ParityMacros", wb.VbaProjectProperties.Name);
            Assert.Equal("CSharpBinding = 1", wb.VbaProjectProperties.Constants);
            Assert.Contains("Sub Hello()", wb.VbaModuleText("Module1")!);
            Assert.Contains(wb.VbaModules, m => m.Name == "Module1");
            Assert.Contains(wb.VbaModules, m => m.Name == "ParityClass" && m.Type == VbaModuleType.Class && m.ReadOnly && m.PrivateModule);
            Assert.True(wb.VbaProjectBytes.Length > 512);
            Assert.False(wb.HasVbaSignature);
            Assert.True(wb.RemoveVbaModule("Module1"));
        }
    }

        [Fact]
        public void ScopedDefinedNamesTrackNativeSemantics()
        {
            using var wb = new Workbook();
            wb.AddWorksheet("S1");
            wb.AddWorksheet("S2");
            wb.AddDefinedName("Rate", "S1!$A$1", 0);
            wb.AddDefinedName("Rate", "S2!$A$1", 1);
            var n1 = wb.GetDefinedName("Rate", 0)!;
            var n2 = wb.GetDefinedName("Rate", 1)!;
            Assert.Equal((ulong)0, n1.LocalSheetId);
            Assert.Equal((ulong)1, n2.LocalSheetId);
            Assert.Equal("S1!$A$1", n1.Value);
            Assert.Equal("S2!$A$1", n2.Value);
        }
}
