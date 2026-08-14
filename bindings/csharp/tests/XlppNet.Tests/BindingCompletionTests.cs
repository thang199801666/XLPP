using System;
using System.IO;
using Xunit;
using XLPP;

namespace XlppNet.Tests
{
    public class BindingCompletionTests
    {
        [Fact]
        public void RichText_RangeCoordinates_And_ScalarAppend_AreBound()
        {
            using var workbook = new Workbook();
            var sheet = workbook.AddWorksheet("Data");
            sheet.AppendScalars(ScalarCellValue.Number(12.5), ScalarCellValue.Boolean(true), ScalarCellValue.Empty());
            Assert.Equal(12.5, sheet["A1"].Value);
            Assert.Equal(true, sheet["B1"].Value);

            var range = sheet.Range(1, 1, 1, 3);
            Assert.Equal("A1:C1", range.Address);

            using var richText = new RichText();
            var first = richText.AddRun("XL");
            first.Bold = true;
            first.Color = "FF3366CC";
            first.FontName = "Aptos";
            first.Size = 14;
            var second = richText.AddRun("++");
            second.Italic = true;
            sheet["A2"].SetRichText(richText);

            Assert.True(sheet["A2"].HasRichText);
            var stored = sheet["A2"].RichText;
            Assert.Equal("XL++", stored.PlainText);
            Assert.Equal(2, stored.RunCount);
            Assert.True(stored.RunAt(0).Bold);
            Assert.Equal("FF3366CC", stored.RunAt(0).Color);
            Assert.True(stored.RunAt(1).Italic);
        }

        [Fact]
        public void EnterpriseInspection_And_ControlledEdits_AreBound()
        {
            using var workbook = new Workbook();
            workbook.AddWorksheet("Data");

            var inspection = workbook.InspectEnterpriseFeatures();
            Assert.Empty(inspection.Features);
            Assert.Empty(inspection.Warnings);
            Assert.False(inspection.Has(EnterpriseFeatureKind.PowerQuery));

            var edit = workbook.SetConnectionRefreshOnLoad("missing", true);
            Assert.Equal(0UL, edit.Matched);
            Assert.Equal(0UL, edit.Modified);
            Assert.False(edit.Success);
        }

        [Fact]
        public void OfficeEncryptionInspection_IsAvailable()
        {
            var path = Path.Combine(Path.GetTempPath(), $"xlpp-csharp-encryption-{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var workbook = new Workbook())
                {
                    workbook.AddWorksheet("Data");
                    Assert.True(workbook.Save(path));
                }

                var info = Workbook.InspectOfficeEncryption(path);
                Assert.False(info.Encrypted);
                Assert.Equal(OfficeEncryptionMode.None, info.Mode);
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void FormulaMetadata_References_And_DrawingModels_MatchNative()
        {
            var reference = CellReference.Parse("$XFD$1048576");
            Assert.Equal(16384UL, reference.Column);
            Assert.Equal("XFD1048576", reference.Address);

            using var workbook = new Workbook();
            var sheet = workbook.AddWorksheet("Data");
            var cell = sheet["A1"];
            cell.Formula = "=B1";
            cell.FormulaMetadata.Type = FormulaType.Shared;
            cell.FormulaMetadata.Reference = "A1:A3";
            cell.FormulaMetadata.SharedIndex = 7;
            cell.FormulaMetadata.CalculateOnLoad = true;
            Assert.Equal(FormulaType.Shared, cell.FormulaMetadata.Type);
            Assert.Equal("A1:A3", cell.FormulaMetadata.Reference);
            Assert.Equal(7UL, cell.FormulaMetadata.SharedIndex);
            Assert.True(cell.FormulaMetadata.CalculateOnLoad);

            var translated = ReferenceTranslator.TranslateFormulaReferences("=SUM(A1:A3)", "Data", "Data",
                StructuralEditKind.InsertRows, 1, 1);
            Assert.Equal("=SUM(A2:A4)", translated.Value);
            Assert.True(translated.Changed);

            var chart = sheet.AddChart(ChartType.Line);
            chart.AnchorInfo = new DrawingAnchorInfo
            {
                Type = DrawingAnchorType.TwoCell,
                From = new DrawingMarker { Row = 2, Column = 3, RowOffsetEmu = 10 },
                To = new DrawingMarker { Row = 12, Column = 9 },
                EditAs = "twoCell"
            };
            Assert.Equal(DrawingAnchorType.TwoCell, chart.AnchorInfo.Type);
            Assert.Equal(2UL, chart.AnchorInfo.From.Row);
            Assert.Equal("twoCell", chart.AnchorInfo.EditAs);
        }

        [Fact]
        public void ChartModels_And_StreamingWorksheetReader_AreUsable()
        {
            var cache = new ChartSeriesCache
            {
                Present = true,
                Numeric = true,
                PointCount = 2,
                Points = { new ChartCachePoint { Index = 0, Value = "1" }, new ChartCachePoint { Index = 1, Value = "2" } }
            };
            Assert.True(cache.Valid());
            Assert.Equal(2UL, cache.EffectivePointCount);

            var palette = new ChartThemePalette { Present = true };
            palette.Colors.Add(new ChartThemeColor { Name = "accent1", SRgb = "336699" });
            Assert.Equal("336699", palette.ResolveFinalRgb(new ChartColor { Kind = ChartColorKind.Scheme, Value = "accent1" }));

            var path = Path.Combine(Path.GetTempPath(), $"xlpp-csharp-stream-{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var writer = new StreamingWorkbookWriter(path))
                {
                    var output = writer.AddWorksheet("Rows");
                    output.AppendRow("one", "two");
                }
                using var reader = new StreamingWorkbookReader(path, new StreamingReaderOptions { ValidateCellReferences = true });
                var input = reader.Worksheet("Rows");
                var count = 0;
                input.ForEachRow(row => { count++; return true; });
                Assert.Equal(1, count);
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Fact]
        public void NestedChartValues_AreMarshalledThroughTheNativeAbi()
        {
            using var workbook = new Workbook();
            var sheet = workbook.AddWorksheet("Charts");
            const string missingImportedChart = "missing-chart";
            var line = new ChartLineFormat
            {
                Present = true, WidthPoints = 1.5, Dash = "dash",
                Color = new ChartColor { Kind = ChartColorKind.SRgb, Value = "336699",
                    Transforms = { new ChartColorTransform { Kind = ChartColorTransformKind.Tint, Value = 25000 } } },
                CustomDash = { new ChartCustomDashStop { Dash = 2, Space = 1 } }
            };
            var fill = new ChartFillFormat
            {
                Present = true, Kind = ChartFillKind.Gradient, GradientAngleDegrees = 45,
                GradientStops = { new ChartGradientStop { Position = 0, Color = new ChartColor { Kind = ChartColorKind.SRgb, Value = "FFFFFF" } } }
            };
            var rich = new ChartRichText { Present = true, Runs = { new ChartTextRun { Text = "Revenue", Bold = true, Color = new ChartColor { Kind = ChartColorKind.Scheme, Value = "accent1" } } } };
            var cache = new ChartSeriesCache { Present = true, Numeric = true, PointCount = 1,
                Points = { new ChartCachePoint { Index = 0, Value = "42" } } };

            Assert.False(sheet.SetChartTitleRichText(missingImportedChart, rich));
            Assert.False(sheet.SetChartAxisScaling(missingImportedChart, 1, new ChartAxisScaling { HasMinimum = true, Minimum = 0, ReverseOrder = true }));
            Assert.False(sheet.SetChartAxisDisplayUnits(missingImportedChart, 1, new ChartDisplayUnits { Present = true, BuiltInUnit = "thousands", ShowLabel = true, LabelRichText = rich }));
            Assert.False(sheet.SetChartAxisLineFormat(missingImportedChart, 1, line));
            Assert.False(sheet.SetChartAxisGridlineFormat(missingImportedChart, 1, true, line));
            Assert.False(sheet.SetChartAreaFillFormat(missingImportedChart, fill));
            Assert.False(sheet.SetChartPlotAreaLineFormat(missingImportedChart, line));
            Assert.False(sheet.SetChartSeriesValueCache(missingImportedChart, 0, cache));
            Assert.False(sheet.SetChartSeriesLineFormat(missingImportedChart, 0, line));
            Assert.False(sheet.SetChartSeriesFillFormat(missingImportedChart, 0, fill));
            Assert.False(sheet.AddChartSeriesTrendline(missingImportedChart, 0, new ChartTrendline { LineFormat = line }));
            Assert.False(sheet.SetChartSeriesErrorBars(missingImportedChart, 0, new ChartErrorBars { Value = 2, LineFormat = line }));
        }
    }
}
