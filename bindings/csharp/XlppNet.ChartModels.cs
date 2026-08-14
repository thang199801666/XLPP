using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace XLPP
{
    public enum ChartColorTransformKind { Alpha, AlphaMod, AlphaOff, Tint, Shade, LumMod, LumOff, SatMod, SatOff }
    public enum ChartColorKind { None, SRgb, Scheme, System, Preset, Unknown }
    public enum ChartFillKind { None, NoFill, Solid, Gradient, Pattern }
    public enum ChartAxisKind { Category, Value, Date, Series }
    public enum ChartTrendlineType { Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage }
    public enum ChartErrorBarDirection { X, Y }
    public enum ChartErrorBarType { Both, Plus, Minus }
    public enum ChartErrorValueType { FixedValue, Percentage, StandardDeviation, StandardError, Custom }

    public sealed class ChartColorTransform
    {
        public ChartColorTransformKind Kind { get; set; }
        public int Value { get; set; } = 100000;
    }

    public sealed class ChartColor
    {
        public ChartColorKind Kind { get; set; }
        public string Value { get; set; } = string.Empty;
        public List<ChartColorTransform> Transforms { get; set; } = new();
        public bool Present => Kind != ChartColorKind.None && Value.Length != 0;
    }

    public sealed class ChartCustomDashStop
    {
        public double Dash { get; set; }
        public double Space { get; set; }
    }

    public sealed class ChartLineFormat
    {
        public bool Present { get; set; }
        public bool NoFill { get; set; }
        public ChartColor Color { get; set; } = new();
        public double WidthPoints { get; set; }
        public string Dash { get; set; } = string.Empty;
        public string Cap { get; set; } = string.Empty;
        public string Compound { get; set; } = string.Empty;
        public string Join { get; set; } = string.Empty;
        public List<ChartCustomDashStop> CustomDash { get; set; } = new();
    }

    public sealed class ChartGradientStop
    {
        public int Position { get; set; }
        public ChartColor Color { get; set; } = new();
    }

    public sealed class ChartFillFormat
    {
        public bool Present { get; set; }
        public bool NoFill { get; set; }
        public ChartColor Color { get; set; } = new();
        public ChartFillKind Kind { get; set; }
        public List<ChartGradientStop> GradientStops { get; set; } = new();
        public double GradientAngleDegrees { get; set; }
        public string Pattern { get; set; } = string.Empty;
        public ChartColor ForegroundColor { get; set; } = new();
        public ChartColor BackgroundColor { get; set; } = new();
    }

    public sealed class ChartTextRun
    {
        public string Text { get; set; } = string.Empty;
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public double FontSizePoints { get; set; }
        public string Typeface { get; set; } = string.Empty;
        public ChartColor Color { get; set; } = new();
    }

    public sealed class ChartTextStyle
    {
        public bool Present { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public double FontSizePoints { get; set; }
        public string Typeface { get; set; } = string.Empty;
        public ChartColor Color { get; set; } = new();
    }

    public sealed class ChartRichText
    {
        public bool Present { get; set; }
        public List<ChartTextRun> Runs { get; set; } = new();
        public string PlainText => string.Concat(Runs.Select(run => run.Text));
    }

    public sealed class ChartCachePoint
    {
        public ulong Index { get; set; }
        public string Value { get; set; } = string.Empty;
    }

    public sealed class ChartSeriesCache
    {
        public bool Present { get; set; }
        public bool Numeric { get; set; }
        public string FormatCode { get; set; } = string.Empty;
        public ulong PointCount { get; set; }
        public List<ChartCachePoint> Points { get; set; } = new();
        public ulong EffectivePointCount => Points.Count == 0 ? PointCount : Math.Max(PointCount, Points.Max(point => point.Index + 1));
        public bool Valid(bool allowSparse = true)
        {
            if (!Present) return true;
            if (Points.Select(point => point.Index).Distinct().Count() != Points.Count) return false;
            if (PointCount != 0 && PointCount < EffectivePointCount) return false;
            return allowSparse || (Points.Count == (int)EffectivePointCount && Points.Select((point, index) => point.Index == (ulong)index).All(value => value));
        }
    }

    public sealed class ChartManualLayout
    {
        public bool Present { get; set; }
        public string Target { get; set; } = string.Empty;
        public string XMode { get; set; } = string.Empty;
        public string YMode { get; set; } = string.Empty;
        public string WidthMode { get; set; } = string.Empty;
        public string HeightMode { get; set; } = string.Empty;
        public bool HasX { get; set; }
        public bool HasY { get; set; }
        public bool HasWidth { get; set; }
        public bool HasHeight { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }

    public sealed class ChartAxisScaling
    {
        public bool HasMinimum { get; set; }
        public bool HasMaximum { get; set; }
        public bool HasLogBase { get; set; }
        public double Minimum { get; set; }
        public double Maximum { get; set; }
        public double LogBase { get; set; }
        public bool ReverseOrder { get; set; }
    }

    public sealed class ChartDataLabelPoint
    {
        public ulong Index { get; set; }
        public bool Deleted { get; set; }
        public bool ShowLegendKey { get; set; }
        public bool ShowValue { get; set; }
        public bool ShowCategoryName { get; set; }
        public bool ShowSeriesName { get; set; }
        public bool ShowPercent { get; set; }
        public bool ShowBubbleSize { get; set; }
        public bool ShowLeaderLines { get; set; }
        public string Position { get; set; } = string.Empty;
        public string Separator { get; set; } = string.Empty;
        public ChartRichText RichText { get; set; } = new();
    }

    public sealed class ChartDataLabels
    {
        public bool Present { get; set; }
        public bool ShowLegendKey { get; set; }
        public bool ShowValue { get; set; }
        public bool ShowCategoryName { get; set; }
        public bool ShowSeriesName { get; set; }
        public bool ShowPercent { get; set; }
        public bool ShowBubbleSize { get; set; }
        public bool ShowLeaderLines { get; set; }
        public bool HasLeaderLines { get; set; }
        public ChartLineFormat LeaderLineFormat { get; set; } = new();
        public string Position { get; set; } = string.Empty;
        public string Separator { get; set; } = string.Empty;
        public List<ChartDataLabelPoint> Points { get; set; } = new();
    }

    public sealed class ChartDisplayUnits
    {
        public bool Present { get; set; }
        public string BuiltInUnit { get; set; } = string.Empty;
        public bool HasCustomUnit { get; set; }
        public double CustomUnit { get; set; }
        public bool ShowLabel { get; set; }
        public ChartRichText LabelRichText { get; set; } = new();
    }

    public sealed class ChartMarkerFormat
    {
        public bool Present { get; set; }
        public string Symbol { get; set; } = string.Empty;
        public int Size { get; set; }
        public ChartFillFormat Fill { get; set; } = new();
        public ChartLineFormat Line { get; set; } = new();
    }

    public sealed class ChartDataPointFormat
    {
        public ulong Index { get; set; }
        public ChartFillFormat Fill { get; set; } = new();
        public ChartLineFormat Line { get; set; } = new();
        public ChartMarkerFormat Marker { get; set; } = new();
    }

    public sealed class ChartDataTable
    {
        public bool Present { get; set; }
        public bool ShowHorizontalBorder { get; set; }
        public bool ShowVerticalBorder { get; set; }
        public bool ShowOutline { get; set; }
        public bool ShowLegendKeys { get; set; }
        public ChartFillFormat Fill { get; set; } = new();
        public ChartLineFormat Line { get; set; } = new();
        public ChartTextStyle TextStyle { get; set; } = new();
    }

    public sealed class ChartUpDownBars
    {
        public bool Present { get; set; }
        public int GapWidth { get; set; } = 150;
        public ChartFillFormat UpFill { get; set; } = new();
        public ChartLineFormat UpLine { get; set; } = new();
        public ChartFillFormat DownFill { get; set; } = new();
        public ChartLineFormat DownLine { get; set; } = new();
    }

    public sealed class ChartProjectedPieOptions
    {
        public bool Present { get; set; }
        public string OfPieType { get; set; } = "pie";
        public int GapWidth { get; set; } = 150;
        public string SplitType { get; set; } = "auto";
        public bool HasSplitPosition { get; set; }
        public double SplitPosition { get; set; }
        public List<int> CustomSplitPoints { get; set; } = new();
        public int SecondPlotSize { get; set; } = 75;
        public bool HasSeriesLines { get; set; }
        public ChartLineFormat SeriesLinesFormat { get; set; } = new();
    }

    public sealed class ChartView3D
    {
        public bool Present { get; set; }
        public bool HasRotationX { get; set; }
        public bool HasRotationY { get; set; }
        public bool HasHeightPercent { get; set; }
        public bool HasDepthPercent { get; set; }
        public bool HasRightAngleAxes { get; set; }
        public bool HasPerspective { get; set; }
        public int RotationX { get; set; }
        public int RotationY { get; set; }
        public int HeightPercent { get; set; } = 100;
        public int DepthPercent { get; set; } = 100;
        public bool RightAngleAxes { get; set; } = true;
        public int Perspective { get; set; } = 30;
    }

    public sealed class ChartWallFormat
    {
        public bool Present { get; set; }
        public bool HasThickness { get; set; }
        public int Thickness { get; set; }
        public ChartFillFormat Fill { get; set; } = new();
        public ChartLineFormat Line { get; set; } = new();
    }

    public sealed class ChartLegendFormat
    {
        public bool Present { get; set; }
        public bool Overlay { get; set; }
        public ChartManualLayout Layout { get; set; } = new();
        public ChartFillFormat Fill { get; set; } = new();
        public ChartLineFormat Line { get; set; } = new();
    }

    public sealed class ChartTrendline
    {
        public ChartTrendlineType Type { get; set; }
        public int Order { get; set; } = 2;
        public int Period { get; set; } = 2;
        public double Forward { get; set; }
        public double Backward { get; set; }
        public bool DisplayEquation { get; set; }
        public bool DisplayRSquared { get; set; }
        public ChartLineFormat LineFormat { get; set; } = new();
    }

    public sealed class ChartErrorBars
    {
        public ChartErrorBarDirection Direction { get; set; } = ChartErrorBarDirection.Y;
        public ChartErrorBarType BarType { get; set; }
        public ChartErrorValueType ValueType { get; set; }
        public double Value { get; set; }
        public bool NoEndCap { get; set; }
        public string PlusReference { get; set; } = string.Empty;
        public string MinusReference { get; set; } = string.Empty;
        public ChartLineFormat LineFormat { get; set; } = new();
    }

    public sealed class ChartAxis
    {
        public ChartAxisKind Kind { get; set; } = ChartAxisKind.Value;
        public ulong Id { get; set; }
        public ulong CrossAxisId { get; set; }
        public string Position { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public ChartRichText TitleRichText { get; set; } = new();
        public bool Secondary { get; set; }
        public string NumberFormat { get; set; } = string.Empty;
        public bool NumberFormatSourceLinked { get; set; } = true;
        public string MajorTickMark { get; set; } = string.Empty;
        public string MinorTickMark { get; set; } = string.Empty;
        public string TickLabelPosition { get; set; } = string.Empty;
        public bool HasMajorUnit { get; set; }
        public bool HasMinorUnit { get; set; }
        public double MajorUnit { get; set; }
        public double MinorUnit { get; set; }
        public string Crosses { get; set; } = string.Empty;
        public string CrossBetween { get; set; } = string.Empty;
        public bool HasCrossesAt { get; set; }
        public double CrossesAt { get; set; }
        public ChartAxisScaling Scaling { get; set; } = new();
        public ChartDisplayUnits DisplayUnits { get; set; } = new();
        public bool HasMajorGridlines { get; set; }
        public bool HasMinorGridlines { get; set; }
        public ChartLineFormat LineFormat { get; set; } = new();
        public ChartLineFormat MajorGridlineFormat { get; set; } = new();
        public ChartLineFormat MinorGridlineFormat { get; set; } = new();
    }

    public sealed class ChartPlot
    {
        public ChartType Type { get; set; }
        public ChartGrouping Grouping { get; set; }
        public List<ulong> AxisIds { get; set; } = new();
        public ulong FirstSeries { get; set; }
        public ulong SeriesCount { get; set; }
        public bool UsesSecondaryAxes { get; set; }
        public ChartDataLabels DataLabels { get; set; } = new();
        public bool HasDropLines { get; set; }
        public ChartLineFormat DropLinesFormat { get; set; } = new();
        public bool HasHighLowLines { get; set; }
        public ChartLineFormat HighLowLinesFormat { get; set; } = new();
        public ChartUpDownBars UpDownBars { get; set; } = new();
        public bool HasGapDepth { get; set; }
        public int GapDepth { get; set; } = 150;
        public bool HasWireframe { get; set; }
        public bool Wireframe { get; set; }
        public string Shape { get; set; } = string.Empty;
        public bool HasFirstSliceAngle { get; set; }
        public int FirstSliceAngle { get; set; }
        public bool HasHoleSize { get; set; }
        public int HoleSize { get; set; } = 10;
        public string RadarStyle { get; set; } = string.Empty;
        public ChartProjectedPieOptions ProjectedPie { get; set; } = new();
        public ChartBarDirection BarDirection { get; set; }
        public ChartScatterStyle ScatterStyle { get; set; } = ChartScatterStyle.Marker;
        public bool HasBubbleScale { get; set; }
        public int BubbleScale { get; set; } = 100;
        public bool ShowNegativeBubbles { get; set; }
        public ChartBubbleSizeRepresents BubbleSizeRepresents { get; set; }
        public bool Bubble3D { get; set; }
        public double HistogramBinWidth { get; set; }
        public int HistogramBinCount { get; set; }
        public bool HistogramAutomaticBins { get; set; } = true;
        public bool HistogramHasUnderflow { get; set; }
        public double HistogramUnderflow { get; set; }
        public bool HistogramHasOverflow { get; set; }
        public double HistogramOverflow { get; set; }
        public bool BoxWhiskerShowInnerPoints { get; set; } = true;
        public bool BoxWhiskerShowOutlierPoints { get; set; } = true;
        public bool BoxWhiskerShowMeanLine { get; set; }
        public bool BoxWhiskerShowMeanMarker { get; set; } = true;
        public bool BoxWhiskerQuartileInclusive { get; set; }
        public bool WaterfallShowConnectorLines { get; set; } = true;
        public string MapProjection { get; set; } = "automatic";
        public string MapArea { get; set; } = "automatic";
        public string MapLabels { get; set; } = "bestFitOnly";
    }

    public sealed class ChartThemeColor { public string Name { get; set; } = string.Empty; public string SRgb { get; set; } = string.Empty; }
    public sealed class ChartThemeFontScheme { public bool Present { get; set; } public string Name { get; set; } = string.Empty; public string MajorLatinTypeface { get; set; } = string.Empty; public string MinorLatinTypeface { get; set; } = string.Empty; }
    public sealed class ChartThemeEffectScheme { public bool Present { get; set; } public string Name { get; set; } = string.Empty; public ulong FillStyleCount { get; set; } public ulong LineStyleCount { get; set; } public ulong EffectStyleCount { get; set; } public ulong BackgroundFillStyleCount { get; set; } }
    public sealed class ChartStyleResources { public bool ChartStylePresent { get; set; } public bool ColorStylePresent { get; set; } public string ChartStylePart { get; set; } = string.Empty; public string ColorStylePart { get; set; } = string.Empty; }

    public readonly record struct ChartResolvedColor(bool Present, int Red, int Green, int Blue, double Alpha)
    {
        public string SRgb => Present ? $"{Math.Clamp(Red, 0, 255):X2}{Math.Clamp(Green, 0, 255):X2}{Math.Clamp(Blue, 0, 255):X2}" : string.Empty;
    }

    public sealed class ChartThemePalette
    {
        public bool Present { get; set; }
        public List<ChartThemeColor> Colors { get; set; } = new();
        public ChartThemeFontScheme FontScheme { get; set; } = new();
        public ChartThemeEffectScheme EffectScheme { get; set; } = new();
        public string BaseColor(string name) => Colors.FirstOrDefault(color => color.Name == name)?.SRgb ?? string.Empty;
        public string ResolveBase(ChartColor color) => color.Kind == ChartColorKind.SRgb ? color.Value : color.Kind == ChartColorKind.Scheme ? BaseColor(color.Value) : string.Empty;
        public ChartResolvedColor Resolve(ChartColor color)
        {
            var hex = ResolveBase(color);
            if (hex.Length != 6 || !int.TryParse(hex[..2], NumberStyles.HexNumber, null, out var r) ||
                !int.TryParse(hex.Substring(2, 2), NumberStyles.HexNumber, null, out var g) ||
                !int.TryParse(hex.Substring(4, 2), NumberStyles.HexNumber, null, out var b)) return default;
            var alpha = 1.0;
            foreach (var transform in color.Transforms)
                if (transform.Kind == ChartColorTransformKind.Alpha) alpha = Math.Clamp(transform.Value / 100000.0, 0, 1);
            return new ChartResolvedColor(true, r, g, b, alpha);
        }
        public string ResolveFinalRgb(ChartColor color) => Resolve(color).SRgb;
    }
}
