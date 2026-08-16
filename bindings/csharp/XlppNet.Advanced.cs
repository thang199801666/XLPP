using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace XLPP
{
    public enum CompressionStrategy { Default = 0, Filtered = 1, HuffmanOnly = 2, Rle = 3, Fixed = 4 }
    public enum FormulaDependencyKind { CellOrRange = 0, DefinedName = 1, Table = 2, ExternalReference = 3, VolatileReference = 4 }
    public enum WorkbookValidationSeverity { Warning = 0, Error = 1 }
    public enum VbaModuleType { Standard = 0, Document = 1, Class = 2 }
    public enum FormulaType { Normal = 0, Shared = 1, Array = 2, DataTable = 3, DynamicArray = 4 }
    public enum DrawingAnchorType { OneCell = 0, TwoCell = 1, Absolute = 2 }
    public enum EnterpriseFeatureKind
    {
        PivotChart = 0, Slicer = 1, SlicerCache = 2, Timeline = 3, TimelineCache = 4,
        OlapPivotCache = 5, DataModel = 6, PowerQuery = 7, SmartArt = 8, ActiveX = 9, VbaUserForm = 10
    }

    public readonly record struct EnterpriseFeatureRelationship(string SourcePart, string Id, string Type,
        string TargetMode, string TargetPart, bool Outgoing);

    public sealed record EnterpriseFeatureInfo(EnterpriseFeatureKind Kind, string PartName, string ContentType,
        string Name, string SourceName, string ConnectionId, string CacheId,
        IReadOnlyList<string> ReferencedPivotTables, IReadOnlyList<EnterpriseFeatureRelationship> Relationships,
        bool Binary, bool SemanticEditable, bool HasRefreshOnLoad, bool RefreshOnLoad);

    public sealed record EnterpriseFeatureInspection(IReadOnlyList<EnterpriseFeatureInfo> Features,
        IReadOnlyList<string> Warnings)
    {
        public int Count(EnterpriseFeatureKind kind) => Features.Count(feature => feature.Kind == kind);
        public bool Has(EnterpriseFeatureKind kind) => Features.Any(feature => feature.Kind == kind);
    }

    public sealed record EnterpriseEditReport(ulong Matched, ulong Modified, IReadOnlyList<string> Warnings, bool Success);

    public readonly record struct CellReference(ulong Row, ulong Column)
    {
        public static CellReference Parse(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) throw new ArgumentException("Cell address cannot be empty", nameof(address));
            var value = address.Replace("$", string.Empty, StringComparison.Ordinal);
            var split = 0;
            while (split < value.Length && char.IsLetter(value[split])) split++;
            if (split == 0 || split == value.Length || !ulong.TryParse(value.AsSpan(split), out var row) || row is 0 or > 1_048_576)
                throw new FormatException($"Invalid Excel cell address: {address}");
            return new CellReference(row, ColumnIndex(value[..split]));
        }

        public static string ColumnName(ulong column)
        {
            if (column is 0 or > 16_384) throw new ArgumentOutOfRangeException(nameof(column));
            var chars = new Stack<char>();
            while (column != 0) { column--; chars.Push((char)('A' + column % 26)); column /= 26; }
            return new string(chars.ToArray());
        }

        public static ulong ColumnIndex(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Column name cannot be empty", nameof(name));
            ulong value = 0;
            foreach (var raw in name) {
                var c = char.ToUpperInvariant(raw);
                if (c is < 'A' or > 'Z') throw new FormatException($"Invalid Excel column: {name}");
                value = checked(value * 26 + (ulong)(c - 'A' + 1));
            }
            if (value > 16_384) throw new ArgumentOutOfRangeException(nameof(name));
            return value;
        }

        public string Address => ColumnName(Column) + Row;
        public override string ToString() => Address;
    }

    public readonly record struct ReferenceTranslationResult(string Value, ulong ReferencesVisited,
        ulong ReferencesChanged, ulong ReferencesInvalidated)
    {
        public bool Changed => ReferencesChanged != 0 || ReferencesInvalidated != 0;
    }

    public sealed class DrawingMarker
    {
        public ulong Row { get; set; } = 1;
        public ulong Column { get; set; } = 1;
        public long RowOffsetEmu { get; set; }
        public long ColumnOffsetEmu { get; set; }
    }

    public sealed class DrawingAnchorInfo
    {
        public DrawingAnchorType Type { get; set; } = DrawingAnchorType.OneCell;
        public DrawingMarker From { get; set; } = new();
        public DrawingMarker To { get; set; } = new();
        public long XEmu { get; set; }
        public long YEmu { get; set; }
        public long WidthEmu { get; set; }
        public long HeightEmu { get; set; }
        public string EditAs { get; set; } = string.Empty;
    }

    public sealed class StreamingReaderOptions
    {
        public ulong MaxFileBytes { get; set; }
        public ulong MaxEntryBytes { get; set; }
        public ulong MaxTotalBytes { get; set; }
        public ulong MaxEntries { get; set; }
        public bool ValidateCellReferences { get; set; } = true;
    }

    public sealed record PivotGrouping(bool DateGrouping, PivotDatePart DatePart, double Start, double End,
        double Interval, bool AutoStart = false, bool AutoEnd = false, string StartDate = "", string EndDate = "");

    public sealed record PivotFilter(string Type, int FieldIndex, int MeasureFieldIndex = -1,
        string Value1 = "", string Value2 = "", double Top10Value = 10,
        bool Top10Percent = false, bool Top10Top = true);

    public sealed class LoadOptions
    {
        public bool Lenient { get; set; }
        public ulong MaxEntries { get; set; }
        public ulong MaxEntryBytes { get; set; }
        public ulong MaxTotalBytes { get; set; }
        public ulong MaxFileBytes { get; set; }
        public string? Password { get; set; }
        public bool VerifyEncryptionIntegrity { get; set; } = true;
        public Func<bool>? Cancel { get; set; }
        public Action<ulong, ulong>? Progress { get; set; }
    }

    public sealed class SaveOptions
    {
        public CompressionLevel CompressionLevel { get; set; } = CompressionLevel.Default;
        public CompressionStrategy CompressionStrategy { get; set; } = CompressionStrategy.Default;
        public ulong ParallelWorkers { get; set; }
        public bool ParallelSheets { get; set; } = true;
        public bool ParallelRows { get; set; }
        public bool StrictNamespace { get; set; }
        public bool SynchronizeChartCaches { get; set; }
        public bool SynchronizeChangedChartCachesOnly { get; set; } = true;
        public bool CalculateFormulasBeforeSave { get; set; }
        public bool AtomicWrite { get; set; } = true;
        public bool ValidateBeforeSave { get; set; } = true;
        public string? EncryptionPassword { get; set; }
        public OfficeEncryptionMode EncryptionMode { get; set; } = OfficeEncryptionMode.AgileAes256Sha512;
        public ulong EncryptionSpinCount { get; set; } = 100000;
        public ulong EncryptionKeyBits { get; set; } = 256;
    }

    public readonly record struct ExternalReferenceValue(CellValueType Type, object? Value)
    {
        public static ExternalReferenceValue Empty() => new(CellValueType.Empty, null);
        public static ExternalReferenceValue Number(double value) => new(CellValueType.Number, value);
        public static ExternalReferenceValue Boolean(bool value) => new(CellValueType.Bool, value);
        public static ExternalReferenceValue Text(string value) => new(CellValueType.String, value);
        public static ExternalReferenceValue Error(CellError value) => new(CellValueType.Error, value);
        public static ExternalReferenceValue Date(DateTime value) => new(CellValueType.Date, value);
    }

    public sealed class CalculationOptions
    {
        public bool RecursiveDependencies { get; set; } = true;
        public bool UpdateCachedValues { get; set; } = true;
        public bool EvaluateVolatileFunctions { get; set; } = true;
        public bool SpillDynamicArrays { get; set; } = true;
        public bool IterativeCalculation { get; set; }
        public ulong MaxIterations { get; set; } = 100;
        public double MaxChange { get; set; } = 0.001;
        public ulong MaxDepth { get; set; } = 512;
        public Func<string, string, string, ExternalReferenceValue?>? ExternalReferenceResolver { get; set; }
    }

    public sealed class LoadDiagnostics
    {
        internal LoadDiagnostics(List<string> warnings, List<string> errors) { Warnings = warnings; Errors = errors; }
        public IReadOnlyList<string> Warnings { get; }
        public IReadOnlyList<string> Errors { get; }
        public bool HadErrors => Errors.Count != 0;
    }

    public sealed class StructuralEditOptions
    {
        public bool Transactional { get; set; } = true;
        public bool UpdateDefinedNames { get; set; } = true;
        public bool RecalculateFormulas { get; set; }
        public bool SynchronizeChartCaches { get; set; } = true;
        public bool ChangedChartCachesOnly { get; set; }
        public bool FailOnInvalidReference { get; set; }
    }

    public sealed class WorksheetRenameOptions
    {
        public bool RecalculateFormulas { get; set; }
        public bool SynchronizeChartCaches { get; set; } = true;
        public bool ChangedChartCachesOnly { get; set; }
    }

    public readonly record struct WorksheetRenameReport(
        ulong WorksheetsVisited, ulong FormulasUpdated, ulong FormulaMetadataUpdated,
        ulong DefinedNamesUpdated, ulong ChartReferencesUpdated, ulong PivotReferencesUpdated,
        ulong HyperlinksUpdated, ulong ReferencesUpdated, ulong FormulasCalculated,
        ulong ChartCachesUpdated, bool Success);

    public sealed class ChartCacheSyncOptions
    {
        public bool SynchronizeTitles { get; set; } = true;
        public bool SynchronizeCategories { get; set; } = true;
        public bool SynchronizeValues { get; set; } = true;
        public bool ChangedReferencesOnly { get; set; }
        public bool ClearUnsupportedReferences { get; set; }
    }

    public readonly record struct ChartCacheSyncReport(
        ulong ChartsVisited, ulong SeriesVisited, ulong ReferencesChecked, ulong ReferencesUnchanged,
        ulong DependenciesRegistered, ulong DependenciesChanged, ulong CachesUpdated,
        ulong CachesCleared, ulong ReferencesSkipped, bool Success);

    public readonly record struct FormulaDependency(
        string DependentSheet, string DependentCell, FormulaDependencyKind Kind,
        string PrecedentSheet, string PrecedentReference, string Symbol);

    public readonly record struct FormulaDependencyReport(
        ulong FormulaCells, ulong Edges, ulong CellOrRangeEdges, ulong DefinedNameEdges,
        ulong TableEdges, ulong ExternalEdges, ulong VolatileReferences, ulong UnresolvedSymbols);

    public sealed class WorkbookValidationOptions
    {
        public bool ValidateWorksheetNames { get; set; } = true;
        public bool ValidateDefinedNames { get; set; } = true;
        public bool ValidateTables { get; set; } = true;
        public bool ValidatePivots { get; set; } = true;
    }

    public readonly record struct WorkbookValidationIssue(
        WorkbookValidationSeverity Severity, string Code, string Message, string Worksheet);

    public sealed class WorkbookValidationReport
    {
        internal WorkbookValidationReport(ulong errors, ulong warnings, List<WorkbookValidationIssue> issues)
        { ErrorCount = errors; WarningCount = warnings; Issues = issues; }
        public ulong ErrorCount { get; }
        public ulong WarningCount { get; }
        public IReadOnlyList<WorkbookValidationIssue> Issues { get; }
        public bool Ok => ErrorCount == 0;
    }

    public readonly record struct VbaModule(string Name, string Source, VbaModuleType Type, bool ReadOnly, bool PrivateModule);

    public readonly record struct VbaProjectProperties(string Name, string Description, string HelpFile, uint HelpContextId, string Constants);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    internal struct NativeLoadOptions
    {
        public int Lenient;
        public ulong MaxEntries, MaxEntryBytes, MaxTotalBytes, MaxFileBytes;
        [MarshalAs(UnmanagedType.LPStr)] public string? Password;
        public int VerifyEncryptionIntegrity;
        public IntPtr Cancel;
        public IntPtr Progress;
        public IntPtr CallbackUser;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    internal struct NativeSaveOptions
    {
        public int CompressionLevel, CompressionStrategy;
        public ulong ParallelWorkers;
        public int ParallelSheets, ParallelRows, StrictNamespace;
        public int SynchronizeChartCaches, SynchronizeChangedChartCachesOnly, CalculateFormulasBeforeSave;
        public int AtomicWrite, ValidateBeforeSave;
        [MarshalAs(UnmanagedType.LPStr)] public string? EncryptionPassword;
        public int EncryptionMode;
        public ulong EncryptionSpinCount, EncryptionKeyBits;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeCalculationOptions
    {
        public int RecursiveDependencies, UpdateCachedValues, EvaluateVolatileFunctions, SpillDynamicArrays, IterativeCalculation;
        public ulong MaxIterations;
        public double MaxChange;
        public ulong MaxDepth;
        public IntPtr ExternalReferenceResolver;
        public IntPtr ExternalReferenceUser;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeStructuralOptions
    {
        public int Transactional, UpdateDefinedNames, RecalculateFormulas;
        public int SynchronizeChartCaches, ChangedChartCachesOnly, FailOnInvalidReference;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeCalculationReport
    {
        public ulong FormulaCellsVisited, FormulaCellsEvaluated, CachedValuesUpdated, DependencyEvaluations;
        public ulong DefinedNamesResolved, CircularReferences, UnsupportedFormulas, EvaluationErrors;
        public ulong DynamicArraysSpilled, SpillCellsUpdated, SpillConflicts, StructuredReferencesResolved;
        public ulong IterativeIterations, IterativeConvergenceFailures, ExternalReferencesResolved, UnresolvedExternalReferences;
        public int Success;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeStructuralReport
    {
        public ulong WorksheetsVisited, CellsMoved, CellsRemoved, FormulasUpdated;
        public ulong FormulaMetadataUpdated, WorksheetReferencesUpdated, DefinedNamesUpdated;
        public ulong ChartReferencesUpdated, PivotReferencesUpdated, DrawingAnchorsUpdated;
        public ulong HyperlinksUpdated, ReferencesInvalidated, FormulasCalculated, ChartCachesUpdated;
        public int Success;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeWorksheetRenameReport
    {
        public ulong WorksheetsVisited, FormulasUpdated, FormulaMetadataUpdated, DefinedNamesUpdated;
        public ulong ChartReferencesUpdated, PivotReferencesUpdated, HyperlinksUpdated, ReferencesUpdated;
        public ulong FormulasCalculated, ChartCachesUpdated;
        public int Success;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeChartCacheSyncOptions
    {
        public int SynchronizeTitles, SynchronizeCategories, SynchronizeValues;
        public int ChangedReferencesOnly, ClearUnsupportedReferences;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeChartCacheSyncReport
    {
        public ulong ChartsVisited, SeriesVisited, ReferencesChecked, ReferencesUnchanged;
        public ulong DependenciesRegistered, DependenciesChanged, CachesUpdated, CachesCleared, ReferencesSkipped;
        public int Success;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeDependencyReport
    {
        public ulong FormulaCells, Edges, CellOrRangeEdges, DefinedNameEdges;
        public ulong TableEdges, ExternalEdges, VolatileReferences, UnresolvedSymbols;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeValidationOptions
    {
        public int ValidateWorksheetNames, ValidateDefinedNames, ValidateTables, ValidatePivots;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeReferenceTranslationReport
    {
        public ulong ReferencesVisited, ReferencesChanged, ReferencesInvalidated;
        public int Changed;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeDrawingAnchor
    {
        public int Type;
        public ulong FromRow, FromColumn;
        public long FromRowOffsetEmu, FromColumnOffsetEmu;
        public ulong ToRow, ToColumn;
        public long ToRowOffsetEmu, ToColumnOffsetEmu;
        public long XEmu, YEmu, WidthEmu, HeightEmu;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeStreamingReaderOptions
    {
        public ulong MaxFileBytes, MaxEntryBytes, MaxTotalBytes, MaxEntries;
        public int ValidateCellReferences;
    }

    internal static partial class Native
    {
        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal delegate int ExternalReferenceResolverCallback(IntPtr user, string workbookToken, string sheetName, string address, IntPtr output);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        internal delegate int CancelCallback(IntPtr user);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        internal delegate void ProgressCallback(IntPtr user, ulong done, ulong total);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_external_value_set_empty(IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_external_value_set_number(IntPtr value, double number);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_external_value_set_bool(IntPtr value, int booleanValue);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern void xlpp_external_value_set_string(IntPtr value, string stringValue);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_external_value_set_error(IntPtr value, int errorCode);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_external_value_set_date(IntPtr value, int year, int month, int day, int hour, int minute, double second, int hasTime);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern IntPtr xlpp_last_error();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_clear_error();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_free_string(IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_load_ex(IntPtr wb, string path, ref NativeLoadOptions options);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_save_ex(IntPtr wb, string path, ref NativeSaveOptions options);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_load_bytes(IntPtr wb, byte[] bytes, ulong size, ref NativeLoadOptions options);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_save_bytes(IntPtr wb, ref NativeSaveOptions options, out IntPtr bytes, out ulong size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_free_bytes(IntPtr bytes);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_calculate_options(IntPtr wb, ref NativeCalculationOptions options, out NativeCalculationReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_workbook_diagnostic_warning_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_workbook_diagnostic_error_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_diagnostic_warning(IntPtr wb, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_diagnostic_error(IntPtr wb, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_structural_edit_ex(IntPtr wb, string sheetName, int kind, ulong index, ulong amount, ref NativeStructuralOptions options, out NativeStructuralReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_rename_sheet(IntPtr wb, string oldName, string newName, int recalc, int syncCaches, int changedOnly, out NativeWorksheetRenameReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_synchronize_chart_caches(IntPtr wb, ref NativeChartCacheSyncOptions options, out NativeChartCacheSyncReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_workbook_reset_chart_cache_tracking(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_workbook_tracked_chart_cache_dependencies(IntPtr wb);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern IntPtr xlpp_workbook_dependency_graph(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_dependency_graph_destroy(IntPtr graph);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_dependency_graph_report(IntPtr graph, out NativeDependencyReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_dependency_graph_edge_count(IntPtr graph);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_dependency_graph_edge_kind(IntPtr graph, ulong index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_dependency_graph_edge_dependent_sheet(IntPtr graph, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_dependency_graph_edge_dependent_cell(IntPtr graph, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_dependency_graph_edge_precedent_sheet(IntPtr graph, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_dependency_graph_edge_precedent_reference(IntPtr graph, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_dependency_graph_edge_symbol(IntPtr graph, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_dependency_graph_depends_on(IntPtr graph, string dependentSheet, string dependentCell, string precedentSheet, string precedentCell);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern IntPtr xlpp_workbook_validate(IntPtr wb, ref NativeValidationOptions options);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_validation_report_destroy(IntPtr report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_validation_error_count(IntPtr report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_validation_warning_count(IntPtr report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_validation_issue_count(IntPtr report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_validation_issue_severity(IntPtr report, ulong index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_validation_issue_code(IntPtr report, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_validation_issue_message(IntPtr report, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_validation_issue_worksheet(IntPtr report, ulong index, IntPtr output, int outputSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_add_vba_project(IntPtr wb, string path);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_set_vba_project(IntPtr wb, byte[] bytes, ulong size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_has_vba_project(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_remove_vba_project(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_set_vba_module_text(IntPtr wb, string moduleName, string source);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_set_vba_class_module_text(IntPtr wb, string moduleName, string source, int readOnly, int privateModule);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_set_vba_document_module_text(IntPtr wb, string moduleName, string source);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_vba_module_text(IntPtr wb, string moduleName, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_workbook_vba_module_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_module_type(IntPtr wb, ulong index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_module_read_only(IntPtr wb, ulong index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_module_private(IntPtr wb, ulong index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_module_name(IntPtr wb, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_module_source(IntPtr wb, ulong index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_remove_vba_module(IntPtr wb, string moduleName);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_workbook_vba_project_bytes(IntPtr wb, IntPtr output, ulong outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_save_vba_project(IntPtr wb, string path);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_has_vba_signature(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_source_editable(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_project_name(IntPtr wb, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_project_description(IntPtr wb, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_project_help_file(IntPtr wb, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern uint xlpp_workbook_vba_project_help_context(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_workbook_vba_project_constants(IntPtr wb, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_workbook_set_vba_project_properties(IntPtr wb, string name, string description, string helpFile, uint helpContext, string constants);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern IntPtr xlpp_workbook_inspect_enterprise_features(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_enterprise_inspection_destroy(IntPtr inspection);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_count(IntPtr inspection);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_warning_count(IntPtr inspection);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_warning_at(IntPtr inspection, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_kind(IntPtr inspection, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_part_name(IntPtr inspection, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_content_type(IntPtr inspection, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_name(IntPtr inspection, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_source_name(IntPtr inspection, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_connection_id(IntPtr inspection, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_cache_id(IntPtr inspection, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_binary(IntPtr inspection, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_semantic_editable(IntPtr inspection, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_has_refresh_on_load(IntPtr inspection, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_refresh_on_load(IntPtr inspection, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_pivot_table_count(IntPtr inspection, int featureIndex);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_pivot_table_at(IntPtr inspection, int featureIndex, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_feature_relationship_count(IntPtr inspection, int featureIndex);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_relationship_source_part(IntPtr inspection, int featureIndex, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_relationship_id(IntPtr inspection, int featureIndex, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_relationship_type(IntPtr inspection, int featureIndex, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_relationship_target_mode(IntPtr inspection, int featureIndex, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_relationship_target_part(IntPtr inspection, int featureIndex, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_relationship_outgoing(IntPtr inspection, int featureIndex, int index);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern IntPtr xlpp_workbook_set_connection_refresh_on_load(IntPtr wb, string connectionId, int enabled);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern IntPtr xlpp_workbook_set_query_table_refresh_on_load(IntPtr wb, string queryName, int enabled);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern IntPtr xlpp_workbook_set_olap_pivot_cache_refresh_on_load(IntPtr wb, string partName, int enabled);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern IntPtr xlpp_workbook_set_pivot_chart_source_name(IntPtr wb, string partName, string sourceName);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_enterprise_edit_report_destroy(IntPtr report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_enterprise_edit_report_matched(IntPtr report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_enterprise_edit_report_modified(IntPtr report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_edit_report_warning_count(IntPtr report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_edit_report_warning_at(IntPtr report, int index, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_enterprise_edit_report_success(IntPtr report);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_translate_formula_references(string formula, string contextSheet,
            string editSheet, int kind, ulong index, ulong amount, IntPtr output, int outputSize,
            out NativeReferenceTranslationReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_translate_range_references(string reference, string contextSheet,
            string editSheet, int kind, ulong index, ulong amount, IntPtr output, int outputSize,
            out NativeReferenceTranslationReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_rename_worksheet_references(string expression, string oldName, string newName,
            IntPtr output, int outputSize, out NativeReferenceTranslationReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern int xlpp_invalidate_worksheet_references(string expression, string removedName,
            IntPtr output, int outputSize, out NativeReferenceTranslationReport report);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_cell_formula_type(IntPtr cell);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_cell_set_formula_type(IntPtr cell, int type);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_cell_formula_reference(IntPtr cell, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern void xlpp_cell_set_formula_reference(IntPtr cell, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_cell_clear_formula_reference(IntPtr cell);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern ulong xlpp_cell_formula_shared_index(IntPtr cell, out int hasValue);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_cell_set_formula_shared_index(IntPtr cell, ulong value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_cell_clear_formula_shared_index(IntPtr cell);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_cell_formula_always_calculate_array(IntPtr cell);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_cell_set_formula_always_calculate_array(IntPtr cell, int value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_cell_formula_calculate_on_load(IntPtr cell);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_cell_set_formula_calculate_on_load(IntPtr cell, int value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_cell_formula_metadata_empty(IntPtr cell);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_image_anchor_info(IntPtr image, out NativeDrawingAnchor value, IntPtr editAs, int editAsSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern void xlpp_image_set_anchor_info(IntPtr image, ref NativeDrawingAnchor value, string editAs);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_image_stable_id(IntPtr image, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_image_source_drawing_part(IntPtr image, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_image_source_media_part(IntPtr image, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_image_source_relationship_id(IntPtr image, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_image_imported(IntPtr image);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_chart_anchor_info(IntPtr chart, out NativeDrawingAnchor value, IntPtr editAs, int editAsSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern void xlpp_chart_set_anchor_info(IntPtr chart, ref NativeDrawingAnchor value, string editAs);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_chart_stable_id(IntPtr chart, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_chart_source_drawing_part(IntPtr chart, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_chart_source_chart_part(IntPtr chart, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern void xlpp_chart_source_relationship_id(IntPtr chart, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        internal static extern int xlpp_chart_imported(IntPtr chart);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        internal static extern IntPtr xlpp_stream_reader_open_ex(string path, ref NativeStreamingReaderOptions options);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern IntPtr xlpp_sheet_image_by_stable_id(IntPtr ws, string stableId);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_move_image(IntPtr ws, string stableId, string anchor);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_move_image_absolute(IntPtr ws, string stableId, long xEmu, long yEmu);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_resize_image(IntPtr ws, string stableId, double width, double height);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_replace_image(IntPtr ws, string stableId, string path);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_image(IntPtr ws, string stableId);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern IntPtr xlpp_sheet_chart_by_stable_id(IntPtr ws, string stableId);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_move_chart(IntPtr ws, string stableId, string anchor);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_move_chart_absolute(IntPtr ws, string stableId, long xEmu, long yEmu);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_resize_chart(IntPtr ws, string stableId, double width, double height);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_title(IntPtr ws, string stableId, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_style(IntPtr ws, string stableId, string style);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_x_axis_title(IntPtr ws, string stableId, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_y_axis_title(IntPtr ws, string stableId, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_title(IntPtr ws, string stableId, ulong axisId, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_number_format(IntPtr ws, string stableId, ulong axisId, string code, int linked);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_ticks(IntPtr ws, string stableId, ulong axisId, string major, string minor, string position);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_units(IntPtr ws, string stableId, ulong axisId, double major, double minor);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_crossing(IntPtr ws, string stableId, ulong axisId, string crosses, string between);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_crosses_at(IntPtr ws, string stableId, ulong axisId, double value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_clear_chart_axis_crosses_at(IntPtr ws, string stableId, ulong axisId);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_clear_chart_axis_display_units(IntPtr ws, string stableId, ulong axisId);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_axis_gridlines(IntPtr ws, string stableId, ulong axisId, int major);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_data_table(IntPtr ws, string stableId);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_plot_drop_lines(IntPtr ws, string stableId, ulong plot);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_plot_high_low_lines(IntPtr ws, string stableId, ulong plot);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_plot_up_down_bars(IntPtr ws, string stableId, ulong plot);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_first_slice_angle(IntPtr ws, string stableId, ulong plot, int degrees);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_doughnut_hole_size(IntPtr ws, string stableId, ulong plot, int percent);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_radar_style(IntPtr ws, string stableId, ulong plot, string style);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_plot_leader_lines(IntPtr ws, string stableId, ulong plot);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_legend(IntPtr ws, string stableId, int show, string position);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_legend_overlay(IntPtr ws, string stableId, int overlay);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_title(IntPtr ws, string stableId, ulong series, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_references(IntPtr ws, string stableId, ulong series, string categories, string values);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_clear_chart_series_caches(IntPtr ws, string stableId, ulong series);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_series_trendline(IntPtr ws, string stableId, ulong series, ulong trendline);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_series_error_bars(IntPtr ws, string stableId, ulong series, int direction);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart(IntPtr ws, string stableId);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern IntPtr xlpp_chart_line_value_create();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern void xlpp_chart_line_value_destroy(IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern void xlpp_chart_line_value_set(IntPtr value, int present, int noFill, double width, string dash, string cap, string compound, string join);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern void xlpp_chart_line_value_set_color(IntPtr value, int kind, string color);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern void xlpp_chart_line_value_add_color_transform(IntPtr value, int kind, int amount);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern void xlpp_chart_line_value_add_custom_dash(IntPtr value, double dash, double space);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern IntPtr xlpp_chart_fill_value_create();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern void xlpp_chart_fill_value_destroy(IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern void xlpp_chart_fill_value_set(IntPtr value, int present, int noFill, int kind, double angle, string pattern);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern void xlpp_chart_fill_value_set_color(IntPtr value, int role, int kind, string color);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern void xlpp_chart_fill_value_add_gradient_stop(IntPtr value, int position, int kind, string color);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern IntPtr xlpp_chart_rich_text_value_create();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern void xlpp_chart_rich_text_value_destroy(IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern void xlpp_chart_rich_text_value_set_present(IntPtr value, int present);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern void xlpp_chart_rich_text_value_add_run(IntPtr value, string text, int bold, int italic, double size, string typeface, int colorKind, string color);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern IntPtr xlpp_chart_cache_value_create();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] internal static extern void xlpp_chart_cache_value_destroy(IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern void xlpp_chart_cache_value_set(IntPtr value, int present, int numeric, string format, ulong count);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern void xlpp_chart_cache_value_add_point(IntPtr value, ulong index, string pointValue);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_title_rich_text(IntPtr ws, string id, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_title_rich_text(IntPtr ws, string id, ulong axis, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_scaling(IntPtr ws, string id, ulong axis, int hasMin, double min, int hasMax, double max, int hasLog, double logBase, int reverse);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_display_units(IntPtr ws, string id, ulong axis, int present, string builtIn, int hasCustom, double custom, int showLabel, IntPtr label);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_line_format(IntPtr ws, string id, ulong axis, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_axis_gridline_format(IntPtr ws, string id, ulong axis, int major, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_area_line_format(IntPtr ws, string id, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_area_fill_format(IntPtr ws, string id, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_area_line_format(IntPtr ws, string id, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_area_fill_format(IntPtr ws, string id, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_area_layout(IntPtr ws, string id, int present, string target, string xm, string ym, string wm, string hm, int hasX, double x, int hasY, double y, int hasW, double w, int hasH, double h);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_view_3d(IntPtr ws, string id, int present, int hrx, int rx, int hry, int ry, int hhp, int hp, int hdp, int dp, int hra, int ra, int hpersp, int persp);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_drop_lines(IntPtr ws, string id, ulong plot, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_high_low_lines(IntPtr ws, string id, ulong plot, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_plot_leader_line_format(IntPtr ws, string id, ulong plot, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_legend_line_format(IntPtr ws, string id, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_legend_fill_format(IntPtr ws, string id, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_legend_layout(IntPtr ws, string id, int present, string target, string xm, string ym, string wm, string hm, int hasX, double x, int hasY, double y, int hasW, double w, int hasH, double h);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_category_cache(IntPtr ws, string id, ulong series, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_value_cache(IntPtr ws, string id, ulong series, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_title_cache(IntPtr ws, string id, ulong series, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_line_format(IntPtr ws, string id, ulong series, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_fill_format(IntPtr ws, string id, ulong series, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_leader_line_format(IntPtr ws, string id, ulong series, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_remove_chart_series_leader_lines(IntPtr ws, string id, ulong series);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_trendline_line_format(IntPtr ws, string id, ulong series, ulong index, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_error_bars_line_format(IntPtr ws, string id, ulong series, int direction, IntPtr value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_trendline(IntPtr ws, string id, ulong series, ulong index, int type, int order, int period, double forward, double backward, int equation, int r2, IntPtr line);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_add_chart_series_trendline(IntPtr ws, string id, ulong series, int type, int order, int period, double forward, double backward, int equation, int r2, IntPtr line);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] internal static extern int xlpp_sheet_set_chart_series_error_bars(IntPtr ws, string id, ulong series, int direction, int barType, int valueType, double value, int noEndCap, string plus, string minus, IntPtr line);
    }

    internal static class AdvancedMarshal
    {
        internal delegate int SizedFill(IntPtr output, int outputSize);
        internal static string FromSizedBuffer(SizedFill fill)
        {
            var required = fill(IntPtr.Zero, 0);
            if (required <= 1) return string.Empty;
            var buffer = Marshal.AllocHGlobal(required);
            try
            {
                fill(buffer, required);
                return Marshal.PtrToStringAnsi(buffer) ?? string.Empty;
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }
    }

    internal static class ChartValueMarshal
    {
        internal static T WithLine<T>(ChartLineFormat value, Func<IntPtr, T> action)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var handle = Native.xlpp_chart_line_value_create();
            if (handle == IntPtr.Zero) throw new OutOfMemoryException("Unable to allocate native chart line value");
            try
            {
                Native.xlpp_chart_line_value_set(handle, value.Present ? 1 : 0, value.NoFill ? 1 : 0,
                    value.WidthPoints, value.Dash ?? string.Empty, value.Cap ?? string.Empty,
                    value.Compound ?? string.Empty, value.Join ?? string.Empty);
                var color = value.Color ?? new ChartColor();
                Native.xlpp_chart_line_value_set_color(handle, (int)color.Kind, color.Value ?? string.Empty);
                foreach (var transform in color.Transforms ?? new())
                    Native.xlpp_chart_line_value_add_color_transform(handle, (int)transform.Kind, transform.Value);
                foreach (var stop in value.CustomDash ?? new())
                    Native.xlpp_chart_line_value_add_custom_dash(handle, stop.Dash, stop.Space);
                return action(handle);
            }
            finally { Native.xlpp_chart_line_value_destroy(handle); }
        }

        internal static T WithFill<T>(ChartFillFormat value, Func<IntPtr, T> action)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var handle = Native.xlpp_chart_fill_value_create();
            if (handle == IntPtr.Zero) throw new OutOfMemoryException("Unable to allocate native chart fill value");
            try
            {
                Native.xlpp_chart_fill_value_set(handle, value.Present ? 1 : 0, value.NoFill ? 1 : 0,
                    (int)value.Kind, value.GradientAngleDegrees, value.Pattern ?? string.Empty);
                SetColor(handle, 0, value.Color);
                SetColor(handle, 1, value.ForegroundColor);
                SetColor(handle, 2, value.BackgroundColor);
                foreach (var stop in value.GradientStops ?? new())
                    Native.xlpp_chart_fill_value_add_gradient_stop(handle, stop.Position, (int)(stop.Color?.Kind ?? ChartColorKind.None), stop.Color?.Value ?? string.Empty);
                return action(handle);
            }
            finally { Native.xlpp_chart_fill_value_destroy(handle); }
        }

        private static void SetColor(IntPtr handle, int role, ChartColor? color) =>
            Native.xlpp_chart_fill_value_set_color(handle, role, (int)(color?.Kind ?? ChartColorKind.None), color?.Value ?? string.Empty);

        internal static T WithRichText<T>(ChartRichText value, Func<IntPtr, T> action)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var handle = Native.xlpp_chart_rich_text_value_create();
            if (handle == IntPtr.Zero) throw new OutOfMemoryException("Unable to allocate native chart rich-text value");
            try
            {
                Native.xlpp_chart_rich_text_value_set_present(handle, value.Present ? 1 : 0);
                foreach (var run in value.Runs ?? new())
                    Native.xlpp_chart_rich_text_value_add_run(handle, run.Text ?? string.Empty, run.Bold ? 1 : 0,
                        run.Italic ? 1 : 0, run.FontSizePoints, run.Typeface ?? string.Empty,
                        (int)(run.Color?.Kind ?? ChartColorKind.None), run.Color?.Value ?? string.Empty);
                return action(handle);
            }
            finally { Native.xlpp_chart_rich_text_value_destroy(handle); }
        }

        internal static T WithCache<T>(ChartSeriesCache value, Func<IntPtr, T> action)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var handle = Native.xlpp_chart_cache_value_create();
            if (handle == IntPtr.Zero) throw new OutOfMemoryException("Unable to allocate native chart cache value");
            try
            {
                Native.xlpp_chart_cache_value_set(handle, value.Present ? 1 : 0, value.Numeric ? 1 : 0,
                    value.FormatCode ?? string.Empty, value.PointCount);
                foreach (var point in value.Points ?? new())
                    Native.xlpp_chart_cache_value_add_point(handle, point.Index, point.Value ?? string.Empty);
                return action(handle);
            }
            finally { Native.xlpp_chart_cache_value_destroy(handle); }
        }
    }

    public static class ReferenceTranslator
    {
        private delegate int TranslateCall(IntPtr output, int outputSize, out NativeReferenceTranslationReport report);

        private static ReferenceTranslationResult Invoke(TranslateCall call)
        {
            var required = call(IntPtr.Zero, 0, out var report);
            if (required <= 0) throw new InvalidOperationException("Reference translation failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            var buffer = Marshal.AllocHGlobal(required);
            try
            {
                call(buffer, required, out report);
                return new ReferenceTranslationResult(Marshal.PtrToStringAnsi(buffer) ?? string.Empty,
                    report.ReferencesVisited, report.ReferencesChanged, report.ReferencesInvalidated);
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }

        public static ReferenceTranslationResult TranslateFormulaReferences(string formula, string contextSheet,
            string editSheet, StructuralEditKind kind, ulong index, ulong amount = 1) =>
            Invoke((IntPtr b, int n, out NativeReferenceTranslationReport r) =>
                Native.xlpp_translate_formula_references(formula, contextSheet, editSheet, (int)kind, index, amount, b, n, out r));

        public static ReferenceTranslationResult TranslateRangeReferences(string reference, string contextSheet,
            string editSheet, StructuralEditKind kind, ulong index, ulong amount = 1) =>
            Invoke((IntPtr b, int n, out NativeReferenceTranslationReport r) =>
                Native.xlpp_translate_range_references(reference, contextSheet, editSheet, (int)kind, index, amount, b, n, out r));

        public static ReferenceTranslationResult RenameWorksheetReferences(string expression, string oldName, string newName) =>
            Invoke((IntPtr b, int n, out NativeReferenceTranslationReport r) =>
                Native.xlpp_rename_worksheet_references(expression, oldName, newName, b, n, out r));

        public static ReferenceTranslationResult InvalidateWorksheetReferences(string expression, string removedName) =>
            Invoke((IntPtr b, int n, out NativeReferenceTranslationReport r) =>
                Native.xlpp_invalidate_worksheet_references(expression, removedName, b, n, out r));
    }

    public sealed class FormulaMetadata
    {
        private readonly IntPtr _cell;
        internal FormulaMetadata(IntPtr cell) => _cell = cell;
        public FormulaType Type { get => (FormulaType)Native.xlpp_cell_formula_type(_cell); set => Native.xlpp_cell_set_formula_type(_cell, (int)value); }
        public string Reference
        {
            get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_cell_formula_reference(_cell, b, n));
            set => Native.xlpp_cell_set_formula_reference(_cell, value ?? string.Empty);
        }
        public void ClearReference() => Native.xlpp_cell_clear_formula_reference(_cell);
        public ulong? SharedIndex
        {
            get { var value = Native.xlpp_cell_formula_shared_index(_cell, out var present); return present != 0 ? value : null; }
            set { if (value.HasValue) Native.xlpp_cell_set_formula_shared_index(_cell, value.Value); else Native.xlpp_cell_clear_formula_shared_index(_cell); }
        }
        public void ClearSharedIndex() => Native.xlpp_cell_clear_formula_shared_index(_cell);
        public bool AlwaysCalculateArray { get => Native.xlpp_cell_formula_always_calculate_array(_cell) != 0; set => Native.xlpp_cell_set_formula_always_calculate_array(_cell, value ? 1 : 0); }
        public bool CalculateOnLoad { get => Native.xlpp_cell_formula_calculate_on_load(_cell) != 0; set => Native.xlpp_cell_set_formula_calculate_on_load(_cell, value ? 1 : 0); }
        public bool Empty => Native.xlpp_cell_formula_metadata_empty(_cell) != 0;
    }

    internal static class DrawingAnchorInterop
    {
        internal static NativeDrawingAnchor ToNative(DrawingAnchorInfo value) => new()
        {
            Type = (int)value.Type,
            FromRow = value.From.Row, FromColumn = value.From.Column,
            FromRowOffsetEmu = value.From.RowOffsetEmu, FromColumnOffsetEmu = value.From.ColumnOffsetEmu,
            ToRow = value.To.Row, ToColumn = value.To.Column,
            ToRowOffsetEmu = value.To.RowOffsetEmu, ToColumnOffsetEmu = value.To.ColumnOffsetEmu,
            XEmu = value.XEmu, YEmu = value.YEmu, WidthEmu = value.WidthEmu, HeightEmu = value.HeightEmu
        };

        internal static DrawingAnchorInfo FromNative(NativeDrawingAnchor value, string editAs) => new()
        {
            Type = (DrawingAnchorType)value.Type,
            From = new DrawingMarker { Row = value.FromRow, Column = value.FromColumn, RowOffsetEmu = value.FromRowOffsetEmu, ColumnOffsetEmu = value.FromColumnOffsetEmu },
            To = new DrawingMarker { Row = value.ToRow, Column = value.ToColumn, RowOffsetEmu = value.ToRowOffsetEmu, ColumnOffsetEmu = value.ToColumnOffsetEmu },
            XEmu = value.XEmu, YEmu = value.YEmu, WidthEmu = value.WidthEmu, HeightEmu = value.HeightEmu, EditAs = editAs
        };
    }

    public partial class Worksheet
    {
        public Image? ImageByStableId(string stableId) { var h = Native.xlpp_sheet_image_by_stable_id(_handle, stableId); return h == IntPtr.Zero ? null : new Image(h); }
        public bool MoveImage(string stableId, string anchor) => Native.xlpp_sheet_move_image(_handle, stableId, anchor) != 0;
        public bool MoveImageAbsolute(string stableId, long xEmu, long yEmu) => Native.xlpp_sheet_move_image_absolute(_handle, stableId, xEmu, yEmu) != 0;
        public bool ResizeImage(string stableId, double widthPixels, double heightPixels) => Native.xlpp_sheet_resize_image(_handle, stableId, widthPixels, heightPixels) != 0;
        public bool ReplaceImage(string stableId, string path) => Native.xlpp_sheet_replace_image(_handle, stableId, path) != 0;
        public bool RemoveImage(string stableId) => Native.xlpp_sheet_remove_image(_handle, stableId) != 0;

        public Chart? ChartByStableId(string stableId) { var h = Native.xlpp_sheet_chart_by_stable_id(_handle, stableId); return h == IntPtr.Zero ? null : new Chart(h); }
        public bool MoveChart(string stableId, string anchor) => Native.xlpp_sheet_move_chart(_handle, stableId, anchor) != 0;
        public bool MoveChartAbsolute(string stableId, long xEmu, long yEmu) => Native.xlpp_sheet_move_chart_absolute(_handle, stableId, xEmu, yEmu) != 0;
        public bool ResizeChart(string stableId, double widthPixels, double heightPixels) => Native.xlpp_sheet_resize_chart(_handle, stableId, widthPixels, heightPixels) != 0;
        public bool SetChartTitle(string stableId, string title) => Native.xlpp_sheet_set_chart_title(_handle, stableId, title) != 0;
        public bool SetChartStyle(string stableId, string style) => Native.xlpp_sheet_set_chart_style(_handle, stableId, style) != 0;
        public bool SetChartXAxisTitle(string stableId, string title) => Native.xlpp_sheet_set_chart_x_axis_title(_handle, stableId, title) != 0;
        public bool SetChartYAxisTitle(string stableId, string title) => Native.xlpp_sheet_set_chart_y_axis_title(_handle, stableId, title) != 0;
        public bool SetChartAxisTitle(string stableId, ulong axisId, string title) => Native.xlpp_sheet_set_chart_axis_title(_handle, stableId, axisId, title) != 0;
        public bool SetChartAxisNumberFormat(string stableId, ulong axisId, string code, bool sourceLinked = false) => Native.xlpp_sheet_set_chart_axis_number_format(_handle, stableId, axisId, code, sourceLinked ? 1 : 0) != 0;
        public bool SetChartAxisTicks(string stableId, ulong axisId, string major, string minor, string labelPosition) => Native.xlpp_sheet_set_chart_axis_ticks(_handle, stableId, axisId, major, minor, labelPosition) != 0;
        public bool SetChartAxisUnits(string stableId, ulong axisId, double majorUnit, double minorUnit = 0) => Native.xlpp_sheet_set_chart_axis_units(_handle, stableId, axisId, majorUnit, minorUnit) != 0;
        public bool SetChartAxisCrossing(string stableId, ulong axisId, string crosses, string crossBetween = "") => Native.xlpp_sheet_set_chart_axis_crossing(_handle, stableId, axisId, crosses, crossBetween) != 0;
        public bool SetChartAxisCrossesAt(string stableId, ulong axisId, double value) => Native.xlpp_sheet_set_chart_axis_crosses_at(_handle, stableId, axisId, value) != 0;
        public bool SetChartTitleRichText(string stableId, ChartRichText value) => ChartValueMarshal.WithRichText(value, h => Native.xlpp_sheet_set_chart_title_rich_text(_handle, stableId, h) != 0);
        public bool SetChartAxisTitleRichText(string stableId, ulong axisId, ChartRichText value) => ChartValueMarshal.WithRichText(value, h => Native.xlpp_sheet_set_chart_axis_title_rich_text(_handle, stableId, axisId, h) != 0);
        public bool SetChartAxisScaling(string stableId, ulong axisId, ChartAxisScaling value) => Native.xlpp_sheet_set_chart_axis_scaling(_handle, stableId, axisId,
            value.HasMinimum ? 1 : 0, value.Minimum, value.HasMaximum ? 1 : 0, value.Maximum,
            value.HasLogBase ? 1 : 0, value.LogBase, value.ReverseOrder ? 1 : 0) != 0;
        public bool SetChartAxisDisplayUnits(string stableId, ulong axisId, ChartDisplayUnits value) => ChartValueMarshal.WithRichText(value.LabelRichText,
            h => Native.xlpp_sheet_set_chart_axis_display_units(_handle, stableId, axisId, value.Present ? 1 : 0,
                value.BuiltInUnit ?? string.Empty, value.HasCustomUnit ? 1 : 0, value.CustomUnit, value.ShowLabel ? 1 : 0, h) != 0);
        public bool SetChartAxisLineFormat(string stableId, ulong axisId, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_axis_line_format(_handle, stableId, axisId, h) != 0);
        public bool SetChartAxisGridlineFormat(string stableId, ulong axisId, bool major, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_axis_gridline_format(_handle, stableId, axisId, major ? 1 : 0, h) != 0);
        public bool ClearChartAxisCrossesAt(string stableId, ulong axisId) => Native.xlpp_sheet_clear_chart_axis_crosses_at(_handle, stableId, axisId) != 0;
        public bool ClearChartAxisDisplayUnits(string stableId, ulong axisId) => Native.xlpp_sheet_clear_chart_axis_display_units(_handle, stableId, axisId) != 0;
        public bool RemoveChartAxisGridlines(string stableId, ulong axisId, bool major = true) => Native.xlpp_sheet_remove_chart_axis_gridlines(_handle, stableId, axisId, major ? 1 : 0) != 0;
        public bool RemoveChartDataTable(string stableId) => Native.xlpp_sheet_remove_chart_data_table(_handle, stableId) != 0;
        public bool SetChartAreaLineFormat(string stableId, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_area_line_format(_handle, stableId, h) != 0);
        public bool SetChartAreaFillFormat(string stableId, ChartFillFormat value) => ChartValueMarshal.WithFill(value, h => Native.xlpp_sheet_set_chart_area_fill_format(_handle, stableId, h) != 0);
        public bool SetChartPlotAreaLineFormat(string stableId, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_plot_area_line_format(_handle, stableId, h) != 0);
        public bool SetChartPlotAreaFillFormat(string stableId, ChartFillFormat value) => ChartValueMarshal.WithFill(value, h => Native.xlpp_sheet_set_chart_plot_area_fill_format(_handle, stableId, h) != 0);
        public bool SetChartPlotAreaLayout(string stableId, ChartManualLayout value) => Native.xlpp_sheet_set_chart_plot_area_layout(_handle, stableId,
            value.Present ? 1 : 0, value.Target ?? string.Empty, value.XMode ?? string.Empty, value.YMode ?? string.Empty,
            value.WidthMode ?? string.Empty, value.HeightMode ?? string.Empty, value.HasX ? 1 : 0, value.X,
            value.HasY ? 1 : 0, value.Y, value.HasWidth ? 1 : 0, value.Width, value.HasHeight ? 1 : 0, value.Height) != 0;
        public bool SetChartView3D(string stableId, ChartView3D value) => Native.xlpp_sheet_set_chart_view_3d(_handle, stableId,
            value.Present ? 1 : 0, value.HasRotationX ? 1 : 0, value.RotationX, value.HasRotationY ? 1 : 0, value.RotationY,
            value.HasHeightPercent ? 1 : 0, value.HeightPercent, value.HasDepthPercent ? 1 : 0, value.DepthPercent,
            value.HasRightAngleAxes ? 1 : 0, value.RightAngleAxes ? 1 : 0, value.HasPerspective ? 1 : 0, value.Perspective) != 0;
        public bool SetChartPlotDropLines(string stableId, ulong plotIndex, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_plot_drop_lines(_handle, stableId, plotIndex, h) != 0);
        public bool SetChartPlotHighLowLines(string stableId, ulong plotIndex, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_plot_high_low_lines(_handle, stableId, plotIndex, h) != 0);
        public bool SetChartPlotLeaderLineFormat(string stableId, ulong plotIndex, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_plot_leader_line_format(_handle, stableId, plotIndex, h) != 0);
        public bool RemoveChartPlotDropLines(string stableId, ulong plotIndex) => Native.xlpp_sheet_remove_chart_plot_drop_lines(_handle, stableId, plotIndex) != 0;
        public bool RemoveChartPlotHighLowLines(string stableId, ulong plotIndex) => Native.xlpp_sheet_remove_chart_plot_high_low_lines(_handle, stableId, plotIndex) != 0;
        public bool RemoveChartPlotUpDownBars(string stableId, ulong plotIndex) => Native.xlpp_sheet_remove_chart_plot_up_down_bars(_handle, stableId, plotIndex) != 0;
        public bool SetChartPlotFirstSliceAngle(string stableId, ulong plotIndex, int degrees) => Native.xlpp_sheet_set_chart_plot_first_slice_angle(_handle, stableId, plotIndex, degrees) != 0;
        public bool SetChartPlotDoughnutHoleSize(string stableId, ulong plotIndex, int percent) => Native.xlpp_sheet_set_chart_plot_doughnut_hole_size(_handle, stableId, plotIndex, percent) != 0;
        public bool SetChartPlotRadarStyle(string stableId, ulong plotIndex, string style) => Native.xlpp_sheet_set_chart_plot_radar_style(_handle, stableId, plotIndex, style) != 0;
        public bool RemoveChartPlotLeaderLines(string stableId, ulong plotIndex) => Native.xlpp_sheet_remove_chart_plot_leader_lines(_handle, stableId, plotIndex) != 0;
        public bool SetChartLegend(string stableId, bool show, string position = "r") => Native.xlpp_sheet_set_chart_legend(_handle, stableId, show ? 1 : 0, position) != 0;
        public bool SetChartLegendOverlay(string stableId, bool overlay) => Native.xlpp_sheet_set_chart_legend_overlay(_handle, stableId, overlay ? 1 : 0) != 0;
        public bool SetChartLegendLineFormat(string stableId, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_legend_line_format(_handle, stableId, h) != 0);
        public bool SetChartLegendFillFormat(string stableId, ChartFillFormat value) => ChartValueMarshal.WithFill(value, h => Native.xlpp_sheet_set_chart_legend_fill_format(_handle, stableId, h) != 0);
        public bool SetChartLegendLayout(string stableId, ChartManualLayout value) => Native.xlpp_sheet_set_chart_legend_layout(_handle, stableId,
            value.Present ? 1 : 0, value.Target ?? string.Empty, value.XMode ?? string.Empty, value.YMode ?? string.Empty,
            value.WidthMode ?? string.Empty, value.HeightMode ?? string.Empty, value.HasX ? 1 : 0, value.X,
            value.HasY ? 1 : 0, value.Y, value.HasWidth ? 1 : 0, value.Width, value.HasHeight ? 1 : 0, value.Height) != 0;
        public bool SetChartSeriesTitle(string stableId, ulong seriesIndex, string title) => Native.xlpp_sheet_set_chart_series_title(_handle, stableId, seriesIndex, title) != 0;
        public bool SetChartSeriesReferences(string stableId, ulong seriesIndex, string categories, string values) => Native.xlpp_sheet_set_chart_series_references(_handle, stableId, seriesIndex, categories, values) != 0;
        public bool ClearChartSeriesCaches(string stableId, ulong seriesIndex) => Native.xlpp_sheet_clear_chart_series_caches(_handle, stableId, seriesIndex) != 0;
        public bool SetChartSeriesCategoryCache(string stableId, ulong seriesIndex, ChartSeriesCache value) => ChartValueMarshal.WithCache(value, h => Native.xlpp_sheet_set_chart_series_category_cache(_handle, stableId, seriesIndex, h) != 0);
        public bool SetChartSeriesValueCache(string stableId, ulong seriesIndex, ChartSeriesCache value) => ChartValueMarshal.WithCache(value, h => Native.xlpp_sheet_set_chart_series_value_cache(_handle, stableId, seriesIndex, h) != 0);
        public bool SetChartSeriesTitleCache(string stableId, ulong seriesIndex, ChartSeriesCache value) => ChartValueMarshal.WithCache(value, h => Native.xlpp_sheet_set_chart_series_title_cache(_handle, stableId, seriesIndex, h) != 0);
        public bool SetChartSeriesLineFormat(string stableId, ulong seriesIndex, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_series_line_format(_handle, stableId, seriesIndex, h) != 0);
        public bool SetChartSeriesFillFormat(string stableId, ulong seriesIndex, ChartFillFormat value) => ChartValueMarshal.WithFill(value, h => Native.xlpp_sheet_set_chart_series_fill_format(_handle, stableId, seriesIndex, h) != 0);
        public bool SetChartSeriesLeaderLineFormat(string stableId, ulong seriesIndex, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_series_leader_line_format(_handle, stableId, seriesIndex, h) != 0);
        public bool RemoveChartSeriesLeaderLines(string stableId, ulong seriesIndex) => Native.xlpp_sheet_remove_chart_series_leader_lines(_handle, stableId, seriesIndex) != 0;
        public bool SetChartSeriesTrendlineLineFormat(string stableId, ulong seriesIndex, ulong trendlineIndex, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_series_trendline_line_format(_handle, stableId, seriesIndex, trendlineIndex, h) != 0);
        public bool SetChartSeriesErrorBarsLineFormat(string stableId, ulong seriesIndex, ChartErrorBarDirection direction, ChartLineFormat value) => ChartValueMarshal.WithLine(value, h => Native.xlpp_sheet_set_chart_series_error_bars_line_format(_handle, stableId, seriesIndex, (int)direction, h) != 0);
        public bool SetChartSeriesTrendline(string stableId, ulong seriesIndex, ulong trendlineIndex, ChartTrendline value) => ChartValueMarshal.WithLine(value.LineFormat,
            h => Native.xlpp_sheet_set_chart_series_trendline(_handle, stableId, seriesIndex, trendlineIndex, (int)value.Type,
                value.Order, value.Period, value.Forward, value.Backward, value.DisplayEquation ? 1 : 0, value.DisplayRSquared ? 1 : 0, h) != 0);
        public bool AddChartSeriesTrendline(string stableId, ulong seriesIndex, ChartTrendline value) => ChartValueMarshal.WithLine(value.LineFormat,
            h => Native.xlpp_sheet_add_chart_series_trendline(_handle, stableId, seriesIndex, (int)value.Type,
                value.Order, value.Period, value.Forward, value.Backward, value.DisplayEquation ? 1 : 0, value.DisplayRSquared ? 1 : 0, h) != 0);
        public bool SetChartSeriesErrorBars(string stableId, ulong seriesIndex, ChartErrorBars value) => ChartValueMarshal.WithLine(value.LineFormat,
            h => Native.xlpp_sheet_set_chart_series_error_bars(_handle, stableId, seriesIndex, (int)value.Direction,
                (int)value.BarType, (int)value.ValueType, value.Value, value.NoEndCap ? 1 : 0,
                value.PlusReference ?? string.Empty, value.MinusReference ?? string.Empty, h) != 0);
        public bool RemoveChartSeriesTrendline(string stableId, ulong seriesIndex, ulong trendlineIndex) => Native.xlpp_sheet_remove_chart_series_trendline(_handle, stableId, seriesIndex, trendlineIndex) != 0;
        public bool RemoveChartSeriesErrorBars(string stableId, ulong seriesIndex, ChartErrorBarDirection direction = ChartErrorBarDirection.Y) => Native.xlpp_sheet_remove_chart_series_error_bars(_handle, stableId, seriesIndex, (int)direction) != 0;
        public bool RemoveChart(string stableId) => Native.xlpp_sheet_remove_chart(_handle, stableId) != 0;
    }

    public sealed class FormulaDependencyGraph : IDisposable
    {
        private IntPtr _handle;
        internal FormulaDependencyGraph(IntPtr handle) => _handle = handle;
        public FormulaDependencyReport Report
        {
            get
            {
                if (Native.xlpp_dependency_graph_report(_handle, out var r) == 0)
                    throw new InvalidOperationException("Dependency graph report is unavailable");
                return new FormulaDependencyReport(r.FormulaCells, r.Edges, r.CellOrRangeEdges, r.DefinedNameEdges,
                    r.TableEdges, r.ExternalEdges, r.VolatileReferences, r.UnresolvedSymbols);
            }
        }
        public IReadOnlyList<FormulaDependency> Edges
        {
            get
            {
                var result = new List<FormulaDependency>();
                var count = Native.xlpp_dependency_graph_edge_count(_handle);
                for (ulong i = 0; i < count; ++i)
                {
                    string ds = AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_dependency_graph_edge_dependent_sheet(_handle,i,b,n));
                    string dc = AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_dependency_graph_edge_dependent_cell(_handle,i,b,n));
                    string ps = AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_dependency_graph_edge_precedent_sheet(_handle,i,b,n));
                    string pr = AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_dependency_graph_edge_precedent_reference(_handle,i,b,n));
                    string sy = AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_dependency_graph_edge_symbol(_handle,i,b,n));
                    result.Add(new FormulaDependency(ds, dc, (FormulaDependencyKind)Native.xlpp_dependency_graph_edge_kind(_handle,i), ps, pr, sy));
                }
                return result;
            }
        }
        public bool DependsOn(string dependentSheet, string dependentCell, string precedentSheet, string precedentCell) =>
            Native.xlpp_dependency_graph_depends_on(_handle, dependentSheet, dependentCell, precedentSheet, precedentCell) != 0;
        public IReadOnlyList<FormulaDependency> PrecedentsOf(string sheet, string cell)
        {
            var result = new List<FormulaDependency>();
            foreach (var edge in Edges) if (edge.DependentSheet == sheet && edge.DependentCell == cell) result.Add(edge);
            return result;
        }
        public IReadOnlyList<FormulaDependency> DependentsOf(string sheet, string cell)
        {
            var result = new List<FormulaDependency>();
            foreach (var edge in Edges)
                if (DependsOn(edge.DependentSheet, edge.DependentCell, sheet, cell)) result.Add(edge);
            return result;
        }
        public void Dispose()
        {
            if (_handle != IntPtr.Zero) { Native.xlpp_dependency_graph_destroy(_handle); _handle = IntPtr.Zero; }
            GC.SuppressFinalize(this);
        }
        ~FormulaDependencyGraph() { Dispose(); }
    }

    public partial class Workbook
    {
        private static NativeLoadOptions BuildNativeLoadOptions(LoadOptions options, out Native.CancelCallback? cancel, out Native.ProgressCallback? progress)
        {
            cancel = options.Cancel is null ? null : _ => options.Cancel() ? 1 : 0;
            progress = options.Progress is null ? null : (_, done, total) => options.Progress(done, total);
            return new NativeLoadOptions { Lenient = options.Lenient ? 1 : 0, MaxEntries = options.MaxEntries,
                MaxEntryBytes = options.MaxEntryBytes, MaxTotalBytes = options.MaxTotalBytes, MaxFileBytes = options.MaxFileBytes,
                Password = options.Password, VerifyEncryptionIntegrity = options.VerifyEncryptionIntegrity ? 1 : 0,
                Cancel = cancel is null ? IntPtr.Zero : Marshal.GetFunctionPointerForDelegate(cancel),
                Progress = progress is null ? IntPtr.Zero : Marshal.GetFunctionPointerForDelegate(progress) };
        }

        public bool Load(string path, LoadOptions options)
        {
            var n = BuildNativeLoadOptions(options, out var cancel, out var progress);
            var ok = Native.xlpp_workbook_load_ex(_handle, path, ref n) != 0;
            GC.KeepAlive(cancel); GC.KeepAlive(progress);
            return ok;
        }

        public bool Load(byte[] bytes, LoadOptions? options = null)
        {
            options ??= new LoadOptions();
            var n = BuildNativeLoadOptions(options, out var cancel, out var progress);
            var ok = Native.xlpp_workbook_load_bytes(_handle, bytes, (ulong)bytes.LongLength, ref n) != 0;
            GC.KeepAlive(cancel); GC.KeepAlive(progress);
            return ok;
        }

        public bool Save(string path, SaveOptions options)
        {
            var n = new NativeSaveOptions { CompressionLevel = (int)options.CompressionLevel, CompressionStrategy = (int)options.CompressionStrategy,
                ParallelWorkers = options.ParallelWorkers, ParallelSheets = options.ParallelSheets ? 1 : 0, ParallelRows = options.ParallelRows ? 1 : 0,
                StrictNamespace = options.StrictNamespace ? 1 : 0, SynchronizeChartCaches = options.SynchronizeChartCaches ? 1 : 0,
                SynchronizeChangedChartCachesOnly = options.SynchronizeChangedChartCachesOnly ? 1 : 0,
                CalculateFormulasBeforeSave = options.CalculateFormulasBeforeSave ? 1 : 0, AtomicWrite = options.AtomicWrite ? 1 : 0,
                ValidateBeforeSave = options.ValidateBeforeSave ? 1 : 0, EncryptionPassword = options.EncryptionPassword,
                EncryptionMode = (int)options.EncryptionMode, EncryptionSpinCount = options.EncryptionSpinCount,
                EncryptionKeyBits = options.EncryptionKeyBits };
            return Native.xlpp_workbook_save_ex(_handle, path, ref n) != 0;
        }

        public byte[] SaveBytes(SaveOptions? options = null)
        {
            options ??= new SaveOptions();
            var n = new NativeSaveOptions { CompressionLevel = (int)options.CompressionLevel, CompressionStrategy = (int)options.CompressionStrategy,
                ParallelWorkers = options.ParallelWorkers, ParallelSheets = options.ParallelSheets ? 1 : 0, ParallelRows = options.ParallelRows ? 1 : 0,
                StrictNamespace = options.StrictNamespace ? 1 : 0, SynchronizeChartCaches = options.SynchronizeChartCaches ? 1 : 0,
                SynchronizeChangedChartCachesOnly = options.SynchronizeChangedChartCachesOnly ? 1 : 0,
                CalculateFormulasBeforeSave = options.CalculateFormulasBeforeSave ? 1 : 0, AtomicWrite = options.AtomicWrite ? 1 : 0,
                ValidateBeforeSave = options.ValidateBeforeSave ? 1 : 0, EncryptionPassword = options.EncryptionPassword,
                EncryptionMode = (int)options.EncryptionMode, EncryptionSpinCount = options.EncryptionSpinCount,
                EncryptionKeyBits = options.EncryptionKeyBits };
            if (Native.xlpp_workbook_save_bytes(_handle, ref n, out var buffer, out var size) == 0)
                throw new InvalidOperationException("Workbook save-to-bytes failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            try
            {
                if (size > int.MaxValue) throw new InvalidOperationException("Managed byte arrays cannot exceed Int32.MaxValue");
                var result = new byte[(int)size];
                if (size != 0) Marshal.Copy(buffer, result, 0, (int)size);
                return result;
            }
            finally { Native.xlpp_free_bytes(buffer); }
        }

        public CalculationReport CalculateFormulas(CalculationOptions options)
        {
            Native.ExternalReferenceResolverCallback? resolverCallback = null;
            var n = new NativeCalculationOptions { RecursiveDependencies = options.RecursiveDependencies ? 1 : 0,
                UpdateCachedValues = options.UpdateCachedValues ? 1 : 0, EvaluateVolatileFunctions = options.EvaluateVolatileFunctions ? 1 : 0,
                SpillDynamicArrays = options.SpillDynamicArrays ? 1 : 0, IterativeCalculation = options.IterativeCalculation ? 1 : 0,
                MaxIterations = options.MaxIterations, MaxChange = options.MaxChange, MaxDepth = options.MaxDepth };
            if (options.ExternalReferenceResolver is not null)
            {
                resolverCallback = (_, workbookToken, sheetName, address, output) =>
                {
                    var resolved = options.ExternalReferenceResolver(workbookToken, sheetName, address);
                    if (!resolved.HasValue) return 0;
                    var value = resolved.Value;
                    switch (value.Type)
                    {
                        case CellValueType.Empty: Native.xlpp_external_value_set_empty(output); break;
                        case CellValueType.Number: Native.xlpp_external_value_set_number(output, Convert.ToDouble(value.Value)); break;
                        case CellValueType.Bool: Native.xlpp_external_value_set_bool(output, Convert.ToBoolean(value.Value) ? 1 : 0); break;
                        case CellValueType.String: Native.xlpp_external_value_set_string(output, Convert.ToString(value.Value) ?? string.Empty); break;
                        case CellValueType.Error: Native.xlpp_external_value_set_error(output, (int)(value.Value is CellError e ? e : CellError.Value)); break;
                        case CellValueType.Date:
                            if (value.Value is not DateTime dt) return 0;
                            Native.xlpp_external_value_set_date(output, dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second + dt.Millisecond / 1000.0, 1);
                            break;
                        default: return 0;
                    }
                    return 1;
                };
                n.ExternalReferenceResolver = Marshal.GetFunctionPointerForDelegate(resolverCallback);
            }
            var calculateOk = Native.xlpp_workbook_calculate_options(_handle, ref n, out var r);
            GC.KeepAlive(resolverCallback);
            if (calculateOk == 0)
                throw new InvalidOperationException("Formula calculation failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            return new CalculationReport(r.FormulaCellsVisited, r.FormulaCellsEvaluated, r.CachedValuesUpdated,
                r.DependencyEvaluations, r.DefinedNamesResolved, r.CircularReferences, r.UnsupportedFormulas,
                r.EvaluationErrors, r.DynamicArraysSpilled, r.SpillCellsUpdated, r.SpillConflicts,
                r.StructuredReferencesResolved, r.IterativeIterations, r.IterativeConvergenceFailures,
                r.ExternalReferencesResolved, r.UnresolvedExternalReferences, r.Success != 0);
        }

        public LoadDiagnostics Diagnostics
        {
            get
            {
                var warnings = new List<string>();
                for (ulong i = 0, n = Native.xlpp_workbook_diagnostic_warning_count(_handle); i < n; ++i)
                    warnings.Add(AdvancedMarshal.FromSizedBuffer((b,s2) => Native.xlpp_workbook_diagnostic_warning(_handle,i,b,s2)));
                var errors = new List<string>();
                for (ulong i = 0, n = Native.xlpp_workbook_diagnostic_error_count(_handle); i < n; ++i)
                    errors.Add(AdvancedMarshal.FromSizedBuffer((b,s2) => Native.xlpp_workbook_diagnostic_error(_handle,i,b,s2)));
                return new LoadDiagnostics(warnings, errors);
            }
        }

        public EnterpriseFeatureInspection InspectEnterpriseFeatures()
        {
            var handle = Native.xlpp_workbook_inspect_enterprise_features(_handle);
            if (handle == IntPtr.Zero)
                throw new InvalidOperationException("Enterprise feature inspection failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            try
            {
                var features = new List<EnterpriseFeatureInfo>();
                var featureCount = Native.xlpp_enterprise_feature_count(handle);
                for (var featureIndex = 0; featureIndex < featureCount; ++featureIndex)
                {
                    var fi = featureIndex;
                    var pivotTables = new List<string>();
                    for (var i = 0; i < Native.xlpp_enterprise_feature_pivot_table_count(handle, fi); ++i)
                    {
                        var pi = i;
                        pivotTables.Add(AdvancedMarshal.FromSizedBuffer((buffer, size) =>
                            Native.xlpp_enterprise_feature_pivot_table_at(handle, fi, pi, buffer, size)));
                    }

                    var relationships = new List<EnterpriseFeatureRelationship>();
                    for (var i = 0; i < Native.xlpp_enterprise_feature_relationship_count(handle, fi); ++i)
                    {
                        var ri = i;
                        relationships.Add(new EnterpriseFeatureRelationship(
                            AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_relationship_source_part(handle, fi, ri, buffer, size)),
                            AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_relationship_id(handle, fi, ri, buffer, size)),
                            AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_relationship_type(handle, fi, ri, buffer, size)),
                            AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_relationship_target_mode(handle, fi, ri, buffer, size)),
                            AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_relationship_target_part(handle, fi, ri, buffer, size)),
                            Native.xlpp_enterprise_relationship_outgoing(handle, fi, ri) != 0));
                    }

                    features.Add(new EnterpriseFeatureInfo(
                        (EnterpriseFeatureKind)Native.xlpp_enterprise_feature_kind(handle, fi),
                        AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_feature_part_name(handle, fi, buffer, size)),
                        AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_feature_content_type(handle, fi, buffer, size)),
                        AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_feature_name(handle, fi, buffer, size)),
                        AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_feature_source_name(handle, fi, buffer, size)),
                        AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_feature_connection_id(handle, fi, buffer, size)),
                        AdvancedMarshal.FromSizedBuffer((buffer, size) => Native.xlpp_enterprise_feature_cache_id(handle, fi, buffer, size)),
                        pivotTables, relationships,
                        Native.xlpp_enterprise_feature_binary(handle, fi) != 0,
                        Native.xlpp_enterprise_feature_semantic_editable(handle, fi) != 0,
                        Native.xlpp_enterprise_feature_has_refresh_on_load(handle, fi) != 0,
                        Native.xlpp_enterprise_feature_refresh_on_load(handle, fi) != 0));
                }

                var warnings = new List<string>();
                for (var i = 0; i < Native.xlpp_enterprise_warning_count(handle); ++i)
                {
                    var wi = i;
                    warnings.Add(AdvancedMarshal.FromSizedBuffer((buffer, size) =>
                        Native.xlpp_enterprise_warning_at(handle, wi, buffer, size)));
                }
                return new EnterpriseFeatureInspection(features, warnings);
            }
            finally { Native.xlpp_enterprise_inspection_destroy(handle); }
        }

        private static EnterpriseEditReport ReadEnterpriseEditReport(IntPtr handle)
        {
            if (handle == IntPtr.Zero)
                throw new InvalidOperationException("Enterprise feature edit failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            try
            {
                var warnings = new List<string>();
                for (var i = 0; i < Native.xlpp_enterprise_edit_report_warning_count(handle); ++i)
                {
                    var wi = i;
                    warnings.Add(AdvancedMarshal.FromSizedBuffer((buffer, size) =>
                        Native.xlpp_enterprise_edit_report_warning_at(handle, wi, buffer, size)));
                }
                return new EnterpriseEditReport(Native.xlpp_enterprise_edit_report_matched(handle),
                    Native.xlpp_enterprise_edit_report_modified(handle), warnings,
                    Native.xlpp_enterprise_edit_report_success(handle) != 0);
            }
            finally { Native.xlpp_enterprise_edit_report_destroy(handle); }
        }

        public EnterpriseEditReport SetConnectionRefreshOnLoad(string connectionId, bool enabled) =>
            ReadEnterpriseEditReport(Native.xlpp_workbook_set_connection_refresh_on_load(_handle, connectionId, enabled ? 1 : 0));
        public EnterpriseEditReport SetQueryTableRefreshOnLoad(string queryName, bool enabled) =>
            ReadEnterpriseEditReport(Native.xlpp_workbook_set_query_table_refresh_on_load(_handle, queryName, enabled ? 1 : 0));
        public EnterpriseEditReport SetOlapPivotCacheRefreshOnLoad(string partName, bool enabled) =>
            ReadEnterpriseEditReport(Native.xlpp_workbook_set_olap_pivot_cache_refresh_on_load(_handle, partName, enabled ? 1 : 0));
        public EnterpriseEditReport SetPivotChartSourceName(string partName, string sourceName) =>
            ReadEnterpriseEditReport(Native.xlpp_workbook_set_pivot_chart_source_name(_handle, partName, sourceName));

        public StructuralEditReport ApplyStructuralEdit(string sheetName, StructuralEditKind kind, ulong index, ulong amount, StructuralEditOptions options)
        {
            var n = new NativeStructuralOptions { Transactional = options.Transactional ? 1 : 0, UpdateDefinedNames = options.UpdateDefinedNames ? 1 : 0,
                RecalculateFormulas = options.RecalculateFormulas ? 1 : 0, SynchronizeChartCaches = options.SynchronizeChartCaches ? 1 : 0,
                ChangedChartCachesOnly = options.ChangedChartCachesOnly ? 1 : 0, FailOnInvalidReference = options.FailOnInvalidReference ? 1 : 0 };
            var ok = Native.xlpp_workbook_structural_edit_ex(_handle, sheetName, (int)kind, index, amount, ref n, out var r);
            if (ok == 0 && r.Success == 0) throw new InvalidOperationException("Structural edit failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            return new StructuralEditReport(r.WorksheetsVisited, r.CellsMoved, r.CellsRemoved, r.FormulasUpdated,
                r.FormulaMetadataUpdated, r.WorksheetReferencesUpdated, r.DefinedNamesUpdated, r.ChartReferencesUpdated,
                r.PivotReferencesUpdated, r.DrawingAnchorsUpdated, r.HyperlinksUpdated, r.ReferencesInvalidated,
                r.FormulasCalculated, r.ChartCachesUpdated, r.Success != 0);
        }

        public WorksheetRenameReport RenameWorksheet(string oldName, string newName, WorksheetRenameOptions? options = null)
        {
            options ??= new WorksheetRenameOptions();
            var ok = Native.xlpp_workbook_rename_sheet(_handle, oldName, newName, options.RecalculateFormulas ? 1 : 0,
                options.SynchronizeChartCaches ? 1 : 0, options.ChangedChartCachesOnly ? 1 : 0, out var r);
            if (ok == 0 && r.Success == 0) throw new InvalidOperationException("Worksheet rename failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            return new WorksheetRenameReport(r.WorksheetsVisited, r.FormulasUpdated, r.FormulaMetadataUpdated,
                r.DefinedNamesUpdated, r.ChartReferencesUpdated, r.PivotReferencesUpdated, r.HyperlinksUpdated,
                r.ReferencesUpdated, r.FormulasCalculated, r.ChartCachesUpdated, r.Success != 0);
        }

        public ChartCacheSyncReport SynchronizeChartCaches(ChartCacheSyncOptions? options = null)
        {
            options ??= new ChartCacheSyncOptions();
            var n = new NativeChartCacheSyncOptions { SynchronizeTitles = options.SynchronizeTitles ? 1 : 0,
                SynchronizeCategories = options.SynchronizeCategories ? 1 : 0, SynchronizeValues = options.SynchronizeValues ? 1 : 0,
                ChangedReferencesOnly = options.ChangedReferencesOnly ? 1 : 0, ClearUnsupportedReferences = options.ClearUnsupportedReferences ? 1 : 0 };
            var ok = Native.xlpp_workbook_synchronize_chart_caches(_handle, ref n, out var r);
            if (ok == 0 && r.Success == 0) throw new InvalidOperationException("Chart cache synchronization failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            return new ChartCacheSyncReport(r.ChartsVisited, r.SeriesVisited, r.ReferencesChecked, r.ReferencesUnchanged,
                r.DependenciesRegistered, r.DependenciesChanged, r.CachesUpdated, r.CachesCleared, r.ReferencesSkipped, r.Success != 0);
        }
        public void ResetChartCacheDependencyTracking() => Native.xlpp_workbook_reset_chart_cache_tracking(_handle);
        public ulong TrackedChartCacheDependencyCount => Native.xlpp_workbook_tracked_chart_cache_dependencies(_handle);

        public FormulaDependencyGraph DependencyGraph()
        {
            var handle = Native.xlpp_workbook_dependency_graph(_handle);
            if (handle == IntPtr.Zero) throw new InvalidOperationException("Dependency graph failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            return new FormulaDependencyGraph(handle);
        }

        public WorkbookValidationReport Validate(WorkbookValidationOptions? options = null)
        {
            options ??= new WorkbookValidationOptions();
            var n = new NativeValidationOptions { ValidateWorksheetNames = options.ValidateWorksheetNames ? 1 : 0,
                ValidateDefinedNames = options.ValidateDefinedNames ? 1 : 0, ValidateTables = options.ValidateTables ? 1 : 0,
                ValidatePivots = options.ValidatePivots ? 1 : 0 };
            var handle = Native.xlpp_workbook_validate(_handle, ref n);
            if (handle == IntPtr.Zero) throw new InvalidOperationException("Workbook validation failed: " + MarshalHelper.PtrToString(Native.xlpp_last_error()));
            try
            {
                var issues = new List<WorkbookValidationIssue>();
                var count = Native.xlpp_validation_issue_count(handle);
                for (ulong i = 0; i < count; ++i)
                    issues.Add(new WorkbookValidationIssue((WorkbookValidationSeverity)Native.xlpp_validation_issue_severity(handle, i),
                        AdvancedMarshal.FromSizedBuffer((b,n2) => Native.xlpp_validation_issue_code(handle,i,b,n2)),
                        AdvancedMarshal.FromSizedBuffer((b,n2) => Native.xlpp_validation_issue_message(handle,i,b,n2)),
                        AdvancedMarshal.FromSizedBuffer((b,n2) => Native.xlpp_validation_issue_worksheet(handle,i,b,n2))));
                return new WorkbookValidationReport(Native.xlpp_validation_error_count(handle), Native.xlpp_validation_warning_count(handle), issues);
            }
            finally { Native.xlpp_validation_report_destroy(handle); }
        }

        public bool AddVbaProject(string path) => Native.xlpp_workbook_add_vba_project(_handle, path) != 0;
        public bool SetVbaProject(byte[] bytes) => Native.xlpp_workbook_set_vba_project(_handle, bytes, (ulong)bytes.LongLength) != 0;
        public bool HasVbaProject => Native.xlpp_workbook_has_vba_project(_handle) != 0;
        public bool RemoveVbaProject() => Native.xlpp_workbook_remove_vba_project(_handle) != 0;
        public bool SetVbaModule(VbaModule module)
            => module.Type switch
            {
                VbaModuleType.Standard => SetVbaModuleText(module.Name, module.Source),
                VbaModuleType.Class => SetVbaClassModuleText(module.Name, module.Source, module.ReadOnly, module.PrivateModule),
                VbaModuleType.Document => SetVbaDocumentModuleText(module.Name, module.Source),
                _ => false
            };
        public bool SetVbaModuleText(string moduleName, string source) => Native.xlpp_workbook_set_vba_module_text(_handle, moduleName, source) != 0;
        public bool SetVbaClassModuleText(string moduleName, string source, bool readOnly = false, bool privateModule = false)
            => Native.xlpp_workbook_set_vba_class_module_text(_handle, moduleName, source, readOnly ? 1 : 0, privateModule ? 1 : 0) != 0;
        public bool SetVbaDocumentModuleText(string moduleName, string source)
            => Native.xlpp_workbook_set_vba_document_module_text(_handle, moduleName, source) != 0;
        public string? VbaModuleText(string moduleName)
        {
            var required = Native.xlpp_workbook_vba_module_text(_handle, moduleName, IntPtr.Zero, 0);
            if (required == 0) return null;
            return AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_workbook_vba_module_text(_handle,moduleName,b,n));
        }
        public IReadOnlyList<VbaModule> VbaModules
        {
            get
            {
                var result = new List<VbaModule>();
                var count = Native.xlpp_workbook_vba_module_count(_handle);
                for (ulong i = 0; i < count; ++i)
                    result.Add(new VbaModule(
                        AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_workbook_vba_module_name(_handle,i,b,n)),
                        AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_workbook_vba_module_source(_handle,i,b,n)),
                        (VbaModuleType)Native.xlpp_workbook_vba_module_type(_handle,i),
                        Native.xlpp_workbook_vba_module_read_only(_handle,i) != 0,
                        Native.xlpp_workbook_vba_module_private(_handle,i) != 0));
                return result;
            }
        }
        public bool RemoveVbaModule(string moduleName) => Native.xlpp_workbook_remove_vba_module(_handle, moduleName) != 0;
        public byte[] VbaProjectBytes
        {
            get
            {
                var required = Native.xlpp_workbook_vba_project_bytes(_handle, IntPtr.Zero, 0);
                if (required == 0) return Array.Empty<byte>();
                var ptr = Marshal.AllocHGlobal(checked((int)required));
                try
                {
                    Native.xlpp_workbook_vba_project_bytes(_handle, ptr, required);
                    var result = new byte[checked((int)required)];
                    Marshal.Copy(ptr, result, 0, result.Length);
                    return result;
                }
                finally { Marshal.FreeHGlobal(ptr); }
            }
        }
        public bool SaveVbaProject(string path) => Native.xlpp_workbook_save_vba_project(_handle, path) != 0;
        public bool HasVbaSignature => Native.xlpp_workbook_has_vba_signature(_handle) != 0;
        public bool VbaSourceEditable => Native.xlpp_workbook_vba_source_editable(_handle) != 0;
        public VbaProjectProperties VbaProjectProperties
        {
            get => new(
                AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_workbook_vba_project_name(_handle,b,n)),
                AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_workbook_vba_project_description(_handle,b,n)),
                AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_workbook_vba_project_help_file(_handle,b,n)),
                Native.xlpp_workbook_vba_project_help_context(_handle),
                AdvancedMarshal.FromSizedBuffer((b,n) => Native.xlpp_workbook_vba_project_constants(_handle,b,n)));
            set => SetVbaProjectProperties(value);
        }
        public void SetVbaProjectProperties(VbaProjectProperties value)
        {
            if (Native.xlpp_workbook_set_vba_project_properties(_handle, value.Name, value.Description, value.HelpFile, value.HelpContextId, value.Constants) == 0)
                throw new InvalidOperationException("Failed to update VBA project properties");
        }
    }
}
