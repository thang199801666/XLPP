using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace XLPP
{
    public enum CompressionStrategy { Default = 0, Filtered = 1, HuffmanOnly = 2, Rle = 3, Fixed = 4 }
    public enum FormulaDependencyKind { CellOrRange = 0, DefinedName = 1, Table = 2, ExternalReference = 3, VolatileReference = 4 }
    public enum WorkbookValidationSeverity { Warning = 0, Error = 1 }
    public enum VbaModuleType { Standard = 0, Document = 1, Class = 2 }

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
