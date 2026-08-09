// XL++ C# Bindings — P/Invoke wrapper with OOP API
// Requires xlpp_capi.dll in the same directory or PATH.

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace XLPP
{
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
    internal struct NativeExternalDataSummary
    {
        public ulong ExternalWorkbooks, Connections, QueryTables, PowerQueryParts, WebQueryParts, UnknownConnectionParts;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeDataModelSummary
    {
        public int Present, HasOlapPivotCaches;
        public ulong ModelParts, ModelRelationships, OlapPivotCacheParts, Warnings;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeStructuralReport
    {
        public ulong WorksheetsVisited, CellsMoved, CellsRemoved, FormulasUpdated, FormulaMetadataUpdated;
        public ulong WorksheetReferencesUpdated, DefinedNamesUpdated, ChartReferencesUpdated, PivotReferencesUpdated;
        public ulong DrawingAnchorsUpdated, HyperlinksUpdated, ReferencesInvalidated, FormulasCalculated, ChartCachesUpdated;
        public int Success;
    }

    public readonly record struct CalculationReport(ulong FormulaCellsVisited, ulong FormulaCellsEvaluated,
        ulong CachedValuesUpdated, ulong DependencyEvaluations, ulong DefinedNamesResolved,
        ulong CircularReferences, ulong UnsupportedFormulas, ulong EvaluationErrors, ulong DynamicArraysSpilled,
        ulong SpillCellsUpdated, ulong SpillConflicts, ulong StructuredReferencesResolved, ulong IterativeIterations,
        ulong IterativeConvergenceFailures, ulong ExternalReferencesResolved, ulong UnresolvedExternalReferences, bool Success);

    public readonly record struct ExternalDataInspection(ulong ExternalWorkbooks, ulong Connections, ulong QueryTables,
        ulong PowerQueryParts, ulong WebQueryParts, ulong UnknownConnectionParts);

    public readonly record struct DataModelInspection(bool Present, bool HasOlapPivotCaches, ulong ModelParts,
        ulong ModelRelationships, ulong OlapPivotCacheParts, ulong Warnings);

    public readonly record struct StructuralEditReport(ulong WorksheetsVisited, ulong CellsMoved, ulong CellsRemoved,
        ulong FormulasUpdated, ulong FormulaMetadataUpdated, ulong WorksheetReferencesUpdated, ulong DefinedNamesUpdated,
        ulong ChartReferencesUpdated, ulong PivotReferencesUpdated, ulong DrawingAnchorsUpdated, ulong HyperlinksUpdated,
        ulong ReferencesInvalidated, ulong FormulasCalculated, ulong ChartCachesUpdated, bool Success);

    public enum StructuralEditKind { InsertRows = 0, DeleteRows = 1, InsertColumns = 2, DeleteColumns = 3 }
    public enum OfficeEncryptionMode { AgileAes256Sha512 = 1, StandardAesSha1 = 2 }

    // === Native interop ===
    internal static partial class Native
    {
        private const string Dll = "xlpp_capi";

        static Native()
        {
            NativeLibrary.SetDllImportResolver(
                typeof(Native).Assembly,
                (name, assembly, searchPath) =>
                {
                    if (name != Dll && name != Dll + ".dll" && name != "lib" + Dll + ".so" &&
                        name != "lib" + Dll + ".dylib")
                        return IntPtr.Zero;

                    if (NativeLibrary.TryLoad(Dll, assembly, searchPath, out var handle))
                        return handle;

                    string nativeName = Dll + ".dll";
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) nativeName = "lib" + Dll + ".so";
                    else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) nativeName = "lib" + Dll + ".dylib";
                    var rid = RuntimeInformation.RuntimeIdentifier;
                    if (!string.IsNullOrEmpty(rid))
                    {
                        var candidate = Path.Combine(AppContext.BaseDirectory, "runtimes", rid, "native", nativeName);
                        if (File.Exists(candidate)) return NativeLibrary.Load(candidate);
                    }

                    var local = Path.Combine(AppContext.BaseDirectory, nativeName);
                    if (File.Exists(local)) return NativeLibrary.Load(local);

                    return IntPtr.Zero;
                });
        }

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_version();

        // Workbook
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_c_abi_version();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_capabilities();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_create();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_workbook_destroy(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_add_sheet(IntPtr wb, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_sheet_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_get_sheet(IntPtr wb, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_sheet_by_name(IntPtr wb, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_remove_sheet(IntPtr wb, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_copy_sheet(IntPtr wb, IntPtr src, string newName);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_sheet_index(IntPtr wb, IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_sheet_names_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_workbook_sheet_name(IntPtr wb, int index, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_load(IntPtr wb, string path);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_save(IntPtr wb, string path);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_save_durable(IntPtr wb, string path, int durableWrite);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_load_password(IntPtr wb, string path, string password, int verifyIntegrity);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_save_encrypted(IntPtr wb, string path, string password, ulong spinCount, int calculateFormulas);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_save_encrypted_ex(IntPtr wb, string path, string password, int encryptionMode, ulong keyBits, ulong spinCount, int calculateFormulas);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_calculate(IntPtr wb, out NativeCalculationReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_calculate_ex(IntPtr wb, int iterative, ulong maxIterations, double maxChange, out NativeCalculationReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_structural_edit(IntPtr wb, string sheetName, int kind, ulong index, ulong amount, int failOnInvalidReference, out NativeStructuralReport report);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_workbook_set_date1904(IntPtr wb, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_date1904(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_workbook_clear(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_strict_namespaces(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_inspect_external_data(IntPtr wb, out NativeExternalDataSummary summary);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_inspect_data_model(IntPtr wb, out NativeDataModelSummary summary);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_properties(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_protection(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_calc_properties(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_custom_properties(IntPtr wb);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_add_named_style(IntPtr wb, string name, out int ok);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_named_style(IntPtr wb, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_named_styles_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_named_style_at(IntPtr wb, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_workbook_apply_named_style(IntPtr wb, IntPtr cell, string name);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_add_defined_name(IntPtr wb, string name, string value, out int ok);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_add_defined_name_scoped(IntPtr wb, string name, string value, ulong localSheetId, out int ok);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_defined_name(IntPtr wb, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_defined_name_scoped(IntPtr wb, string name, int hasLocalSheetId, ulong localSheetId);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_defined_names_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_defined_name_at(IntPtr wb, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_preserved_relationships_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_preserved_relationship_source_part(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_preserved_relationship_id(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_preserved_relationship_type(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_preserved_relationship_target(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_preserved_relationship_target_mode(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern int xlpp_workbook_preserved_parts_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern int xlpp_workbook_preserved_part_name(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern ulong xlpp_workbook_preserved_part_data_size(IntPtr wb, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern ulong xlpp_workbook_preserved_part_data(IntPtr wb, int index, IntPtr output, ulong size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern int xlpp_workbook_preserved_part_override_type(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern int xlpp_workbook_preserved_part_extension(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern int xlpp_workbook_preserved_part_default_type(IntPtr wb, int index, IntPtr output, int size);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern int xlpp_workbook_preserved_part_compress(IntPtr wb, int index);

        // Properties
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_title(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_creator(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_subject(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_description(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_keywords(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_category(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_last_modified_by(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_title(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_creator(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_subject(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_description(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_keywords(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_category(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_last_modified_by(IntPtr p);

        // Worksheet
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_cell(IntPtr ws, string address);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_cell_rc(IntPtr ws, ulong row, ulong col);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_sheet_has_cell(IntPtr ws, string address);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_has_cell_rc(IntPtr ws, ulong row, ulong col);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_merge_cells(IntPtr ws, string range);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_unmerge_cells(IntPtr ws, string range);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_sheet_is_merged(IntPtr ws, string cell);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_merged_count(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_merged_at(IntPtr ws, int index, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_freeze_panes(IntPtr ws, string cell);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_range(IntPtr ws, string address);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_dimensions(IntPtr ws, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_empty(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_sheet_row_count(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_sheet_col_count(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_clear_freeze_panes(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_frozen_pane(IntPtr ws, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_row_dimension(IntPtr ws, ulong row);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_col_dimension(IntPtr ws, ulong col);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_sheet_max_row(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_sheet_max_col(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_name(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_vba_code_name(IntPtr ws, IntPtr output, int outputSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_sheet_set_vba_code_name(IntPtr ws, string codeName);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_rename(IntPtr ws, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_append_row(IntPtr ws, string[] values, int count);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_append_doubles(IntPtr ws, double[] values, int count);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_insert_rows(IntPtr ws, ulong index, ulong amount);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_delete_rows(IntPtr ws, ulong index, ulong amount);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_insert_cols(IntPtr ws, ulong index, ulong amount);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_delete_cols(IntPtr ws, ulong index, ulong amount);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_set_print_area(IntPtr ws, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_print_area(IntPtr ws, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_set_print_titles_rows(IntPtr ws, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_print_titles_rows(IntPtr ws, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_set_print_titles_cols(IntPtr ws, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_print_titles_cols(IntPtr ws, IntPtr outPtr, int outSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_auto_filter(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_conditional_formatting(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_data_validations(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_page_setup(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_page_margins(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_print_options(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_header_footer(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_protection(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_view(IntPtr ws);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_add_table(IntPtr ws, string name, string reference, out int ok);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_table(IntPtr ws, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_table_count(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_table_at(IntPtr ws, int index);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_add_image(IntPtr ws, string path, string anchor, out int ok);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_image_count(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_image_at(IntPtr ws, int index);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_add_chart(IntPtr ws, int type);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_chart_count(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_chart_at(IntPtr ws, int index);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_add_pivot(IntPtr ws, string name, string location);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_pivot_count(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_pivot_at(IntPtr ws, int index);

        // Cell
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_value_type(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_get_bool(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_cell_get_number(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_get_string(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_is_empty(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_has_value(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_is_numeric(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_is_string(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_is_bool(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_is_date(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_is_error(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_error_code(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_date(IntPtr c, out int year, out int month, out int day, out int hour, out int minute, out double second);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_address(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_cell_row(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_cell_column(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cell_set_string(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_number(IntPtr c, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_bool(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_empty(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_error(IntPtr c, int errorCode);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_date(IntPtr c, int year, int month, int day, int hour, int minute, double second, int hasTime);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_clear(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cell_set_formula(IntPtr c, string f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_get_formula(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_has_formula(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_shared_formula(IntPtr c, string f, uint sharedIndex, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_array_formula(IntPtr c, string f, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_dynamic_array_formula(IntPtr c, string f, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_clear_formula(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_style(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_font(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_fill(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_border(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_alignment(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cell_set_number_format(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_number_format(IntPtr c, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cell_set_named_style(IntPtr c, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_named_style(IntPtr c, IntPtr outPtr, int outSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_has_hyperlink(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_hyperlink(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cell_set_hyperlink(IntPtr c, string url);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_set_hyperlink_full(IntPtr c, string url, string display, string tooltip, int external);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_clear_hyperlink(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_has_comment(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_comment(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cell_set_comment(IntPtr c, string text, string author);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_clear_comment(IntPtr c);

        // CellRange
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_range_min_row(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_range_min_col(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_range_max_row(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_range_max_col(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_range_row_count(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_range_col_count(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_range_address(IntPtr r, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_range_cell(IntPtr r, ulong relRow, ulong relCol);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_range_set_value(IntPtr r, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_range_set_string(IntPtr r, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_range_clear(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_range_values(IntPtr r, IntPtr outPtr, ref int outCount);

        // Hyperlink / Comment
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_hyperlink_set_target(IntPtr h, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_hyperlink_set_display(IntPtr h, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_hyperlink_set_tooltip(IntPtr h, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_hyperlink_set_external(IntPtr h, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_hyperlink_target(IntPtr h);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_hyperlink_display(IntPtr h);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_hyperlink_tooltip(IntPtr h);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_hyperlink_external(IntPtr h);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_comment_set_text(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_comment_set_author(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_comment_text(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_comment_author(IntPtr c);

        // Named style / Defined name
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_namedstyle_set_name(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_namedstyle_name(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_namedstyle_style(IntPtr s);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_definedname_set_value(IntPtr d, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_definedname_name(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_definedname_value(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_definedname_set_local_sheet_id(IntPtr d, ulong v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_definedname_clear_local_sheet_id(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_definedname_has_local_sheet_id(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_definedname_local_sheet_id(IntPtr d, out int hasValue);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_definedname_set_hidden(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_definedname_hidden(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_definedname_set_comment(IntPtr d, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_definedname_comment(IntPtr d);

        // Styles
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_font_set_name(IntPtr f, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_set_size(IntPtr f, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_set_bold(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_set_italic(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_set_underline(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_set_strike(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_font_set_color(IntPtr f, string argb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_font_get_name(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_font_get_size(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_font_get_bold(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_font_get_italic(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_font_get_underline(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_font_get_strike(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_get_color(IntPtr f, IntPtr outPtr, int outSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_fill_set_pattern(IntPtr f, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_fill_set_fg_color(IntPtr f, string argb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_fill_set_bg_color(IntPtr f, string argb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_fill_get_pattern(IntPtr f, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_fill_get_fg_color(IntPtr f, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_fill_get_bg_color(IntPtr f, IntPtr outPtr, int outSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_left(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_right(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_top(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_bottom(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_diagonal(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_borderside_set_style(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_borderside_set_color(IntPtr s, string argb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_borderside_get_style(IntPtr s, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_borderside_get_color(IntPtr s, IntPtr outPtr, int outSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_alignment_set_horizontal(IntPtr a, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_alignment_set_vertical(IntPtr a, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_alignment_set_wrap_text(IntPtr a, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_alignment_set_shrink_to_fit(IntPtr a, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_alignment_set_text_rotation(IntPtr a, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_alignment_set_indent(IntPtr a, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_alignment_get_horizontal(IntPtr a, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_alignment_get_vertical(IntPtr a, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_alignment_get_wrap_text(IntPtr a);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_alignment_get_shrink_to_fit(IntPtr a);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_alignment_get_text_rotation(IntPtr a);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_alignment_get_indent(IntPtr a);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_style_font(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_style_fill(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_style_border(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_style_alignment(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_style_set_number_format(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_style_number_format(IntPtr s, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_style_set_num_fmt_id(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_style_num_fmt_id(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_style_set_locked(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_style_locked(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_style_set_hidden(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_style_hidden(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_style_is_default(IntPtr s);

        // Page setup / margins / print / header footer
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagesetup_set_orientation(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pagesetup_orientation(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagesetup_set_paper_size(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pagesetup_paper_size(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagesetup_set_scale(IntPtr p, uint v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint xlpp_pagesetup_scale(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagesetup_set_fit_to_width(IntPtr p, uint v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint xlpp_pagesetup_fit_to_width(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagesetup_set_fit_to_height(IntPtr p, uint v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint xlpp_pagesetup_fit_to_height(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagesetup_set_fit_to_page(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pagesetup_fit_to_page(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagesetup_set_black_and_white(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pagesetup_black_and_white(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagesetup_set_draft(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pagesetup_draft(IntPtr p);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagemargins_set_left(IntPtr m, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_pagemargins_left(IntPtr m);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagemargins_set_right(IntPtr m, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_pagemargins_right(IntPtr m);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagemargins_set_top(IntPtr m, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_pagemargins_top(IntPtr m);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagemargins_set_bottom(IntPtr m, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_pagemargins_bottom(IntPtr m);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagemargins_set_header(IntPtr m, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_pagemargins_header(IntPtr m);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pagemargins_set_footer(IntPtr m, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_pagemargins_footer(IntPtr m);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_printopts_set_horizontal_centered(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_printopts_horizontal_centered(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_printopts_set_vertical_centered(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_printopts_vertical_centered(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_printopts_set_headings(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_printopts_headings(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_printopts_set_grid_lines(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_printopts_grid_lines(IntPtr p);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_headerfooter_set_odd_header(IntPtr h, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_headerfooter_set_odd_footer(IntPtr h, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_headerfooter_set_even_header(IntPtr h, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_headerfooter_set_even_footer(IntPtr h, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_headerfooter_set_different_odd_even(IntPtr h, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_headerfooter_different_odd_even(IntPtr h);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_headerfooter_set_different_first(IntPtr h, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_headerfooter_different_first(IntPtr h);

        // Protection
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wbprotection_set_lock_structure(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wbprotection_lock_structure(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wbprotection_set_lock_windows(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wbprotection_lock_windows(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wbprotection_set_lock_revision(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wbprotection_lock_revision(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_wbprotection_set_password_hash(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_wbprotection_password_hash(IntPtr p);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_enabled(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_enabled(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_wssprotection_set_password_hash(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_wssprotection_password_hash(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_select_locked(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_select_locked(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_select_unlocked(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_select_unlocked(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_format_cells(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_format_cells(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_format_columns(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_format_columns(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_format_rows(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_format_rows(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_insert_rows(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_insert_rows(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_insert_columns(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_insert_columns(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_delete_rows(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_delete_rows(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_delete_columns(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_delete_columns(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_sort(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_sort(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_wssprotection_set_auto_filter(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_wssprotection_auto_filter(IntPtr p);

        // SheetView
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_workbook_view_id(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_workbook_view_id(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheetview_set_tab_color(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_clear_tab_color(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_tab_color(IntPtr s, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_zoom_scale(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_zoom_scale(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_zoom_scale_normal(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_zoom_scale_normal(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_show_grid_lines(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_show_grid_lines(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_tab_selected(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_tab_selected(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_right_to_left(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_right_to_left(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_show_outline_symbols(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_show_outline_symbols(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheetview_set_pane(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_pane(IntPtr s, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheetview_set_top_left_cell(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_top_left_cell(IntPtr s, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_x_split(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_x_split(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheetview_set_y_split(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheetview_y_split(IntPtr s);

        // AutoFilter / FilterColumn / SortState
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_autofilter_set_reference(IntPtr f, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_autofilter_reference(IntPtr f, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_autofilter_enabled(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_autofilter_clear(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_autofilter_column(IntPtr f, ulong columnId);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_autofilter_sort_state(IntPtr f);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_filtercol_column_id(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_filtercol_add_value(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_filtercol_clear_values(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_filtercol_value_count(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_filtercol_value_at(IntPtr c, int index, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_filtercol_set_and_mode(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_filtercol_and_mode(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_filtercol_set_include_blank(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_filtercol_include_blank(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sortstate_set_reference(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sortstate_reference(IntPtr s, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sortstate_set_case_sensitive(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sortstate_case_sensitive(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sortstate_add_condition(IntPtr s, string reference, int descending);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sortstate_clear(IntPtr s);

        // Conditional formatting
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cfcollection_entry_count(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cfcollection_add_entry(IntPtr c, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cfcollection_entry_at(IntPtr c, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cfcollection_clear(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cfcollection_empty(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cfentry_reference(IntPtr e, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cfentry_set_reference(IntPtr e, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cfentry_rule_count(IntPtr e);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cfentry_add_rule(IntPtr e, int type);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cfentry_rule_at(IntPtr e, int index);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cfrule_add_formula(IntPtr r, string f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cfrule_formula_count(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cfrule_formula_at(IntPtr r, int index, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cfrule_set_operator(IntPtr r, int op);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cfrule_operator(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cfrule_type(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cfrule_set_priority(IntPtr r, ulong v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_cfrule_priority(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cfrule_set_stop_if_true(IntPtr r, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cfrule_stop_if_true(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cfrule_set_differential_style(IntPtr r, IntPtr s);

        // Data validation
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_dvcollection_count(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_dvcollection_add(IntPtr c, int type, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_dvcollection_at(IntPtr c, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_dvcollection_clear(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_datavalidation_type(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_set_type(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_datavalidation_operator(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_set_operator(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_datavalidation_error_style(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_set_error_style(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_datavalidation_set_formula1(IntPtr d, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_formula1(IntPtr d, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_datavalidation_set_formula2(IntPtr d, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_formula2(IntPtr d, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_datavalidation_set_reference(IntPtr d, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_reference(IntPtr d, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_set_allow_blank(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_datavalidation_allow_blank(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_set_show_drop_down(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_datavalidation_show_drop_down(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_set_show_input_message(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_datavalidation_set_show_error_message(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_datavalidation_set_prompt_title(IntPtr d, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_datavalidation_set_prompt(IntPtr d, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_datavalidation_set_error_title(IntPtr d, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_datavalidation_set_error(IntPtr d, string v);

        // Table
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_table_set_display_name(IntPtr t, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_table_name(IntPtr t, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_table_display_name(IntPtr t, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_table_set_reference(IntPtr t, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_table_reference(IntPtr t, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_table_set_show_header_row(IntPtr t, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_table_show_header_row(IntPtr t);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_table_set_show_totals_row(IntPtr t, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_table_show_totals_row(IntPtr t);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_table_column_count(IntPtr t);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_table_column_at(IntPtr t, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_table_add_column(IntPtr t, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_table_style_info(IntPtr t);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_tablecolumn_set_name(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_tablecolumn_name(IntPtr c, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_tablecolumn_id(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_tablestyle_set_name(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_tablestyle_name(IntPtr s, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_tablestyle_set_show_first(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_tablestyle_show_first(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_tablestyle_set_show_last(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_tablestyle_show_last(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_tablestyle_set_show_row_stripes(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_tablestyle_show_row_stripes(IntPtr s);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_tablestyle_set_show_column_stripes(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_tablestyle_show_column_stripes(IntPtr s);

        // Chart
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_set_grouping(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_grouping(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chart_set_title(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_title(IntPtr c, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chart_set_x_axis_title(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chart_set_y_axis_title(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chart_set_style(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_set_width(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_width(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_set_height(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_height(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_set_show_legend(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_show_legend(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chart_set_legend_position(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_series_count(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_chart_series_at(IntPtr c, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_chart_add_series(IntPtr c, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_type(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_is_modern(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_plot_count(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_chart_add_plot(IntPtr c, int type, ulong firstSeries, ulong seriesCount, int secondaryAxes);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_plot_set_grouping(IntPtr c, int plotIndex, int grouping);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_plot_set_bar_direction(IntPtr c, int plotIndex, int direction);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_plot_set_scatter_style(IntPtr c, int plotIndex, int style);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_plot_set_bubble_options(IntPtr c, int plotIndex, int scale, int showNegative, int sizeRepresents, int bubble3d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_plot_set_histogram_bins(IntPtr c, int plotIndex, double binWidth, int binCount, int automaticBins);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_plot_set_histogram_bounds(IntPtr c, int plotIndex, int hasUnderflow, double underflow, int hasOverflow, double overflow);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_plot_set_box_whisker_options(IntPtr c, int plotIndex, int innerPoints, int outliers, int meanLine, int meanMarker, int quartileInclusive);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chart_plot_set_waterfall_connector_lines(IntPtr c, int plotIndex, int enabled);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chartseries_set_title(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chartseries_set_values_reference(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chartseries_set_categories_reference(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_chartseries_set_bubble_size_reference(IntPtr s, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chartseries_set_smooth(IntPtr s, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_chartseries_clear_smooth(IntPtr s);

        // Pivot
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_set_name(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivottable_name(IntPtr p, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_set_location(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivottable_location(IntPtr p, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_pivottable_cache(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_add_row_field(IntPtr p, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_add_column_field(IntPtr p, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_add_page_field(IntPtr p, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivottable_add_data_field(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_add_data_field_named(IntPtr p, string name, string subtotal);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivottable_set_layout(IntPtr p, int layout);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pivottable_layout(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivottable_set_flags(IntPtr p, int showEmptyRow, int showEmptyColumn, int showDrill, int enableDrill, int multipleFieldFilters, int showValuesRow, int subtotalHiddenItems);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivottable_set_page_layout(IntPtr p, int pageWrap, int overThenDown);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_set_style(IntPtr p, string styleName, string dataCaption);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivottable_set_display_options(IntPtr p, int rowGrandTotals, int columnGrandTotals, int preserveFormatting, int useAutoFormatting, int showRowHeaders, int showColumnHeaders, int showRowStripes, int showColumnStripes, int showLastColumn);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pivottable_data_field_count(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_configure_data_field(IntPtr p, int index, string subtotal, string caption, uint numberFormatId, string showDataAs, int baseField, int baseItem);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pivottable_row_field_count(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pivottable_column_field_count(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pivottable_page_field_count(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_pivottable_row_field_at(IntPtr p, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_pivottable_column_field_at(IntPtr p, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_pivottable_page_field_at(IntPtr p, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivottable_add_filter(IntPtr p, string type, int fieldIndex, int measureFieldIndex, string value1, string value2, double top10Value, int top10Percent, int top10Top);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivotcache_set_cache_id(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_pivotcache_cache_id(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_pivotcache_set_source_data(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_pivotcache_source_data(IntPtr c, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotcache_set_refresh_on_load(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotcache_set_save_data(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotcache_set_enable_refresh(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotcache_set_missing_items_limit(IntPtr c, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotcache_set_advanced_flags(IntPtr c, int backgroundQuery, int optimizeMemory, int upgradeOnRefresh, int supportSubquery, int supportAdvancedDrill);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_pivotcache_set_refreshed_by(IntPtr c, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotfield_set_repeat_item_labels(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotfield_set_compact(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotfield_set_outline(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotfield_set_show_dropdowns(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotfield_set_behavior(IntPtr f, int showAll, int sortType, int subtotalTop, int insertBlankRow, int includeNewItemsInFilter, int multipleItemSelectionAllowed, int selectedItemIndex, int insertPageBreak, int defaultSubtotal);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_pivotfield_add_subtotal(IntPtr f, string subtotal);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotfield_set_item_hidden(IntPtr f, int itemIndex, int hidden);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotfield_set_numeric_grouping(IntPtr f, int autoStart, int autoEnd, double start, double end, double interval);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_pivotfield_set_date_grouping(IntPtr f, int datePart, int autoStart, int autoEnd, string startDate, string endDate);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_pivotfield_clear_grouping(IntPtr f);

        // Image
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_image_set_anchor(IntPtr img, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_image_anchor(IntPtr img, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_image_extension(IntPtr img, IntPtr outPtr, int outSize);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_image_set_width(IntPtr img, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_image_width(IntPtr img);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_image_set_height(IntPtr img, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_image_height(IntPtr img);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_image_set_name(IntPtr img, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_image_name(IntPtr img, IntPtr outPtr, int outSize);

        // Row/Column dimension
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_rowdim_set_height(IntPtr d, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_rowdim_has_height(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_rowdim_height(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_rowdim_set_hidden(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_rowdim_hidden(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_rowdim_set_outline_level(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_rowdim_outline_level(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_rowdim_set_collapsed(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_rowdim_collapsed(IntPtr d);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_coldim_set_width(IntPtr d, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_coldim_has_width(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_coldim_width(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_coldim_set_hidden(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_coldim_hidden(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_coldim_set_best_fit(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_coldim_best_fit(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_coldim_set_outline_level(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_coldim_outline_level(IntPtr d);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_coldim_set_collapsed(IntPtr d, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_coldim_collapsed(IntPtr d);

        // Calc properties
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_calcprops_set_calc_id(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_calcprops_calc_id(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_calcprops_set_calc_mode(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_calcprops_calc_mode(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_calcprops_set_calc_on_save(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_calcprops_calc_on_save(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_calcprops_set_full_calc_on_load(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_calcprops_full_calc_on_load(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_calcprops_set_full_precision(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_calcprops_full_precision(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_calcprops_set_iterate(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_calcprops_iterate(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_calcprops_set_iterate_count(IntPtr p, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_calcprops_iterate_count(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_calcprops_set_iterate_delta(IntPtr p, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_calcprops_iterate_delta(IntPtr p);

        // Custom properties
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_customprops_add(IntPtr c, string name, string value, int type);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_customprops_count(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_customprops_at(IntPtr c, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_customprop_name(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_customprop_value(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_customprop_type(IntPtr p);

        // Streaming writer
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_stream_create(string path);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_stream_destroy(IntPtr w);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern ulong xlpp_stream_add_sheet(IntPtr w, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_stream_append_row(IntPtr w, ulong sheetIndex, string[] values, int count);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_stream_append_doubles(IntPtr w, ulong sheetIndex, double[] values, int count);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_stream_row_count(IntPtr w, ulong sheetIndex);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_stream_sheet_count(IntPtr w);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_stream_set_date1904(IntPtr w, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_stream_set_compression_level(IntPtr w, int level);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_stream_set_parallel_workers(IntPtr w, ulong workers);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_stream_close(IntPtr w);

        // Streaming reader
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_stream_reader_open(string path);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_stream_reader_destroy(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_stream_reader_sheet_count(IntPtr r);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_stream_reader_sheet_name(IntPtr r, int index, IntPtr outPtr, int outSize);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        public delegate int StreamRowCallback(IntPtr user, ulong rowNumber, int cellCount,
            IntPtr addresses, IntPtr numbers, IntPtr valueTypes, IntPtr strings,
            IntPtr formulas, IntPtr styleIndexes);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_stream_reader_read_sheet(IntPtr r, int index,
            StreamRowCallback callback, IntPtr user);
    }

    // === Value types ===
    public enum CellValueType { Empty = 0, Bool = 1, Number = 2, String = 3, Error = 4, Date = 5 }
    public enum CellError { Null = 0, DivisionByZero = 1, Value = 2, Reference = 3, Name = 4, Number = 5, NotAvailable = 6, GettingData = 7, Spill = 8, Calculation = 9 }
    public enum CompressionLevel { Store = 0, Fastest = 1, Default = 2, Best = 3 }
    public enum PageOrientation { Default = 0, Portrait = 1, Landscape = 2 }
    public enum PaperSize { Default = 0, Letter = 1, Legal = 5, A4 = 9, A3 = 8 }
    public enum DataValidationType { None = 0, Whole = 1, Decimal = 2, List = 3, Date = 4, Time = 5, TextLength = 6, Custom = 7 }
    public enum DataValidationOperator { Between = 0, NotBetween = 1, Equal = 2, NotEqual = 3, LessThan = 4, LessThanOrEqual = 5, GreaterThan = 6, GreaterThanOrEqual = 7 }
    public enum DataValidationErrorStyle { Stop = 0, Warning = 1, Information = 2 }
    public enum ChartType { Bar = 0, Line = 1, Pie = 2, Scatter = 3, Doughnut = 4, Radar = 5, Area = 6, Bubble = 7, Stock = 8, Bar3D = 9, Line3D = 10, Area3D = 11, Pie3D = 12, Surface = 13, Surface3D = 14, PieOfPie = 15, BarOfPie = 16, HorizontalBar = 17, Histogram = 18, Pareto = 19, BoxWhisker = 20, Waterfall = 21, Funnel = 22, Treemap = 23, Sunburst = 24, FilledMap = 25 }
    public enum ChartBarDirection { Column = 0, Bar = 1 }
    public enum ChartScatterStyle { None = 0, Line = 1, LineMarker = 2, Marker = 3, Smooth = 4, SmoothMarker = 5 }
    public enum ChartBubbleSizeRepresents { Area = 0, Width = 1 }
    public enum PivotLayout { Compact = 0, Outline = 1, Tabular = 2 }
    public enum PivotDatePart { Seconds = 0, Minutes = 1, Hours = 2, Days = 3, Months = 4, Quarters = 5, Years = 6 }
    public enum ChartGrouping { Standard = 0, Stacked = 1, PercentStacked = 2, Clustered = 3 }
    public enum ConditionalRuleType { Formula = 0, CellIs = 1, DataBar = 2, ColorScale = 3, IconSet = 4 }
    public enum ConditionalOperator { Equal = 0, NotEqual = 1, LessThan = 2, LessThanOrEqual = 3, GreaterThan = 4, GreaterThanOrEqual = 5, Between = 6, NotBetween = 7 }

    // === OOP wrappers ===
    internal static class MarshalHelper
    {
        public static string PtrToString(IntPtr ptr)
        {
            if (ptr == IntPtr.Zero) return "";
            return Marshal.PtrToStringAnsi(ptr) ?? "";
        }

        public static string BufString(IntPtr buf)
        {
            return PtrToString(buf);
        }

        public static string FromBuffer(Action<IntPtr, int> fill)
        {
            var buf = Marshal.AllocHGlobal(4096);
            try
            {
                fill(buf, 4096);
                return PtrToString(buf);
            }
            finally { Marshal.FreeHGlobal(buf); }
        }
    }

    public class Font
    {
        internal IntPtr _handle;
        internal Font(IntPtr h) => _handle = h;
        public void SetName(string v) => Native.xlpp_font_set_name(_handle, v);
        public void SetSize(double v) => Native.xlpp_font_set_size(_handle, v);
        public void SetBold(bool v) => Native.xlpp_font_set_bold(_handle, v ? 1 : 0);
        public void SetItalic(bool v) => Native.xlpp_font_set_italic(_handle, v ? 1 : 0);
        public void SetUnderline(bool v) => Native.xlpp_font_set_underline(_handle, v ? 1 : 0);
        public void SetStrike(bool v) => Native.xlpp_font_set_strike(_handle, v ? 1 : 0);
        public void SetColor(string argb) => Native.xlpp_font_set_color(_handle, argb);
        public string Name => MarshalHelper.PtrToString(Native.xlpp_font_get_name(_handle));
        public double Size => Native.xlpp_font_get_size(_handle);
        public bool Bold => Native.xlpp_font_get_bold(_handle) != 0;
        public bool Italic => Native.xlpp_font_get_italic(_handle) != 0;
        public bool Underline => Native.xlpp_font_get_underline(_handle) != 0;
        public bool Strike => Native.xlpp_font_get_strike(_handle) != 0;
        public string Color => MarshalHelper.FromBuffer((b, n) => Native.xlpp_font_get_color(_handle, b, n));
    }

    public class Fill
    {
        internal IntPtr _handle;
        internal Fill(IntPtr h) => _handle = h;
        public void SetPattern(string v) => Native.xlpp_fill_set_pattern(_handle, v);
        public void SetForegroundColor(string argb) => Native.xlpp_fill_set_fg_color(_handle, argb);
        public void SetBackgroundColor(string argb) => Native.xlpp_fill_set_bg_color(_handle, argb);
        public string Pattern => MarshalHelper.FromBuffer((b, n) => Native.xlpp_fill_get_pattern(_handle, b, n));
        public string ForegroundColor => MarshalHelper.FromBuffer((b, n) => Native.xlpp_fill_get_fg_color(_handle, b, n));
        public string BackgroundColor => MarshalHelper.FromBuffer((b, n) => Native.xlpp_fill_get_bg_color(_handle, b, n));
    }

    public class BorderSide
    {
        internal IntPtr _handle;
        internal BorderSide(IntPtr h) => _handle = h;
        public void SetStyle(string v) => Native.xlpp_borderside_set_style(_handle, v);
        public void SetColor(string argb) => Native.xlpp_borderside_set_color(_handle, argb);
        public string Style => MarshalHelper.FromBuffer((b, n) => Native.xlpp_borderside_get_style(_handle, b, n));
        public string Color => MarshalHelper.FromBuffer((b, n) => Native.xlpp_borderside_get_color(_handle, b, n));
    }

    public class Border
    {
        internal IntPtr _handle;
        internal Border(IntPtr h) => _handle = h;
        public BorderSide Left => new(Native.xlpp_border_left(_handle));
        public BorderSide Right => new(Native.xlpp_border_right(_handle));
        public BorderSide Top => new(Native.xlpp_border_top(_handle));
        public BorderSide Bottom => new(Native.xlpp_border_bottom(_handle));
        public BorderSide Diagonal => new(Native.xlpp_border_diagonal(_handle));
    }

    public class Alignment
    {
        internal IntPtr _handle;
        internal Alignment(IntPtr h) => _handle = h;
        public void SetHorizontal(string v) => Native.xlpp_alignment_set_horizontal(_handle, v);
        public void SetVertical(string v) => Native.xlpp_alignment_set_vertical(_handle, v);
        public void SetWrapText(bool v) => Native.xlpp_alignment_set_wrap_text(_handle, v ? 1 : 0);
        public void SetShrinkToFit(bool v) => Native.xlpp_alignment_set_shrink_to_fit(_handle, v ? 1 : 0);
        public void SetTextRotation(int v) => Native.xlpp_alignment_set_text_rotation(_handle, v);
        public void SetIndent(int v) => Native.xlpp_alignment_set_indent(_handle, v);
        public string Horizontal => MarshalHelper.FromBuffer((b, n) => Native.xlpp_alignment_get_horizontal(_handle, b, n));
        public string Vertical => MarshalHelper.FromBuffer((b, n) => Native.xlpp_alignment_get_vertical(_handle, b, n));
        public bool WrapText => Native.xlpp_alignment_get_wrap_text(_handle) != 0;
        public bool ShrinkToFit => Native.xlpp_alignment_get_shrink_to_fit(_handle) != 0;
        public int TextRotation => Native.xlpp_alignment_get_text_rotation(_handle);
        public int Indent => Native.xlpp_alignment_get_indent(_handle);
    }

    public class Style
    {
        internal IntPtr _handle;
        internal Style(IntPtr h) => _handle = h;
        public Font Font => new(Native.xlpp_style_font(_handle));
        public Fill Fill => new(Native.xlpp_style_fill(_handle));
        public Border Border => new(Native.xlpp_style_border(_handle));
        public Alignment Alignment => new(Native.xlpp_style_alignment(_handle));
        public string NumberFormat
        {
            get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_style_number_format(_handle, b, n));
            set => Native.xlpp_style_set_number_format(_handle, value);
        }
        public int NumFmtId { get => Native.xlpp_style_num_fmt_id(_handle); set => Native.xlpp_style_set_num_fmt_id(_handle, value); }
        public bool Locked { get => Native.xlpp_style_locked(_handle) != 0; set => Native.xlpp_style_set_locked(_handle, value ? 1 : 0); }
        public bool Hidden { get => Native.xlpp_style_hidden(_handle) != 0; set => Native.xlpp_style_set_hidden(_handle, value ? 1 : 0); }
        public bool IsDefault => Native.xlpp_style_is_default(_handle) != 0;
    }

    public class Properties
    {
        internal IntPtr _handle;
        internal Properties(IntPtr h) => _handle = h;
        public string Title { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_title(_handle)); set => Native.xlpp_properties_set_title(_handle, value); }
        public string Creator { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_creator(_handle)); set => Native.xlpp_properties_set_creator(_handle, value); }
        public string Subject { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_subject(_handle)); set => Native.xlpp_properties_set_subject(_handle, value); }
        public string Description { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_description(_handle)); set => Native.xlpp_properties_set_description(_handle, value); }
        public string Keywords { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_keywords(_handle)); set => Native.xlpp_properties_set_keywords(_handle, value); }
        public string Category { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_category(_handle)); set => Native.xlpp_properties_set_category(_handle, value); }
        public string LastModifiedBy { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_last_modified_by(_handle)); set => Native.xlpp_properties_set_last_modified_by(_handle, value); }
    }

    public class WorkbookProtection
    {
        internal IntPtr _handle;
        internal WorkbookProtection(IntPtr h) => _handle = h;
        public bool LockStructure { get => Native.xlpp_wbprotection_lock_structure(_handle) != 0; set => Native.xlpp_wbprotection_set_lock_structure(_handle, value ? 1 : 0); }
        public bool LockWindows { get => Native.xlpp_wbprotection_lock_windows(_handle) != 0; set => Native.xlpp_wbprotection_set_lock_windows(_handle, value ? 1 : 0); }
        public bool LockRevision { get => Native.xlpp_wbprotection_lock_revision(_handle) != 0; set => Native.xlpp_wbprotection_set_lock_revision(_handle, value ? 1 : 0); }
        public string PasswordHash { get => MarshalHelper.PtrToString(Native.xlpp_wbprotection_password_hash(_handle)); set => Native.xlpp_wbprotection_set_password_hash(_handle, value); }
    }

    public class WorksheetProtection
    {
        internal IntPtr _handle;
        internal WorksheetProtection(IntPtr h) => _handle = h;
        public bool Enabled { get => Native.xlpp_wssprotection_enabled(_handle) != 0; set => Native.xlpp_wssprotection_set_enabled(_handle, value ? 1 : 0); }
        public string PasswordHash { get => MarshalHelper.PtrToString(Native.xlpp_wssprotection_password_hash(_handle)); set => Native.xlpp_wssprotection_set_password_hash(_handle, value); }
        public bool SelectLockedCells { get => Native.xlpp_wssprotection_select_locked(_handle) != 0; set => Native.xlpp_wssprotection_set_select_locked(_handle, value ? 1 : 0); }
        public bool SelectUnlockedCells { get => Native.xlpp_wssprotection_select_unlocked(_handle) != 0; set => Native.xlpp_wssprotection_set_select_unlocked(_handle, value ? 1 : 0); }
        public bool FormatCells { get => Native.xlpp_wssprotection_format_cells(_handle) != 0; set => Native.xlpp_wssprotection_set_format_cells(_handle, value ? 1 : 0); }
        public bool FormatColumns { get => Native.xlpp_wssprotection_format_columns(_handle) != 0; set => Native.xlpp_wssprotection_set_format_columns(_handle, value ? 1 : 0); }
        public bool FormatRows { get => Native.xlpp_wssprotection_format_rows(_handle) != 0; set => Native.xlpp_wssprotection_set_format_rows(_handle, value ? 1 : 0); }
        public bool InsertRows { get => Native.xlpp_wssprotection_insert_rows(_handle) != 0; set => Native.xlpp_wssprotection_set_insert_rows(_handle, value ? 1 : 0); }
        public bool InsertColumns { get => Native.xlpp_wssprotection_insert_columns(_handle) != 0; set => Native.xlpp_wssprotection_set_insert_columns(_handle, value ? 1 : 0); }
        public bool DeleteRows { get => Native.xlpp_wssprotection_delete_rows(_handle) != 0; set => Native.xlpp_wssprotection_set_delete_rows(_handle, value ? 1 : 0); }
        public bool DeleteColumns { get => Native.xlpp_wssprotection_delete_columns(_handle) != 0; set => Native.xlpp_wssprotection_set_delete_columns(_handle, value ? 1 : 0); }
        public bool Sort { get => Native.xlpp_wssprotection_sort(_handle) != 0; set => Native.xlpp_wssprotection_set_sort(_handle, value ? 1 : 0); }
        public bool AutoFilter { get => Native.xlpp_wssprotection_auto_filter(_handle) != 0; set => Native.xlpp_wssprotection_set_auto_filter(_handle, value ? 1 : 0); }
    }

    public class SheetView
    {
        internal IntPtr _handle;
        internal SheetView(IntPtr h) => _handle = h;
        public int WorkbookViewId { get => Native.xlpp_sheetview_workbook_view_id(_handle); set => Native.xlpp_sheetview_set_workbook_view_id(_handle, value); }
        public string TabColor { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheetview_tab_color(_handle, b, n)); set => Native.xlpp_sheetview_set_tab_color(_handle, value); }
        public void ClearTabColor() => Native.xlpp_sheetview_clear_tab_color(_handle);
        public int ZoomScale { get => Native.xlpp_sheetview_zoom_scale(_handle); set => Native.xlpp_sheetview_set_zoom_scale(_handle, value); }
        public int ZoomScaleNormal { get => Native.xlpp_sheetview_zoom_scale_normal(_handle); set => Native.xlpp_sheetview_set_zoom_scale_normal(_handle, value); }
        public bool ShowGridLines { get => Native.xlpp_sheetview_show_grid_lines(_handle) != 0; set => Native.xlpp_sheetview_set_show_grid_lines(_handle, value ? 1 : 0); }
        public bool TabSelected { get => Native.xlpp_sheetview_tab_selected(_handle) != 0; set => Native.xlpp_sheetview_set_tab_selected(_handle, value ? 1 : 0); }
        public bool RightToLeft { get => Native.xlpp_sheetview_right_to_left(_handle) != 0; set => Native.xlpp_sheetview_set_right_to_left(_handle, value ? 1 : 0); }
        public bool ShowOutlineSymbols { get => Native.xlpp_sheetview_show_outline_symbols(_handle) != 0; set => Native.xlpp_sheetview_set_show_outline_symbols(_handle, value ? 1 : 0); }
        public string Pane { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheetview_pane(_handle, b, n)); set => Native.xlpp_sheetview_set_pane(_handle, value); }
        public string TopLeftCell { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheetview_top_left_cell(_handle, b, n)); set => Native.xlpp_sheetview_set_top_left_cell(_handle, value); }
        public int XSplit { get => Native.xlpp_sheetview_x_split(_handle); set => Native.xlpp_sheetview_set_x_split(_handle, value); }
        public int YSplit { get => Native.xlpp_sheetview_y_split(_handle); set => Native.xlpp_sheetview_set_y_split(_handle, value); }
    }

    public class RowDimension
    {
        internal IntPtr _handle;
        internal RowDimension(IntPtr h) => _handle = h;
        public double? Height
        {
            get => Native.xlpp_rowdim_has_height(_handle) != 0 ? Native.xlpp_rowdim_height(_handle) : null;
            set => Native.xlpp_rowdim_set_height(_handle, value ?? 0);
        }
        public bool Hidden { get => Native.xlpp_rowdim_hidden(_handle) != 0; set => Native.xlpp_rowdim_set_hidden(_handle, value ? 1 : 0); }
        public int OutlineLevel { get => Native.xlpp_rowdim_outline_level(_handle); set => Native.xlpp_rowdim_set_outline_level(_handle, value); }
        public bool Collapsed { get => Native.xlpp_rowdim_collapsed(_handle) != 0; set => Native.xlpp_rowdim_set_collapsed(_handle, value ? 1 : 0); }
    }

    public class ColumnDimension
    {
        internal IntPtr _handle;
        internal ColumnDimension(IntPtr h) => _handle = h;
        public double? Width
        {
            get => Native.xlpp_coldim_has_width(_handle) != 0 ? Native.xlpp_coldim_width(_handle) : null;
            set => Native.xlpp_coldim_set_width(_handle, value ?? 0);
        }
        public bool Hidden { get => Native.xlpp_coldim_hidden(_handle) != 0; set => Native.xlpp_coldim_set_hidden(_handle, value ? 1 : 0); }
        public bool BestFit { get => Native.xlpp_coldim_best_fit(_handle) != 0; set => Native.xlpp_coldim_set_best_fit(_handle, value ? 1 : 0); }
        public int OutlineLevel { get => Native.xlpp_coldim_outline_level(_handle); set => Native.xlpp_coldim_set_outline_level(_handle, value); }
        public bool Collapsed { get => Native.xlpp_coldim_collapsed(_handle) != 0; set => Native.xlpp_coldim_set_collapsed(_handle, value ? 1 : 0); }
    }

    public class PageSetup
    {
        internal IntPtr _handle;
        internal PageSetup(IntPtr h) => _handle = h;
        public PageOrientation Orientation { get => (PageOrientation)Native.xlpp_pagesetup_orientation(_handle); set => Native.xlpp_pagesetup_set_orientation(_handle, (int)value); }
        public PaperSize PaperSize { get => (PaperSize)Native.xlpp_pagesetup_paper_size(_handle); set => Native.xlpp_pagesetup_set_paper_size(_handle, (int)value); }
        public uint Scale { get => Native.xlpp_pagesetup_scale(_handle); set => Native.xlpp_pagesetup_set_scale(_handle, value); }
        public uint FitToWidth { get => Native.xlpp_pagesetup_fit_to_width(_handle); set => Native.xlpp_pagesetup_set_fit_to_width(_handle, value); }
        public uint FitToHeight { get => Native.xlpp_pagesetup_fit_to_height(_handle); set => Native.xlpp_pagesetup_set_fit_to_height(_handle, value); }
        public bool FitToPage { get => Native.xlpp_pagesetup_fit_to_page(_handle) != 0; set => Native.xlpp_pagesetup_set_fit_to_page(_handle, value ? 1 : 0); }
        public bool BlackAndWhite { get => Native.xlpp_pagesetup_black_and_white(_handle) != 0; set => Native.xlpp_pagesetup_set_black_and_white(_handle, value ? 1 : 0); }
        public bool Draft { get => Native.xlpp_pagesetup_draft(_handle) != 0; set => Native.xlpp_pagesetup_set_draft(_handle, value ? 1 : 0); }
    }

    public class PageMargins
    {
        internal IntPtr _handle;
        internal PageMargins(IntPtr h) => _handle = h;
        public double Left { get => Native.xlpp_pagemargins_left(_handle); set => Native.xlpp_pagemargins_set_left(_handle, value); }
        public double Right { get => Native.xlpp_pagemargins_right(_handle); set => Native.xlpp_pagemargins_set_right(_handle, value); }
        public double Top { get => Native.xlpp_pagemargins_top(_handle); set => Native.xlpp_pagemargins_set_top(_handle, value); }
        public double Bottom { get => Native.xlpp_pagemargins_bottom(_handle); set => Native.xlpp_pagemargins_set_bottom(_handle, value); }
        public double Header { get => Native.xlpp_pagemargins_header(_handle); set => Native.xlpp_pagemargins_set_header(_handle, value); }
        public double Footer { get => Native.xlpp_pagemargins_footer(_handle); set => Native.xlpp_pagemargins_set_footer(_handle, value); }
    }

    public class PrintOptions
    {
        internal IntPtr _handle;
        internal PrintOptions(IntPtr h) => _handle = h;
        public bool HorizontalCentered { get => Native.xlpp_printopts_horizontal_centered(_handle) != 0; set => Native.xlpp_printopts_set_horizontal_centered(_handle, value ? 1 : 0); }
        public bool VerticalCentered { get => Native.xlpp_printopts_vertical_centered(_handle) != 0; set => Native.xlpp_printopts_set_vertical_centered(_handle, value ? 1 : 0); }
        public bool Headings { get => Native.xlpp_printopts_headings(_handle) != 0; set => Native.xlpp_printopts_set_headings(_handle, value ? 1 : 0); }
        public bool GridLines { get => Native.xlpp_printopts_grid_lines(_handle) != 0; set => Native.xlpp_printopts_set_grid_lines(_handle, value ? 1 : 0); }
    }

    public class HeaderFooter
    {
        internal IntPtr _handle;
        internal HeaderFooter(IntPtr h) => _handle = h;
        public string OddHeader { set => Native.xlpp_headerfooter_set_odd_header(_handle, value); }
        public string OddFooter { set => Native.xlpp_headerfooter_set_odd_footer(_handle, value); }
        public string EvenHeader { set => Native.xlpp_headerfooter_set_even_header(_handle, value); }
        public string EvenFooter { set => Native.xlpp_headerfooter_set_even_footer(_handle, value); }
        public bool DifferentOddEven { get => Native.xlpp_headerfooter_different_odd_even(_handle) != 0; set => Native.xlpp_headerfooter_set_different_odd_even(_handle, value ? 1 : 0); }
        public bool DifferentFirst { get => Native.xlpp_headerfooter_different_first(_handle) != 0; set => Native.xlpp_headerfooter_set_different_first(_handle, value ? 1 : 0); }
    }

    public class AutoFilter
    {
        internal IntPtr _handle;
        internal AutoFilter(IntPtr h) => _handle = h;
        public string Reference { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_autofilter_reference(_handle, b, n)); set => Native.xlpp_autofilter_set_reference(_handle, value); }
        public bool Enabled => Native.xlpp_autofilter_enabled(_handle) != 0;
        public void Clear() => Native.xlpp_autofilter_clear(_handle);
        public FilterColumn Column(ulong columnId) => new(Native.xlpp_autofilter_column(_handle, columnId));
        public SortState SortState => new(Native.xlpp_autofilter_sort_state(_handle));
    }

    public class FilterColumn
    {
        internal IntPtr _handle;
        internal FilterColumn(IntPtr h) => _handle = h;
        public ulong ColumnId => Native.xlpp_filtercol_column_id(_handle);
        public void AddValue(string v) => Native.xlpp_filtercol_add_value(_handle, v);
        public void ClearValues() => Native.xlpp_filtercol_clear_values(_handle);
        public int ValueCount => Native.xlpp_filtercol_value_count(_handle);
        public string ValueAt(int index) => MarshalHelper.FromBuffer((b, n) => Native.xlpp_filtercol_value_at(_handle, index, b, n));
        public bool AndMode { get => Native.xlpp_filtercol_and_mode(_handle) != 0; set => Native.xlpp_filtercol_set_and_mode(_handle, value ? 1 : 0); }
        public bool IncludeBlank { get => Native.xlpp_filtercol_include_blank(_handle) != 0; set => Native.xlpp_filtercol_set_include_blank(_handle, value ? 1 : 0); }
    }

    public class SortState
    {
        internal IntPtr _handle;
        internal SortState(IntPtr h) => _handle = h;
        public string Reference { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sortstate_reference(_handle, b, n)); set => Native.xlpp_sortstate_set_reference(_handle, value); }
        public bool CaseSensitive { get => Native.xlpp_sortstate_case_sensitive(_handle) != 0; set => Native.xlpp_sortstate_set_case_sensitive(_handle, value ? 1 : 0); }
        public void AddCondition(string reference, bool descending = false) => Native.xlpp_sortstate_add_condition(_handle, reference, descending ? 1 : 0);
        public void Clear() => Native.xlpp_sortstate_clear(_handle);
    }

    public class Hyperlink
    {
        internal IntPtr _handle;
        internal Hyperlink(IntPtr h) => _handle = h;
        public string Target { get => MarshalHelper.PtrToString(Native.xlpp_hyperlink_target(_handle)); set => Native.xlpp_hyperlink_set_target(_handle, value); }
        public string Display { get => MarshalHelper.PtrToString(Native.xlpp_hyperlink_display(_handle)); set => Native.xlpp_hyperlink_set_display(_handle, value); }
        public string Tooltip { get => MarshalHelper.PtrToString(Native.xlpp_hyperlink_tooltip(_handle)); set => Native.xlpp_hyperlink_set_tooltip(_handle, value); }
        public bool External { get => Native.xlpp_hyperlink_external(_handle) != 0; set => Native.xlpp_hyperlink_set_external(_handle, value ? 1 : 0); }
    }

    public class Comment
    {
        internal IntPtr _handle;
        internal Comment(IntPtr h) => _handle = h;
        public string Text { get => MarshalHelper.PtrToString(Native.xlpp_comment_text(_handle)); set => Native.xlpp_comment_set_text(_handle, value); }
        public string Author { get => MarshalHelper.PtrToString(Native.xlpp_comment_author(_handle)); set => Native.xlpp_comment_set_author(_handle, value); }
    }

    public class NamedStyle
    {
        internal IntPtr _handle;
        internal NamedStyle(IntPtr h) => _handle = h;
        public string Name { get => MarshalHelper.PtrToString(Native.xlpp_namedstyle_name(_handle)); set => Native.xlpp_namedstyle_set_name(_handle, value); }
        public Style Style => new(Native.xlpp_namedstyle_style(_handle));
    }

    public class DefinedName
    {
        internal IntPtr _handle;
        internal DefinedName(IntPtr h) => _handle = h;
        public string Name => MarshalHelper.PtrToString(Native.xlpp_definedname_name(_handle));
        public string Value { get => MarshalHelper.PtrToString(Native.xlpp_definedname_value(_handle)); set => Native.xlpp_definedname_set_value(_handle, value); }
        public bool HasLocalSheetId => Native.xlpp_definedname_has_local_sheet_id(_handle) != 0;
        public ulong? LocalSheetId {
            get {
                var value = Native.xlpp_definedname_local_sheet_id(_handle, out var hasValue);
                return hasValue != 0 ? value : null;
            }
            set {
                if (value.HasValue) Native.xlpp_definedname_set_local_sheet_id(_handle, value.Value);
                else Native.xlpp_definedname_clear_local_sheet_id(_handle);
            }
        }
        public void SetLocalSheetId(ulong v) => Native.xlpp_definedname_set_local_sheet_id(_handle, v);
        public void ClearLocalSheetId() => Native.xlpp_definedname_clear_local_sheet_id(_handle);
        public bool Hidden { get => Native.xlpp_definedname_hidden(_handle) != 0; set => Native.xlpp_definedname_set_hidden(_handle, value ? 1 : 0); }
        public string Comment { get => MarshalHelper.PtrToString(Native.xlpp_definedname_comment(_handle)); set => Native.xlpp_definedname_set_comment(_handle, value); }
    }

    public class CalcProperties
    {
        internal IntPtr _handle;
        internal CalcProperties(IntPtr h) => _handle = h;
        public int CalcId { get => Native.xlpp_calcprops_calc_id(_handle); set => Native.xlpp_calcprops_set_calc_id(_handle, value); }
        public string CalcMode { get => MarshalHelper.PtrToString(Native.xlpp_calcprops_calc_mode(_handle)); set => Native.xlpp_calcprops_set_calc_mode(_handle, value); }
        public bool CalcOnSave { get => Native.xlpp_calcprops_calc_on_save(_handle) != 0; set => Native.xlpp_calcprops_set_calc_on_save(_handle, value ? 1 : 0); }
        public bool FullCalcOnLoad { get => Native.xlpp_calcprops_full_calc_on_load(_handle) != 0; set => Native.xlpp_calcprops_set_full_calc_on_load(_handle, value ? 1 : 0); }
        public bool FullPrecision { get => Native.xlpp_calcprops_full_precision(_handle) != 0; set => Native.xlpp_calcprops_set_full_precision(_handle, value ? 1 : 0); }
        public bool Iterate { get => Native.xlpp_calcprops_iterate(_handle) != 0; set => Native.xlpp_calcprops_set_iterate(_handle, value ? 1 : 0); }
        public int IterateCount { get => Native.xlpp_calcprops_iterate_count(_handle); set => Native.xlpp_calcprops_set_iterate_count(_handle, value); }
        public double IterateDelta { get => Native.xlpp_calcprops_iterate_delta(_handle); set => Native.xlpp_calcprops_set_iterate_delta(_handle, value); }
    }

    public class CustomProperty
    {
        internal IntPtr _handle;
        internal CustomProperty(IntPtr h) => _handle = h;
        public string Name => MarshalHelper.PtrToString(Native.xlpp_customprop_name(_handle));
        public string Value => MarshalHelper.PtrToString(Native.xlpp_customprop_value(_handle));
        public string Type => MarshalHelper.PtrToString(Native.xlpp_customprop_type(_handle));
    }

    public class CustomProperties
    {
        internal IntPtr _handle;
        internal CustomProperties(IntPtr h) => _handle = h;
        public CustomProperty Add(string name, string value, int type = 0) => new(Native.xlpp_customprops_add(_handle, name, value, type));
        public int Count => Native.xlpp_customprops_count(_handle);
        public CustomProperty At(int index) => new(Native.xlpp_customprops_at(_handle, index));
    }

    public class TableColumn
    {
        internal IntPtr _handle;
        internal TableColumn(IntPtr h) => _handle = h;
        public ulong Id => Native.xlpp_tablecolumn_id(_handle);
        public string Name { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_tablecolumn_name(_handle, b, n)); set => Native.xlpp_tablecolumn_set_name(_handle, value); }
    }

    public class TableStyleInfo
    {
        internal IntPtr _handle;
        internal TableStyleInfo(IntPtr h) => _handle = h;
        public string Name { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_tablestyle_name(_handle, b, n)); set => Native.xlpp_tablestyle_set_name(_handle, value); }
        public bool ShowFirstColumn { get => Native.xlpp_tablestyle_show_first(_handle) != 0; set => Native.xlpp_tablestyle_set_show_first(_handle, value ? 1 : 0); }
        public bool ShowLastColumn { get => Native.xlpp_tablestyle_show_last(_handle) != 0; set => Native.xlpp_tablestyle_set_show_last(_handle, value ? 1 : 0); }
        public bool ShowRowStripes { get => Native.xlpp_tablestyle_show_row_stripes(_handle) != 0; set => Native.xlpp_tablestyle_set_show_row_stripes(_handle, value ? 1 : 0); }
        public bool ShowColumnStripes { get => Native.xlpp_tablestyle_show_column_stripes(_handle) != 0; set => Native.xlpp_tablestyle_set_show_column_stripes(_handle, value ? 1 : 0); }
    }

    public class Table
    {
        internal IntPtr _handle;
        internal Table(IntPtr h) => _handle = h;
        public string Name => MarshalHelper.FromBuffer((b, n) => Native.xlpp_table_name(_handle, b, n));
        public string DisplayName { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_table_display_name(_handle, b, n)); set => Native.xlpp_table_set_display_name(_handle, value); }
        public string Reference { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_table_reference(_handle, b, n)); set => Native.xlpp_table_set_reference(_handle, value); }
        public bool ShowHeaderRow { get => Native.xlpp_table_show_header_row(_handle) != 0; set => Native.xlpp_table_set_show_header_row(_handle, value ? 1 : 0); }
        public bool ShowTotalsRow { get => Native.xlpp_table_show_totals_row(_handle) != 0; set => Native.xlpp_table_set_show_totals_row(_handle, value ? 1 : 0); }
        public int ColumnCount => Native.xlpp_table_column_count(_handle);
        public TableColumn ColumnAt(int index) => new(Native.xlpp_table_column_at(_handle, index));
        public TableColumn AddColumn(string name) => new(Native.xlpp_table_add_column(_handle, name));
        public TableStyleInfo StyleInfo => new(Native.xlpp_table_style_info(_handle));
    }

    public class ChartSeries
    {
        internal IntPtr _handle;
        internal ChartSeries(IntPtr h) => _handle = h;
        public string Title { set => Native.xlpp_chartseries_set_title(_handle, value); }
        public string ValuesReference { set => Native.xlpp_chartseries_set_values_reference(_handle, value); }
        public string CategoriesReference { set => Native.xlpp_chartseries_set_categories_reference(_handle, value); }
        public string BubbleSizeReference { set => Native.xlpp_chartseries_set_bubble_size_reference(_handle, value); }
        public bool Smooth { set => Native.xlpp_chartseries_set_smooth(_handle, value ? 1 : 0); }
        public void ClearSmooth() => Native.xlpp_chartseries_clear_smooth(_handle);
    }

    public class Chart
    {
        internal IntPtr _handle;
        internal Chart(IntPtr h) => _handle = h;
        public ChartGrouping Grouping { get => (ChartGrouping)Native.xlpp_chart_grouping(_handle); set => Native.xlpp_chart_set_grouping(_handle, (int)value); }
        public ChartType Type => (ChartType)Native.xlpp_chart_type(_handle);
        public bool IsModern => Native.xlpp_chart_is_modern(_handle) != 0;
        public int PlotCount => Native.xlpp_chart_plot_count(_handle);
        public int AddPlot(ChartType type, ulong firstSeries, ulong seriesCount, bool secondaryAxes = false) => Native.xlpp_chart_add_plot(_handle, (int)type, firstSeries, seriesCount, secondaryAxes ? 1 : 0);
        public void SetPlotGrouping(int plotIndex, ChartGrouping grouping) => Native.xlpp_chart_plot_set_grouping(_handle, plotIndex, (int)grouping);
        public void SetBarDirection(int plotIndex, ChartBarDirection direction) => Native.xlpp_chart_plot_set_bar_direction(_handle, plotIndex, (int)direction);
        public void SetScatterStyle(int plotIndex, ChartScatterStyle style) => Native.xlpp_chart_plot_set_scatter_style(_handle, plotIndex, (int)style);
        public void SetBubbleOptions(int plotIndex, int scale = 100, bool showNegative = false, ChartBubbleSizeRepresents sizeRepresents = ChartBubbleSizeRepresents.Area, bool bubble3D = false) => Native.xlpp_chart_plot_set_bubble_options(_handle, plotIndex, scale, showNegative ? 1 : 0, (int)sizeRepresents, bubble3D ? 1 : 0);
        public void SetHistogramBins(int plotIndex, double binWidth = 0, int binCount = 0, bool automatic = true) => Native.xlpp_chart_plot_set_histogram_bins(_handle, plotIndex, binWidth, binCount, automatic ? 1 : 0);
        public void SetHistogramBounds(int plotIndex, double? underflow = null, double? overflow = null) => Native.xlpp_chart_plot_set_histogram_bounds(_handle, plotIndex, underflow.HasValue ? 1 : 0, underflow ?? 0, overflow.HasValue ? 1 : 0, overflow ?? 0);
        public void SetBoxWhiskerOptions(int plotIndex, bool innerPoints = true, bool outliers = true, bool meanLine = false, bool meanMarker = true, bool quartileInclusive = false) => Native.xlpp_chart_plot_set_box_whisker_options(_handle, plotIndex, innerPoints ? 1 : 0, outliers ? 1 : 0, meanLine ? 1 : 0, meanMarker ? 1 : 0, quartileInclusive ? 1 : 0);
        public void SetWaterfallConnectorLines(int plotIndex, bool enabled) => Native.xlpp_chart_plot_set_waterfall_connector_lines(_handle, plotIndex, enabled ? 1 : 0);
        public string Title { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_chart_title(_handle, b, n)); set => Native.xlpp_chart_set_title(_handle, value); }
        public void SetXAxisTitle(string v) => Native.xlpp_chart_set_x_axis_title(_handle, v);
        public void SetYAxisTitle(string v) => Native.xlpp_chart_set_y_axis_title(_handle, v);
        public void SetStyle(string v) => Native.xlpp_chart_set_style(_handle, v);
        public int Width { get => Native.xlpp_chart_width(_handle); set => Native.xlpp_chart_set_width(_handle, value); }
        public int Height { get => Native.xlpp_chart_height(_handle); set => Native.xlpp_chart_set_height(_handle, value); }
        public bool ShowLegend { get => Native.xlpp_chart_show_legend(_handle) != 0; set => Native.xlpp_chart_set_show_legend(_handle, value ? 1 : 0); }
        public void SetLegendPosition(string v) => Native.xlpp_chart_set_legend_position(_handle, v);
        public int SeriesCount => Native.xlpp_chart_series_count(_handle);
        public ChartSeries SeriesAt(int index) => new(Native.xlpp_chart_series_at(_handle, index));
        public ChartSeries AddSeries(string title) => new(Native.xlpp_chart_add_series(_handle, title));
    }

    public class PivotCache
    {
        internal IntPtr _handle;
        internal PivotCache(IntPtr h) => _handle = h;
        public int CacheId { get => Native.xlpp_pivotcache_cache_id(_handle); set => Native.xlpp_pivotcache_set_cache_id(_handle, value); }
        public string SourceData { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_pivotcache_source_data(_handle, b, n)); set => Native.xlpp_pivotcache_set_source_data(_handle, value); }
        public bool RefreshOnLoad { set => Native.xlpp_pivotcache_set_refresh_on_load(_handle, value ? 1 : 0); }
        public bool SaveData { set => Native.xlpp_pivotcache_set_save_data(_handle, value ? 1 : 0); }
        public bool EnableRefresh { set => Native.xlpp_pivotcache_set_enable_refresh(_handle, value ? 1 : 0); }
        public int MissingItemsLimit { set => Native.xlpp_pivotcache_set_missing_items_limit(_handle, value); }
        public string RefreshedBy { set => Native.xlpp_pivotcache_set_refreshed_by(_handle, value); }
        public void SetAdvancedFlags(bool backgroundQuery = false, bool optimizeMemory = false, bool upgradeOnRefresh = false, bool supportSubquery = true, bool supportAdvancedDrill = true) => Native.xlpp_pivotcache_set_advanced_flags(_handle, backgroundQuery ? 1 : 0, optimizeMemory ? 1 : 0, upgradeOnRefresh ? 1 : 0, supportSubquery ? 1 : 0, supportAdvancedDrill ? 1 : 0);
    }

    public class PivotField
    {
        internal IntPtr _handle;
        internal PivotField(IntPtr h) => _handle = h;
        public bool RepeatItemLabels { set => Native.xlpp_pivotfield_set_repeat_item_labels(_handle, value ? 1 : 0); }
        public bool Compact { set => Native.xlpp_pivotfield_set_compact(_handle, value ? 1 : 0); }
        public bool Outline { set => Native.xlpp_pivotfield_set_outline(_handle, value ? 1 : 0); }
        public bool ShowDropDowns { set => Native.xlpp_pivotfield_set_show_dropdowns(_handle, value ? 1 : 0); }
        public void SetBehavior(bool showAll = false, int sortType = 0, bool subtotalTop = true, bool insertBlankRow = false, bool includeNewItemsInFilter = false, bool multipleItemSelectionAllowed = false, int selectedItemIndex = -1, bool insertPageBreak = false, bool defaultSubtotal = false) => Native.xlpp_pivotfield_set_behavior(_handle, showAll ? 1 : 0, sortType, subtotalTop ? 1 : 0, insertBlankRow ? 1 : 0, includeNewItemsInFilter ? 1 : 0, multipleItemSelectionAllowed ? 1 : 0, selectedItemIndex, insertPageBreak ? 1 : 0, defaultSubtotal ? 1 : 0);
        public void AddSubtotal(string subtotal) => Native.xlpp_pivotfield_add_subtotal(_handle, subtotal);
        public void SetItemHidden(int itemIndex, bool hidden = true) => Native.xlpp_pivotfield_set_item_hidden(_handle, itemIndex, hidden ? 1 : 0);
        public void SetNumericGrouping(double start, double end, double interval = 1, bool autoStart = false, bool autoEnd = false) => Native.xlpp_pivotfield_set_numeric_grouping(_handle, autoStart ? 1 : 0, autoEnd ? 1 : 0, start, end, interval);
        public void SetDateGrouping(PivotDatePart part, string startDate = "", string endDate = "", bool autoStart = true, bool autoEnd = true) => Native.xlpp_pivotfield_set_date_grouping(_handle, (int)part, autoStart ? 1 : 0, autoEnd ? 1 : 0, startDate, endDate);
        public void ClearGrouping() => Native.xlpp_pivotfield_clear_grouping(_handle);
    }

    public class PivotTable
    {
        internal IntPtr _handle;
        internal PivotTable(IntPtr h) => _handle = h;
        public string Name { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_pivottable_name(_handle, b, n)); set => Native.xlpp_pivottable_set_name(_handle, value); }
        public string Location { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_pivottable_location(_handle, b, n)); set => Native.xlpp_pivottable_set_location(_handle, value); }
        public PivotLayout Layout { get => (PivotLayout)Native.xlpp_pivottable_layout(_handle); set => Native.xlpp_pivottable_set_layout(_handle, (int)value); }
        public PivotCache Cache => new(Native.xlpp_pivottable_cache(_handle));
        public void AddRowField(string name) => Native.xlpp_pivottable_add_row_field(_handle, name);
        public void AddColumnField(string name) => Native.xlpp_pivottable_add_column_field(_handle, name);
        public void AddPageField(string name) => Native.xlpp_pivottable_add_page_field(_handle, name);
        public void AddDataField() => Native.xlpp_pivottable_add_data_field(_handle);
        public void AddDataField(string name, string subtotal = "sum") => Native.xlpp_pivottable_add_data_field_named(_handle, name, subtotal);
        public int RowFieldCount => Native.xlpp_pivottable_row_field_count(_handle);
        public int ColumnFieldCount => Native.xlpp_pivottable_column_field_count(_handle);
        public int PageFieldCount => Native.xlpp_pivottable_page_field_count(_handle);
        public int DataFieldCount => Native.xlpp_pivottable_data_field_count(_handle);
        public PivotField RowFieldAt(int index) => new(Native.xlpp_pivottable_row_field_at(_handle, index));
        public PivotField ColumnFieldAt(int index) => new(Native.xlpp_pivottable_column_field_at(_handle, index));
        public PivotField PageFieldAt(int index) => new(Native.xlpp_pivottable_page_field_at(_handle, index));
        public void SetBehaviorFlags(bool showEmptyRow = false, bool showEmptyColumn = false, bool showDrill = true, bool enableDrill = true, bool multipleFieldFilters = false, bool showValuesRow = false, bool subtotalHiddenItems = false) => Native.xlpp_pivottable_set_flags(_handle, showEmptyRow ? 1 : 0, showEmptyColumn ? 1 : 0, showDrill ? 1 : 0, enableDrill ? 1 : 0, multipleFieldFilters ? 1 : 0, showValuesRow ? 1 : 0, subtotalHiddenItems ? 1 : 0);
        public void SetPageLayout(int pageWrap, bool overThenDown = false) => Native.xlpp_pivottable_set_page_layout(_handle, pageWrap, overThenDown ? 1 : 0);
        public void SetStyle(string styleName = "PivotStyleLight16", string dataCaption = "Values") => Native.xlpp_pivottable_set_style(_handle, styleName, dataCaption);
        public void SetDisplayOptions(bool rowGrandTotals = true, bool columnGrandTotals = true, bool preserveFormatting = true, bool useAutoFormatting = true, bool showRowHeaders = true, bool showColumnHeaders = true, bool showRowStripes = false, bool showColumnStripes = false, bool showLastColumn = true) => Native.xlpp_pivottable_set_display_options(_handle, rowGrandTotals ? 1 : 0, columnGrandTotals ? 1 : 0, preserveFormatting ? 1 : 0, useAutoFormatting ? 1 : 0, showRowHeaders ? 1 : 0, showColumnHeaders ? 1 : 0, showRowStripes ? 1 : 0, showColumnStripes ? 1 : 0, showLastColumn ? 1 : 0);
        public void ConfigureDataField(int index, string subtotal = "sum", string caption = "", uint numberFormatId = 0, string showDataAs = "normal", int baseField = 0, int baseItem = 0) => Native.xlpp_pivottable_configure_data_field(_handle, index, subtotal, caption, numberFormatId, showDataAs, baseField, baseItem);
        public void AddFilter(string type, int fieldIndex, int measureFieldIndex = -1, string value1 = "", string value2 = "", double top10Value = 10, bool top10Percent = false, bool top10Top = true) => Native.xlpp_pivottable_add_filter(_handle, type, fieldIndex, measureFieldIndex, value1, value2, top10Value, top10Percent ? 1 : 0, top10Top ? 1 : 0);
    }

    public class Image
    {
        internal IntPtr _handle;
        internal Image(IntPtr h) => _handle = h;
        public string Anchor { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_image_anchor(_handle, b, n)); set => Native.xlpp_image_set_anchor(_handle, value); }
        public string Extension => MarshalHelper.FromBuffer((b, n) => Native.xlpp_image_extension(_handle, b, n));
        public double Width { get => Native.xlpp_image_width(_handle); set => Native.xlpp_image_set_width(_handle, value); }
        public double Height { get => Native.xlpp_image_height(_handle); set => Native.xlpp_image_set_height(_handle, value); }
        public string Name { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_image_name(_handle, b, n)); set => Native.xlpp_image_set_name(_handle, value); }
    }

    public class ConditionalFormattingCollection
    {
        internal IntPtr _handle;
        internal ConditionalFormattingCollection(IntPtr h) => _handle = h;
        public int EntryCount => Native.xlpp_cfcollection_entry_count(_handle);
        public ConditionalFormattingEntry AddEntry(string reference) => new(Native.xlpp_cfcollection_add_entry(_handle, reference));
        public ConditionalFormattingEntry EntryAt(int index) => new(Native.xlpp_cfcollection_entry_at(_handle, index));
        public void Clear() => Native.xlpp_cfcollection_clear(_handle);
        public bool IsEmpty => Native.xlpp_cfcollection_empty(_handle) != 0;
    }

    public class ConditionalFormattingEntry
    {
        internal IntPtr _handle;
        internal ConditionalFormattingEntry(IntPtr h) => _handle = h;
        public string Reference { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_cfentry_reference(_handle, b, n)); set => Native.xlpp_cfentry_set_reference(_handle, value); }
        public int RuleCount => Native.xlpp_cfentry_rule_count(_handle);
        public ConditionalRule AddRule(ConditionalRuleType type) => new(Native.xlpp_cfentry_add_rule(_handle, (int)type));
        public ConditionalRule RuleAt(int index) => new(Native.xlpp_cfentry_rule_at(_handle, index));
    }

    public class ConditionalRule
    {
        internal IntPtr _handle;
        internal ConditionalRule(IntPtr h) => _handle = h;
        public void AddFormula(string f) => Native.xlpp_cfrule_add_formula(_handle, f);
        public int FormulaCount => Native.xlpp_cfrule_formula_count(_handle);
        public string FormulaAt(int index) => MarshalHelper.FromBuffer((b, n) => Native.xlpp_cfrule_formula_at(_handle, index, b, n));
        public ConditionalOperator Operator { get => (ConditionalOperator)Native.xlpp_cfrule_operator(_handle); set => Native.xlpp_cfrule_set_operator(_handle, (int)value); }
        public ConditionalRuleType Type => (ConditionalRuleType)Native.xlpp_cfrule_type(_handle);
        public ulong Priority { get => Native.xlpp_cfrule_priority(_handle); set => Native.xlpp_cfrule_set_priority(_handle, value); }
        public bool StopIfTrue { get => Native.xlpp_cfrule_stop_if_true(_handle) != 0; set => Native.xlpp_cfrule_set_stop_if_true(_handle, value ? 1 : 0); }
        public void SetDifferentialStyle(Style s) => Native.xlpp_cfrule_set_differential_style(_handle, s._handle);
    }

    public class DataValidationCollection
    {
        internal IntPtr _handle;
        internal DataValidationCollection(IntPtr h) => _handle = h;
        public int Count => Native.xlpp_dvcollection_count(_handle);
        public DataValidation Add(DataValidationType type, string reference) => new(Native.xlpp_dvcollection_add(_handle, (int)type, reference));
        public DataValidation At(int index) => new(Native.xlpp_dvcollection_at(_handle, index));
        public void Clear() => Native.xlpp_dvcollection_clear(_handle);
    }

    public class DataValidation
    {
        internal IntPtr _handle;
        internal DataValidation(IntPtr h) => _handle = h;
        public DataValidationType Type { get => (DataValidationType)Native.xlpp_datavalidation_type(_handle); set => Native.xlpp_datavalidation_set_type(_handle, (int)value); }
        public DataValidationOperator Operator { get => (DataValidationOperator)Native.xlpp_datavalidation_operator(_handle); set => Native.xlpp_datavalidation_set_operator(_handle, (int)value); }
        public DataValidationErrorStyle ErrorStyle { get => (DataValidationErrorStyle)Native.xlpp_datavalidation_error_style(_handle); set => Native.xlpp_datavalidation_set_error_style(_handle, (int)value); }
        public string Formula1 { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_datavalidation_formula1(_handle, b, n)); set => Native.xlpp_datavalidation_set_formula1(_handle, value); }
        public string Formula2 { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_datavalidation_formula2(_handle, b, n)); set => Native.xlpp_datavalidation_set_formula2(_handle, value); }
        public string Reference { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_datavalidation_reference(_handle, b, n)); set => Native.xlpp_datavalidation_set_reference(_handle, value); }
        public bool AllowBlank { get => Native.xlpp_datavalidation_allow_blank(_handle) != 0; set => Native.xlpp_datavalidation_set_allow_blank(_handle, value ? 1 : 0); }
        public bool ShowDropDown { get => Native.xlpp_datavalidation_show_drop_down(_handle) != 0; set => Native.xlpp_datavalidation_set_show_drop_down(_handle, value ? 1 : 0); }
        public void SetShowInputMessage(bool v) => Native.xlpp_datavalidation_set_show_input_message(_handle, v ? 1 : 0);
        public void SetShowErrorMessage(bool v) => Native.xlpp_datavalidation_set_show_error_message(_handle, v ? 1 : 0);
        public void SetPromptTitle(string v) => Native.xlpp_datavalidation_set_prompt_title(_handle, v);
        public void SetPrompt(string v) => Native.xlpp_datavalidation_set_prompt(_handle, v);
        public void SetErrorTitle(string v) => Native.xlpp_datavalidation_set_error_title(_handle, v);
        public void SetError(string v) => Native.xlpp_datavalidation_set_error(_handle, v);
    }

    public class Cell
    {
        internal IntPtr _handle;
        internal Cell(IntPtr h) => _handle = h;

        public CellValueType ValueType => (CellValueType)Native.xlpp_cell_value_type(_handle);
        public bool IsEmpty => Native.xlpp_cell_is_empty(_handle) != 0;
        public bool HasValue => Native.xlpp_cell_has_value(_handle) != 0;
        public bool IsNumeric => Native.xlpp_cell_is_numeric(_handle) != 0;
        public bool IsString => Native.xlpp_cell_is_string(_handle) != 0;
        public bool IsBool => Native.xlpp_cell_is_bool(_handle) != 0;
        public bool IsDate => Native.xlpp_cell_is_date(_handle) != 0;
        public bool IsError => Native.xlpp_cell_is_error(_handle) != 0;
        public string Address => MarshalHelper.PtrToString(Native.xlpp_cell_address(_handle));
        public ulong Row => Native.xlpp_cell_row(_handle);
        public ulong Column => Native.xlpp_cell_column(_handle);

        public CellError? Error
        {
            get
            {
                var code = Native.xlpp_cell_error_code(_handle);
                return code < 0 ? null : (CellError)code;
            }
        }

        public DateTime? Date
        {
            get
            {
                if (Native.xlpp_cell_date(_handle, out var y, out var mo, out var d, out var h, out var mi, out var s) == 0)
                    return null;
                return new DateTime((int)y, (int)mo, (int)d, (int)h, (int)mi, (int)s);
            }
        }

        public object? Value
        {
            get => ValueType switch
            {
                CellValueType.Bool => Native.xlpp_cell_get_bool(_handle) != 0,
                CellValueType.Number => Native.xlpp_cell_get_number(_handle),
                CellValueType.String => MarshalHelper.PtrToString(Native.xlpp_cell_get_string(_handle)),
                CellValueType.Date => Date,
                CellValueType.Error => Error?.ToString(),
                _ => null
            };
            set
            {
                if (value == null) return;
                switch (value)
                {
                    case string s: Native.xlpp_cell_set_string(_handle, s); break;
                    case double d: Native.xlpp_cell_set_number(_handle, d); break;
                    case int i: Native.xlpp_cell_set_number(_handle, i); break;
                    case long l: Native.xlpp_cell_set_number(_handle, l); break;
                    case float f: Native.xlpp_cell_set_number(_handle, f); break;
                    case bool b: Native.xlpp_cell_set_bool(_handle, b ? 1 : 0); break;
                    case DateTime dt: SetDate(dt); break;
                }
            }
        }

        public void SetDate(DateTime dt)
        {
            bool hasTime = dt.Hour != 0 || dt.Minute != 0 || dt.Second != 0;
            Native.xlpp_cell_set_date(_handle, dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second, hasTime ? 1 : 0);
        }

        public void SetError(CellError error) => Native.xlpp_cell_set_error(_handle, (int)error);
        public void SetEmpty() => Native.xlpp_cell_set_empty(_handle);
        public void Clear() => Native.xlpp_cell_clear(_handle);

        public string? Formula
        {
            get => MarshalHelper.PtrToString(Native.xlpp_cell_get_formula(_handle));
            set => Native.xlpp_cell_set_formula(_handle, value ?? "");
        }

        public bool HasFormula => Native.xlpp_cell_has_formula(_handle) != 0;
        public void SetSharedFormula(string f, uint sharedIndex, string reference = "") => Native.xlpp_cell_set_shared_formula(_handle, f, sharedIndex, reference);
        public void SetArrayFormula(string f, string reference) => Native.xlpp_cell_set_array_formula(_handle, f, reference);
        public void SetDynamicArrayFormula(string f, string reference) => Native.xlpp_cell_set_dynamic_array_formula(_handle, f, reference);
        public void ClearFormula() => Native.xlpp_cell_clear_formula(_handle);

        public Font Font => new(Native.xlpp_cell_font(_handle));
        public Fill Fill => new(Native.xlpp_cell_fill(_handle));
        public Border Border => new(Native.xlpp_cell_border(_handle));
        public Alignment Alignment => new(Native.xlpp_cell_alignment(_handle));
        public Style Style => new(Native.xlpp_cell_style(_handle));
        public string NumberFormat
        {
            get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_cell_number_format(_handle, b, n));
            set => Native.xlpp_cell_set_number_format(_handle, value);
        }
        public void SetNamedStyle(string name) => Native.xlpp_cell_set_named_style(_handle, name);
        public string NamedStyle => MarshalHelper.FromBuffer((b, n) => Native.xlpp_cell_named_style(_handle, b, n));

        public bool HasHyperlink => Native.xlpp_cell_has_hyperlink(_handle) != 0;
        public Hyperlink Hyperlink => new(Native.xlpp_cell_hyperlink(_handle));
        public void SetHyperlink(string url) => Native.xlpp_cell_set_hyperlink(_handle, url);
        public void SetHyperlink(string url, string display, string tooltip, bool external = true) =>
            Native.xlpp_cell_set_hyperlink_full(_handle, url, display, tooltip, external ? 1 : 0);
        public void ClearHyperlink() => Native.xlpp_cell_clear_hyperlink(_handle);

        public bool HasComment => Native.xlpp_cell_has_comment(_handle) != 0;
        public Comment Comment => new(Native.xlpp_cell_comment(_handle));
        public void SetComment(string text, string author) => Native.xlpp_cell_set_comment(_handle, text, author);
        public void ClearComment() => Native.xlpp_cell_clear_comment(_handle);
    }

    public class CellRange
    {
        internal IntPtr _handle;
        internal CellRange(IntPtr h) => _handle = h;
        public ulong MinRow => Native.xlpp_range_min_row(_handle);
        public ulong MinColumn => Native.xlpp_range_min_col(_handle);
        public ulong MaxRow => Native.xlpp_range_max_row(_handle);
        public ulong MaxColumn => Native.xlpp_range_max_col(_handle);
        public ulong RowCount => Native.xlpp_range_row_count(_handle);
        public ulong ColumnCount => Native.xlpp_range_col_count(_handle);
        public string Address => MarshalHelper.FromBuffer((b, n) => Native.xlpp_range_address(_handle, b, n));
        public Cell Cell(ulong relRow, ulong relCol) => new(Native.xlpp_range_cell(_handle, relRow, relCol));
        public void SetValue(double v) => Native.xlpp_range_set_value(_handle, v);
        public void SetString(string v) => Native.xlpp_range_set_string(_handle, v);
        public void Clear() => Native.xlpp_range_clear(_handle);
        public double[] Values
        {
            get
            {
                int count = 0;
                Native.xlpp_range_values(_handle, IntPtr.Zero, ref count);
                if (count <= 0) return Array.Empty<double>();
                var buf = Marshal.AllocHGlobal(count * sizeof(double));
                try
                {
                    Native.xlpp_range_values(_handle, buf, ref count);
                    var arr = new double[count];
                    Marshal.Copy(buf, arr, 0, count);
                    return arr;
                }
                finally { Marshal.FreeHGlobal(buf); }
            }
        }
    }

    public class Worksheet
    {
        internal IntPtr _handle;
        internal Worksheet(IntPtr h) => _handle = h;

        public string Name => MarshalHelper.PtrToString(Native.xlpp_sheet_name(_handle));
        public string VbaCodeName
        {
            get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheet_vba_code_name(_handle, b, n));
            set
            {
                if (Native.xlpp_sheet_set_vba_code_name(_handle, value) == 0)
                    throw new ArgumentException("Invalid or duplicate VBA worksheet code name", nameof(value));
            }
        }
        public void Rename(string name) => Native.xlpp_sheet_rename(_handle, name);

        public Cell this[string address] => new(Native.xlpp_sheet_cell(_handle, address));
        public Cell Cell(ulong row, ulong col) => new(Native.xlpp_sheet_cell_rc(_handle, row, col));
        public Cell Cell(string address) => this[address];
        public bool HasCell(string address) => Native.xlpp_sheet_has_cell(_handle, address) != 0;
        public bool HasCell(ulong row, ulong col) => Native.xlpp_sheet_has_cell_rc(_handle, row, col) != 0;

        public ulong MaxRow => Native.xlpp_sheet_max_row(_handle);
        public ulong MaxColumn => Native.xlpp_sheet_max_col(_handle);
        public ulong RowCount => Native.xlpp_sheet_row_count(_handle);
        public ulong ColumnCount => Native.xlpp_sheet_col_count(_handle);
        public bool IsEmpty => Native.xlpp_sheet_empty(_handle) != 0;
        public string Dimensions => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheet_dimensions(_handle, b, n));

        public void MergeCells(string range) => Native.xlpp_sheet_merge_cells(_handle, range);
        public void UnmergeCells(string range) => Native.xlpp_sheet_unmerge_cells(_handle, range);
        public bool IsMerged(string cell) => Native.xlpp_sheet_is_merged(_handle, cell) != 0;
        public int MergedCount => Native.xlpp_sheet_merged_count(_handle);
        public string MergedAt(int index) => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheet_merged_at(_handle, index, b, n));
        public IReadOnlyList<string> MergedRanges { get { var r = new List<string>(); for (var i = 0; i < MergedCount; ++i) r.Add(MergedAt(i)); return r; } }
        public void FreezePanes(string cell) => Native.xlpp_sheet_freeze_panes(_handle, cell);
        public void ClearFreezePanes() => Native.xlpp_sheet_clear_freeze_panes(_handle);
        public string FrozenPane => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheet_frozen_pane(_handle, b, n));

        public RowDimension RowDimension(ulong row) => new(Native.xlpp_sheet_row_dimension(_handle, row));
        public ColumnDimension ColumnDimension(ulong col) => new(Native.xlpp_sheet_col_dimension(_handle, col));

        public CellRange Range(string address) => new(Native.xlpp_sheet_range(_handle, address));

        public void AppendRow(params string[] values) => Native.xlpp_sheet_append_row(_handle, values, values.Length);
        public void AppendRow(params double[] values) => Native.xlpp_sheet_append_doubles(_handle, values, values.Length);

        public void InsertRows(ulong index, ulong amount = 1) => Native.xlpp_sheet_insert_rows(_handle, index, amount);
        public void DeleteRows(ulong index, ulong amount = 1) => Native.xlpp_sheet_delete_rows(_handle, index, amount);
        public void InsertColumns(ulong index, ulong amount = 1) => Native.xlpp_sheet_insert_cols(_handle, index, amount);
        public void DeleteColumns(ulong index, ulong amount = 1) => Native.xlpp_sheet_delete_cols(_handle, index, amount);

        public string PrintArea { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheet_print_area(_handle, b, n)); set => Native.xlpp_sheet_set_print_area(_handle, value); }
        public string PrintTitlesRows { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheet_print_titles_rows(_handle, b, n)); set => Native.xlpp_sheet_set_print_titles_rows(_handle, value); }
        public string PrintTitlesCols { get => MarshalHelper.FromBuffer((b, n) => Native.xlpp_sheet_print_titles_cols(_handle, b, n)); set => Native.xlpp_sheet_set_print_titles_cols(_handle, value); }

        public AutoFilter AutoFilter => new(Native.xlpp_sheet_auto_filter(_handle));
        public ConditionalFormattingCollection ConditionalFormatting => new(Native.xlpp_sheet_conditional_formatting(_handle));
        public DataValidationCollection DataValidations => new(Native.xlpp_sheet_data_validations(_handle));
        public PageSetup PageSetup => new(Native.xlpp_sheet_page_setup(_handle));
        public PageMargins PageMargins => new(Native.xlpp_sheet_page_margins(_handle));
        public PrintOptions PrintOptions => new(Native.xlpp_sheet_print_options(_handle));
        public HeaderFooter HeaderFooter => new(Native.xlpp_sheet_header_footer(_handle));
        public WorksheetProtection Protection => new(Native.xlpp_sheet_protection(_handle));
        public SheetView SheetView => new(Native.xlpp_sheet_view(_handle));

        public Table AddTable(string name, string reference)
        {
            var h = Native.xlpp_sheet_add_table(_handle, name, reference, out var ok);
            return ok != 0 && h != IntPtr.Zero ? new Table(h) : throw new ArgumentException($"Failed to add table '{name}'");
        }
        public Table? GetTable(string name)
        {
            var h = Native.xlpp_sheet_table(_handle, name);
            return h == IntPtr.Zero ? null : new Table(h);
        }
        public int TableCount => Native.xlpp_sheet_table_count(_handle);
        public Table TableAt(int index) => new(Native.xlpp_sheet_table_at(_handle, index));
        public IReadOnlyList<Table> Tables { get { var r = new List<Table>(); for (var i = 0; i < TableCount; ++i) r.Add(TableAt(i)); return r; } }

        public Image AddImage(string path, string anchor)
        {
            var h = Native.xlpp_sheet_add_image(_handle, path, anchor, out var ok);
            return ok != 0 && h != IntPtr.Zero ? new Image(h) : throw new InvalidOperationException($"Failed to add image '{path}'");
        }
        public int ImageCount => Native.xlpp_sheet_image_count(_handle);
        public Image ImageAt(int index) => new(Native.xlpp_sheet_image_at(_handle, index));
        public IReadOnlyList<Image> Images { get { var r = new List<Image>(); for (var i = 0; i < ImageCount; ++i) r.Add(ImageAt(i)); return r; } }

        public Chart AddChart(ChartType type)
        {
            Native.xlpp_sheet_add_chart(_handle, (int)type);
            return ChartAt(ChartCount - 1);
        }
        public int ChartCount => Native.xlpp_sheet_chart_count(_handle);
        public Chart ChartAt(int index) => new(Native.xlpp_sheet_chart_at(_handle, index));
        public IReadOnlyList<Chart> Charts { get { var r = new List<Chart>(); for (var i = 0; i < ChartCount; ++i) r.Add(ChartAt(i)); return r; } }

        public PivotTable AddPivotTable(string name, string location)
        {
            Native.xlpp_sheet_add_pivot(_handle, name, location);
            return PivotAt(PivotCount - 1);
        }
        public int PivotCount => Native.xlpp_sheet_pivot_count(_handle);
        public PivotTable PivotAt(int index) => new(Native.xlpp_sheet_pivot_at(_handle, index));
        public IReadOnlyList<PivotTable> PivotTables { get { var r = new List<PivotTable>(); for (var i = 0; i < PivotCount; ++i) r.Add(PivotAt(i)); return r; } }
    }

    internal sealed class WorkbookSafeHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        private WorkbookSafeHandle() : base(true) { }
        internal WorkbookSafeHandle(IntPtr value) : base(true) => SetHandle(value);
        protected override bool ReleaseHandle()
        {
            Native.xlpp_workbook_destroy(handle);
            return true;
        }
    }

    public partial class Workbook : IDisposable
    {
        private readonly WorkbookSafeHandle _safeHandle;
        internal IntPtr _handle => _safeHandle.DangerousGetHandle();

        public static string Version => MarshalHelper.PtrToString(Native.xlpp_version());
        public static ulong CAbiVersion => Native.xlpp_c_abi_version();
        public static ulong Capabilities => Native.xlpp_capabilities();

        public Workbook()
        {
            var handle = Native.xlpp_workbook_create();
            if (handle == IntPtr.Zero)
                throw new InvalidOperationException("Failed to create XL++ Workbook");
            _safeHandle = new WorkbookSafeHandle(handle);
        }

        public void Dispose() => _safeHandle.Dispose();

        public Worksheet AddWorksheet(string name)
        {
            var h = Native.xlpp_workbook_add_sheet(_handle, name);
            if (h == IntPtr.Zero)
                throw new ArgumentException($"Failed to add worksheet '{name}'");
            return new Worksheet(h);
        }

        public Worksheet CopyWorksheet(Worksheet src, string newName)
        {
            var h = Native.xlpp_workbook_copy_sheet(_handle, src._handle, newName);
            if (h == IntPtr.Zero)
                throw new ArgumentException($"Failed to copy worksheet '{newName}'");
            return new Worksheet(h);
        }

        public Worksheet? GetWorksheet(string name)
        {
            var h = Native.xlpp_workbook_sheet_by_name(_handle, name);
            return h == IntPtr.Zero ? null : new Worksheet(h);
        }

        public Worksheet this[int index]
        {
            get
            {
                var h = Native.xlpp_workbook_get_sheet(_handle, index);
                return h == IntPtr.Zero
                    ? throw new ArgumentOutOfRangeException(nameof(index), "Worksheet index out of range")
                    : new Worksheet(h);
            }
        }

        public Worksheet this[string name]
        {
            get
            {
                var ws = GetWorksheet(name)
                    ?? throw new KeyNotFoundException($"Worksheet '{name}' not found");
                return ws;
            }
        }

        public int SheetCount => Native.xlpp_workbook_sheet_count(_handle);
        public bool RemoveWorksheet(string name) => Native.xlpp_workbook_remove_sheet(_handle, name) != 0;
        public int IndexOf(Worksheet ws) => Native.xlpp_workbook_sheet_index(_handle, ws._handle);

        public List<string> SheetNames
        {
            get
            {
                var result = new List<string>();
                var count = Native.xlpp_workbook_sheet_names_count(_handle);
                for (int i = 0; i < count; ++i)
                    result.Add(MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_sheet_name(_handle, i, b, n)));
                return result;
            }
        }

        public bool Load(string path) => Native.xlpp_workbook_load(_handle, path) != 0;
        public ExternalDataInspection InspectExternalData()
        {
            if (Native.xlpp_workbook_inspect_external_data(_handle, out var r) == 0)
                throw new InvalidOperationException("External-data inspection failed");
            return new ExternalDataInspection(r.ExternalWorkbooks, r.Connections, r.QueryTables, r.PowerQueryParts,
                r.WebQueryParts, r.UnknownConnectionParts);
        }

        public DataModelInspection InspectDataModel()
        {
            if (Native.xlpp_workbook_inspect_data_model(_handle, out var r) == 0)
                throw new InvalidOperationException("Data-model inspection failed");
            return new DataModelInspection(r.Present != 0, r.HasOlapPivotCaches != 0, r.ModelParts,
                r.ModelRelationships, r.OlapPivotCacheParts, r.Warnings);
        }

        public bool Save(string path) => Native.xlpp_workbook_save(_handle, path) != 0;
        public bool Save(string path, bool durableWrite) => Native.xlpp_workbook_save_durable(_handle, path, durableWrite ? 1 : 0) != 0;
        public bool Load(string path, string password, bool verifyIntegrity = true) => Native.xlpp_workbook_load_password(_handle, path, password, verifyIntegrity ? 1 : 0) != 0;
        public bool SaveEncrypted(string path, string password, ulong spinCount = 100000, bool calculateFormulas = false) =>
            Native.xlpp_workbook_save_encrypted(_handle, path, password, spinCount, calculateFormulas ? 1 : 0) != 0;
        public bool SaveEncrypted(string path, string password, OfficeEncryptionMode mode, ulong keyBits = 256, ulong spinCount = 100000, bool calculateFormulas = false) =>
            Native.xlpp_workbook_save_encrypted_ex(_handle, path, password, (int)mode, keyBits, spinCount, calculateFormulas ? 1 : 0) != 0;

        public CalculationReport CalculateFormulas()
        {
            if (Native.xlpp_workbook_calculate(_handle, out var r) == 0)
                throw new InvalidOperationException("Formula calculation failed");
            return new CalculationReport(r.FormulaCellsVisited, r.FormulaCellsEvaluated, r.CachedValuesUpdated,
                r.DependencyEvaluations, r.DefinedNamesResolved, r.CircularReferences, r.UnsupportedFormulas,
                r.EvaluationErrors, r.DynamicArraysSpilled, r.SpillCellsUpdated, r.SpillConflicts,
                r.StructuredReferencesResolved, r.IterativeIterations, r.IterativeConvergenceFailures,
                r.ExternalReferencesResolved, r.UnresolvedExternalReferences, r.Success != 0);
        }

        public CalculationReport CalculateFormulasIterative(ulong maxIterations = 100, double maxChange = 0.001)
        {
            if (Native.xlpp_workbook_calculate_ex(_handle, 1, maxIterations, maxChange, out var r) == 0)
                throw new InvalidOperationException("Iterative formula calculation failed");
            return new CalculationReport(r.FormulaCellsVisited, r.FormulaCellsEvaluated, r.CachedValuesUpdated,
                r.DependencyEvaluations, r.DefinedNamesResolved, r.CircularReferences, r.UnsupportedFormulas,
                r.EvaluationErrors, r.DynamicArraysSpilled, r.SpillCellsUpdated, r.SpillConflicts,
                r.StructuredReferencesResolved, r.IterativeIterations, r.IterativeConvergenceFailures,
                r.ExternalReferencesResolved, r.UnresolvedExternalReferences, r.Success != 0);
        }

        public StructuralEditReport ApplyStructuralEdit(string sheetName, StructuralEditKind kind, ulong index, ulong amount = 1, bool failOnInvalidReference = false)
        {
            var ok = Native.xlpp_workbook_structural_edit(_handle, sheetName, (int)kind, index, amount, failOnInvalidReference ? 1 : 0, out var r);
            if (ok == 0 && r.Success == 0)
                throw new InvalidOperationException("Structural edit failed");
            return new StructuralEditReport(r.WorksheetsVisited, r.CellsMoved, r.CellsRemoved, r.FormulasUpdated,
                r.FormulaMetadataUpdated, r.WorksheetReferencesUpdated, r.DefinedNamesUpdated, r.ChartReferencesUpdated,
                r.PivotReferencesUpdated, r.DrawingAnchorsUpdated, r.HyperlinksUpdated, r.ReferencesInvalidated,
                r.FormulasCalculated, r.ChartCachesUpdated, r.Success != 0);
        }
        public StructuralEditReport InsertRows(string sheetName, ulong index, ulong amount = 1) => ApplyStructuralEdit(sheetName, StructuralEditKind.InsertRows, index, amount);
        public StructuralEditReport DeleteRows(string sheetName, ulong index, ulong amount = 1, bool failOnInvalidReference = false) => ApplyStructuralEdit(sheetName, StructuralEditKind.DeleteRows, index, amount, failOnInvalidReference);
        public StructuralEditReport InsertColumns(string sheetName, ulong index, ulong amount = 1) => ApplyStructuralEdit(sheetName, StructuralEditKind.InsertColumns, index, amount);
        public StructuralEditReport DeleteColumns(string sheetName, ulong index, ulong amount = 1, bool failOnInvalidReference = false) => ApplyStructuralEdit(sheetName, StructuralEditKind.DeleteColumns, index, amount, failOnInvalidReference);
        public bool Date1904 { get => Native.xlpp_workbook_date1904(_handle) != 0; set => Native.xlpp_workbook_set_date1904(_handle, value ? 1 : 0); }
        public bool StrictNamespaces => Native.xlpp_workbook_strict_namespaces(_handle) != 0;
        public void Clear() => Native.xlpp_workbook_clear(_handle);

        public Properties Properties
        {
            get
            {
                var h = Native.xlpp_workbook_properties(_handle);
                return new Properties(h);
            }
        }

        public WorkbookProtection Protection => new(Native.xlpp_workbook_protection(_handle));
        public CalcProperties CalcProperties => new(Native.xlpp_workbook_calc_properties(_handle));
        public CustomProperties CustomProperties => new(Native.xlpp_workbook_custom_properties(_handle));

        public NamedStyle AddNamedStyle(string name)
        {
            var h = Native.xlpp_workbook_add_named_style(_handle, name, out var ok);
            return ok != 0 && h != IntPtr.Zero ? new NamedStyle(h) : throw new ArgumentException($"Failed to add named style '{name}'");
        }

        public NamedStyle? GetNamedStyle(string name)
        {
            var h = Native.xlpp_workbook_named_style(_handle, name);
            return h == IntPtr.Zero ? null : new NamedStyle(h);
        }

        public int NamedStyleCount => Native.xlpp_workbook_named_styles_count(_handle);
        public NamedStyle NamedStyleAt(int index) => new(Native.xlpp_workbook_named_style_at(_handle, index));
        public void ApplyNamedStyle(Cell cell, string name) => Native.xlpp_workbook_apply_named_style(_handle, cell._handle, name);

        public DefinedName AddDefinedName(string name, string value)
        {
            var h = Native.xlpp_workbook_add_defined_name(_handle, name, value, out var ok);
            return ok != 0 && h != IntPtr.Zero ? new DefinedName(h) : throw new ArgumentException($"Failed to add defined name '{name}'");
        }

        public DefinedName AddDefinedName(string name, string value, ulong localSheetId)
        {
            var h = Native.xlpp_workbook_add_defined_name_scoped(_handle, name, value, localSheetId, out var ok);
            return ok != 0 && h != IntPtr.Zero ? new DefinedName(h) : throw new ArgumentException($"Failed to add scoped defined name '{name}'");
        }

        public DefinedName? GetDefinedName(string name)
        {
            var h = Native.xlpp_workbook_defined_name(_handle, name);
            return h == IntPtr.Zero ? null : new DefinedName(h);
        }

        public DefinedName? GetDefinedName(string name, ulong? localSheetId)
        {
            var h = Native.xlpp_workbook_defined_name_scoped(_handle, name, localSheetId.HasValue ? 1 : 0, localSheetId.GetValueOrDefault());
            return h == IntPtr.Zero ? null : new DefinedName(h);
        }

        public int DefinedNameCount => Native.xlpp_workbook_defined_names_count(_handle);
        public DefinedName DefinedNameAt(int index) => new(Native.xlpp_workbook_defined_name_at(_handle, index));
        public IReadOnlyList<Worksheet> Worksheets { get { var r = new List<Worksheet>(); for (var i = 0; i < SheetCount; ++i) r.Add(this[i]); return r; } }
        public IReadOnlyList<NamedStyle> NamedStyles { get { var r = new List<NamedStyle>(); for (var i = 0; i < NamedStyleCount; ++i) r.Add(NamedStyleAt(i)); return r; } }
        public IReadOnlyList<DefinedName> DefinedNames { get { var r = new List<DefinedName>(); for (var i = 0; i < DefinedNameCount; ++i) r.Add(DefinedNameAt(i)); return r; } }
        public int Index(Worksheet worksheet) => IndexOf(worksheet);
        public IReadOnlyList<PreservedRelationship> PreservedRelationships
        {
            get
            {
                var r = new List<PreservedRelationship>();
                var count = Native.xlpp_workbook_preserved_relationships_count(_handle);
                for (var i = 0; i < count; ++i)
                    r.Add(new PreservedRelationship(
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_relationship_source_part(_handle, i, b, n)),
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_relationship_id(_handle, i, b, n)),
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_relationship_type(_handle, i, b, n)),
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_relationship_target(_handle, i, b, n)),
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_relationship_target_mode(_handle, i, b, n))));
                return r;
            }
        }
        public IReadOnlyList<PreservedPart> PreservedParts
        {
            get
            {
                var r = new List<PreservedPart>();
                var count = Native.xlpp_workbook_preserved_parts_count(_handle);
                for (var i = 0; i < count; ++i)
                {
                    var size = Native.xlpp_workbook_preserved_part_data_size(_handle, i);
                    var data = new byte[checked((int)size)];
                    if (size != 0) { var handle = GCHandle.Alloc(data, GCHandleType.Pinned); try { Native.xlpp_workbook_preserved_part_data(_handle, i, handle.AddrOfPinnedObject(), size); } finally { handle.Free(); } }
                    r.Add(new PreservedPart(
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_part_name(_handle, i, b, n)), data,
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_part_override_type(_handle, i, b, n)),
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_part_extension(_handle, i, b, n)),
                        MarshalHelper.FromBuffer((b, n) => Native.xlpp_workbook_preserved_part_default_type(_handle, i, b, n)),
                        Native.xlpp_workbook_preserved_part_compress(_handle, i) != 0));
                }
                return r;
            }
        }
    }

    public sealed record PreservedRelationship(string SourcePart, string Id, string Type, string Target, string TargetMode);
    public sealed record PreservedPart(string Name, byte[] Data, string OverrideType, string Extension, string DefaultType, bool Compress);

    public sealed class StreamingWorkbookWriter : IDisposable
    {
        internal IntPtr Handle { get; private set; }

        public StreamingWorkbookWriter(string path)
        {
            Handle = Native.xlpp_stream_create(path);
            if (Handle == IntPtr.Zero)
                throw new InvalidOperationException("Failed to create XL++ streaming writer");
        }

        public StreamingWorksheetWriter AddWorksheet(string name)
        {
            var index = Native.xlpp_stream_add_sheet(Handle, name);
            if (index == ulong.MaxValue)
                throw new ArgumentException($"Failed to add worksheet '{name}'");
            return new StreamingWorksheetWriter(this, index);
        }

        public ulong SheetCount => Native.xlpp_stream_sheet_count(Handle);
        public ulong RowCount(StreamingWorksheetWriter sheet) => Native.xlpp_stream_row_count(Handle, sheet.Index);
        public bool Date1904 { set => Native.xlpp_stream_set_date1904(Handle, value ? 1 : 0); }
        public void SetCompressionLevel(CompressionLevel level) => Native.xlpp_stream_set_compression_level(Handle, (int)level);
        public void SetParallelWorkers(ulong workers) => Native.xlpp_stream_set_parallel_workers(Handle, workers);

        public void Close()
        {
            if (Handle == IntPtr.Zero) return;
            Native.xlpp_stream_close(Handle);
        }

        public void Dispose()
        {
            Close();
            if (Handle != IntPtr.Zero)
            {
                Native.xlpp_stream_destroy(Handle);
                Handle = IntPtr.Zero;
            }
        }
    }

    public sealed class StreamingWorksheetWriter
    {
        internal ulong Index { get; }
        private readonly StreamingWorkbookWriter _workbook;

        internal StreamingWorksheetWriter(StreamingWorkbookWriter workbook, ulong index)
        {
            _workbook = workbook;
            Index = index;
        }

        public void AppendRow(params string[] values) => Native.xlpp_stream_append_row(_workbook.Handle, Index, values, values.Length);
        public void AppendRow(params double[] values) => Native.xlpp_stream_append_doubles(_workbook.Handle, Index, values, values.Length);
    }

    public sealed class StreamingRow
    {
        public ulong RowNumber { get; }
        public List<StreamingCell> Cells { get; } = new();
        public StreamingRow(ulong rowNumber) => RowNumber = rowNumber;
    }

    public sealed class StreamingCell
    {
        public string Address { get; }
        public object? Value { get; }
        public string Formula { get; }
        public int StyleIndex { get; }
        public StreamingCell(string address, object? value, string formula, int styleIndex)
        {
            Address = address;
            Value = value;
            Formula = formula;
            StyleIndex = styleIndex;
        }
    }

    public sealed class StreamingWorkbookReader : IDisposable
    {
        internal IntPtr _handle;
        private Native.StreamRowCallback? _callback; // keep delegate alive during P/Invoke

        public StreamingWorkbookReader(string path)
        {
            _handle = Native.xlpp_stream_reader_open(path);
            if (_handle == IntPtr.Zero)
                throw new InvalidOperationException("Failed to open XL++ streaming reader");
        }

        public void Dispose()
        {
            if (_handle != IntPtr.Zero)
            {
                Native.xlpp_stream_reader_destroy(_handle);
                _handle = IntPtr.Zero;
            }
        }

        public int SheetCount => Native.xlpp_stream_reader_sheet_count(_handle);

        public List<string> SheetNames
        {
            get
            {
                var result = new List<string>();
                for (int i = 0; i < SheetCount; ++i)
                    result.Add(MarshalHelper.FromBuffer((b, n) => Native.xlpp_stream_reader_sheet_name(_handle, i, b, n)));
                return result;
            }
        }

        public List<StreamingRow> ReadSheet(int sheetIndex)
        {
            var rows = new List<StreamingRow>();
            _callback = (user, rowNumber, cellCount, addrPtr, numPtr, typePtr, strPtr, formulaPtr, stylePtr) =>
            {
                var row = new StreamingRow(rowNumber);
                var addr = new IntPtr[cellCount];
                var strings = new IntPtr[cellCount];
                var formulas = new IntPtr[cellCount];
                var types = new int[cellCount];
                var nums = new double[cellCount];
                var styles = new int[cellCount];
                Marshal.Copy(addrPtr, addr, 0, cellCount);
                Marshal.Copy(strPtr, strings, 0, cellCount);
                Marshal.Copy(formulaPtr, formulas, 0, cellCount);
                Marshal.Copy(typePtr, types, 0, cellCount);
                Marshal.Copy(numPtr, nums, 0, cellCount);
                Marshal.Copy(stylePtr, styles, 0, cellCount);
                for (int i = 0; i < cellCount; ++i)
                {
                    var address = MarshalHelper.PtrToString(addr[i]);
                    var formula = MarshalHelper.PtrToString(formulas[i]);
                    object? value = types[i] switch
                    {
                        (int)CellValueType.Bool => nums[i] != 0,
                        (int)CellValueType.Number => nums[i],
                        (int)CellValueType.String => MarshalHelper.PtrToString(strings[i]),
                        (int)CellValueType.Error => "#ERR",
                        _ => null
                    };
                    row.Cells.Add(new StreamingCell(address, value, formula, styles[i]));
                }
                rows.Add(row);
                return 1;
            };
            Native.xlpp_stream_reader_read_sheet(_handle, sheetIndex, _callback, IntPtr.Zero);
            return rows;
        }
    }
}
