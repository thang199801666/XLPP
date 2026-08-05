// XL++ C# Bindings — P/Invoke wrapper with OOP API
// Requires xlpp_capi.dll in the same directory or PATH.

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace XLPP
{
    // === Native interop ===
    internal static class Native
    {
        private const string Dll = "xlpp_capi";

        static Native()
        {
            // Resolve the native library from the NuGet runtimes/<rid>/native
            // folder (automatic on .NET Core 3.0+) and fall back to the
            // application base directory for manual deployment.
            NativeLibrary.SetDllImportResolver(
                typeof(Native).Assembly,
                (name, assembly, searchPath) =>
                {
                    if (name != Dll && name != Dll + ".dll" && name != "lib" + Dll + ".so" &&
                        name != "lib" + Dll + ".dylib")
                        return IntPtr.Zero;

                    // 1) Standard .NET resolution (deps.json / runtimes/<rid>/native).
                    if (NativeLibrary.TryLoad(Dll, assembly, searchPath, out var handle))
                        return handle;

                    // 2) NuGet runtimes/<rid>/native next to the app (project-ref dev).
                    string nativeName = Dll + ".dll";
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) nativeName = "lib" + Dll + ".so";
                    else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) nativeName = "lib" + Dll + ".dylib";
                    var rid = RuntimeInformation.RuntimeIdentifier;
                    if (!string.IsNullOrEmpty(rid))
                    {
                        var candidate = Path.Combine(AppContext.BaseDirectory, "runtimes", rid, "native", nativeName);
                        if (File.Exists(candidate)) return NativeLibrary.Load(candidate);
                    }

                    // 3) Native library copied next to the assembly.
                    var local = Path.Combine(AppContext.BaseDirectory, nativeName);
                    if (File.Exists(local)) return NativeLibrary.Load(local);

                    return IntPtr.Zero;
                });
        }

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_version();

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_last_error();

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_clear_error();

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_create();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_workbook_destroy(IntPtr wb);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_add_sheet(IntPtr wb, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_sheet_count(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_workbook_date1904(IntPtr wb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_workbook_set_date1904(IntPtr wb, int enabled);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_get_sheet(IntPtr wb, int index);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_workbook_sheet_by_name(IntPtr wb, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_remove_sheet(IntPtr wb, string name);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_load(IntPtr wb, string path);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_workbook_save(IntPtr wb, string path);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_workbook_properties(IntPtr wb);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_title(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_creator(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_title(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_creator(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_subject(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_subject(IntPtr p);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_properties_set_description(IntPtr p, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_properties_get_description(IntPtr p);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_cell(IntPtr ws, string address);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_cell_rc(IntPtr ws, ulong row, ulong col);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_sheet_has_cell(IntPtr ws, string address);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_merge_cells(IntPtr ws, string range);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_unmerge_cells(IntPtr ws, string range);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_freeze_panes(IntPtr ws, string cell);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_clear_freeze_panes(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern int xlpp_sheet_is_merged(IntPtr ws, string address);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_append_mixed(IntPtr ws, string[] values, int[] types, int count);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_sheet_max_row(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_sheet_max_col(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_add_table(IntPtr ws, string name, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_table_name(IntPtr table);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_table_reference(IntPtr table);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_table_column_count(IntPtr table);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_table_add_column(IntPtr table, string name);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_table_display_name(IntPtr table);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_table_set_display_name(IntPtr table, string value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_table_style_name(IntPtr table);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_table_set_style_name(IntPtr table, string value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_table_show_row_stripes(IntPtr table);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_table_set_show_row_stripes(IntPtr table, int value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_add_list_validation(IntPtr ws, string reference, string formula);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_validation_set_allow_blank(IntPtr validation, int value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_validation_set_prompt(IntPtr validation, string title, string text);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_validation_set_error(IntPtr validation, string title, string text);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern IntPtr xlpp_chart_create(int type);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_chart_destroy(IntPtr chart);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_chart_set_title(IntPtr chart, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_chart_set_x_axis_title(IntPtr chart, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_chart_set_y_axis_title(IntPtr chart, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_chart_set_style(IntPtr chart, string style);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_chart_set_grouping(IntPtr chart, int grouping);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern void xlpp_chart_set_size(IntPtr chart, int width, int height);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_chart_set_legend(IntPtr chart, int show, string position);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern IntPtr xlpp_chart_add_series(IntPtr chart, string title);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_chart_series_set_values_reference(IntPtr series, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)] public static extern void xlpp_chart_series_set_categories_reference(IntPtr series, string reference);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] public static extern int xlpp_sheet_add_chart(IntPtr ws, IntPtr chart);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_add_formula_rule(IntPtr ws, string reference, string formula);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_add_data_bar_rule(IntPtr ws, string reference, string color);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_conditional_rule_set_priority(IntPtr rule, ulong priority);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_conditional_rule_set_stop_if_true(IntPtr rule, int value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_conditional_rule_set_font_color(IntPtr rule, string argb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_conditional_rule_set_fill_color(IntPtr rule, string argb);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_sheet_dimensions(IntPtr ws, StringBuilder output, int outputSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_frozen_pane(IntPtr ws, StringBuilder output, int outputSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_merged_range_count(IntPtr ws);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_sheet_merged_range(IntPtr ws, int index, StringBuilder output, int outputSize);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_sheet_set_print_area(IntPtr ws, string value);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_sheet_print_area(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr xlpp_sheet_name(IntPtr ws);
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
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cell_set_formula(IntPtr c, string f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_get_formula(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_has_formula(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_clear_formula(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_has_hyperlink(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_clear_hyperlink(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_cell_has_comment(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_cell_clear_comment(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_get_number_format(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_cell_set_number_format(IntPtr c, string value);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_font(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_fill(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_border(IntPtr c);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_cell_alignment(IntPtr c);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_font_set_name(IntPtr f, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_set_size(IntPtr f, double v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_set_bold(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_font_set_italic(IntPtr f, int v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_font_set_color(IntPtr f, string argb);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_font_get_name(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern double xlpp_font_get_size(IntPtr f);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int xlpp_font_get_bold(IntPtr f);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_left(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_right(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_top(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_border_bottom(IntPtr b);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_borderside_set_style(IntPtr s, string v);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern void xlpp_alignment_set_horizontal(IntPtr a, string v);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void xlpp_alignment_set_wrap_text(IntPtr a, int v);
    }

    // === Value types ===
    public enum CellValueType { Empty = 0, Bool = 1, Number = 2, String = 3, Error = 4, Date = 5 }
    public enum ChartType { Bar = 0, Line = 1, Pie = 2, Scatter = 3, Doughnut = 4, Radar = 5, Area = 6, Bubble = 7 }

    // === OOP wrappers ===
    internal static class MarshalHelper
    {
        public static string PtrToString(IntPtr ptr)
        {
            if (ptr == IntPtr.Zero) return "";
            return Marshal.PtrToStringAnsi(ptr) ?? "";
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
        public void SetColor(string argb) => Native.xlpp_font_set_color(_handle, argb);
        public string Name => MarshalHelper.PtrToString(Native.xlpp_font_get_name(_handle));
        public double Size => Native.xlpp_font_get_size(_handle);
        public bool Bold => Native.xlpp_font_get_bold(_handle) != 0;
    }

    public class Fill
    {
        internal IntPtr _handle;
        internal Fill(IntPtr h) => _handle = h;
    }

    public class BorderSide
    {
        internal IntPtr _handle;
        internal BorderSide(IntPtr h) => _handle = h;
        public void SetStyle(string v) => Native.xlpp_borderside_set_style(_handle, v);
    }

    public class Border
    {
        internal IntPtr _handle;
        internal Border(IntPtr h) => _handle = h;
        public BorderSide Left => new(Native.xlpp_border_left(_handle));
        public BorderSide Right => new(Native.xlpp_border_right(_handle));
        public BorderSide Top => new(Native.xlpp_border_top(_handle));
        public BorderSide Bottom => new(Native.xlpp_border_bottom(_handle));
    }

    public class Alignment
    {
        internal IntPtr _handle;
        internal Alignment(IntPtr h) => _handle = h;
        public void SetHorizontal(string v) => Native.xlpp_alignment_set_horizontal(_handle, v);
        public void SetWrapText(bool v) => Native.xlpp_alignment_set_wrap_text(_handle, v ? 1 : 0);
    }

    public class Properties
    {
        internal IntPtr _handle;
        internal Properties(IntPtr h) => _handle = h;
        public string Title { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_title(_handle)); set => Native.xlpp_properties_set_title(_handle, value ?? ""); }
        public string Creator { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_creator(_handle)); set => Native.xlpp_properties_set_creator(_handle, value ?? ""); }
        public string Subject { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_subject(_handle)); set => Native.xlpp_properties_set_subject(_handle, value ?? ""); }
        public string Description { get => MarshalHelper.PtrToString(Native.xlpp_properties_get_description(_handle)); set => Native.xlpp_properties_set_description(_handle, value ?? ""); }
    }

    public class Cell
    {
        internal IntPtr _handle;
        internal Cell(IntPtr h) => _handle = h;

        public CellValueType ValueType => (CellValueType)Native.xlpp_cell_value_type(_handle);
        public bool IsEmpty => Native.xlpp_cell_is_empty(_handle) != 0;
        public string Address => MarshalHelper.PtrToString(Native.xlpp_cell_address(_handle));
        public ulong Row => Native.xlpp_cell_row(_handle);
        public ulong Column => Native.xlpp_cell_column(_handle);

        public object? Value
        {
            get => ValueType switch
            {
                CellValueType.Bool => Native.xlpp_cell_get_bool(_handle) != 0,
                CellValueType.Number => Native.xlpp_cell_get_number(_handle),
                CellValueType.String => MarshalHelper.PtrToString(Native.xlpp_cell_get_string(_handle)),
                CellValueType.Empty => null,
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
                }
            }
        }

        public string? Formula
        {
            get => MarshalHelper.PtrToString(Native.xlpp_cell_get_formula(_handle));
            set => Native.xlpp_cell_set_formula(_handle, value ?? "");
        }

        public bool HasFormula => Native.xlpp_cell_has_formula(_handle) != 0;
        public void ClearFormula() => Native.xlpp_cell_clear_formula(_handle);
        public bool HasHyperlink => Native.xlpp_cell_has_hyperlink(_handle) != 0;
        public void ClearHyperlink() => Native.xlpp_cell_clear_hyperlink(_handle);
        public bool HasComment => Native.xlpp_cell_has_comment(_handle) != 0;
        public void ClearComment() => Native.xlpp_cell_clear_comment(_handle);
        public string NumberFormat
        {
            get => MarshalHelper.PtrToString(Native.xlpp_cell_get_number_format(_handle));
            set => Native.xlpp_cell_set_number_format(_handle, value ?? "General");
        }

        public Font Font => new(Native.xlpp_cell_font(_handle));
        public Fill Fill => new(Native.xlpp_cell_fill(_handle));
        public Border Border => new(Native.xlpp_cell_border(_handle));
        public Alignment Alignment => new(Native.xlpp_cell_alignment(_handle));
    }

    public sealed class Chart : IDisposable
    {
        internal IntPtr _handle;
        public Chart(ChartType type = ChartType.Bar) { _handle = Native.xlpp_chart_create((int)type); if (_handle == IntPtr.Zero) throw new ArgumentException($"Invalid chart type: {Workbook.LastError}"); }
        public void Dispose() { if (_handle != IntPtr.Zero) { Native.xlpp_chart_destroy(_handle); _handle = IntPtr.Zero; } }
        public string Title { set => Native.xlpp_chart_set_title(_handle, value ?? ""); }
        public string XAxisTitle { set => Native.xlpp_chart_set_x_axis_title(_handle, value ?? ""); }
        public string YAxisTitle { set => Native.xlpp_chart_set_y_axis_title(_handle, value ?? ""); }
        public string Style { set => Native.xlpp_chart_set_style(_handle, value ?? ""); }
        public int Grouping { set => Native.xlpp_chart_set_grouping(_handle, value); }
        public void SetSize(int width, int height) => Native.xlpp_chart_set_size(_handle, width, height);
        public void SetLegend(bool show, string position = "r") => Native.xlpp_chart_set_legend(_handle, show ? 1 : 0, position);
        public ChartSeries AddSeries(string title) => new(Native.xlpp_chart_add_series(_handle, title));
        internal IntPtr Detach() { var handle = _handle; _handle = IntPtr.Zero; return handle; }
    }

    public sealed class ChartSeries
    {
        internal IntPtr _handle;
        internal ChartSeries(IntPtr handle) => _handle = handle;
        public string ValuesReference { set => Native.xlpp_chart_series_set_values_reference(_handle, value ?? ""); }
        public string CategoriesReference { set => Native.xlpp_chart_series_set_categories_reference(_handle, value ?? ""); }
    }

    public class Worksheet
    {
        internal IntPtr _handle;
        internal Worksheet(IntPtr h) => _handle = h;

        public string Name => MarshalHelper.PtrToString(Native.xlpp_sheet_name(_handle));
        public void Rename(string name) => Native.xlpp_sheet_rename(_handle, name);

        public Cell this[string address] => new(Native.xlpp_sheet_cell(_handle, address));
        public Cell Cell(ulong row, ulong col) => new(Native.xlpp_sheet_cell_rc(_handle, row, col));
        public Cell Cell(string address) => this[address];
        public bool HasCell(string address) => Native.xlpp_sheet_has_cell(_handle, address) != 0;

        public ulong MaxRow => Native.xlpp_sheet_max_row(_handle);
        public ulong MaxColumn => Native.xlpp_sheet_max_col(_handle);

        private static string Buffer(int capacity, Func<StringBuilder, int> read)
        {
            var buffer = new StringBuilder(capacity);
            var length = read(buffer);
            return length < 0 ? string.Empty : buffer.ToString();
        }

        public string Dimensions => Buffer(256, b => { Native.xlpp_sheet_dimensions(_handle, b, b.Capacity); return b.Length; });
        public string? FrozenPane => Buffer(128, b => Native.xlpp_sheet_frozen_pane(_handle, b, b.Capacity)) is { Length: > 0 } value ? value : null;
        public string? PrintArea
        {
            get => MarshalHelper.PtrToString(Native.xlpp_sheet_print_area(_handle));
            set => Native.xlpp_sheet_set_print_area(_handle, value ?? "");
        }

        public IReadOnlyList<string> MergedRanges
        {
            get
            {
                var ranges = new List<string>();
                for (var i = 0; i < Native.xlpp_sheet_merged_range_count(_handle); i++)
                {
                    var value = Buffer(128, b => Native.xlpp_sheet_merged_range(_handle, i, b, b.Capacity));
                    if (value.Length > 0) ranges.Add(value);
                }
                return ranges;
            }
        }

        public void MergeCells(string range) => Native.xlpp_sheet_merge_cells(_handle, range);
        public void UnmergeCells(string range) => Native.xlpp_sheet_unmerge_cells(_handle, range);
        public void FreezePanes(string cell) => Native.xlpp_sheet_freeze_panes(_handle, cell);
        public void ClearFreezePanes() => Native.xlpp_sheet_clear_freeze_panes(_handle);
        public bool IsMerged(string address) => Native.xlpp_sheet_is_merged(_handle, address) != 0;
        public void AppendMixed(string[] values, int[] types)
        {
            if (values.Length != types.Length) throw new ArgumentException("Values and types must have the same length");
            Native.xlpp_sheet_append_mixed(_handle, values, types, values.Length);
        }

        public void AppendRow(params string[] values) =>
            Native.xlpp_sheet_append_row(_handle, values, values.Length);
        public void AppendRow(params double[] values) =>
            Native.xlpp_sheet_append_doubles(_handle, values, values.Length);

        public void InsertRows(ulong index, ulong amount = 1) => Native.xlpp_sheet_insert_rows(_handle, index, amount);
        public void DeleteRows(ulong index, ulong amount = 1) => Native.xlpp_sheet_delete_rows(_handle, index, amount);
        public void InsertColumns(ulong index, ulong amount = 1) => Native.xlpp_sheet_insert_cols(_handle, index, amount);
        public void DeleteColumns(ulong index, ulong amount = 1) => Native.xlpp_sheet_delete_cols(_handle, index, amount);
        public Table AddTable(string name, string reference)
        {
            var handle = Native.xlpp_sheet_add_table(_handle, name, reference);
            if (handle == IntPtr.Zero) throw new InvalidOperationException($"Failed to add table: {Workbook.LastError}");
            return new Table(handle);
        }
        public void AddChart(Chart chart)
        {
            if (chart == null || Native.xlpp_sheet_add_chart(_handle, chart.Detach()) == 0)
                throw new InvalidOperationException($"Failed to add chart: {Workbook.LastError}");
        }
        public DataValidation AddListValidation(string reference, string formula)
        {
            var handle = Native.xlpp_sheet_add_list_validation(_handle, reference, formula);
            if (handle == IntPtr.Zero) throw new InvalidOperationException($"Failed to add validation: {Workbook.LastError}");
            return new DataValidation(handle);
        }
        public ConditionalRule AddFormulaRule(string reference, string formula)
        {
            var handle = Native.xlpp_sheet_add_formula_rule(_handle, reference, formula);
            if (handle == IntPtr.Zero) throw new InvalidOperationException($"Failed to add conditional rule: {Workbook.LastError}");
            return new ConditionalRule(handle);
        }
        public ConditionalRule AddDataBarRule(string reference, string color = "FF638EC6")
        {
            var handle = Native.xlpp_sheet_add_data_bar_rule(_handle, reference, color);
            if (handle == IntPtr.Zero) throw new InvalidOperationException($"Failed to add data bar rule: {Workbook.LastError}");
            return new ConditionalRule(handle);
        }
    }

    public sealed class Table
    {
        internal IntPtr _handle;
        internal Table(IntPtr handle) => _handle = handle;
        public string Name => MarshalHelper.PtrToString(Native.xlpp_table_name(_handle));
        public string DisplayName { get => MarshalHelper.PtrToString(Native.xlpp_table_display_name(_handle)); set => Native.xlpp_table_set_display_name(_handle, value ?? ""); }
        public string Reference => MarshalHelper.PtrToString(Native.xlpp_table_reference(_handle));
        public int ColumnCount => Native.xlpp_table_column_count(_handle);
        public void AddColumn(string name) => Native.xlpp_table_add_column(_handle, name);
        public string StyleName { get => MarshalHelper.PtrToString(Native.xlpp_table_style_name(_handle)); set => Native.xlpp_table_set_style_name(_handle, value ?? ""); }
        public bool ShowRowStripes { get => Native.xlpp_table_show_row_stripes(_handle) != 0; set => Native.xlpp_table_set_show_row_stripes(_handle, value ? 1 : 0); }
    }

    public sealed class DataValidation
    {
        internal IntPtr _handle;
        internal DataValidation(IntPtr handle) => _handle = handle;
        public bool AllowBlank { set => Native.xlpp_validation_set_allow_blank(_handle, value ? 1 : 0); }
        public void SetPrompt(string title, string text) => Native.xlpp_validation_set_prompt(_handle, title, text);
        public void SetError(string title, string text) => Native.xlpp_validation_set_error(_handle, title, text);
    }

    public sealed class ConditionalRule
    {
        internal IntPtr _handle;
        internal ConditionalRule(IntPtr handle) => _handle = handle;
        public ulong Priority { set => Native.xlpp_conditional_rule_set_priority(_handle, value); }
        public bool StopIfTrue { set => Native.xlpp_conditional_rule_set_stop_if_true(_handle, value ? 1 : 0); }
        public void SetFontColor(string argb) => Native.xlpp_conditional_rule_set_font_color(_handle, argb);
        public void SetFillColor(string argb) => Native.xlpp_conditional_rule_set_fill_color(_handle, argb);
    }

    public class Workbook : IDisposable
    {
        internal IntPtr _handle;

    public static string Version => MarshalHelper.PtrToString(Native.xlpp_version());
        public static string LastError => MarshalHelper.PtrToString(Native.xlpp_last_error());

        public bool Date1904
        {
            get => Native.xlpp_workbook_date1904(_handle) != 0;
            set => Native.xlpp_workbook_set_date1904(_handle, value ? 1 : 0);
        }

        public Workbook()
        {
            _handle = Native.xlpp_workbook_create();
            if (_handle == IntPtr.Zero)
                throw new InvalidOperationException("Failed to create XL++ Workbook");
        }

        public void Dispose()
        {
            if (_handle != IntPtr.Zero)
            {
                Native.xlpp_workbook_destroy(_handle);
                _handle = IntPtr.Zero;
            }
        }

        public Worksheet AddWorksheet(string name)
        {
            var h = Native.xlpp_workbook_add_sheet(_handle, name);
            if (h == IntPtr.Zero)
                throw new ArgumentException($"Failed to add worksheet '{name}': {LastError}");
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

        public bool Load(string path) => Native.xlpp_workbook_load(_handle, path) != 0;
        public bool Save(string path) => Native.xlpp_workbook_save(_handle, path) != 0;

        public Properties Properties
        {
            get
            {
                var h = Native.xlpp_workbook_properties(_handle);
                return new Properties(h);
            }
        }
    }
}
