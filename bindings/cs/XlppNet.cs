// XL++ C# Bindings — P/Invoke wrapper with OOP API
// Requires xlpp_capi.dll in the same directory or PATH.

using System;
using System.Runtime.InteropServices;

namespace XLPP
{
    // === Native interop ===
    internal static class Native
    {
        private const string Dll = "xlpp_capi";

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr xlpp_version();
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern string xlpp_version_str();  // simulated

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
        public static extern ulong xlpp_sheet_max_row(IntPtr ws);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern ulong xlpp_sheet_max_col(IntPtr ws);

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
        public string Title { set => Native.xlpp_properties_set_title(_handle, value); }
        public string Creator { set => Native.xlpp_properties_set_creator(_handle, value); }
    }

    public class Cell
    {
        internal IntPtr _handle;
        internal Cell(IntPtr h) => _handle = h;

        public CellValueType ValueType => (CellValueType)Native.xlpp_cell_value_type(_handle);
        public bool IsEmpty => Native.xlpp_cell_is_empty(_handle) != 0;

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

        public Font Font => new(Native.xlpp_cell_font(_handle));
        public Fill Fill => new(Native.xlpp_cell_fill(_handle));
        public Border Border => new(Native.xlpp_cell_border(_handle));
        public Alignment Alignment => new(Native.xlpp_cell_alignment(_handle));
    }

    public class Worksheet
    {
        internal IntPtr _handle;
        internal Worksheet(IntPtr h) => _handle = h;

        public Cell this[string address] => new(Native.xlpp_sheet_cell(_handle, address));
        public Cell Cell(ulong row, ulong col) => new(Native.xlpp_sheet_cell_rc(_handle, row, col));
        public Cell Cell(string address) => this[address];
        public bool HasCell(string address) => Native.xlpp_sheet_has_cell(_handle, address) != 0;

        public ulong MaxRow => Native.xlpp_sheet_max_row(_handle);
        public ulong MaxColumn => Native.xlpp_sheet_max_col(_handle);

        public void MergeCells(string range) => Native.xlpp_sheet_merge_cells(_handle, range);
        public void UnmergeCells(string range) => Native.xlpp_sheet_unmerge_cells(_handle, range);
        public void FreezePanes(string cell) => Native.xlpp_sheet_freeze_panes(_handle, cell);
    }

    public class Workbook : IDisposable
    {
        internal IntPtr _handle;

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
                return new Worksheet(h);
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
