cmake_policy(SET CMP0057 NEW)

if(NOT DEFINED XLPP_SOURCE_DIR)
    message(FATAL_ERROR "XLPP_SOURCE_DIR is required")
endif()

function(xlpp_require file needle label)
    file(READ "${XLPP_SOURCE_DIR}/${file}" _text)
    string(FIND "${_text}" "${needle}" _pos)
    if(_pos EQUAL -1)
        message(FATAL_ERROR "Binding parity failure: ${label} missing '${needle}' in ${file}")
    endif()
endfunction()

# Keep public package versions synchronized with the native project version.
file(READ "${XLPP_SOURCE_DIR}/VERSION" _version)
string(STRIP "${_version}" _version)
foreach(_file "CMakeLists.txt" "vcpkg.json" "bindings/csharp/XlppNet.csproj" "bindings/python/src/xlpp_bindings.cpp" "bindings/c/xlpp_capi.cpp")
    xlpp_require("${_file}" "${_version}" "version synchronization")
endforeach()

# Scoped defined names are an Excel semantic rather than a naming convention;
# keep explicit bridge checks for this overload family.
xlpp_require("bindings/c/xlpp_capi.h" "xlpp_workbook_defined_name_scoped" "C ABI scoped defined-name lookup")
xlpp_require("bindings/c/xlpp_capi.h" "xlpp_workbook_add_defined_name_scoped" "C ABI scoped defined-name creation")
xlpp_require("bindings/python/src/xlpp_bindings.cpp" "local_sheet_id" "Python scoped defined-name lookup")
xlpp_require("bindings/csharp/XlppNet.cs" "GetDefinedName(string name, ulong? localSheetId)" "C# scoped defined-name lookup")

# Full C# P/Invoke symbol consistency. Every managed native declaration must
# exist in the C header, and every Native.xlpp_* call must have a declaration.
file(READ "${XLPP_SOURCE_DIR}/bindings/c/xlpp_capi.h" _capi_header)
file(READ "${XLPP_SOURCE_DIR}/bindings/csharp/XlppNet.cs" _csharp_base)
file(READ "${XLPP_SOURCE_DIR}/bindings/csharp/XlppNet.Advanced.cs" _csharp_advanced)
set(_csharp_all "${_csharp_base}\n${_csharp_advanced}")

string(REGEX MATCHALL "extern[^;\n]+xlpp_[A-Za-z0-9_]+" _extern_matches "${_csharp_all}")
set(_extern_symbols "")
foreach(_match IN LISTS _extern_matches)
    string(REGEX MATCH "xlpp_[A-Za-z0-9_]+$" _symbol "${_match}")
    list(APPEND _extern_symbols "${_symbol}")
    string(FIND "${_capi_header}" "${_symbol}" _header_pos)
    if(_header_pos EQUAL -1)
        message(FATAL_ERROR "Binding parity failure: C# P/Invoke '${_symbol}' is absent from xlpp_capi.h")
    endif()
endforeach()
list(REMOVE_DUPLICATES _extern_symbols)

# C# is the managed projection of the stable C ABI. Require every exported C
# function to have a P/Invoke declaration so additive native APIs cannot be
# silently omitted from the managed package.
string(REGEX MATCHALL "XLPP_API[^;\n]+xlpp_[A-Za-z0-9_]+[ \t]*\\(" _capi_export_matches "${_capi_header}")
foreach(_match IN LISTS _capi_export_matches)
    string(REGEX MATCH "xlpp_[A-Za-z0-9_]+[ \t]*\\($" _symbol_with_paren "${_match}")
    string(REGEX MATCH "xlpp_[A-Za-z0-9_]+" _symbol "${_symbol_with_paren}")
    list(FIND _extern_symbols "${_symbol}" _extern_index)
    if(_extern_index EQUAL -1)
        message(FATAL_ERROR "Binding parity failure: C ABI export '${_symbol}' has no C# P/Invoke declaration")
    endif()
endforeach()

string(REGEX MATCHALL "Native\\.xlpp_[A-Za-z0-9_]+" _native_call_matches "${_csharp_all}")
foreach(_match IN LISTS _native_call_matches)
    string(REGEX MATCH "xlpp_[A-Za-z0-9_]+" _symbol "${_match}")
    list(FIND _extern_symbols "${_symbol}" _extern_index)
    if(_extern_index EQUAL -1)
        message(FATAL_ERROR "Binding parity failure: C# call Native.${_symbol} has no P/Invoke declaration")
    endif()
endforeach()

message(STATUS "Binding bridge consistency PASS for XLPP ${_version}")
