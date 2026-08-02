# Unit Tests

add_executable(XLPP_UnitTests
    XLPP.UnitTests/main.cpp
)

target_include_directories(XLPP_UnitTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(XLPP_UnitTests PRIVATE xlpp_static)

add_test(NAME XLPP_UnitTests COMMAND XLPP_UnitTests)

# --- Standalone public-header include test ---
# Verifies every public header can be included on its own (no missing
# transitive includes). Generates one tiny .cpp per header at configure time.
file(GLOB_RECURSE XLPP_PUBLIC_HEADERS
    ${CMAKE_CURRENT_SOURCE_DIR}/../include/XLPP/*.h
)

set(_standalone_sources)
foreach(_hdr IN LISTS XLPP_PUBLIC_HEADERS)
    file(RELATIVE_PATH _rel ${CMAKE_CURRENT_SOURCE_DIR}/../include ${_hdr})
    string(REPLACE "/" "_" _safe ${_rel})
    string(REPLACE "\\" "_" _safe ${_safe})
    string(REGEX REPLACE "\\.h$" "" _base ${_safe})
    set(_gen ${CMAKE_CURRENT_BINARY_DIR}/standalone/${_safe}.cpp)
    file(WRITE ${_gen}
        "#include <${_rel}>\n"
        "static int verify_${_base}_compiles() { return 0; }\n")
    list(APPEND _standalone_sources ${_gen})
endforeach()

add_library(XLPP_HeaderCheck OBJECT ${_standalone_sources})
target_include_directories(XLPP_HeaderCheck PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../include
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)
target_compile_definitions(XLPP_HeaderCheck PRIVATE XLPP_STATIC)
