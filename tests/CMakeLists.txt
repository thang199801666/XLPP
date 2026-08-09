# Unit Tests

add_executable(XLPP_UnitTests
    unit/TestMain.cpp
    unit/Model/ModelWorkbookTests.cpp
    unit/Formula/FormulaDependencyTests.cpp
    unit/Regression/RegressionRegistry.cpp
    unit/Regression/PackageStreamingTests.cpp
    unit/Regression/WorkbookTests.cpp
    unit/Regression/DrawingsTests.cpp
    unit/Regression/ChartsTests.cpp
    unit/Regression/PivotTests.cpp
    unit/Regression/CompatibilityTests.cpp
    unit/Regression/RefinementPhaseTests.cpp
)

target_include_directories(XLPP_UnitTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_compile_definitions(XLPP_UnitTests PRIVATE
    XLPP_STATIC
    NOMINMAX
    XLPP_TEST_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
)

target_link_libraries(XLPP_UnitTests PRIVATE xlpp_static)

# The regression runner intentionally covers almost every public subsystem in
# one process. Precompiling the umbrella public header keeps incremental test
# builds practical while standalone-header checks continue to guard include
# hygiene independently.
target_precompile_headers(XLPP_UnitTests PRIVATE <XLPP/XLPP.h>)

# Regression coverage is split by domain so test compilation can run in parallel.
# Release optimization provides little value for assertion-heavy test code; keep
# production XLPP objects optimized while compiling regression-only TUs at -O0.
set(XLPP_REGRESSION_SOURCES
    unit/Regression/PackageStreamingTests.cpp
    unit/Regression/WorkbookTests.cpp
    unit/Regression/DrawingsTests.cpp
    unit/Regression/ChartsTests.cpp
    unit/Regression/PivotTests.cpp
    unit/Regression/CompatibilityTests.cpp
    unit/Regression/RefinementPhaseTests.cpp
)
if(MSVC)
    set_source_files_properties(${XLPP_REGRESSION_SOURCES} PROPERTIES COMPILE_OPTIONS "/Od" SKIP_PRECOMPILE_HEADERS ON)
elseif(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
    set_source_files_properties(${XLPP_REGRESSION_SOURCES} PROPERTIES COMPILE_OPTIONS "-O0" SKIP_PRECOMPILE_HEADERS ON)
endif()

add_test(NAME XLPP_UnitTests COMMAND XLPP_UnitTests)

add_test(NAME ArchitectureBoundaryTests
    COMMAND ${CMAKE_COMMAND}
        -DXLPP_SOURCE_DIR=${CMAKE_SOURCE_DIR}
        -P ${CMAKE_SOURCE_DIR}/cmake/check_architecture.cmake)

add_test(NAME BindingParityTests
    COMMAND ${CMAKE_COMMAND}
        -DXLPP_SOURCE_DIR=${CMAKE_SOURCE_DIR}
        -P ${CMAKE_SOURCE_DIR}/cmake/check_binding_parity.cmake)

find_package(Python3 QUIET COMPONENTS Interpreter)
if(Python3_Interpreter_FOUND)
    add_test(NAME BindingManifestTests
        COMMAND ${Python3_EXECUTABLE}
            ${CMAKE_SOURCE_DIR}/tools/generate_binding_manifest.py
            --root ${CMAKE_SOURCE_DIR}
            --check ${CMAKE_SOURCE_DIR}/bindings/PARITY_MANIFEST.json)
else()
    message(STATUS "Python3 not found; BindingManifestTests will be skipped")
endif()

# Standalone examples. Each executable writes its own workbook.
add_executable(XLPP_WriteValues
    XLPP.Examples/write_values.cpp
    XLPP.Examples/standalone_values_main.cpp
)
add_executable(XLPP_WriteChart
    XLPP.Examples/write_chart.cpp
    XLPP.Examples/standalone_chart_main.cpp
)
add_executable(XLPP_WriteImage
    XLPP.Examples/write_image.cpp
    XLPP.Examples/standalone_image_main.cpp
)
foreach(_example XLPP_WriteValues XLPP_WriteChart XLPP_WriteImage)
    target_include_directories(${_example} PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../include)
    target_compile_definitions(${_example} PRIVATE XLPP_STATIC NOMINMAX)
    target_link_libraries(${_example} PRIVATE xlpp_static)
endforeach()

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
target_compile_definitions(XLPP_HeaderCheck PRIVATE XLPP_STATIC NOMINMAX)

# --- Package consumer tests ---
# Installs the project to a temp prefix, then builds and runs a consumer that
# uses find_package(XLPP CONFIG) plus one that uses add_subdirectory().
set(XLPP_PACKAGE_TEST_SCRIPT ${CMAKE_CURRENT_BINARY_DIR}/package_test.cmake)
configure_file(
    ${CMAKE_CURRENT_SOURCE_DIR}/package/package_test.cmake.in
    ${XLPP_PACKAGE_TEST_SCRIPT}
    @ONLY
)
add_test(NAME PackageConsumerTests COMMAND ${CMAKE_COMMAND} -P ${XLPP_PACKAGE_TEST_SCRIPT})
