# Unit Tests

add_executable(XLPP_UnitTests
    XLPP.UnitTests/main.cpp
)

target_include_directories(XLPP_UnitTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_compile_definitions(XLPP_UnitTests PRIVATE
    XLPP_TEST_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
)

target_link_libraries(XLPP_UnitTests PRIVATE xlpp_static)

add_test(NAME XLPP_UnitTests COMMAND XLPP_UnitTests)

# P1P density/correctness regression kept in a small translation unit so new
# core tests do not further increase the legacy monolithic unit-test compile cost.
add_executable(XLPP_P1P_LazyFormulaTests
    XLPP.UnitTests/p1p_lazy_formula.cpp
)
target_link_libraries(XLPP_P1P_LazyFormulaTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1P_LazyFormulaTests COMMAND XLPP_P1P_LazyFormulaTests)

# P1Q malformed-input and streaming relationship hardening regression.
add_executable(XLPP_P1Q_CoreHardeningTests
    XLPP.UnitTests/p1q_core_hardening.cpp
)
target_include_directories(XLPP_P1Q_CoreHardeningTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)
target_link_libraries(XLPP_P1Q_CoreHardeningTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1Q_CoreHardeningTests COMMAND XLPP_P1Q_CoreHardeningTests)

# P1R transactional load, stream-I/O and model-materialization hardening.
add_executable(XLPP_P1R_TransactionalLimitsTests
    XLPP.UnitTests/p1r_transactional_limits.cpp
)
target_include_directories(XLPP_P1R_TransactionalLimitsTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)
target_link_libraries(XLPP_P1R_TransactionalLimitsTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1R_TransactionalLimitsTests COMMAND XLPP_P1R_TransactionalLimitsTests)

# P1S read/write, chart generation and typed PivotCache parity regression.
add_executable(XLPP_P1S_ThreePillarTests
    XLPP.UnitTests/p1s_three_pillars.cpp
)
target_include_directories(XLPP_P1S_ThreePillarTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)
target_link_libraries(XLPP_P1S_ThreePillarTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1S_ThreePillarTests COMMAND XLPP_P1S_ThreePillarTests)

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
target_compile_definitions(XLPP_HeaderCheck PRIVATE XLPP_STATIC)

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

# P1T first-class chartsheets / workbook tab order and stable Pivot item identity.
add_executable(XLPP_P1T_ThreePillarTests
    XLPP.UnitTests/p1t_three_pillars.cpp
)
target_include_directories(XLPP_P1T_ThreePillarTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)
target_compile_definitions(XLPP_P1T_ThreePillarTests PRIVATE
    XLPP_TEST_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
)
target_link_libraries(XLPP_P1T_ThreePillarTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1T_ThreePillarTests COMMAND XLPP_P1T_ThreePillarTests)

# P1U template (.xltx/.xltm) and Chartsheet production semantics.
add_executable(XLPP_P1U_TemplateChartsheetTests
    XLPP.UnitTests/p1u_template_chartsheet.cpp
)
target_include_directories(XLPP_P1U_TemplateChartsheetTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)
target_compile_definitions(XLPP_P1U_TemplateChartsheetTests PRIVATE
    XLPP_TEST_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
)
target_link_libraries(XLPP_P1U_TemplateChartsheetTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1U_TemplateChartsheetTests COMMAND XLPP_P1U_TemplateChartsheetTests)

# P1V deeper template/Chartsheet page-setup, protection and custom-view semantics.
add_executable(XLPP_P1V_ChartsheetDepthTests
    XLPP.UnitTests/p1v_chartsheet_depth.cpp
)
target_include_directories(XLPP_P1V_ChartsheetDepthTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)
target_compile_definitions(XLPP_P1V_ChartsheetDepthTests PRIVATE
    XLPP_TEST_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
)
target_link_libraries(XLPP_P1V_ChartsheetDepthTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1V_ChartsheetDepthTests COMMAND XLPP_P1V_ChartsheetDepthTests)

add_executable(XLPP_P1W_ChartsheetPackageTests
    XLPP.UnitTests/p1w_chartsheet_package_ownership.cpp
)
target_include_directories(XLPP_P1W_ChartsheetPackageTests PRIVATE
    ${PROJECT_SOURCE_DIR}/src/XLPP
)
target_compile_definitions(XLPP_P1W_ChartsheetPackageTests PRIVATE
    XLPP_TEST_SOURCE_DIR="${PROJECT_SOURCE_DIR}"
)
target_link_libraries(XLPP_P1W_ChartsheetPackageTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1W_ChartsheetPackageTests COMMAND XLPP_P1W_ChartsheetPackageTests)

# P1X Chartsheet package-writer decomposition and relationship-collision hardening.
add_executable(XLPP_P1X_ChartsheetWriterTests
    XLPP.UnitTests/p1x_chartsheet_writer_decomposition.cpp
)
target_include_directories(XLPP_P1X_ChartsheetWriterTests PRIVATE
    ${PROJECT_SOURCE_DIR}/src/XLPP
)
target_link_libraries(XLPP_P1X_ChartsheetWriterTests PRIVATE xlpp_static)
add_test(NAME XLPP_P1X_ChartsheetWriterTests COMMAND XLPP_P1X_ChartsheetWriterTests)
