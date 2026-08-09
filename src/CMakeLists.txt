# XLPP core modules
#
# The public deliverable remains a single XLPP::xlpp / XLPP::XLPP static
# library.  Internally, sources are compiled as domain object libraries so
# dependency boundaries are visible and incremental builds can scale without
# turning Workbook into a build-time god target.

set(XLPP_INTERNAL_ROOT ${CMAKE_CURRENT_SOURCE_DIR}/XLPP)

file(GLOB_RECURSE XLPP_PUBLIC_HEADERS
    CONFIGURE_DEPENDS
    ${CMAKE_CURRENT_SOURCE_DIR}/../include/XLPP/*.h
)

function(xlpp_add_core_module target)
    set(_sources ${ARGN})
    add_library(${target} OBJECT ${_sources})
    set_target_properties(${target} PROPERTIES POSITION_INDEPENDENT_CODE ON)

    target_include_directories(${target}
        PRIVATE
            ${CMAKE_CURRENT_SOURCE_DIR}
            ${XLPP_INTERNAL_ROOT}
            ${CMAKE_CURRENT_SOURCE_DIR}/../include
    )
    target_include_directories(${target} SYSTEM PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/../third_party/zlib
    )
    target_compile_definitions(${target} PRIVATE XLPP_STATIC NOMINMAX)

    if(XLPP_ENABLE_STRICT_WARNINGS)
        if(MSVC)
            target_compile_options(${target} PRIVATE /W4 /permissive- /we4700 /we4715)
        elseif(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
            target_compile_options(${target} PRIVATE
                -Wall -Wextra -Wpedantic -Wshadow -Wconversion -Wsign-conversion
                -Wnon-virtual-dtor -Wold-style-cast -Woverloaded-virtual
                -Werror=return-type -Werror=uninitialized -Werror=array-bounds)
            if(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
                target_compile_options(${target} PRIVATE
                    -Werror=maybe-uninitialized -Werror=stringop-overflow)
            endif()
        endif()
    endif()

    if(MSVC)
        target_compile_options(${target} PRIVATE /MP)
    endif()
endfunction()

file(GLOB_RECURSE XLPP_MODEL_WORKBOOK_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Model/Workbook/*.cpp)
file(GLOB_RECURSE XLPP_MODEL_WORKSHEET_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Model/Worksheet/*.cpp)
set(XLPP_FORMULA_ENGINE_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/Calculation.cpp)
set(XLPP_FORMULA_PARSER_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/FormulaParser.cpp)
set(XLPP_FORMULA_FUNCTION_DISPATCH_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/FormulaFunctions.cpp
    ${XLPP_INTERNAL_ROOT}/Formula/FormulaFunctionSupport.cpp
    ${XLPP_INTERNAL_ROOT}/Formula/FunctionRegistry.cpp)
set(XLPP_FORMULA_FUNCTION_MATH_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/FormulaFunctionsMath.cpp)
set(XLPP_FORMULA_FUNCTION_TEXT_DATE_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/FormulaFunctionsTextDate.cpp)
set(XLPP_FORMULA_FUNCTION_LOOKUP_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/FormulaFunctionsLookup.cpp)
set(XLPP_FORMULA_FUNCTION_DYNAMIC_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/FormulaFunctionsDynamic.cpp)
set(XLPP_FORMULA_DEPENDENCY_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/DependencyGraph.cpp)
set(XLPP_FORMULA_REFERENCE_SOURCES
    ${XLPP_INTERNAL_ROOT}/Formula/ReferenceTranslator.cpp)
file(GLOB_RECURSE XLPP_DEPENDENCY_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Dependencies/*.cpp)
file(GLOB_RECURSE XLPP_PACKAGE_ZIP_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Package/Zip/*.cpp)
file(GLOB_RECURSE XLPP_PACKAGE_OPC_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Package/Opc/*.cpp)
file(GLOB_RECURSE XLPP_PACKAGE_XML_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Package/Xml/*.cpp)
file(GLOB_RECURSE XLPP_IO_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/IO/*.cpp)
file(GLOB_RECURSE XLPP_PRESERVATION_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Preservation/*.cpp)
file(GLOB_RECURSE XLPP_OOXML_COMMON_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/OOXML/Common/*.cpp)
file(GLOB_RECURSE XLPP_OOXML_WORKBOOK_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/OOXML/Workbook/*.cpp)
list(FILTER XLPP_OOXML_WORKBOOK_SOURCES EXCLUDE REGEX "/WorkbookCodec\\.cpp$")
file(GLOB_RECURSE XLPP_OOXML_WORKSHEET_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/OOXML/Worksheet/*.cpp)
set(XLPP_OOXML_CHART_SUPPORT_SOURCES
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ChartFormatCodec.cpp
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ChartMutationSupport.cpp
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ChartXmlSupport.cpp)
set(XLPP_OOXML_CHART_READER_SOURCES
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ChartReader.cpp)
set(XLPP_OOXML_CHART_SERIALIZER_SOURCES
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ChartSerializer.cpp)
set(XLPP_OOXML_CHART_XML_EDITOR_SOURCES
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ImportedChartLayoutEditor.cpp
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ImportedChartSeriesEditor.cpp
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ImportedChartPlotEditor.cpp
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ImportedChartXmlEditor.cpp)
set(XLPP_OOXML_CHART_PACKAGE_PATCHER_SOURCES
    ${XLPP_INTERNAL_ROOT}/OOXML/Charts/ImportedChartPatcher.cpp)
file(GLOB_RECURSE XLPP_OOXML_DRAWING_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/OOXML/Drawings/*.cpp)
file(GLOB_RECURSE XLPP_OOXML_PIVOT_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/OOXML/Pivot/*.cpp)
file(GLOB_RECURSE XLPP_OOXML_STYLE_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/OOXML/Styles/*.cpp)
file(GLOB_RECURSE XLPP_OOXML_TABLE_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/OOXML/Tables/*.cpp)
file(GLOB_RECURSE XLPP_OOXML_COMMENT_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/OOXML/Comments/*.cpp)
file(GLOB_RECURSE XLPP_CHART_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Charts/*.cpp)
file(GLOB_RECURSE XLPP_STREAMING_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Streaming/*.cpp)
file(GLOB_RECURSE XLPP_ENCRYPTION_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Encryption/*.cpp)
file(GLOB_RECURSE XLPP_VBA_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/VBA/*.cpp)
file(GLOB_RECURSE XLPP_VALIDATION_SOURCES CONFIGURE_DEPENDS
    ${XLPP_INTERNAL_ROOT}/Validation/*.cpp)

xlpp_add_core_module(xlpp_model_workbook_obj ${XLPP_MODEL_WORKBOOK_SOURCES})
xlpp_add_core_module(xlpp_model_worksheet_obj ${XLPP_MODEL_WORKSHEET_SOURCES})
xlpp_add_core_module(xlpp_formula_engine_obj ${XLPP_FORMULA_ENGINE_SOURCES})
xlpp_add_core_module(xlpp_formula_parser_obj ${XLPP_FORMULA_PARSER_SOURCES})
xlpp_add_core_module(xlpp_formula_function_dispatch_obj ${XLPP_FORMULA_FUNCTION_DISPATCH_SOURCES})
xlpp_add_core_module(xlpp_formula_function_math_obj ${XLPP_FORMULA_FUNCTION_MATH_SOURCES})
xlpp_add_core_module(xlpp_formula_function_text_date_obj ${XLPP_FORMULA_FUNCTION_TEXT_DATE_SOURCES})
xlpp_add_core_module(xlpp_formula_function_lookup_obj ${XLPP_FORMULA_FUNCTION_LOOKUP_SOURCES})
xlpp_add_core_module(xlpp_formula_function_dynamic_obj ${XLPP_FORMULA_FUNCTION_DYNAMIC_SOURCES})
xlpp_add_core_module(xlpp_formula_dependency_obj ${XLPP_FORMULA_DEPENDENCY_SOURCES})
xlpp_add_core_module(xlpp_formula_reference_obj ${XLPP_FORMULA_REFERENCE_SOURCES})
xlpp_add_core_module(xlpp_dependencies_obj ${XLPP_DEPENDENCY_SOURCES})
xlpp_add_core_module(xlpp_package_zip_obj ${XLPP_PACKAGE_ZIP_SOURCES})
xlpp_add_core_module(xlpp_package_opc_obj ${XLPP_PACKAGE_OPC_SOURCES})
xlpp_add_core_module(xlpp_package_xml_obj ${XLPP_PACKAGE_XML_SOURCES})
xlpp_add_core_module(xlpp_io_obj ${XLPP_IO_SOURCES})
xlpp_add_core_module(xlpp_preservation_obj ${XLPP_PRESERVATION_SOURCES})
xlpp_add_core_module(xlpp_ooxml_common_obj ${XLPP_OOXML_COMMON_SOURCES})
xlpp_add_core_module(xlpp_ooxml_workbook_obj ${XLPP_OOXML_WORKBOOK_SOURCES})
xlpp_add_core_module(xlpp_ooxml_worksheet_obj ${XLPP_OOXML_WORKSHEET_SOURCES})
xlpp_add_core_module(xlpp_ooxml_chart_support_obj ${XLPP_OOXML_CHART_SUPPORT_SOURCES})
xlpp_add_core_module(xlpp_ooxml_chart_reader_obj ${XLPP_OOXML_CHART_READER_SOURCES})
xlpp_add_core_module(xlpp_ooxml_chart_serializer_obj ${XLPP_OOXML_CHART_SERIALIZER_SOURCES})
xlpp_add_core_module(xlpp_ooxml_chart_xml_editor_obj ${XLPP_OOXML_CHART_XML_EDITOR_SOURCES})
xlpp_add_core_module(xlpp_ooxml_chart_package_patcher_obj ${XLPP_OOXML_CHART_PACKAGE_PATCHER_SOURCES})
xlpp_add_core_module(xlpp_ooxml_drawings_obj ${XLPP_OOXML_DRAWING_SOURCES})
xlpp_add_core_module(xlpp_ooxml_pivot_obj ${XLPP_OOXML_PIVOT_SOURCES})
xlpp_add_core_module(xlpp_ooxml_styles_obj ${XLPP_OOXML_STYLE_SOURCES})
xlpp_add_core_module(xlpp_ooxml_tables_obj ${XLPP_OOXML_TABLE_SOURCES})
xlpp_add_core_module(xlpp_ooxml_comments_obj ${XLPP_OOXML_COMMENT_SOURCES})
xlpp_add_core_module(xlpp_charts_obj ${XLPP_CHART_SOURCES})
xlpp_add_core_module(xlpp_streaming_obj ${XLPP_STREAMING_SOURCES})
xlpp_add_core_module(xlpp_encryption_obj ${XLPP_ENCRYPTION_SOURCES})
xlpp_add_core_module(xlpp_vba_obj ${XLPP_VBA_SOURCES})
xlpp_add_core_module(xlpp_validation_obj ${XLPP_VALIDATION_SOURCES})

set(XLPP_CORE_OBJECT_TARGETS
    xlpp_model_workbook_obj
    xlpp_model_worksheet_obj
    xlpp_formula_engine_obj
    xlpp_formula_parser_obj
    xlpp_formula_function_dispatch_obj
    xlpp_formula_function_math_obj
    xlpp_formula_function_text_date_obj
    xlpp_formula_function_lookup_obj
    xlpp_formula_function_dynamic_obj
    xlpp_formula_dependency_obj
    xlpp_formula_reference_obj
    xlpp_dependencies_obj
    xlpp_package_zip_obj
    xlpp_package_opc_obj
    xlpp_package_xml_obj
    xlpp_io_obj
    xlpp_preservation_obj
    xlpp_ooxml_common_obj
    xlpp_ooxml_workbook_obj
    xlpp_ooxml_worksheet_obj
    xlpp_ooxml_chart_support_obj
    xlpp_ooxml_chart_reader_obj
    xlpp_ooxml_chart_serializer_obj
    xlpp_ooxml_chart_xml_editor_obj
    xlpp_ooxml_chart_package_patcher_obj
    xlpp_ooxml_drawings_obj
    xlpp_ooxml_pivot_obj
    xlpp_ooxml_styles_obj
    xlpp_ooxml_tables_obj
    xlpp_ooxml_comments_obj
    xlpp_charts_obj
    xlpp_streaming_obj
    xlpp_encryption_obj
    xlpp_vba_obj
    xlpp_validation_obj
)

add_library(xlpp_static STATIC ${XLPP_PUBLIC_HEADERS})
foreach(_module IN LISTS XLPP_CORE_OBJECT_TARGETS)
    target_sources(xlpp_static PRIVATE $<TARGET_OBJECTS:${_module}>)
endforeach()
set_target_properties(xlpp_static PROPERTIES POSITION_INDEPENDENT_CODE ON)

add_library(XLPP::xlpp ALIAS xlpp_static)

target_include_directories(xlpp_static
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/../include>
        $<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>
)
target_link_libraries(xlpp_static PUBLIC ${ZLIB_TARGET})
if(WIN32)
    target_link_libraries(xlpp_static PUBLIC bcrypt)
endif()
target_compile_definitions(xlpp_static PUBLIC XLPP_STATIC NOMINMAX)
