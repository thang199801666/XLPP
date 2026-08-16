# XLPP Static Library

set(XLPP_SOURCES
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/Workbook.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookStylesIO.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookDrawingIO.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookChartReader.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookPivotWrite.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookPartXml.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookDrawingEdit.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookChartEditing.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookSheetWriter.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookUserForm.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookPivotRead.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Legacy/XlsBinaryReader.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Legacy/XlsBinaryWriter.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Legacy/XlsbBinaryReader.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookChartSerializer.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookIO.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookChartsheetIO.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookChartsheetPackage.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookPackageReader.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookVba.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookPivot.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookChartCache.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookChartStyle.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookStructuralEdit.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookModelValidation.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Workbook/WorkbookSheetOperations.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Formula/ReferenceTransformer.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Formula/DependencyGraph.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Pivot/PivotStructuralEdit.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Encryption/OfficeCrypto.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Vba/VbaProjectBinary.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Worksheet/Worksheet.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Worksheet/WorksheetStructuralEdit.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Worksheet/Range.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Packaging/ZipArchive.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Packaging/ZipArchiveReader.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Packaging/RelationshipGraph.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/XML/XmlUtilities.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/XML/XmlPullReader.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Streaming/StreamingWorkbookWriter.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Streaming/StreamingWorkbookReader.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/Streaming/SharedStringsReader.cpp
)

file(GLOB_RECURSE XLPP_INTERNAL_HEADERS
    CONFIGURE_DEPENDS
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/*.h
)

file(GLOB_RECURSE XLPP_PUBLIC_HEADERS
    CONFIGURE_DEPENDS
    ${CMAKE_CURRENT_SOURCE_DIR}/../include/XLPP/*.h
)

add_library(xlpp_static STATIC
    ${XLPP_SOURCES}
    ${XLPP_INTERNAL_HEADERS}
    ${XLPP_PUBLIC_HEADERS}
)

set_target_properties(xlpp_static PROPERTIES POSITION_INDEPENDENT_CODE ON)

add_library(XLPP::xlpp ALIAS xlpp_static)

target_include_directories(xlpp_static BEFORE PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP
    ${CMAKE_CURRENT_SOURCE_DIR}/../third_party/zlib
)
target_include_directories(xlpp_static
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/../include>
        $<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>
)

target_link_libraries(xlpp_static PUBLIC ${ZLIB_TARGET})
if(WIN32)
    target_link_libraries(xlpp_static PUBLIC bcrypt)
else()
    target_link_libraries(xlpp_static PUBLIC OpenSSL::Crypto)
endif()

target_compile_definitions(xlpp_static PUBLIC XLPP_STATIC)

# MSVC-specific: enable parallel compilation
if(MSVC)
    target_compile_options(xlpp_static PRIVATE /MP)
endif()
