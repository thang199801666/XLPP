# XLPP Static Library

file(GLOB_RECURSE XLPP_SOURCES
    CONFIGURE_DEPENDS
    ${CMAKE_CURRENT_SOURCE_DIR}/XLPP/*.cpp
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

target_include_directories(xlpp_static
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/../include>
        $<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>
    PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}
        ${CMAKE_CURRENT_SOURCE_DIR}/../third_party/zlib
)

target_link_libraries(xlpp_static PUBLIC ${ZLIB_TARGET})

target_compile_definitions(xlpp_static PUBLIC XLPP_STATIC)

# MSVC-specific: enable parallel compilation
if(MSVC)
    target_compile_options(xlpp_static PRIVATE /MP)
endif()
