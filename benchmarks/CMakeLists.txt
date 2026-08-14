# Native XL++ benchmark is always available when XLPP_BUILD_BENCHMARKS=ON.
# External comparison libraries are optional so performance-regression CI does
# not depend on a system-wide libxlsxwriter installation.
add_executable(xlpp_native_benchmark bench_native.cpp)
target_include_directories(xlpp_native_benchmark PRIVATE
    ${PROJECT_SOURCE_DIR}/include
)
target_link_libraries(xlpp_native_benchmark PRIVATE xlpp_static)
target_compile_features(xlpp_native_benchmark PRIVATE cxx_std_20)

# A smaller benchmark dedicated to core allocation/geometry/save hot paths.
# It is intentionally independent of external libraries and prints stable
# machine-readable records suitable for performance-regression CI.
add_executable(xlpp_core_hotpath_benchmark bench_core_hotpath.cpp)
target_include_directories(xlpp_core_hotpath_benchmark PRIVATE
    ${PROJECT_SOURCE_DIR}/include
)
target_link_libraries(xlpp_core_hotpath_benchmark PRIVATE xlpp_static)
target_compile_features(xlpp_core_hotpath_benchmark PRIVATE cxx_std_20)
if(WIN32)
    target_link_libraries(xlpp_core_hotpath_benchmark PRIVATE psapi)
endif()

add_executable(xlpp_cell_density_benchmark bench_cell_density.cpp)
target_include_directories(xlpp_cell_density_benchmark PRIVATE
    ${PROJECT_SOURCE_DIR}/include
)
target_link_libraries(xlpp_cell_density_benchmark PRIVATE xlpp_static)
target_compile_features(xlpp_cell_density_benchmark PRIVATE cxx_std_20)
if(WIN32)
    target_link_libraries(xlpp_cell_density_benchmark PRIVATE psapi)
endif()

find_library(LIBXLSXWRITER_LIBRARY NAMES xlsxwriter)
if(LIBXLSXWRITER_LIBRARY)
    find_package(ZLIB REQUIRED)

    add_library(XLPP::xlsxwriter UNKNOWN IMPORTED)
    if(WIN32)
        set_target_properties(XLPP::xlsxwriter PROPERTIES
            IMPORTED_IMPLIB "${LIBXLSXWRITER_LIBRARY}"
            IMPORTED_LOCATION "${LIBXLSXWRITER_LIBRARY}"
        )
    else()
        set_target_properties(XLPP::xlsxwriter PROPERTIES
            IMPORTED_LOCATION "${LIBXLSXWRITER_LIBRARY}"
        )
    endif()
    set_target_properties(XLPP::xlsxwriter PROPERTIES
        INTERFACE_LINK_LIBRARIES ZLIB::ZLIB
    )

    add_executable(xlpp_external_benchmark bench_external.cpp)
    target_include_directories(xlpp_external_benchmark PRIVATE
        ${PROJECT_SOURCE_DIR}/include
    )
    target_link_libraries(xlpp_external_benchmark PRIVATE
        xlpp_static
        XLPP::xlsxwriter
    )
    target_compile_features(xlpp_external_benchmark PRIVATE cxx_std_20)
else()
    message(STATUS "libxlsxwriter not found: building XL++ native benchmark only")
endif()
