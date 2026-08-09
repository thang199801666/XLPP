option(XLPP_BENCHMARK_EXTERNAL "Build competitor benchmark requiring libxlsxwriter" OFF)

add_executable(xlpp_native_benchmark bench_native.cpp)
target_include_directories(xlpp_native_benchmark PRIVATE ${PROJECT_SOURCE_DIR}/include)
target_link_libraries(xlpp_native_benchmark PRIVATE xlpp_static)
target_compile_features(xlpp_native_benchmark PRIVATE cxx_std_20)

if(XLPP_BENCHMARK_EXTERNAL)
    find_library(LIBXLSXWRITER_LIBRARY NAMES xlsxwriter REQUIRED)
    find_package(ZLIB REQUIRED)
    add_library(XLPP::xlsxwriter UNKNOWN IMPORTED)
    if(WIN32)
        set_target_properties(XLPP::xlsxwriter PROPERTIES
            IMPORTED_IMPLIB "${LIBXLSXWRITER_LIBRARY}"
            IMPORTED_LOCATION "${LIBXLSXWRITER_LIBRARY}")
    else()
        set_target_properties(XLPP::xlsxwriter PROPERTIES IMPORTED_LOCATION "${LIBXLSXWRITER_LIBRARY}")
    endif()
    set_target_properties(XLPP::xlsxwriter PROPERTIES INTERFACE_LINK_LIBRARIES ZLIB::ZLIB)
    add_executable(xlpp_external_benchmark bench_external.cpp)
    target_include_directories(xlpp_external_benchmark PRIVATE ${PROJECT_SOURCE_DIR}/include)
    target_link_libraries(xlpp_external_benchmark PRIVATE xlpp_static XLPP::xlsxwriter)
    target_compile_features(xlpp_external_benchmark PRIVATE cxx_std_20)
endif()
