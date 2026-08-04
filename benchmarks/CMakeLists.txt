find_path(LIBXLSXWRITER_INCLUDE_DIR xlsxwriter.h REQUIRED)
find_library(LIBXLSXWRITER_LIBRARY NAMES xlsxwriter REQUIRED)

add_executable(xlpp_external_benchmark bench_external.cpp)
target_include_directories(xlpp_external_benchmark PRIVATE
    ${PROJECT_SOURCE_DIR}/include
    ${LIBXLSXWRITER_INCLUDE_DIR}
)
target_link_libraries(xlpp_external_benchmark PRIVATE
    xlpp_static
    ${LIBXLSXWRITER_LIBRARY}
)
target_compile_features(xlpp_external_benchmark PRIVATE cxx_std_20)
