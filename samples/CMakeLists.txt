# Sample Programs

add_executable(XLPP_Sample
    XLPP.Sample/main.cpp
)

target_include_directories(XLPP_Sample PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(XLPP_Sample PRIVATE xlpp_static)
