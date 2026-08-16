add_executable(xlpp-package-validator
    xlpp-package-validator/main.cpp
)

target_include_directories(xlpp-package-validator PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(xlpp-package-validator PRIVATE xlpp_static)

add_executable(xlpp-bench-writer
    xlpp-bench-writer/main.cpp
)

target_include_directories(xlpp-bench-writer PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(xlpp-bench-writer PRIVATE xlpp_static)
