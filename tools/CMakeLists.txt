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

add_executable(calc-debug
    calc-debug/main.cpp
)

target_include_directories(calc-debug PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(calc-debug PRIVATE xlpp_static)

add_executable(xlpp-fuzz-harness
    xlpp-fuzz-harness/main.cpp
)

target_include_directories(xlpp-fuzz-harness PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(xlpp-fuzz-harness PRIVATE xlpp_static)
