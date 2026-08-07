add_executable(xlpp-package-validator
    xlpp-package-validator/main.cpp
)

target_include_directories(xlpp-package-validator PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(xlpp-package-validator PRIVATE xlpp_static)
