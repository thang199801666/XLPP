# Unit Tests

add_executable(XLPP_UnitTests
    XLPP.UnitTests/main.cpp
)

target_include_directories(XLPP_UnitTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(XLPP_UnitTests PRIVATE xlpp_static)

add_test(NAME XLPP_UnitTests COMMAND XLPP_UnitTests)
