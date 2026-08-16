# C API Shared Library (for P/Invoke, FFI)

add_library(xlpp_capi SHARED
    xlpp_capi.cpp
    xlpp_capi_chart.cpp
    xlpp_capi.h
)

target_include_directories(xlpp_capi PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../../src/XLPP
)

target_link_libraries(xlpp_capi PRIVATE xlpp_static)

target_compile_definitions(xlpp_capi PRIVATE XLPP_CAPI_EXPORTS XLPP_STATIC)

if(WIN32)
    set_target_properties(xlpp_capi PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
    )
endif()

install(TARGETS xlpp_capi
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
)

if(XLPP_BUILD_TESTS)
    add_executable(xlpp_capi_smoke capi_smoke.cpp)
    target_link_libraries(xlpp_capi_smoke PRIVATE xlpp_capi)
    target_include_directories(xlpp_capi_smoke PRIVATE ${CMAKE_CURRENT_SOURCE_DIR})
    add_test(NAME XLPP_CApiSmoke COMMAND xlpp_capi_smoke)
endif()
