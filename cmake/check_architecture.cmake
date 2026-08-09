if(NOT DEFINED XLPP_SOURCE_DIR)
    message(FATAL_ERROR "XLPP_SOURCE_DIR is required")
endif()

function(xlpp_assert_no_include root description)
    set(forbidden ${ARGN})
    file(GLOB_RECURSE sources
        "${root}/*.h" "${root}/*.hpp" "${root}/*.cpp")
    foreach(source IN LISTS sources)
        file(READ "${source}" content)
        foreach(pattern IN LISTS forbidden)
            string(FIND "${content}" "${pattern}" hit)
            if(NOT hit EQUAL -1)
                file(RELATIVE_PATH relative "${XLPP_SOURCE_DIR}" "${source}")
                message(FATAL_ERROR
                    "Architecture boundary violation (${description}): ${relative} contains '${pattern}'")
            endif()
        endforeach()
    endforeach()
endfunction()

# Bytes/package primitives must stay independent of the semantic workbook model.
xlpp_assert_no_include(
    "${XLPP_SOURCE_DIR}/src/XLPP/Package"
    "Package -> semantic model"
    "#include <XLPP/Workbook/"
    "#include <XLPP/Worksheet/"
    "#include <XLPP/Chart/"
    "#include <XLPP/Pivot/"
)

# Model implementations are format-neutral and must not depend on OOXML/package codecs.
xlpp_assert_no_include(
    "${XLPP_SOURCE_DIR}/src/XLPP/Model"
    "Model -> serialization layer"
    "#include \"OOXML/"
    "#include \"Package/"
)


# Semantic subsystems below the codec layer must remain independently testable
# and reusable without pulling serialization/package implementation details back
# into the model.  These checks complement the broader Model rule above and
# prevent dependency creep as formula, dependency and validation code grows.
foreach(domain Formula Dependencies Validation)
    xlpp_assert_no_include(
        "${XLPP_SOURCE_DIR}/src/XLPP/${domain}"
        "${domain} -> serialization layer"
        "#include \"OOXML/"
        "#include \"Package/"
        "#include <XLPP/OOXML/"
        "#include <XLPP/Package/"
    )
endforeach()

message(STATUS "XLPP architecture boundary checks passed")
