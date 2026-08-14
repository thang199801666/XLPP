#pragma once
#include "XLPP/Formula/ReferenceTransformer.h"
#include <XLPP/Pivot/PivotTable.h>
#include <XLPP/Workbook/StructuralEdit.h>
#include <string_view>

namespace xlpp::internal {

bool rewritePivotSourceForStructuralEdit(PivotTable& pivot,
                                         std::string_view ownerSheetName,
                                         const StructuralEditSpec& edit,
                                         StructuralEditReport& report);

} // namespace xlpp::internal
