#pragma once

#include <XLPP/Formula/ReferenceTranslator.h>

#include <cstddef>

namespace xlpp {

struct WorksheetStructuralEditReport {
    std::size_t cellsMoved{0};
    std::size_t cellsRemoved{0};
    std::size_t formulasUpdated{0};
    std::size_t formulaMetadataUpdated{0};
    std::size_t worksheetReferencesUpdated{0};
    std::size_t referencesInvalidated{0};
    std::size_t drawingAnchorsUpdated{0};
    std::size_t chartReferencesUpdated{0};
    std::size_t pivotReferencesUpdated{0};
    std::size_t hyperlinksUpdated{0};
};

} // namespace xlpp
