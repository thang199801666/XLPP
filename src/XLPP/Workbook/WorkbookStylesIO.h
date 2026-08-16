#pragma once
#include <XLPP/Styles/Style.h>
#include <string>
#include <unordered_map>
#include <vector>

namespace xlpp {
class NamedStyle;

namespace internal {

// Deduplicating catalogs used by the shared-styles (xl/styles.xml) writer and
// reader. Styles are keyed by their semantic hash with a linear check against
// the bucket occupant to resolve collisions.
struct DxfCatalog {
    std::vector<xlpp::Style> items;
    std::unordered_map<std::size_t, std::size_t> index_;

    std::size_t id(const xlpp::Style& style) {
        const auto h = style.hash();
        const auto it = index_.find(h);
        if (it != index_.end() && items[it->second] == style) return it->second;
        const auto pos = items.size();
        items.push_back(style);
        index_[h] = pos;
        return pos;
    }

    std::size_t find(const xlpp::Style& style) const {
        const auto h = style.hash();
        const auto it = index_.find(h);
        if (it != index_.end() && items[it->second] == style) return it->second;
        return 0;
    }
};

struct StyleCatalog {
    std::vector<xlpp::Style> items{xlpp::Style{}};
    std::unordered_map<std::size_t, std::size_t> index_;

    StyleCatalog() { index_[xlpp::Style{}.hash()] = 0; }

    std::size_t id(const xlpp::Style& style) {
        const auto h = style.hash();
        const auto it = index_.find(h);
        if (it != index_.end() && items[it->second] == style) return it->second;
        const auto pos = items.size();
        items.push_back(style);
        index_[h] = pos;
        return pos;
    }

    std::size_t find(const xlpp::Style& style) const {
        const auto h = style.hash();
        const auto it = index_.find(h);
        if (it != index_.end() && items[it->second] == style) return it->second;
        return 0;
    }
};

// Serializes the shared styles part (xl/styles.xml).
std::string stylesXml(const StyleCatalog& catalog,
                      const std::vector<xlpp::NamedStyle>& namedStyles,
                      const DxfCatalog& dxfs,
                      bool strict);

// Parses the shared styles part into a StyleCatalog.
StyleCatalog parseStyleCatalog(const std::string& xml);

// Parses the differential-styles (<dxfs>) section into a list of styles.
std::vector<xlpp::Style> parseDifferentialStyles(const std::string& xml);

} // namespace internal
} // namespace xlpp
