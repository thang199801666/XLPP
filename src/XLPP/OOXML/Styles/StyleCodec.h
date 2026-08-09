#pragma once
#include <XLPP/Core/StableVector.h>

#include <XLPP/Styles/NamedStyle.h>
#include <XLPP/Styles/Style.h>

#include <cstddef>
#include <string>
#include <unordered_map>
#include <vector>

namespace xlpp::internal::ooxml {

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

std::string stylesXml(const StyleCatalog& catalog,
                      const xlpp::StableVector<xlpp::NamedStyle>& namedStyles,
                      const DxfCatalog& dxfs,
                      bool strict);
StyleCatalog parseStyleCatalog(const std::string& xml);
std::vector<xlpp::Style> parseDifferentialStyles(const std::string& xml);

} // namespace xlpp::internal::ooxml
