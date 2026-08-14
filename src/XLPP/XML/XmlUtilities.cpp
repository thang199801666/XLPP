#include "XmlUtilities.h"
#include "SimdScan.h"
#include <algorithm>

namespace xlpp::internal {

std::string xmlEscape(std::string_view s) {
    std::string out; out.reserve(s.size());
    for(char c:s) switch(c){case '&':out+="&amp;";break;case '<':out+="&lt;";break;case '>':out+="&gt;";break;case '"':out+="&quot;";break;case '\'':out+="&apos;";break;default:out+=c;}
    return out;
}

void writeXmlEscaped(std::ostringstream& out, std::string_view s) {
    std::size_t start = 0;
    const std::string_view specials = "&<>\"'";
    for (;;) {
        const std::size_t pos = s.find_first_of(specials, start);
        if (pos == std::string_view::npos) {
            if (start < s.size()) out.write(s.data() + start, static_cast<std::streamsize>(s.size() - start));
            return;
        }
        if (pos > start) out.write(s.data() + start, static_cast<std::streamsize>(pos - start));
        switch (s[pos]) {
            case '&': out << "&amp;"; break;
            case '<': out << "&lt;"; break;
            case '>': out << "&gt;"; break;
            case '"': out << "&quot;"; break;
            case '\'': out << "&apos;"; break;
        }
        start = pos + 1;
    }
}

std::string xmlUnescape(std::string_view s) {
    std::string out(s);
    const std::pair<const char*,const char*> r[]={{"&amp;","&"},{"&lt;","<"},{"&gt;",">"},{"&quot;","\""},{"&apos;","'"}};
    for(auto [a,b]:r){size_t p=0; while((p=out.find(a,p))!=std::string::npos){out.replace(p,std::char_traits<char>::length(a),b);p+=std::char_traits<char>::length(b);}}
    // XML 1.0 §2.11 line-ending normalization: CRLF and bare CR are read as LF.
    // xlsx producers (e.g. openpyxl) write \r\n in element text.
    for (std::size_t i = 0; i < out.size();) {
        if (out[i] == '\r') {
            if (i + 1 < out.size() && out[i + 1] == '\n') out.erase(i, 1);
            else out[i] = '\n';
        }
        ++i;
    }
    return out;
}

std::string_view attrView(std::string_view tag, std::string_view name) {
    return simd::xmlAttribute(tag, name);
}

std::string attribute(std::string_view tag, std::string_view name) {
    auto v = attrView(tag, name);
    return v.empty() ? std::string{} : xmlUnescape(v);
}

std::vector<std::string> tags(std::string_view xml, std::string_view tagName) {
    std::vector<std::string> out;
    tagsForEach(xml, tagName, [&](std::string_view element) { out.emplace_back(element); });
    return out;
}

std::string tagText(std::string_view xml, std::string_view tagName) {
    auto v = tagTextView(xml, tagName);
    return v.empty() ? std::string{} : xmlUnescape(v);
}

} // namespace xlpp::internal
