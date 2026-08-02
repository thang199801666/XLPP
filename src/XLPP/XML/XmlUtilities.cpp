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
    if (tagName.empty()) return out;

    const std::size_t tnLen = tagName.size();
    const char* const data = xml.data();
    const char* const endPtr = data + xml.size();
    const char* p = data;

    for (;;) {
        while (p < endPtr && *p != '<') ++p;
        if (p >= endPtr) return out;
        if (static_cast<std::size_t>(endPtr - p) < tnLen + 2) return out;

        if (p[1] != '/') {
            bool nameMatch = true;
            for (std::size_t j = 0; j < tnLen; ++j) {
                if (p[1 + j] != tagName[j]) { nameMatch = false; break; }
            }
            if (nameMatch) {
                const char* after = p + 1 + tnLen;
                if (after < endPtr) {
                    char b = *after;
                    if (b == ' ' || b == '\t' || b == '\r' || b == '\n' || b == '>' || b == '/') {
                        const char* gt = after;
                        for (; gt < endPtr && *gt != '>'; ++gt) {}
                        if (gt >= endPtr) return out;

                        if (gt > after && *(gt - 1) == '/') {
                            out.emplace_back(p, static_cast<std::size_t>(gt - p + 1));
                            p = gt + 1;
                            continue;
                        }

                        const char* findStart = gt + 1;
                        for (;;) {
                            while (findStart < endPtr && !(findStart[0] == '<' && findStart + 1 < endPtr && findStart[1] == '/')) ++findStart;
                            if (findStart >= endPtr) return out;
                            if (static_cast<std::size_t>(endPtr - findStart) < tnLen + 3) return out;

                            bool closeMatch = true;
                            for (std::size_t j = 0; j < tnLen; ++j) {
                                if (findStart[2 + j] != tagName[j]) { closeMatch = false; break; }
                            }
                            if (closeMatch && findStart[2 + tnLen] == '>') {
                                out.emplace_back(p, static_cast<std::size_t>(findStart + tnLen + 3 - p));
                                p = findStart + tnLen + 3;
                                break;
                            }
                            ++findStart;
                        }
                        continue;
                    }
                }
            }
        }
        ++p;
    }
}

std::string tagText(std::string_view xml, std::string_view tagName) {
    auto v = tagTextView(xml, tagName);
    return v.empty() ? std::string{} : xmlUnescape(v);
}

} // namespace xlpp::internal
