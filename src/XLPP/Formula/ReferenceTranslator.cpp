#include <XLPP/Formula/ReferenceTranslator.h>
#include <XLPP/Cell/CellReference.h>

#include <algorithm>
#include <cctype>
#include <limits>
#include <optional>
#include <string>
#include <string_view>

namespace xlpp {
namespace {
constexpr std::size_t kMaxExcelRow = 1048576;
constexpr std::size_t kMaxExcelColumn = 16384;

bool iequals(std::string_view a, std::string_view b) {
    if (a.size() != b.size()) return false;
    for (std::size_t i = 0; i < a.size(); ++i) {
        if (std::tolower(static_cast<unsigned char>(a[i])) !=
            std::tolower(static_cast<unsigned char>(b[i]))) return false;
    }
    return true;
}

bool nameChar(char c) {
    const auto u = static_cast<unsigned char>(c);
    return std::isalnum(u) || c == '_' || c == '.' || c == '\\';
}

struct SheetPrefix {
    std::string raw;
    std::string decoded;
    bool present{false};
    bool external{false};
    std::size_t end{0};
};

std::optional<SheetPrefix> parseSheetPrefix(std::string_view text, std::size_t pos) {
    if (pos >= text.size()) return SheetPrefix{};
    if (text[pos] == '\'') {
        std::string decoded;
        std::size_t i = pos + 1;
        while (i < text.size()) {
            if (text[i] == '\'') {
                if (i + 1 < text.size() && text[i + 1] == '\'') {
                    decoded.push_back('\'');
                    i += 2;
                    continue;
                }
                if (i + 1 < text.size() && text[i + 1] == '!') {
                    SheetPrefix result;
                    result.raw = std::string(text.substr(pos, i + 2 - pos));
                    result.decoded = std::move(decoded);
                    result.present = true;
                    result.external = result.decoded.find('[') != std::string::npos;
                    result.end = i + 2;
                    return result;
                }
                return std::nullopt;
            }
            decoded.push_back(text[i++]);
        }
        return std::nullopt;
    }

    // Only treat an unquoted prefix as a sheet qualifier if an exclamation
    // mark is actually present. This avoids confusing function/name tokens.
    std::size_t i = pos;
    while (i < text.size()) {
        const char c = text[i];
        if (c == '!') {
            if (i == pos) return std::nullopt;
            SheetPrefix result;
            result.raw = std::string(text.substr(pos, i + 1 - pos));
            result.decoded = std::string(text.substr(pos, i - pos));
            result.present = true;
            result.external = result.decoded.find('[') != std::string::npos;
            result.end = i + 1;
            return result;
        }
        if (!(nameChar(c) || c == ' ' || c == '-' || c == ':')) break;
        ++i;
    }
    return SheetPrefix{};
}

struct CellRefToken {
    std::size_t row{0};
    std::size_t column{0};
    bool absRow{false};
    bool absColumn{false};
    std::size_t end{0};
};

std::optional<CellRefToken> parseCell(std::string_view text, std::size_t pos) {
    CellRefToken result;
    std::size_t i = pos;
    if (i < text.size() && text[i] == '$') { result.absColumn = true; ++i; }
    const std::size_t letters = i;
    std::string col;
    while (i < text.size() && std::isalpha(static_cast<unsigned char>(text[i])) && col.size() < 4) {
        col.push_back(text[i++]);
    }
    if (i == letters || col.size() > 3) return std::nullopt;
    try { result.column = CellReference::columnIndex(col); }
    catch (...) { return std::nullopt; }
    if (result.column == 0 || result.column > kMaxExcelColumn) return std::nullopt;
    if (i < text.size() && text[i] == '$') { result.absRow = true; ++i; }
    const std::size_t digits = i;
    std::size_t row = 0;
    while (i < text.size() && std::isdigit(static_cast<unsigned char>(text[i]))) {
        const unsigned d = static_cast<unsigned>(text[i] - '0');
        if (row > (std::numeric_limits<std::size_t>::max() - d) / 10) return std::nullopt;
        row = row * 10 + d;
        ++i;
    }
    if (i == digits || row == 0 || row > kMaxExcelRow) return std::nullopt;
    result.row = row;
    result.end = i;
    return result;
}

std::string cellText(const CellRefToken& value) {
    std::string out;
    if (value.absColumn) out.push_back('$');
    out += CellReference::columnName(value.column);
    if (value.absRow) out.push_back('$');
    out += std::to_string(value.row);
    return out;
}

struct IntervalResult {
    std::size_t first{0};
    std::size_t last{0};
    bool invalid{false};
    bool changed{false};
};

IntervalResult translateInterval(std::size_t first, std::size_t last,
                                 std::size_t index, std::size_t amount, bool insert) {
    if (first > last) std::swap(first, last);
    IntervalResult r{first, last, false, false};
    if (insert) {
        if (index <= first) {
            r.first += amount;
            r.last += amount;
            r.changed = true;
        } else if (index <= last) {
            r.last += amount;
            r.changed = true;
        }
        return r;
    }

    const std::size_t deleteFirst = index;
    const std::size_t deleteLast = index + amount - 1;
    if (last < deleteFirst) return r;
    if (first > deleteLast) {
        r.first -= amount;
        r.last -= amount;
        r.changed = true;
        return r;
    }

    // Overlap: keep the surviving cells on either side. Structural deletion
    // makes the survivors contiguous, so the resulting interval is singular.
    const bool hasBefore = first < deleteFirst;
    const bool hasAfter = last > deleteLast;
    if (!hasBefore && !hasAfter) {
        r.invalid = true;
        r.changed = true;
        return r;
    }
    r.first = hasBefore ? first : deleteFirst;
    r.last = hasAfter ? last - amount : deleteFirst - 1;
    if (r.first > r.last) r.invalid = true;
    r.changed = true;
    return r;
}

struct TokenTranslation {
    std::string text;
    std::size_t end{0};
    bool matched{false};
    bool changed{false};
    bool invalid{false};
};

bool appliesTo(const SheetPrefix& prefix, std::string_view contextSheet, const StructuralEdit& edit) {
    if (prefix.external) return false;
    if (prefix.present) return iequals(prefix.decoded, edit.sheetName);
    return iequals(contextSheet, edit.sheetName);
}

TokenTranslation parseCellReferenceToken(std::string_view text, std::size_t pos,
                                         std::string_view contextSheet, const StructuralEdit& edit) {
    TokenTranslation result;
    const std::size_t start = pos;
    auto prefixOpt = parseSheetPrefix(text, pos);
    if (!prefixOpt) return result;
    SheetPrefix prefix = *prefixOpt;
    if (prefix.present) pos = prefix.end;

    auto first = parseCell(text, pos);
    if (!first) return result;
    const std::size_t firstEnd = first->end;

    // LOG10( and similar names can look exactly like a valid cell address.
    if (!prefix.present && firstEnd < text.size() && text[firstEnd] == '(') return result;
    if (start > 0 && !prefix.present && nameChar(text[start - 1])) return result;

    std::optional<CellRefToken> second;
    std::size_t end = firstEnd;
    if (end < text.size() && text[end] == ':') {
        second = parseCell(text, end + 1);
        if (!second) return result;
        end = second->end;
    }
    if (end < text.size() && nameChar(text[end])) return result;

    result.matched = true;
    result.end = end;
    result.text = std::string(text.substr(start, end - start));
    if (!appliesTo(prefix, contextSheet, edit)) return result;

    CellRefToken a = *first;
    CellRefToken b = second.value_or(a);
    if (edit.kind == StructuralEditKind::InsertRows || edit.kind == StructuralEditKind::DeleteRows) {
        const auto tr = translateInterval(a.row, b.row, edit.index, edit.amount,
                                          edit.kind == StructuralEditKind::InsertRows);
        if (tr.invalid) {
            result.text = prefix.raw + "#REF!";
            result.invalid = result.changed = true;
            return result;
        }
        if (tr.changed) {
            const bool ascending = a.row <= b.row;
            a.row = ascending ? tr.first : tr.last;
            b.row = ascending ? tr.last : tr.first;
            result.changed = true;
        }
    } else {
        const auto tr = translateInterval(a.column, b.column, edit.index, edit.amount,
                                          edit.kind == StructuralEditKind::InsertColumns);
        if (tr.invalid) {
            result.text = prefix.raw + "#REF!";
            result.invalid = result.changed = true;
            return result;
        }
        if (tr.changed) {
            const bool ascending = a.column <= b.column;
            a.column = ascending ? tr.first : tr.last;
            b.column = ascending ? tr.last : tr.first;
            result.changed = true;
        }
    }
    if (result.changed) {
        result.text = prefix.raw + cellText(a);
        if (second) result.text += ":" + cellText(b);
    }
    return result;
}

struct AxisToken {
    std::size_t value{0};
    bool absolute{false};
    std::size_t end{0};
};

std::optional<AxisToken> parseRowOnly(std::string_view text, std::size_t pos) {
    AxisToken r;
    std::size_t i = pos;
    if (i < text.size() && text[i] == '$') { r.absolute = true; ++i; }
    const std::size_t begin = i;
    while (i < text.size() && std::isdigit(static_cast<unsigned char>(text[i]))) {
        r.value = r.value * 10 + static_cast<std::size_t>(text[i] - '0');
        ++i;
    }
    if (i == begin || r.value == 0 || r.value > kMaxExcelRow) return std::nullopt;
    r.end = i;
    return r;
}

std::optional<AxisToken> parseColumnOnly(std::string_view text, std::size_t pos) {
    AxisToken r;
    std::size_t i = pos;
    if (i < text.size() && text[i] == '$') { r.absolute = true; ++i; }
    std::string col;
    while (i < text.size() && std::isalpha(static_cast<unsigned char>(text[i])) && col.size() < 4)
        col.push_back(text[i++]);
    if (col.empty() || col.size() > 3) return std::nullopt;
    try { r.value = CellReference::columnIndex(col); } catch (...) { return std::nullopt; }
    if (r.value == 0 || r.value > kMaxExcelColumn) return std::nullopt;
    r.end = i;
    return r;
}

std::string axisText(const AxisToken& a, bool column) {
    std::string s;
    if (a.absolute) s.push_back('$');
    s += column ? CellReference::columnName(a.value) : std::to_string(a.value);
    return s;
}

TokenTranslation parseWholeAxisToken(std::string_view text, std::size_t pos,
                                     std::string_view contextSheet, const StructuralEdit& edit,
                                     bool column) {
    TokenTranslation result;
    const std::size_t start = pos;
    auto prefixOpt = parseSheetPrefix(text, pos);
    if (!prefixOpt) return result;
    SheetPrefix prefix = *prefixOpt;
    if (prefix.present) pos = prefix.end;
    auto a = column ? parseColumnOnly(text, pos) : parseRowOnly(text, pos);
    if (!a || a->end >= text.size() || text[a->end] != ':') return result;
    auto b = column ? parseColumnOnly(text, a->end + 1) : parseRowOnly(text, a->end + 1);
    if (!b) return result;
    const std::size_t end = b->end;
    if (end < text.size() && nameChar(text[end])) return result;
    if (start > 0 && !prefix.present && nameChar(text[start - 1])) return result;
    result.matched = true;
    result.end = end;
    result.text = std::string(text.substr(start, end - start));
    if (!appliesTo(prefix, contextSheet, edit)) return result;
    const bool axisAffected = column
        ? (edit.kind == StructuralEditKind::InsertColumns || edit.kind == StructuralEditKind::DeleteColumns)
        : (edit.kind == StructuralEditKind::InsertRows || edit.kind == StructuralEditKind::DeleteRows);
    if (!axisAffected) return result;
    const bool insert = edit.kind == StructuralEditKind::InsertColumns || edit.kind == StructuralEditKind::InsertRows;
    const auto tr = translateInterval(a->value, b->value, edit.index, edit.amount, insert);
    if (tr.invalid) {
        result.text = prefix.raw + "#REF!";
        result.invalid = result.changed = true;
        return result;
    }
    if (tr.changed) {
        const bool ascending = a->value <= b->value;
        a->value = ascending ? tr.first : tr.last;
        b->value = ascending ? tr.last : tr.first;
        result.text = prefix.raw + axisText(*a, column) + ":" + axisText(*b, column);
        result.changed = true;
    }
    return result;
}



std::string quotedSheetQualifier(std::string_view decoded) {
    std::string out{"'"};
    out.reserve(decoded.size() + 3);
    for (const char c : decoded) {
        if (c == '\'') out += "''";
        else out.push_back(c);
    }
    out += "'!";
    return out;
}

bool renameDecodedQualifier(std::string& decoded, std::string_view oldName,
                            std::string_view newName) {
    bool changed = false;
    std::string rewritten;
    rewritten.reserve(decoded.size() + newName.size());
    std::size_t begin = 0;
    for (;;) {
        const auto colon = decoded.find(':', begin);
        const auto end = colon == std::string::npos ? decoded.size() : colon;
        const auto component = std::string_view(decoded).substr(begin, end - begin);
        if (iequals(component, oldName)) {
            rewritten.append(newName);
            changed = true;
        } else {
            rewritten.append(component);
        }
        if (colon == std::string::npos) break;
        rewritten.push_back(':');
        begin = colon + 1;
    }
    if (changed) decoded = std::move(rewritten);
    return changed;
}

ReferenceTranslationResult translateImpl(std::string_view input, std::string_view contextSheet,
                                         const StructuralEdit& edit, bool rangeMode) {
    ReferenceTranslationResult report;
    report.value.reserve(input.size() + 16);
    std::size_t i = 0;
    while (i < input.size()) {
        if (!rangeMode && input[i] == '"') {
            const std::size_t begin = i++;
            while (i < input.size()) {
                if (input[i] == '"') {
                    if (i + 1 < input.size() && input[i + 1] == '"') { i += 2; continue; }
                    ++i;
                    break;
                }
                ++i;
            }
            report.value.append(input.substr(begin, i - begin));
            continue;
        }

        TokenTranslation token;
        if (rangeMode) {
            token = parseWholeAxisToken(input, i, contextSheet, edit, true);
            if (!token.matched) token = parseWholeAxisToken(input, i, contextSheet, edit, false);
        }
        if (!token.matched) token = parseCellReferenceToken(input, i, contextSheet, edit);
        if (token.matched) {
            ++report.referencesVisited;
            if (token.invalid) ++report.referencesInvalidated;
            else if (token.changed) ++report.referencesChanged;
            report.value += token.text;
            i = token.end;
            continue;
        }
        report.value.push_back(input[i++]);
    }
    return report;
}
} // namespace

ReferenceTranslationResult translateFormulaReferences(std::string_view formula,
                                                       std::string_view contextSheet,
                                                       const StructuralEdit& edit) {
    if (edit.sheetName.empty() || edit.index == 0 || edit.amount == 0)
        return {std::string(formula), 0, 0, 0};
    return translateImpl(formula, contextSheet, edit, false);
}

ReferenceTranslationResult translateRangeReferences(std::string_view reference,
                                                     std::string_view contextSheet,
                                                     const StructuralEdit& edit) {
    if (edit.sheetName.empty() || edit.index == 0 || edit.amount == 0)
        return {std::string(reference), 0, 0, 0};
    return translateImpl(reference, contextSheet, edit, true);
}


ReferenceTranslationResult renameWorksheetReferences(std::string_view expression,
                                                      std::string_view oldWorksheetName,
                                                      std::string_view newWorksheetName) {
    ReferenceTranslationResult report;
    report.value.reserve(expression.size() + 16);
    if (oldWorksheetName.empty() || newWorksheetName.empty() ||
        oldWorksheetName == newWorksheetName) {
        report.value.assign(expression);
        return report;
    }

    std::size_t i = 0;
    while (i < expression.size()) {
        // Excel formula string literals use doubled quotes. Sheet-looking text
        // inside them is data and must never be rewritten.
        if (expression[i] == '"') {
            const auto begin = i++;
            while (i < expression.size()) {
                if (expression[i] == '"') {
                    if (i + 1 < expression.size() && expression[i + 1] == '"') {
                        i += 2;
                        continue;
                    }
                    ++i;
                    break;
                }
                ++i;
            }
            report.value.append(expression.substr(begin, i - begin));
            continue;
        }

        const auto prefix = parseSheetPrefix(expression, i);
        if (prefix && prefix->present) {
            ++report.referencesVisited;
            if (!prefix->external) {
                auto decoded = prefix->decoded;
                if (renameDecodedQualifier(decoded, oldWorksheetName, newWorksheetName)) {
                    report.value += quotedSheetQualifier(decoded);
                    ++report.referencesChanged;
                    i = prefix->end;
                    continue;
                }
            }
            report.value += prefix->raw;
            i = prefix->end;
            continue;
        }
        report.value.push_back(expression[i++]);
    }
    return report;
}


ReferenceTranslationResult invalidateWorksheetReferences(std::string_view expression,
                                                          std::string_view removedWorksheetName) {
    ReferenceTranslationResult report;
    report.value.reserve(expression.size());
    if (removedWorksheetName.empty()) {
        report.value.assign(expression);
        return report;
    }

    std::size_t i = 0;
    while (i < expression.size()) {
        if (expression[i] == '"') {
            const auto begin = i++;
            while (i < expression.size()) {
                if (expression[i] == '"') {
                    if (i + 1 < expression.size() && expression[i + 1] == '"') {
                        i += 2;
                        continue;
                    }
                    ++i;
                    break;
                }
                ++i;
            }
            report.value.append(expression.substr(begin, i - begin));
            continue;
        }
        const auto prefix = parseSheetPrefix(expression, i);
        if (prefix && prefix->present) {
            ++report.referencesVisited;
            bool targetsRemovedSheet = false;
            if (!prefix->external) {
                std::size_t begin = 0;
                for (;;) {
                    const auto colon = prefix->decoded.find(':', begin);
                    const auto end = colon == std::string::npos ? prefix->decoded.size() : colon;
                    if (iequals(std::string_view(prefix->decoded).substr(begin, end - begin), removedWorksheetName)) {
                        targetsRemovedSheet = true;
                        break;
                    }
                    if (colon == std::string::npos) break;
                    begin = colon + 1;
                }
            }
            if (targetsRemovedSheet) {
                report.value += "#REF!";
                ++report.referencesInvalidated;
            } else {
                report.value += prefix->raw;
            }
            i = prefix->end;
            continue;
        }
        report.value.push_back(expression[i++]);
    }
    return report;
}

} // namespace xlpp
