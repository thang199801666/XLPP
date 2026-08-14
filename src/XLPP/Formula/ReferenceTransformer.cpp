#include "ReferenceTransformer.h"
#include "XLPP/Worksheet/WorksheetName.h"
#include <XLPP/Cell/CellReference.h>
#include <algorithm>
#include <cctype>
#include <limits>
#include <sstream>
#include <vector>

namespace xlpp::internal {
namespace {
constexpr std::size_t kMaxExcelRow = 1048576;
constexpr std::size_t kMaxExcelColumn = 16384;

enum class RefKind { Cell, Row, Column };

struct CoordToken {
    RefKind kind{RefKind::Cell};
    std::size_t row{0};
    std::size_t column{0};
    bool absRow{false};
    bool absColumn{false};
    std::size_t length{0};
};

struct ParsedReference {
    std::string qualifierText;
    std::string qualifierSheet;
    std::string qualifierFirstSheet;
    std::string qualifierLastSheet;
    bool qualified{false};
    bool external{false};
    bool threeD{false};
    CoordToken first;
    std::optional<CoordToken> second;
    std::size_t length{0};
};

bool isIdentifierChar(char c) noexcept {
    const unsigned char u = static_cast<unsigned char>(c);
    return std::isalnum(u) != 0 || c == '_' || c == '.';
}

std::string unescapeSheetName(std::string_view token) {
    if (token.size() >= 2 && token.front() == '\'' && token.back() == '\'') {
        std::string result;
        result.reserve(token.size() - 2);
        for (std::size_t i = 1; i + 1 < token.size(); ++i) {
            if (token[i] == '\'' && i + 1 < token.size() - 1 && token[i + 1] == '\'') {
                result.push_back('\'');
                ++i;
            } else {
                result.push_back(token[i]);
            }
        }
        return result;
    }
    return std::string(token);
}

bool parseUnsigned(std::string_view text, std::size_t& value) noexcept {
    if (text.empty()) return false;
    std::size_t result = 0;
    for (char c : text) {
        if (c < '0' || c > '9') return false;
        const auto digit = static_cast<std::size_t>(c - '0');
        if (result > (std::numeric_limits<std::size_t>::max() - digit) / 10) return false;
        result = result * 10 + digit;
    }
    if (result == 0) return false;
    value = result;
    return true;
}

std::optional<CoordToken> parseCoordinate(std::string_view text, std::size_t pos, bool allowBare) {
    if (pos >= text.size()) return std::nullopt;
    const std::size_t begin = pos;
    bool leadingDollar = false;
    if (text[pos] == '$') { leadingDollar = true; ++pos; }
    if (pos >= text.size()) return std::nullopt;

    if (std::isdigit(static_cast<unsigned char>(text[pos])) != 0) {
        const std::size_t digitBegin = pos;
        while (pos < text.size() && std::isdigit(static_cast<unsigned char>(text[pos])) != 0) ++pos;
        if (!allowBare) return std::nullopt;
        std::size_t row = 0;
        if (!parseUnsigned(text.substr(digitBegin, pos - digitBegin), row) || row > kMaxExcelRow) return std::nullopt;
        return CoordToken{RefKind::Row, row, 0, leadingDollar, false, pos - begin};
    }

    if (std::isalpha(static_cast<unsigned char>(text[pos])) == 0) return std::nullopt;
    const std::size_t lettersBegin = pos;
    while (pos < text.size() && std::isalpha(static_cast<unsigned char>(text[pos])) != 0 && pos - lettersBegin < 3) ++pos;
    if (pos < text.size() && std::isalpha(static_cast<unsigned char>(text[pos])) != 0) return std::nullopt;
    const auto letters = text.substr(lettersBegin, pos - lettersBegin);
    std::size_t column = 0;
    try { column = CellReference::columnIndex(letters); } catch (...) { return std::nullopt; }
    if (column == 0 || column > kMaxExcelColumn) return std::nullopt;

    bool absRow = false;
    if (pos < text.size() && text[pos] == '$') { absRow = true; ++pos; }
    const std::size_t digitBegin = pos;
    while (pos < text.size() && std::isdigit(static_cast<unsigned char>(text[pos])) != 0) ++pos;
    if (digitBegin == pos) {
        if (!allowBare) return std::nullopt;
        return CoordToken{RefKind::Column, 0, column, false, leadingDollar, pos - begin};
    }
    std::size_t row = 0;
    if (!parseUnsigned(text.substr(digitBegin, pos - digitBegin), row) || row > kMaxExcelRow) return std::nullopt;
    return CoordToken{RefKind::Cell, row, column, absRow, leadingDollar, pos - begin};
}

std::optional<ParsedReference> parseReferenceAt(std::string_view text, std::size_t start) {
    if (start >= text.size()) return std::nullopt;
    ParsedReference result;
    std::size_t pos = start;

    // Optional worksheet qualifier. External workbook qualifiers are parsed so
    // the caller can deliberately leave them untouched.
    if (text[pos] == '\'') {
        const std::size_t qBegin = pos;
        ++pos;
        bool closed = false;
        while (pos < text.size()) {
            if (text[pos] != '\'') { ++pos; continue; }
            if (pos + 1 < text.size() && text[pos + 1] == '\'') { pos += 2; continue; }
            ++pos; closed = true; break;
        }
        if (closed && pos < text.size() && text[pos] == '!') {
            result.qualified = true;
            result.qualifierText = std::string(text.substr(qBegin, pos - qBegin + 1));
            const auto quoted = text.substr(qBegin, pos - qBegin);
            result.qualifierSheet = unescapeSheetName(quoted);
            result.external = result.qualifierSheet.find('[') != std::string::npos;
            if (!result.external) {
                const auto rangeSep = result.qualifierSheet.find(':');
                if (rangeSep != std::string::npos && rangeSep != 0 && rangeSep + 1 < result.qualifierSheet.size()) {
                    result.threeD = true;
                    result.qualifierFirstSheet = result.qualifierSheet.substr(0, rangeSep);
                    result.qualifierLastSheet = result.qualifierSheet.substr(rangeSep + 1);
                }
            }
            ++pos;
        } else {
            pos = start;
        }
    } else if (text[pos] == '[') {
        // [Book.xlsx]Sheet!A1 -- keep external references untouched.
        const auto bang = text.find('!', pos);
        if (bang != std::string_view::npos) {
            result.qualified = true;
            result.external = true;
            result.qualifierText = std::string(text.substr(pos, bang - pos + 1));
            result.qualifierSheet = std::string(text.substr(pos, bang - pos));
            pos = bang + 1;
        }
    } else if (std::isalpha(static_cast<unsigned char>(text[pos])) != 0 || text[pos] == '_') {
        const std::size_t qBegin = pos;
        while (pos < text.size() && isIdentifierChar(text[pos])) ++pos;
        const std::size_t firstEnd = pos;
        if (pos < text.size() && text[pos] == ':') {
            ++pos;
            const std::size_t secondBegin = pos;
            if (pos < text.size() && (std::isalpha(static_cast<unsigned char>(text[pos])) != 0 || text[pos] == '_'))
                while (pos < text.size() && isIdentifierChar(text[pos])) ++pos;
            if (secondBegin != pos && pos < text.size() && text[pos] == '!') {
                result.qualified = true;
                result.threeD = true;
                result.qualifierText = std::string(text.substr(qBegin, pos - qBegin + 1));
                result.qualifierFirstSheet = std::string(text.substr(qBegin, firstEnd - qBegin));
                result.qualifierLastSheet = std::string(text.substr(secondBegin, pos - secondBegin));
                result.qualifierSheet = result.qualifierFirstSheet + ":" + result.qualifierLastSheet;
                ++pos;
            } else {
                pos = start;
            }
        } else if (pos < text.size() && text[pos] == '!') {
            result.qualified = true;
            result.qualifierText = std::string(text.substr(qBegin, pos - qBegin + 1));
            result.qualifierSheet = std::string(text.substr(qBegin, pos - qBegin));
            ++pos;
        } else {
            pos = start;
        }
    }

    const auto first = parseCoordinate(text, pos, true);
    if (!first) return std::nullopt;
    result.first = *first;
    pos += first->length;

    if (pos < text.size() && text[pos] == ':') {
        const auto second = parseCoordinate(text, pos + 1, true);
        if (!second || second->kind != first->kind) return std::nullopt;
        result.second = *second;
        pos += 1 + second->length;
    } else if (first->kind != RefKind::Cell) {
        // Bare A or 1 is a name/number, not an A1 area.
        return std::nullopt;
    }

    // Avoid interpreting a function name such as LOG10( as cell LOG10.
    if (!result.qualified && !result.second && pos < text.size() && text[pos] == '(') return std::nullopt;
    result.length = pos - start;
    return result;
}

std::optional<std::pair<std::size_t, std::size_t>> transformInterval(std::size_t first,
                                                                    std::size_t last,
                                                                    const StructuralEditSpec& edit,
                                                                    std::size_t limit,
                                                                    bool& changed) noexcept {
    if (first > last) std::swap(first, last);
    const std::size_t cutBegin = edit.index;
    const std::size_t cutEnd = edit.index + edit.amount; // exclusive
    if (edit.action == StructuralAction::Insert) {
        if (edit.index <= first) {
            if (first > limit - std::min(edit.amount, limit)) {
                changed = true;
                return std::nullopt;
            }
            first += edit.amount;
            last = last > limit - std::min(edit.amount, limit) ? limit : last + edit.amount;
            changed = true;
        } else if (edit.index <= last) {
            last = last > limit - std::min(edit.amount, limit) ? limit : last + edit.amount;
            changed = true;
        }
        return first <= last && first <= limit ? std::optional<std::pair<std::size_t, std::size_t>>{{first, last}}
                                               : std::nullopt;
    }

    if (cutEnd <= first) {
        first -= edit.amount;
        last -= edit.amount;
        changed = true;
        return std::pair{first, last};
    }
    if (cutBegin > last) return std::pair{first, last};

    const bool hasBefore = first < cutBegin;
    const bool hasAfter = last >= cutEnd;
    if (!hasBefore && !hasAfter) { changed = true; return std::nullopt; }

    const std::size_t newFirst = hasBefore ? first : cutBegin;
    const std::size_t newLast = hasAfter ? last - edit.amount : cutBegin - 1;
    changed = true;
    if (newFirst == 0 || newFirst > newLast) return std::nullopt;
    return std::pair{newFirst, newLast};
}

std::string formatCoord(const CoordToken& token, std::size_t row, std::size_t column) {
    std::ostringstream out;
    if (token.kind == RefKind::Row) {
        if (token.absRow) out << '$';
        out << row;
    } else if (token.kind == RefKind::Column) {
        if (token.absColumn) out << '$';
        out << CellReference::columnName(column);
    } else {
        if (token.absColumn) out << '$';
        out << CellReference::columnName(column);
        if (token.absRow) out << '$';
        out << row;
    }
    return out.str();
}

struct TransformParsedResult {
    std::string text;
    bool changed{false};
    bool invalidated{false};
};

TransformParsedResult transformParsed(const ParsedReference& parsed, const StructuralEditSpec& edit) {
    TransformParsedResult result;
    const auto& a = parsed.first;
    const auto& b = parsed.second ? *parsed.second : parsed.first;

    if ((edit.axis == StructuralAxis::Row && a.kind == RefKind::Column) ||
        (edit.axis == StructuralAxis::Column && a.kind == RefKind::Row)) {
        // Whole columns are unaffected by row edits and vice versa.
        result.text = parsed.qualifierText + formatCoord(a, a.row, a.column);
        if (parsed.second) result.text += ":" + formatCoord(b, b.row, b.column);
        return result;
    }

    bool changed = false;
    if (edit.axis == StructuralAxis::Row) {
        const auto interval = transformInterval(a.row, b.row, edit, kMaxExcelRow, changed);
        if (!interval) {
            result.text = parsed.qualifierText + "#REF!";
            result.changed = true;
            result.invalidated = true;
            return result;
        }
        result.text = parsed.qualifierText + formatCoord(a, interval->first, a.column);
        if (parsed.second) result.text += ":" + formatCoord(b, interval->second, b.column);
    } else {
        const auto interval = transformInterval(a.column, b.column, edit, kMaxExcelColumn, changed);
        if (!interval) {
            result.text = parsed.qualifierText + "#REF!";
            result.changed = true;
            result.invalidated = true;
            return result;
        }
        result.text = parsed.qualifierText + formatCoord(a, a.row, interval->first);
        if (parsed.second) result.text += ":" + formatCoord(b, b.row, interval->second);
    }
    result.changed = changed;
    return result;
}

bool qualifierTargetsEdit(const ParsedReference& parsed,
                          std::string_view ownerSheetName,
                          const StructuralEditSpec& edit) {
    if (parsed.external || parsed.threeD) return false;
    if (parsed.qualified) return internal::worksheetNamesEquivalent(parsed.qualifierSheet, edit.targetSheetName);
    return !ownerSheetName.empty() && internal::worksheetNamesEquivalent(ownerSheetName, edit.targetSheetName);
}

std::string quoteSheetQualifier(std::string_view sheetName) {
    std::string out{"'"};
    out.reserve(sheetName.size() + 3);
    for (const char c : sheetName) {
        out.push_back(c);
        if (c == '\'') out.push_back('\'');
    }
    out += "'!";
    return out;
}

std::string quoteSheetRangeQualifier(std::string_view first, std::string_view last) {
    std::string out{"'"};
    out.reserve(first.size() + last.size() + 4);
    auto appendName = [&](std::string_view value) {
        for (const char c : value) {
            out.push_back(c);
            if (c == '\'') out.push_back('\'');
        }
    };
    appendName(first);
    out.push_back(':');
    appendName(last);
    out += "'!";
    return out;
}

bool threeDEndpointMatches(const ParsedReference& parsed, std::string_view sheetName) noexcept {
    return parsed.threeD &&
           (internal::worksheetNamesEquivalent(parsed.qualifierFirstSheet, sheetName) ||
            internal::worksheetNamesEquivalent(parsed.qualifierLastSheet, sheetName));
}

bool isFormulaBoundaryBefore(std::string_view text, std::size_t pos) noexcept;

ReferenceRewriteResult rewriteQualifiedWorksheetReferences(std::string_view formula,
                                                            std::string_view targetSheet,
                                                            std::string_view replacementQualifier,
                                                            bool invalidate) {
    ReferenceRewriteResult out;
    out.text.reserve(formula.size() + 16);
    bool inString = false;
    int bracketDepth = 0;
    std::size_t i = 0;
    while (i < formula.size()) {
        const char c = formula[i];
        if (c == '"') {
            out.text.push_back(c);
            if (inString && i + 1 < formula.size() && formula[i + 1] == '"') {
                out.text.push_back('"'); i += 2; continue;
            }
            inString = !inString; ++i; continue;
        }
        if (inString) { out.text.push_back(c); ++i; continue; }
        if (c == ']' && bracketDepth > 0) { --bracketDepth; out.text.push_back(c); ++i; continue; }
        if (bracketDepth > 0) { out.text.push_back(c); ++i; continue; }
        if (c == '[') {
            const auto external = parseReferenceAt(formula, i);
            if (external && external->external) {
                out.text.append(formula.substr(i, external->length));
                i += external->length;
                continue;
            }
            ++bracketDepth; out.text.push_back(c); ++i; continue;
        }
        if (!isFormulaBoundaryBefore(formula, i)) {
            out.text.push_back(c); ++i; continue;
        }
        const auto parsed = parseReferenceAt(formula, i);
        const bool badTrailingBoundary = parsed && i + parsed->length < formula.size() &&
            (isIdentifierChar(formula[i + parsed->length]) || formula[i + parsed->length] == '$');
        if (!parsed || badTrailingBoundary) {
            out.text.push_back(c); ++i; continue;
        }
        const bool directMatch = parsed->qualified && !parsed->threeD &&
                                 internal::worksheetNamesEquivalent(parsed->qualifierSheet, targetSheet);
        const bool threeDMatch = threeDEndpointMatches(*parsed, targetSheet);
        if (parsed->external || !parsed->qualified || (!directMatch && !threeDMatch)) {
            out.text.append(formula.substr(i, parsed->length));
            i += parsed->length;
            continue;
        }

        ++out.referencesRewritten;
        if (invalidate) {
            out.text += "#REF!";
            ++out.referencesInvalidated;
        } else if (threeDMatch) {
            auto firstSheet = parsed->qualifierFirstSheet;
            auto lastSheet = parsed->qualifierLastSheet;
            std::string replacement(replacementQualifier);
            if (replacement.size() >= 3 && replacement.front() == '\'' && replacement.ends_with("'!"))
                replacement = unescapeSheetName(replacement.substr(0, replacement.size() - 1));
            else if (!replacement.empty() && replacement.back() == '!')
                replacement.pop_back();
            if (internal::worksheetNamesEquivalent(firstSheet, targetSheet)) firstSheet = replacement;
            if (internal::worksheetNamesEquivalent(lastSheet, targetSheet)) lastSheet = replacement;
            out.text += quoteSheetRangeQualifier(firstSheet, lastSheet);
            const auto coordinateOffset = parsed->qualifierText.size();
            out.text.append(formula.substr(i + coordinateOffset, parsed->length - coordinateOffset));
        } else {
            out.text += replacementQualifier;
            const auto coordinateOffset = parsed->qualifierText.size();
            out.text.append(formula.substr(i + coordinateOffset, parsed->length - coordinateOffset));
        }
        i += parsed->length;
    }
    return out;
}

bool isFormulaBoundaryBefore(std::string_view text, std::size_t pos) noexcept {
    if (pos == 0) return true;
    const char c = text[pos - 1];
    return !isIdentifierChar(c) && c != '$';
}

} // namespace

ReferenceRewriteResult rewriteFormulaReferences(std::string_view formula,
                                                 std::string_view ownerSheetName,
                                                 const StructuralEditSpec& edit) {
    ReferenceRewriteResult out;
    out.text.reserve(formula.size() + 16);
    bool inString = false;
    int bracketDepth = 0;
    std::size_t i = 0;
    while (i < formula.size()) {
        const char c = formula[i];
        if (c == '"') {
            out.text.push_back(c);
            if (inString && i + 1 < formula.size() && formula[i + 1] == '"') {
                out.text.push_back('"'); i += 2; continue;
            }
            inString = !inString; ++i; continue;
        }
        if (inString) { out.text.push_back(c); ++i; continue; }
        if (c == ']' && bracketDepth > 0) { --bracketDepth; out.text.push_back(c); ++i; continue; }
        if (bracketDepth > 0) { out.text.push_back(c); ++i; continue; }
        if (c == '[') {
            // Distinguish an external workbook qualifier ([Book]Sheet!A1)
            // from a structured-reference bracket (Table[Column]).
            const auto external = parseReferenceAt(formula, i);
            if (external && external->external) {
                out.text.append(formula.substr(i, external->length));
                i += external->length;
                continue;
            }
            ++bracketDepth; out.text.push_back(c); ++i; continue;
        }
        if (!isFormulaBoundaryBefore(formula, i)) {
            out.text.push_back(c); ++i; continue;
        }

        const auto parsed = parseReferenceAt(formula, i);
        const bool badTrailingBoundary = parsed && i + parsed->length < formula.size() &&
                                         (isIdentifierChar(formula[i + parsed->length]) || formula[i + parsed->length] == '$');
        if (!parsed || badTrailingBoundary) {
            out.text.push_back(c); ++i; continue;
        }
        if (parsed->threeD && threeDEndpointMatches(*parsed, edit.targetSheetName)) {
            ++out.referencesSkippedUnsupported;
            out.text.append(formula.substr(i, parsed->length));
            i += parsed->length;
            continue;
        }
        if (!qualifierTargetsEdit(*parsed, ownerSheetName, edit)) {
            out.text.append(formula.substr(i, parsed->length));
            i += parsed->length;
            continue;
        }

        const auto transformed = transformParsed(*parsed, edit);
        if (transformed.changed) {
            ++out.referencesRewritten;
            if (transformed.invalidated) ++out.referencesInvalidated;
            out.text += transformed.text;
        } else {
            out.text.append(formula.substr(i, parsed->length));
        }
        i += parsed->length;
    }
    return out;
}

ReferenceRewriteResult renameWorksheetReferences(std::string_view formula,
                                                  std::string_view oldSheetName,
                                                  std::string_view newSheetName) {
    if (oldSheetName.empty() || oldSheetName == newSheetName) return {std::string(formula), 0, 0};
    return rewriteQualifiedWorksheetReferences(formula, oldSheetName, quoteSheetQualifier(newSheetName), false);
}

ReferenceRewriteResult invalidateWorksheetReferences(std::string_view formula,
                                                      std::string_view removedSheetName) {
    if (removedSheetName.empty()) return {std::string(formula), 0, 0};
    return rewriteQualifiedWorksheetReferences(formula, removedSheetName, {}, true);
}

ReferenceRewriteResult rewriteReferenceExpression(std::string_view reference,
                                                   std::string_view ownerSheetName,
                                                   const StructuralEditSpec& edit) {
    ReferenceRewriteResult out;
    const bool hasEquals = !reference.empty() && reference.front() == '=';
    const std::size_t start = hasEquals ? 1 : 0;
    const auto parsed = parseReferenceAt(reference, start);
    if (!parsed || start + parsed->length != reference.size() ||
        !qualifierTargetsEdit(*parsed, ownerSheetName, edit)) {
        out.text = std::string(reference);
        return out;
    }
    const auto transformed = transformParsed(*parsed, edit);
    out.text = hasEquals ? "=" : "";
    out.text += transformed.text;
    if (transformed.changed) {
        out.referencesRewritten = 1;
        out.referencesInvalidated = transformed.invalidated ? 1 : 0;
    }
    return out;
}

ReferenceRewriteResult rewriteReferenceList(std::string_view references,
                                             const StructuralEditSpec& edit) {
    ReferenceRewriteResult out;

    struct AreaToken {
        std::string separatorBefore;
        std::string text;
    };
    std::vector<AreaToken> tokens;
    std::string separator;
    std::string token;
    bool inSingleQuote = false;

    auto flushToken = [&]() {
        if (token.empty()) return;
        tokens.push_back({std::move(separator), std::move(token)});
        separator.clear();
        token.clear();
    };

    for (std::size_t i = 0; i < references.size(); ++i) {
        const char c = references[i];
        if (c == '\'') {
            token.push_back(c);
            if (inSingleQuote && i + 1 < references.size() && references[i + 1] == '\'') {
                token.push_back('\'');
                ++i;
            } else {
                inSingleQuote = !inSingleQuote;
            }
            continue;
        }
        if (!inSingleQuote && (std::isspace(static_cast<unsigned char>(c)) || c == ',')) {
            flushToken();
            separator.push_back(c);
            continue;
        }
        token.push_back(c);
    }
    flushToken();

    bool wroteArea = false;
    for (const auto& item : tokens) {
        const auto parsed = parseReferenceAt(item.text, 0);
        if (!parsed || parsed->length != item.text.size()) {
            if (wroteArea) out.text += item.separatorBefore.empty() ? " " : item.separatorBefore;
            out.text += item.text;
            wroteArea = true;
            continue;
        }

        // sqref/print-area expressions passed here are local to the edited
        // worksheet even if callers happened to retain a local qualifier.
        ParsedReference local = *parsed;
        local.qualified = false;
        local.qualifierText.clear();
        local.qualifierSheet.clear();
        local.external = false;
        const auto transformed = transformParsed(local, edit);
        if (transformed.changed) ++out.referencesRewritten;
        if (transformed.invalidated) {
            ++out.referencesInvalidated;
            continue;
        }
        if (wroteArea) out.text += item.separatorBefore.empty() ? " " : item.separatorBefore;
        out.text += transformed.text;
        wroteArea = true;
    }
    return out;
}

std::size_t transformPosition(std::size_t coordinate, const StructuralEditSpec& edit) noexcept {
    const std::size_t limit = edit.axis == StructuralAxis::Row ? kMaxExcelRow : kMaxExcelColumn;
    if (edit.action == StructuralAction::Insert) {
        if (coordinate < edit.index) return coordinate;
        if (edit.amount >= limit || coordinate > limit - edit.amount) return limit;
        return coordinate + edit.amount;
    }
    const auto end = edit.index + edit.amount;
    if (coordinate >= end) return coordinate - edit.amount;
    if (coordinate >= edit.index) return std::min(edit.index, limit);
    return coordinate;
}

void transformDrawingAnchor(DrawingAnchorInfo& anchor, const StructuralEditSpec& edit) noexcept {
    if (anchor.type == DrawingAnchorType::Absolute) return;
    if (edit.axis == StructuralAxis::Row) {
        anchor.from.row = transformPosition(anchor.from.row, edit);
        if (anchor.type == DrawingAnchorType::TwoCell) anchor.to.row = transformPosition(anchor.to.row, edit);
    } else {
        anchor.from.column = transformPosition(anchor.from.column, edit);
        if (anchor.type == DrawingAnchorType::TwoCell) anchor.to.column = transformPosition(anchor.to.column, edit);
    }
}

} // namespace xlpp::internal
