#include <XLPP/Formula/DependencyGraph.h>
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Cell/Cell.h>
#include <cctype>
#include <string>

namespace xlpp {

namespace {

// Scans a formula for A1 references, defined-name symbols, table references
// and external-workbook references. String literals and function names are
// skipped. Returns tokens in source order.
struct ReferenceToken {
    FormulaDependencyKind kind{FormulaDependencyKind::CellOrRange};
    std::string sheet;       // explicit sheet qualifier ("" = context sheet)
    std::string reference;   // A1-style cell/range reference
    std::string symbol;      // defined-name / table name / external book
};

bool isColLetter(char c) { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'); }

// Reads a maximal run of [A-Za-z] starting at pos.
std::size_t readWord(std::string_view text, std::size_t pos) {
    std::size_t end = pos;
    while (end < text.size() && isColLetter(text[end])) ++end;
    return end;
}

// Reads a maximal run of [0-9] starting at pos.
std::size_t readDigits(std::string_view text, std::size_t pos) {
    std::size_t end = pos;
    while (end < text.size() && std::isdigit(static_cast<unsigned char>(text[end]))) ++end;
    return end;
}

bool isReferenceBoundaryAfter(std::string_view text, std::size_t end) {
    if (end >= text.size()) return true;
    const char c = text[end];
    return !isColLetter(c) && !std::isdigit(static_cast<unsigned char>(c)) && c != '_' && c != '.';
}

// Attempts to parse an A1 cell/range reference starting at pos (already on a
// letter or digit). Returns (reference, nextPos) or empty on failure.
std::pair<std::string, std::size_t> parseCellRef(std::string_view text, std::size_t pos) {
    const auto start = pos;
    if (pos < text.size() && text[pos] == '$') ++pos;
    const auto colEnd = readWord(text, pos);
    // Excel columns are at most three letters (XFD); longer letter runs are
    // sheet names, function names or defined-name symbols, not columns.
    if (colEnd - pos > 3) return {};
    const bool hasColumn = colEnd > pos;
    pos = colEnd;
    if (pos < text.size() && text[pos] == '$') ++pos;
    const auto rowEnd = readDigits(text, pos);
    const bool hasRow = rowEnd > pos;
    if (!hasColumn && !hasRow) return {};
    if (hasColumn && !hasRow) {
        // Column-only range A:B or single column A.
        if (colEnd + 1 < text.size() && text[colEnd] == ':' && isColLetter(text[colEnd + 1])) {
            pos = colEnd + 1;
            if (pos < text.size() && text[pos] == '$') ++pos;
            const auto end2 = readWord(text, pos);
            if (end2 > pos) return {std::string(text.substr(start, end2 - start)), end2};
        }
        return {std::string(text.substr(start, colEnd - start)), colEnd};
    }
    if (!hasColumn) {
        // Row-only range 1:5 or single row.
        if (rowEnd + 1 < text.size() && text[rowEnd] == ':' && std::isdigit(static_cast<unsigned char>(text[rowEnd + 1]))) {
            pos = rowEnd + 1;
            const auto end2 = readDigits(text, pos);
            if (end2 > pos) return {std::string(text.substr(start, end2 - start)), end2};
        }
        return {std::string(text.substr(start, rowEnd - start)), rowEnd};
    }
    // Cell, optionally followed by :cell for a range.
    auto next = rowEnd;
    if (next < text.size() && text[next] == ':') {
        std::size_t afterColon = next + 1;
        if (afterColon < text.size() && text[afterColon] == '$') ++afterColon;
        const auto c2 = readWord(text, afterColon);
        if (c2 > afterColon) {
            std::size_t r2pos = c2;
            if (r2pos < text.size() && text[r2pos] == '$') ++r2pos;
            const auto r2 = readDigits(text, r2pos);
            if (r2 > r2pos) {
                return {std::string(text.substr(start, r2 - start)), r2};
            }
        }
    }
    return {std::string(text.substr(start, rowEnd - start)), rowEnd};
}

// Scans formula text and appends reference tokens.
void scanFormula(std::string_view formula, std::string_view contextSheet,
                 std::vector<ReferenceToken>& out) {
    std::size_t i = 0;
    while (i < formula.size()) {
        const char c = formula[i];
        if (c == '"') {
            // Skip string literal.
            ++i;
            while (i < formula.size() && formula[i] != '"') {
                if (formula[i] == '"' && i + 1 < formula.size() && formula[i + 1] == '"') { i += 2; continue; }
                ++i;
            }
            ++i;
            continue;
        }
        if (c == '\'') {
            // Skip quoted sheet qualifier (may be followed by '!').
            ++i;
            std::string sheet;
            while (i < formula.size()) {
                if (formula[i] == '\'' && i + 1 < formula.size() && formula[i + 1] == '\'') { sheet.push_back('\''); i += 2; continue; }
                if (formula[i] == '\'') { ++i; break; }
                sheet.push_back(formula[i]);
                ++i;
            }
            std::string ref;
            std::size_t after = i;
            if (i < formula.size() && formula[i] == '!') {
                ++i;
                auto [r, np] = parseCellRef(formula, i);
                if (!r.empty()) { ref = r; after = np; }
            }
            if (!ref.empty()) {
                out.push_back({FormulaDependencyKind::CellOrRange, sheet, std::move(ref), {}});
                i = after;
                continue;
            }
            continue;
        }
        if (c == '[') {
            // External workbook [book] or table ref [Col].
            const auto close = formula.find(']', i);
            if (close != std::string_view::npos) {
                const auto inner = formula.substr(i + 1, close - i - 1);
                out.push_back({FormulaDependencyKind::ExternalReference, {}, {}, std::string(inner)});
                i = close + 1;
                // Check for !sheet!ref external chain; record the trailing ref if present.
                if (i < formula.size() && formula[i] == '!') {
                    ++i;
                    if (i < formula.size() && formula[i] == '\'') { ++i; while (i < formula.size() && formula[i] != '\'') ++i; ++i; }
                    if (i < formula.size() && formula[i] == '!') {
                        ++i;
                        auto [r, np] = parseCellRef(formula, i);
                        if (!r.empty()) i = np;
                    }
                }
                continue;
            }
        }
        if (isColLetter(c) || (std::isdigit(static_cast<unsigned char>(c)))) {
            // Try A1 cell/range reference.
            auto [ref, next] = parseCellRef(formula, i);
            if (!ref.empty() && isReferenceBoundaryAfter(formula, next)) {
                out.push_back({FormulaDependencyKind::CellOrRange, std::string(contextSheet), std::move(ref), {}});
                i = next;
                continue;
            }
            // Not a reference: could be a function name, defined name, or table ref.
            if (isColLetter(c)) {
                auto wordEnd = readWord(formula, i);
                const std::string word(formula.substr(i, wordEnd - i));
                // Function call -> skip. Table column refs contain '[' handled above.
                if (wordEnd < formula.size() && formula[wordEnd] == '(') {
                    // function name
                    i = wordEnd + 1;
                    continue;
                }
                // Plain symbol -> defined name (or trailing "()" empty).
                if (wordEnd < formula.size() && formula[wordEnd] == '!') {
                    // Sheet!A1 — qualifier already handled for quoted; handle bare sheet names.
                    std::string sheetQualifier = word;
                    ++wordEnd;
                    auto [r, np] = parseCellRef(formula, wordEnd);
                    if (!r.empty()) {
                        out.push_back({FormulaDependencyKind::CellOrRange, sheetQualifier, std::move(r), {}});
                        i = np;
                        continue;
                    }
                }
                out.push_back({FormulaDependencyKind::DefinedName, std::string(contextSheet), {}, word});
                i = wordEnd;
                continue;
            }
        }
        ++i;
    }
}

} // namespace

FormulaDependencyGraph buildFormulaDependencyGraph(const Workbook& workbook) {
    FormulaDependencyGraph graph;
    std::vector<FormulaDependency> edges;
    const auto sheetNames = workbook.sheetNames();
    for (const auto& sheetName : sheetNames) {
        const auto* sheet = workbook.worksheet(sheetName);
        if (!sheet) continue;
        for (const auto& [key, cell] : sheet->cells()) {
            if (!cell.hasFormula()) continue;
            std::vector<ReferenceToken> tokens;
            scanFormula(cell.formula(), sheetName, tokens);
            for (auto& token : tokens) {
                FormulaDependency dep;
                dep.dependentSheet = sheetName;
                dep.dependentCell = cell.address();
                dep.kind = token.kind;
                dep.precedentSheet = token.sheet.empty() ? sheetName : token.sheet;
                dep.precedentReference = std::move(token.reference);
                dep.symbol = std::move(token.symbol);
                edges.push_back(std::move(dep));
            }
            graph.report_.formulaCells++;
        }
    }
    graph.edges_ = std::move(edges);
    for (const auto& dep : graph.edges_) {
        graph.report_.edges++;
        switch (dep.kind) {
            case FormulaDependencyKind::CellOrRange: graph.report_.cellOrRangeEdges++; break;
            case FormulaDependencyKind::DefinedName: graph.report_.definedNameEdges++; break;
            case FormulaDependencyKind::Table: graph.report_.tableEdges++; break;
            case FormulaDependencyKind::ExternalReference: graph.report_.externalEdges++; break;
            case FormulaDependencyKind::VolatileReference: graph.report_.volatileReferences++; break;
        }
    }
    return graph;
}

std::vector<FormulaDependency> FormulaDependencyGraph::precedentsOf(const std::string& sheet, const std::string& cell) const {
    std::vector<FormulaDependency> result;
    for (const auto& dep : edges_)
        if (dep.dependentSheet == sheet && dep.dependentCell == cell) result.push_back(dep);
    return result;
}

std::vector<FormulaDependency> FormulaDependencyGraph::dependentsOf(const std::string& sheet, const std::string& cell) const {
    std::vector<FormulaDependency> result;
    for (const auto& dep : edges_)
        if (dep.precedentSheet == sheet && dep.precedentReference == cell) result.push_back(dep);
    return result;
}

bool FormulaDependencyGraph::dependsOn(const std::string& dependentSheet, const std::string& dependentCell,
                                       const std::string& precedentSheet, const std::string& precedentCell) const {
    for (const auto& dep : edges_)
        if (dep.dependentSheet == dependentSheet && dep.dependentCell == dependentCell
            && dep.precedentSheet == precedentSheet && dep.precedentReference == precedentCell)
            return true;
    return false;
}

} // namespace xlpp
