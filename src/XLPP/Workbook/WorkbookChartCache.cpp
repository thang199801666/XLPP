#include <XLPP/Workbook/Workbook.h>
#include "../Worksheet/WorksheetName.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <iomanip>
#include <optional>
#include <regex>
#include <set>
#include <sstream>
#include <stdexcept>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace {
struct ParsedChartReference {
    const xlpp::Worksheet* sheet{nullptr};
    xlpp::CellReference first{};
    xlpp::CellReference last{};
    std::string normalized;
};

std::string trimChartReference(std::string value) {
    const auto first = value.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) return {};
    const auto last = value.find_last_not_of(" \t\r\n");
    return value.substr(first, last - first + 1);
}

bool asciiEqualNoCase(std::string_view lhs, std::string_view rhs) {
    if (lhs.size() != rhs.size()) return false;
    for (std::size_t i = 0; i < lhs.size(); ++i) {
        const auto a = static_cast<unsigned char>(lhs[i]);
        const auto b = static_cast<unsigned char>(rhs[i]);
        if (std::tolower(a) != std::tolower(b)) return false;
    }
    return true;
}

const xlpp::DefinedName* findDefinedNameForSheet(const xlpp::Workbook& workbook,
                                                 const xlpp::Worksheet& sheet,
                                                 const std::string& name) {
    std::optional<std::size_t> sheetIndex;
    try { sheetIndex = workbook.index(sheet); } catch (...) {}
    const xlpp::DefinedName* global = nullptr;
    for (const auto& item : workbook.definedNames()) {
        if (!asciiEqualNoCase(item.name(), name)) continue;
        if (item.localSheetId()) {
            if (sheetIndex && *item.localSheetId() == *sheetIndex) return &item;
        } else if (!global) {
            global = &item;
        }
    }
    return global;
}

struct StructuredTableMatch {
    const xlpp::Worksheet* sheet{nullptr};
    const xlpp::Table* table{nullptr};
};

std::optional<StructuredTableMatch> findStructuredTable(const xlpp::Workbook& workbook,
                                                         const xlpp::Worksheet& preferredSheet,
                                                         std::string_view name,
                                                         bool restrictToPreferred) {
    const auto findInSheet = [&](const xlpp::Worksheet& sheet) -> const xlpp::Table* {
        for (const auto& table : sheet.tables()) {
            if (asciiEqualNoCase(table.name(), name) || asciiEqualNoCase(table.displayName(), name)) return &table;
        }
        return nullptr;
    };
    if (const auto* table = findInSheet(preferredSheet)) return StructuredTableMatch{&preferredSheet, table};
    if (restrictToPreferred) return std::nullopt;
    for (const auto& sheet : workbook.worksheets()) {
        if (&sheet == &preferredSheet) continue;
        if (const auto* table = findInSheet(sheet)) return StructuredTableMatch{&sheet, table};
    }
    return std::nullopt;
}

bool parseTableReferenceBounds(const xlpp::Table& table, xlpp::CellReference& first,
                               xlpp::CellReference& last, std::string& reason) {
    try {
        const auto ref = trimChartReference(table.reference());
        const auto colon = ref.find(':');
        if (colon != std::string::npos && ref.find(':', colon + 1) != std::string::npos) {
            reason = "table has an invalid range: " + table.reference();
            return false;
        }
        first = xlpp::CellReference::parse(colon == std::string::npos ? ref : ref.substr(0, colon));
        last = xlpp::CellReference::parse(colon == std::string::npos ? ref : ref.substr(colon + 1));
        if (first.row > last.row) std::swap(first.row, last.row);
        if (first.column > last.column) std::swap(first.column, last.column);
        return true;
    } catch (const std::exception& ex) {
        reason = "table " + table.name() + " has an invalid range: " + ex.what();
        return false;
    }
}

bool isStructuredEscapeTarget(char ch) {
    return ch == '[' || ch == ']' || ch == '#' || ch == '\'' || ch == '@';
}


std::string structuredTokenText(std::string token) {
    token = trimChartReference(std::move(token));
    std::string decoded;
    decoded.reserve(token.size());
    for (std::size_t i = 0; i < token.size(); ++i) {
        if (token[i] == '\'' && i + 1 < token.size() && isStructuredEscapeTarget(token[i + 1])) {
            decoded.push_back(token[i + 1]);
            ++i;
        } else {
            decoded.push_back(token[i]);
        }
    }
    return trimChartReference(std::move(decoded));
}

struct StructuredTokenScan {
    std::vector<std::string> innermostTokens;
    bool thisRowSelector{false};
    bool balanced{true};
};

StructuredTokenScan scanStructuredTokens(const std::string& structured) {
    struct Frame { std::size_t open{0}; bool nested{false}; };
    StructuredTokenScan result;
    std::vector<Frame> stack;
    for (std::size_t i = 0; i < structured.size(); ++i) {
        const char ch = structured[i];
        if (ch == '\'' && i + 1 < structured.size() && isStructuredEscapeTarget(structured[i + 1])) {
            ++i;
            continue;
        }
        if (ch == '@') result.thisRowSelector = true;
        if (ch == '[') {
            if (!stack.empty()) stack.back().nested = true;
            stack.push_back({i, false});
            continue;
        }
        if (ch != ']') continue;
        if (stack.empty()) { result.balanced = false; return result; }
        const auto frame = stack.back();
        stack.pop_back();
        if (!frame.nested && i > frame.open + 1) {
            auto token = trimChartReference(structured.substr(frame.open + 1, i - frame.open - 1));
            if (!token.empty()) result.innermostTokens.push_back(std::move(token));
        }
    }
    result.balanced = stack.empty();
    return result;
}

std::size_t structuredReferenceEnd(const std::string& text, std::size_t open) {
    if (open >= text.size() || text[open] != '[') return std::string::npos;
    int depth = 0;
    for (std::size_t i = open; i < text.size(); ++i) {
        if (text[i] == '\'' && i + 1 < text.size() && isStructuredEscapeTarget(text[i + 1])) {
            ++i;
            continue;
        }
        if (text[i] == '[') ++depth;
        else if (text[i] == ']') {
            --depth;
            if (depth == 0) return i + 1;
            if (depth < 0) return std::string::npos;
        }
    }
    return std::string::npos;
}

bool resolveStructuredTableReference(const xlpp::Workbook& workbook, const xlpp::Worksheet& preferredSheet,
                                     std::string expression, bool restrictToPreferred,
                                     ParsedChartReference& parsed, std::string& reason,
                                     const xlpp::CellReference* currentCell) {
    expression = trimChartReference(std::move(expression));
    const auto firstBracket = expression.find('[');
    if (firstBracket == std::string::npos || expression.back() != ']') {
        reason = "malformed structured table reference";
        return false;
    }
    auto tableName = trimChartReference(expression.substr(0, firstBracket));
    std::optional<StructuredTableMatch> match;
    if (!tableName.empty()) {
        match = findStructuredTable(workbook, preferredSheet, tableName, restrictToPreferred);
        if (!match) {
            reason = "table not found: " + tableName;
            return false;
        }
    } else {
        if (!currentCell) {
            reason = "implicit structured reference requires a current table row";
            return false;
        }
        for (const auto& table : preferredSheet.tables()) {
            xlpp::CellReference tableFirst, tableLast; std::string tableReason;
            if (!parseTableReferenceBounds(table, tableFirst, tableLast, tableReason)) continue;
            if (currentCell->row >= tableFirst.row && currentCell->row <= tableLast.row &&
                currentCell->column >= tableFirst.column && currentCell->column <= tableLast.column) {
                match = StructuredTableMatch{&preferredSheet, &table};
                tableName = table.name();
                break;
            }
        }
        if (!match) {
            reason = "implicit structured reference is outside a table";
            return false;
        }
    }

    xlpp::CellReference tableFirst, tableLast;
    if (!parseTableReferenceBounds(*match->table, tableFirst, tableLast, reason)) return false;
    const auto width = tableLast.column - tableFirst.column + 1;
    const auto headerRows = match->table->showHeaderRow() ? std::size_t{1} : std::size_t{0};
    const auto totalRows = match->table->showTotalsRow() ? std::size_t{1} : std::size_t{0};
    const auto dataFirstRow = tableFirst.row + headerRows;
    const auto dataLastRow = tableLast.row >= totalRows ? tableLast.row - totalRows : std::size_t{0};

    // Collect innermost specifiers with escape-aware bracket handling. Excel
    // uses a leading apostrophe to escape [, ], #, ' and @ in column headers.
    const auto structured = expression.substr(firstBracket);
    const auto scanned = scanStructuredTokens(structured);
    if (!scanned.balanced) { reason = "unterminated structured-reference token"; return false; }
    if (scanned.innermostTokens.empty()) {
        reason = "structured reference has no selector";
        return false;
    }

    bool hasExplicitRowSelector = false;
    bool includeHeader = false, includeData = false, includeTotals = false;
    bool thisRow = scanned.thisRowSelector;
    std::vector<std::string> columnTokens;
    for (auto rawToken : scanned.innermostTokens) {
        if (!rawToken.empty() && rawToken.front() == '@') {
            thisRow = true;
            auto token = structuredTokenText(rawToken.substr(1));
            if (!token.empty()) columnTokens.push_back(std::move(token));
            continue;
        }
        auto token = structuredTokenText(std::move(rawToken));
        if (asciiEqualNoCase(token, "#All")) {
            hasExplicitRowSelector = true; includeHeader = includeData = includeTotals = true;
        } else if (asciiEqualNoCase(token, "#Data")) {
            hasExplicitRowSelector = true; includeData = true;
        } else if (asciiEqualNoCase(token, "#Headers")) {
            hasExplicitRowSelector = true; includeHeader = true;
        } else if (asciiEqualNoCase(token, "#Totals")) {
            hasExplicitRowSelector = true; includeTotals = true;
        } else if (asciiEqualNoCase(token, "#This Row")) {
            hasExplicitRowSelector = true; thisRow = true;
        } else {
            columnTokens.push_back(std::move(token));
        }
    }
    if (!hasExplicitRowSelector && !thisRow) includeData = true;
    if (thisRow && (includeHeader || includeData || includeTotals)) {
        reason = "#This Row cannot be combined with another special item specifier";
        return false;
    }

    auto firstRow = dataFirstRow;
    auto lastRow = dataLastRow;
    if (thisRow) {
        if (!currentCell) { reason = "#This Row requires a current table row"; return false; }
        if (currentCell->row < dataFirstRow || currentCell->row > dataLastRow) {
            reason = "#This Row is outside the table data body";
            return false;
        }
        firstRow = lastRow = currentCell->row;
    } else {
        if (includeHeader && !headerRows) { reason = "table has no header row"; return false; }
        if (includeTotals && !totalRows) { reason = "table has no totals row"; return false; }
        // #Headers + #Totals without #Data is a non-contiguous multi-area
        // selection for a non-empty table and cannot be represented by the
        // single rectangular ParsedChartReference model.
        if (includeHeader && includeTotals && !includeData && dataFirstRow <= dataLastRow) {
            reason = "structured reference resolves to non-contiguous header/totals areas";
            return false;
        }
        bool selected = false;
        if (includeHeader) { firstRow = lastRow = tableFirst.row; selected = true; }
        if (includeData && dataFirstRow <= dataLastRow) {
            if (!selected) firstRow = dataFirstRow;
            lastRow = dataLastRow;
            selected = true;
        }
        if (includeTotals) {
            if (!selected) firstRow = tableLast.row;
            lastRow = tableLast.row;
            selected = true;
        }
        if (!selected) { reason = "structured reference resolves to an empty table body"; return false; }
    }
    if (firstRow == 0 || lastRow == 0 || firstRow > lastRow) {
        reason = "structured reference resolves to an empty table body";
        return false;
    }

    auto columnNameAt = [&](std::size_t offset) -> std::string {
        if (offset < match->table->columns().size() && !match->table->columns()[offset].name().empty())
            return match->table->columns()[offset].name();
        if (match->table->showHeaderRow()) {
            if (const auto* cell = match->sheet->tryCell(tableFirst.row, tableFirst.column + offset)) {
                if (const auto* text = std::get_if<std::string>(&cell->value()); text && !text->empty()) return *text;
            }
        }
        return "Column" + std::to_string(offset + 1);
    };
    auto findColumn = [&](const std::string& name) -> std::optional<std::size_t> {
        for (std::size_t i = 0; i < width; ++i) if (asciiEqualNoCase(columnNameAt(i), name)) return i;
        return std::nullopt;
    };

    std::size_t firstColumn = tableFirst.column;
    std::size_t lastColumn = tableLast.column;
    if (!columnTokens.empty()) {
        const auto firstOffset = findColumn(columnTokens.front());
        const auto lastOffset = findColumn(columnTokens.back());
        if (!firstOffset) { reason = "table column not found: " + columnTokens.front(); return false; }
        if (!lastOffset) { reason = "table column not found: " + columnTokens.back(); return false; }
        firstColumn = tableFirst.column + std::min(*firstOffset, *lastOffset);
        lastColumn = tableFirst.column + std::max(*firstOffset, *lastOffset);
    }

    parsed.sheet = match->sheet;
    parsed.first = {firstRow, firstColumn};
    parsed.last = {lastRow, lastColumn};
    parsed.normalized = expression;
    return true;
}

bool parseChartReferenceImpl(const xlpp::Workbook& workbook, const xlpp::Worksheet& owner,
                             std::string reference, ParsedChartReference& parsed, std::string& reason,
                             bool allowTwoDimensional, std::set<std::string>& resolvingNames,
                             std::size_t nameDepth, const xlpp::CellReference* currentCell);

bool parseSignedIntegerLiteral(std::string text, long long& value) {
    text = trimChartReference(std::move(text));
    if (text.empty()) return false;
    try {
        std::size_t used = 0;
        value = std::stoll(text, &used);
        return used == text.size();
    } catch (...) { return false; }
}

bool splitReferenceFunction(std::string expression, std::string_view functionName,
                            std::vector<std::string>& args) {
    expression = trimChartReference(std::move(expression));
    if (!expression.empty() && expression.front() == '=') expression = trimChartReference(expression.substr(1));
    if (expression.size() <= functionName.size() + 2 ||
        !asciiEqualNoCase(std::string_view(expression).substr(0, functionName.size()), functionName) ||
        expression[functionName.size()] != '(' || expression.back() != ')') return false;
    const auto body = expression.substr(functionName.size() + 1, expression.size() - functionName.size() - 2);
    std::size_t start = 0; int depth = 0; bool inString = false; bool inQuote = false;
    for (std::size_t i = 0; i <= body.size(); ++i) {
        const char ch = i < body.size() ? body[i] : ',';
        if (i < body.size()) {
            if (ch == '"' && !inQuote) {
                if (inString && i + 1 < body.size() && body[i + 1] == '"') { ++i; continue; }
                inString = !inString; continue;
            }
            if (!inString && ch == '\'') {
                if (inQuote && i + 1 < body.size() && body[i + 1] == '\'') { ++i; continue; }
                inQuote = !inQuote; continue;
            }
            if (inString || inQuote) continue;
            if (ch == '(') { ++depth; continue; }
            if (ch == ')') { --depth; if (depth < 0) return false; continue; }
        }
        if ((ch == ',' || ch == ';') && depth == 0 && !inString && !inQuote) {
            args.push_back(trimChartReference(body.substr(start, i - start)));
            start = i + 1;
        }
    }
    return depth == 0 && !inString && !inQuote;
}

bool isBoundedDynamicReference(std::string expression) {
    expression = trimChartReference(std::move(expression));
    if (!expression.empty() && expression.front() == '=') expression = trimChartReference(expression.substr(1));
    const auto open = expression.find('(');
    if (open == std::string::npos) return false;
    const auto name = trimChartReference(expression.substr(0, open));
    return asciiEqualNoCase(name, "OFFSET") || asciiEqualNoCase(name, "INDEX");
}


std::size_t topLevelReferenceColon(const std::string& expression) {
    int parentheses = 0, brackets = 0;
    bool quotedSheet = false, stringLiteral = false;
    for (std::size_t i = 0; i < expression.size(); ++i) {
        const char ch = expression[i];
        if (ch == '"' && !quotedSheet) {
            if (stringLiteral && i + 1 < expression.size() && expression[i + 1] == '"') { ++i; continue; }
            stringLiteral = !stringLiteral;
            continue;
        }
        if (stringLiteral) continue;
        if (ch == '\'') {
            if (quotedSheet && i + 1 < expression.size() && expression[i + 1] == '\'') { ++i; continue; }
            if (i + 1 < expression.size() && isStructuredEscapeTarget(expression[i + 1]) && brackets > 0) {
                ++i;
                continue;
            }
            quotedSheet = !quotedSheet;
            continue;
        }
        if (quotedSheet) continue;
        if (ch == '(') ++parentheses;
        else if (ch == ')') { if (parentheses > 0) --parentheses; }
        else if (ch == '[') ++brackets;
        else if (ch == ']') { if (brackets > 0) --brackets; }
        else if (ch == ':' && parentheses == 0 && brackets == 0) return i;
    }
    return std::string::npos;
}

bool resolveBoundedDynamicReference(const xlpp::Workbook& workbook, const xlpp::Worksheet& owner,
                                    std::string expression, ParsedChartReference& parsed, std::string& reason,
                                    bool allowTwoDimensional, std::set<std::string>& resolvingNames,
                                    std::size_t nameDepth, const xlpp::CellReference* currentCell) {
    std::vector<std::string> args;
    if (splitReferenceFunction(expression, "OFFSET", args)) {
        if (args.size() < 3 || args.size() > 5) { reason = "OFFSET requires 3 to 5 arguments"; return false; }
        ParsedChartReference base;
        if (!parseChartReferenceImpl(workbook, owner, args[0], base, reason, true, resolvingNames, nameDepth + 1, currentCell)) {
            reason = "OFFSET base: " + reason; return false;
        }
        long long rowOffset = 0, columnOffset = 0;
        if (!parseSignedIntegerLiteral(args[1], rowOffset) || !parseSignedIntegerLiteral(args[2], columnOffset)) {
            reason = "OFFSET row/column offsets must be integer literals"; return false;
        }
        long long height = static_cast<long long>(base.last.row - base.first.row + 1);
        long long width = static_cast<long long>(base.last.column - base.first.column + 1);
        if (args.size() >= 4 && !args[3].empty() && !parseSignedIntegerLiteral(args[3], height)) {
            reason = "OFFSET height must be an integer literal"; return false;
        }
        if (args.size() >= 5 && !args[4].empty() && !parseSignedIntegerLiteral(args[4], width)) {
            reason = "OFFSET width must be an integer literal"; return false;
        }
        if (height <= 0 || width <= 0) { reason = "OFFSET height/width must be positive"; return false; }
        const auto row = static_cast<long long>(base.first.row) + rowOffset;
        const auto column = static_cast<long long>(base.first.column) + columnOffset;
        const auto lastRow = row + height - 1;
        const auto lastColumn = column + width - 1;
        if (row < 1 || column < 1 || lastRow > 1048576 || lastColumn > 16384) {
            reason = "OFFSET result is outside the Excel grid"; return false;
        }
        parsed.sheet = base.sheet;
        parsed.first = {static_cast<std::size_t>(row), static_cast<std::size_t>(column)};
        parsed.last = {static_cast<std::size_t>(lastRow), static_cast<std::size_t>(lastColumn)};
        if (!allowTwoDimensional && parsed.first.row != parsed.last.row && parsed.first.column != parsed.last.column) {
            reason = "OFFSET resolves to a two-dimensional range that cannot be materialized into a chart cache";
            return false;
        }
        parsed.normalized = trimChartReference(std::move(expression));
        return true;
    }

    args.clear();
    if (splitReferenceFunction(expression, "INDEX", args)) {
        if (args.size() < 2 || args.size() > 3) { reason = "INDEX reference form requires 2 or 3 arguments"; return false; }
        ParsedChartReference base;
        if (!parseChartReferenceImpl(workbook, owner, args[0], base, reason, true, resolvingNames, nameDepth + 1, currentCell)) {
            reason = "INDEX base: " + reason; return false;
        }
        long long rowNumber = 0, columnNumber = 1;
        if (!parseSignedIntegerLiteral(args[1], rowNumber) || rowNumber < 0) {
            reason = "INDEX row number must be a non-negative integer literal"; return false;
        }
        if (args.size() >= 3 && !args[2].empty()) {
            if (!parseSignedIntegerLiteral(args[2], columnNumber) || columnNumber < 0) {
                reason = "INDEX column number must be a non-negative integer literal"; return false;
            }
        } else if (base.first.column != base.last.column) {
            columnNumber = 0; // INDEX(2-D-range,row) denotes the selected row.
        }
        const auto height = static_cast<long long>(base.last.row - base.first.row + 1);
        const auto width = static_cast<long long>(base.last.column - base.first.column + 1);
        if (rowNumber > height || columnNumber > width) { reason = "INDEX selection is outside the reference"; return false; }
        parsed.sheet = base.sheet;
        if (rowNumber == 0 && columnNumber == 0) {
            parsed.first = base.first; parsed.last = base.last;
        } else if (rowNumber == 0) {
            const auto column = base.first.column + static_cast<std::size_t>(columnNumber - 1);
            parsed.first = {base.first.row, column}; parsed.last = {base.last.row, column};
        } else if (columnNumber == 0) {
            const auto row = base.first.row + static_cast<std::size_t>(rowNumber - 1);
            parsed.first = {row, base.first.column}; parsed.last = {row, base.last.column};
        } else {
            const auto row = base.first.row + static_cast<std::size_t>(rowNumber - 1);
            const auto column = base.first.column + static_cast<std::size_t>(columnNumber - 1);
            parsed.first = parsed.last = {row, column};
        }
        if (!allowTwoDimensional && parsed.first.row != parsed.last.row && parsed.first.column != parsed.last.column) {
            reason = "INDEX resolves to a two-dimensional range that cannot be materialized into a chart cache";
            return false;
        }
        parsed.normalized = trimChartReference(std::move(expression));
        return true;
    }
    reason = "unsupported dynamic reference expression";
    return false;
}

bool parseChartReferenceImpl(const xlpp::Workbook& workbook, const xlpp::Worksheet& owner,
                             std::string reference, ParsedChartReference& parsed, std::string& reason,
                             bool allowTwoDimensional, std::set<std::string>& resolvingNames,
                             std::size_t nameDepth, const xlpp::CellReference* currentCell) {
    reference = trimChartReference(std::move(reference));
    if (!reference.empty() && reference.front() == '=') reference.erase(reference.begin());
    reference = trimChartReference(std::move(reference));
    if (reference.empty()) { reason = "empty reference"; return false; }
    if (nameDepth >= 64) { reason = "reference resolution depth exceeded"; return false; }

    // P0Z: Excel's reference form allows INDEX results to act as endpoints of
    // the range operator, e.g. INDEX(A1:A10,2):INDEX(A1:A10,5). Resolve the
    // endpoint geometry without evaluating worksheet values.
    const auto referenceColon = topLevelReferenceColon(reference);
    if (referenceColon != std::string::npos) {
        auto leftExpression = trimChartReference(reference.substr(0, referenceColon));
        auto rightExpression = trimChartReference(reference.substr(referenceColon + 1));
        if (isBoundedDynamicReference(leftExpression) || isBoundedDynamicReference(rightExpression)) {
            ParsedChartReference left, right;
            std::string leftReason, rightReason;
            if (!parseChartReferenceImpl(workbook, owner, leftExpression, left, leftReason, true,
                                         resolvingNames, nameDepth + 1, currentCell)) {
                reason = "left range endpoint: " + leftReason; return false;
            }
            if (!parseChartReferenceImpl(workbook, owner, rightExpression, right, rightReason, true,
                                         resolvingNames, nameDepth + 1, currentCell)) {
                reason = "right range endpoint: " + rightReason; return false;
            }
            if (!left.sheet || !right.sheet || left.sheet != right.sheet) {
                reason = "range endpoints resolve to different worksheets"; return false;
            }
            if (left.first.row != left.last.row || left.first.column != left.last.column ||
                right.first.row != right.last.row || right.first.column != right.last.column) {
                reason = "INDEX/OFFSET range endpoints must each resolve to a single cell"; return false;
            }
            parsed.sheet = left.sheet;
            parsed.first = {std::min(left.first.row, right.first.row), std::min(left.first.column, right.first.column)};
            parsed.last = {std::max(left.first.row, right.first.row), std::max(left.first.column, right.first.column)};
            if (!allowTwoDimensional && parsed.first.row != parsed.last.row && parsed.first.column != parsed.last.column) {
                reason = "dynamic endpoint range resolves to a two-dimensional range that cannot be materialized into a chart cache";
                return false;
            }
            parsed.normalized = reference;
            return true;
        }
    }
    if (isBoundedDynamicReference(reference))
        return resolveBoundedDynamicReference(workbook, owner, reference, parsed, reason, allowTwoDimensional,
                                              resolvingNames, nameDepth, currentCell);
    std::size_t bang = std::string::npos;
    bool quoted = false;
    int structuredDepth = 0;
    for (std::size_t i = 0; i < reference.size(); ++i) {
        const char ch = reference[i];
        if (structuredDepth > 0) {
            if (ch == '\'' && i + 1 < reference.size() && isStructuredEscapeTarget(reference[i + 1])) { ++i; continue; }
            if (ch == '[') ++structuredDepth;
            else if (ch == ']') --structuredDepth;
            continue;
        }
        if (!quoted && ch == '[') { structuredDepth = 1; continue; }
        if (ch == '\'') {
            if (quoted && i + 1 < reference.size() && reference[i + 1] == '\'') { ++i; continue; }
            quoted = !quoted; continue;
        }
        if (!quoted && ch == '!') { bang = i; break; }
    }
    if (quoted) { reason = "unterminated worksheet quote"; return false; }
    if (structuredDepth != 0) { reason = "unterminated structured reference"; return false; }
    const auto firstBracketInReference = reference.find('[');
    if (firstBracketInReference != std::string::npos && bang != std::string::npos && firstBracketInReference < bang) {
        reason = "external workbook references are not synchronized";
        return false;
    }

    std::string sheetName = owner.name();
    std::string range = reference;
    if (bang != std::string::npos) {
        auto token = trimChartReference(reference.substr(0, bang));
        range = trimChartReference(reference.substr(bang + 1));
        if (token.size() >= 2 && token.front() == '\'' && token.back() == '\'') {
            std::string unquoted;
            for (std::size_t i = 1; i + 1 < token.size(); ++i) {
                if (token[i] == '\'' && i + 1 < token.size() - 1 && token[i + 1] == '\'') {
                    unquoted.push_back('\''); ++i;
                } else {
                    unquoted.push_back(token[i]);
                }
            }
            sheetName = std::move(unquoted);
        } else {
            sheetName = std::move(token);
        }
    }
    const auto* source = workbook.worksheet(sheetName);
    if (!source) { reason = "worksheet not found: " + sheetName; return false; }

    if (range.find('[') != std::string::npos || range.find(']') != std::string::npos) {
        if (!resolveStructuredTableReference(workbook, *source, range, bang != std::string::npos,
                                             parsed, reason, currentCell)) return false;
        if (!allowTwoDimensional && parsed.first.row != parsed.last.row && parsed.first.column != parsed.last.column) {
            reason = "structured reference resolves to a two-dimensional range that cannot be materialized into a chart cache";
            return false;
        }
        return true;
    }
    if (range.find(',') != std::string::npos || range.find(';') != std::string::npos) {
        reason = "union references are not synchronized"; return false;
    }

    // Formula dependency traversal also accepts whole-column/whole-row
    // references. They are deliberately not materialized as chart caches.
    if (allowTwoDimensional) {
        const auto colon = range.find(':');
        if (colon != std::string::npos && range.find(':', colon + 1) == std::string::npos) {
            auto left = trimChartReference(range.substr(0, colon));
            auto right = trimChartReference(range.substr(colon + 1));
            const auto stripDollar = [](std::string value) {
                if (!value.empty() && value.front() == '$') value.erase(value.begin());
                return value;
            };
            left = stripDollar(std::move(left));
            right = stripDollar(std::move(right));
            const auto allLetters = [](const std::string& value) {
                return !value.empty() && std::all_of(value.begin(), value.end(), [](unsigned char ch) {
                    return std::isalpha(ch) != 0;
                });
            };
            const auto allDigits = [](const std::string& value) {
                return !value.empty() && std::all_of(value.begin(), value.end(), [](unsigned char ch) {
                    return std::isdigit(ch) != 0;
                });
            };
            try {
                if (allLetters(left) && allLetters(right)) {
                    auto first = xlpp::CellReference::parse(left + "1");
                    auto last = xlpp::CellReference::parse(right + "1048576");
                    if (first.column > last.column) std::swap(first.column, last.column);
                    parsed.sheet = source;
                    parsed.first = {1, first.column};
                    parsed.last = {1048576, last.column};
                    parsed.normalized = reference;
                    return true;
                }
                if (allDigits(left) && allDigits(right)) {
                    const auto firstRow = static_cast<std::size_t>(std::stoull(left));
                    const auto lastRow = static_cast<std::size_t>(std::stoull(right));
                    if (firstRow == 0 || lastRow == 0 || firstRow > 1048576 || lastRow > 1048576) {
                        reason = "whole-row reference is outside the Excel grid";
                        return false;
                    }
                    parsed.sheet = source;
                    parsed.first = {std::min(firstRow, lastRow), 1};
                    parsed.last = {std::max(firstRow, lastRow), 16384};
                    parsed.normalized = reference;
                    return true;
                }
            } catch (const std::exception& ex) {
                reason = ex.what();
                return false;
            }
        }
    }

    std::string cellParseReason;
    try {
        const auto colon = range.find(':');
        if (colon != std::string::npos && range.find(':', colon + 1) != std::string::npos) {
            reason = "invalid range reference"; return false;
        }
        auto first = xlpp::CellReference::parse(colon == std::string::npos ? range : range.substr(0, colon));
        auto last = xlpp::CellReference::parse(colon == std::string::npos ? range : range.substr(colon + 1));
        if (first.row > last.row) std::swap(first.row, last.row);
        if (first.column > last.column) std::swap(first.column, last.column);
        if (!allowTwoDimensional && first.row != last.row && first.column != last.column) {
            reason = "two-dimensional ranges are not synchronized into chart caches"; return false;
        }
        parsed.sheet = source;
        parsed.first = first;
        parsed.last = last;
        parsed.normalized = reference;
        return true;
    } catch (const std::exception& ex) {
        cellParseReason = ex.what();
    }

    // A chart or formula may refer to a workbook/local defined name instead of
    // spelling out its A1 range. Names may reduce to A1/structured references
    // or to bounded OFFSET/INDEX reference forms whose geometry is statically
    // knowable. Calculation-dependent dynamic expressions remain diagnostic-only.
    const auto* defined = findDefinedNameForSheet(workbook, *source, range);
    if (!defined) {
        reason = cellParseReason.empty() ? "reference is neither A1 nor a resolvable defined name" : cellParseReason;
        return false;
    }
    if (nameDepth >= 64) {
        reason = "defined-name resolution depth exceeded for " + defined->name();
        return false;
    }

    std::string key = source->name() + "!" + defined->name();
    std::transform(key.begin(), key.end(), key.begin(), [](unsigned char ch) {
        return static_cast<char>(std::tolower(ch));
    });
    if (!resolvingNames.insert(key).second) {
        reason = "defined-name cycle detected at " + defined->name();
        return false;
    }

    ParsedChartReference resolved;
    const bool ok = parseChartReferenceImpl(workbook, *source, defined->value(), resolved, reason,
                                            allowTwoDimensional, resolvingNames, nameDepth + 1, currentCell);
    resolvingNames.erase(key);
    if (!ok) {
        reason = "defined name " + defined->name() + ": " + reason;
        return false;
    }
    parsed = std::move(resolved);
    parsed.normalized = reference;
    return true;
}

bool parseChartReference(const xlpp::Workbook& workbook, const xlpp::Worksheet& owner,
                         std::string reference, ParsedChartReference& parsed, std::string& reason,
                         bool allowTwoDimensional = false, const xlpp::CellReference* currentCell = nullptr) {
    std::set<std::string> resolvingNames;
    return parseChartReferenceImpl(workbook, owner, std::move(reference), parsed, reason,
                                   allowTwoDimensional, resolvingNames, 0, currentCell);
}

std::string chartCacheNumber(double value) {
    if (value == 0.0) value = 0.0; // normalize negative zero
    std::ostringstream out; out << std::setprecision(15) << value; return out.str();
}

enum class ChartCacheKind { String, Numeric, Automatic };

xlpp::ChartSeriesCache buildChartCache(const ParsedChartReference& ref, ChartCacheKind requested,
                                       bool date1904, const xlpp::ChartSeriesCache& existing,
                                       bool preserveFormulaCachedValues = true,
                                       std::size_t* reusedFormulaPoints = nullptr,
                                       std::vector<std::string>* warnings = nullptr) {
    struct SourceValue { std::size_t index; const xlpp::Cell* cell; };
    std::vector<SourceValue> cells;
    if (ref.first.row == ref.last.row) {
        cells.reserve(ref.last.column - ref.first.column + 1);
        for (std::size_t col = ref.first.column, index = 0; col <= ref.last.column; ++col, ++index)
            cells.push_back({index, ref.sheet->tryCell(ref.first.row, col)});
    } else {
        cells.reserve(ref.last.row - ref.first.row + 1);
        for (std::size_t row = ref.first.row, index = 0; row <= ref.last.row; ++row, ++index)
            cells.push_back({index, ref.sheet->tryCell(row, ref.first.column)});
    }
    bool numeric = requested == ChartCacheKind::Numeric;
    if (requested == ChartCacheKind::Automatic) {
        numeric = true;
        bool sawValue = false;
        for (const auto& source : cells) {
            if (!source.cell || !source.cell->hasValue()) continue;
            sawValue = true;
            const auto& value = source.cell->value();
            if (!(std::holds_alternative<double>(value) || std::holds_alternative<xlpp::DateTime>(value) || std::holds_alternative<bool>(value))) {
                numeric = false; break;
            }
        }
        if (!sawValue) numeric = existing.present ? existing.numeric : false;
    }
    xlpp::ChartSeriesCache cache; cache.present = true; cache.numeric = numeric; cache.pointCount = cells.size();
    if (numeric) {
        cache.formatCode = existing.numeric && !existing.formatCode.empty() ? existing.formatCode : "General";
        if (cache.formatCode == "General") {
            for (const auto& source : cells) if (source.cell && source.cell->hasValue() && source.cell->numberFormat() != "General") {
                cache.formatCode = source.cell->numberFormat(); break;
            }
        }
    }
    const auto existingPoint = [&](std::size_t index) -> const xlpp::ChartCachePoint* {
        const auto it = std::find_if(existing.points.begin(), existing.points.end(), [&](const auto& point) { return point.index == index; });
        return it == existing.points.end() ? nullptr : &*it;
    };
    for (const auto& source : cells) {
        if (!source.cell || !source.cell->hasValue()) {
            if (source.cell && source.cell->hasFormula() && preserveFormulaCachedValues) {
                if (const auto* cached = existingPoint(source.index)) {
                    cache.points.push_back(*cached);
                    if (reusedFormulaPoints) ++*reusedFormulaPoints;
                }
            }
            continue; // sparse cache: preserve index, omit blank point
        }
        const auto& value = source.cell->value();
        std::string text;
        if (numeric) {
            if (const auto* number = std::get_if<double>(&value)) text = chartCacheNumber(*number);
            else if (const auto* date = std::get_if<xlpp::DateTime>(&value)) text = chartCacheNumber(xlpp::toExcelSerial(*date, date1904));
            else if (const auto* boolean = std::get_if<bool>(&value)) text = *boolean ? "1" : "0";
            else {
                if (warnings) warnings->push_back("Skipped non-numeric cache point at " + source.cell->address() + " in " + ref.sheet->name());
                continue;
            }
        } else {
            if (const auto* string = std::get_if<std::string>(&value)) text = *string;
            else if (const auto* number = std::get_if<double>(&value)) text = chartCacheNumber(*number);
            else if (const auto* date = std::get_if<xlpp::DateTime>(&value)) text = chartCacheNumber(xlpp::toExcelSerial(*date, date1904));
            else if (const auto* boolean = std::get_if<bool>(&value)) text = *boolean ? "TRUE" : "FALSE";
            else if (const auto* error = std::get_if<xlpp::CellError>(&value)) text = xlpp::toString(*error);
        }
        cache.points.push_back({source.index, std::move(text)});
    }
    return cache;
}

bool chartRangeDirectlyTouchesTrackedCells(const ParsedChartReference& ref) {
    if (!ref.sheet) return false;
    constexpr std::uint64_t columnMask = (std::uint64_t{1} << 20) - 1;
    const auto inside = [&](std::uint64_t key) {
        const auto row = static_cast<std::size_t>(key >> 20);
        const auto column = static_cast<std::size_t>(key & columnMask);
        return row >= ref.first.row && row <= ref.last.row &&
               column >= ref.first.column && column <= ref.last.column;
    };
    for (const auto key : ref.sheet->trackedCellKeys()) {
        if (inside(key)) return true;
    }
    // Retained Cell& handles can be mutated after the initial cell()/range()
    // access. Iterate materialized cells rather than every coordinate in the
    // region, which keeps large 2-D dependency ranges bounded by sheet density.
    for (const auto& entry : ref.sheet->cells()) {
        if (entry.second.mutationRevision() != 0 && inside(entry.first)) return true;
    }
    return false;
}

std::string formulaWithoutStringLiterals(const std::string& formula) {
    std::string sanitized = formula;
    bool inString = false;
    for (std::size_t i = 0; i < sanitized.size(); ++i) {
        if (sanitized[i] != '"') { if (inString) sanitized[i] = ' '; continue; }
        if (inString && i + 1 < sanitized.size() && sanitized[i + 1] == '"') {
            sanitized[i] = sanitized[i + 1] = ' '; ++i; continue;
        }
        inString = !inString; sanitized[i] = ' ';
    }
    return sanitized;
}

std::vector<ParsedChartReference> simpleFormulaPrecedents(const xlpp::Workbook& workbook,
                                                          const xlpp::Worksheet& owner,
                                                          const std::string& formula,
                                                          xlpp::ChartCacheSyncReport* report,
                                                          const xlpp::CellReference* currentCell = nullptr) {
    static const std::regex refPattern(
        R"((?:'(?:(?:'')|[^'])+'|[A-Za-z_][A-Za-z0-9_.]*)!(?:\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]+:\$?[0-9]+)|(?:\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]+:\$?[0-9]+))");
    static const std::regex namePattern(R"(\b[A-Za-z_][A-Za-z0-9_.]*\b)");

    const auto sanitized = formulaWithoutStringLiterals(formula);
    std::string referenceScan = sanitized;
    std::string nameScan = sanitized;
    std::vector<ParsedChartReference> result;
    std::set<std::string> unique;

    const auto append = [&](ParsedChartReference parsed) {
        if (!parsed.sheet) return;
        const auto key = parsed.sheet->name() + "!" +
            std::to_string(parsed.first.row) + ":" + std::to_string(parsed.first.column) + "-" +
            std::to_string(parsed.last.row) + ":" + std::to_string(parsed.last.column);
        if (unique.insert(key).second) result.push_back(std::move(parsed));
    };

    const auto maskSpan = [&](std::size_t pos, std::size_t length) {
        std::fill(referenceScan.begin() + static_cast<std::ptrdiff_t>(pos),
                  referenceScan.begin() + static_cast<std::ptrdiff_t>(pos + length), ' ');
        std::fill(nameScan.begin() + static_cast<std::ptrdiff_t>(pos),
                  nameScan.begin() + static_cast<std::ptrdiff_t>(pos + length), ' ');
    };
    const auto processStructured = [&](std::size_t pos, std::size_t endPos) {
        const auto expression = sanitized.substr(pos, endPos - pos);
        if (report) ++report->structuredReferencesVisited;
        ParsedChartReference parsed; std::string reason;
        if (!parseChartReference(workbook, owner, expression, parsed, reason, true, currentCell)) {
            if (report) {
                ++report->structuredReferencesSkipped;
                report->formulaDependencyDiagnostics.push_back(
                    "Skipped structured reference '" + expression + "' in " + owner.name() + ": " + reason);
            }
        } else {
            if (report) ++report->structuredReferencesResolved;
            append(std::move(parsed));
        }
        maskSpan(pos, endPos - pos);
    };

    // Named structured references: Table1[Sales],
    // Table1[[#Data],[Sales]:[Margin]], etc.
    for (std::size_t i = 0; i < sanitized.size();) {
        const auto isStart = [](unsigned char ch) { return std::isalpha(ch) != 0 || ch == '_'; };
        const auto isContinue = [](unsigned char ch) { return std::isalnum(ch) != 0 || ch == '_' || ch == '.'; };
        if (!isStart(static_cast<unsigned char>(sanitized[i])) ||
            (i > 0 && isContinue(static_cast<unsigned char>(sanitized[i - 1])))) { ++i; continue; }
        std::size_t nameEnd = i + 1;
        while (nameEnd < sanitized.size() && isContinue(static_cast<unsigned char>(sanitized[nameEnd]))) ++nameEnd;
        if (nameEnd >= sanitized.size() || sanitized[nameEnd] != '[') { i = nameEnd; continue; }
        const auto endPos = structuredReferenceEnd(sanitized, nameEnd);
        if (endPos == std::string::npos) { ++i; continue; }
        processStructured(i, endPos);
        i = endPos;
    }

    // Implicit row-scoped table references such as [@Sales] are meaningful
    // only while traversing a formula cell that lives inside a table.
    if (currentCell) {
        for (std::size_t i = 0; i < referenceScan.size();) {
            if (referenceScan[i] != '[') { ++i; continue; }
            if (i > 0) {
                const auto prev = static_cast<unsigned char>(referenceScan[i - 1]);
                if (std::isalnum(prev) || prev == '_' || prev == '.') { ++i; continue; }
            }
            const auto endPos = structuredReferenceEnd(sanitized, i);
            if (endPos == std::string::npos) { ++i; continue; }
            const auto expression = sanitized.substr(i, endPos - i);
            if (expression.find('@') == std::string::npos && expression.find("#This Row") == std::string::npos) {
                i = endPos; continue;
            }
            processStructured(i, endPos);
            i = endPos;
        }
    }

    for (std::sregex_iterator it(referenceScan.begin(), referenceScan.end(), refPattern), end; it != end; ++it) {
        const auto pos = static_cast<std::size_t>(it->position());
        const auto length = static_cast<std::size_t>(it->length());
        if (pos > 0) {
            const auto c = referenceScan[pos - 1];
            if (std::isalnum(static_cast<unsigned char>(c)) || c == '_' || c == '.') continue;
        }
        if (pos + length < referenceScan.size()) {
            const auto c = referenceScan[pos + length];
            if (std::isalnum(static_cast<unsigned char>(c)) || c == '_' || c == '.') continue;
        }
        std::fill(nameScan.begin() + static_cast<std::ptrdiff_t>(pos),
                  nameScan.begin() + static_cast<std::ptrdiff_t>(pos + length), ' ');
        if (report) ++report->formulaReferencesVisited;
        ParsedChartReference parsed; std::string reason;
        if (!parseChartReference(workbook, owner, it->str(), parsed, reason, true, currentCell)) {
            if (report) {
                ++report->formulaReferencesSkipped;
                report->formulaDependencyDiagnostics.push_back(
                    "Skipped formula reference '" + it->str() + "' in " + owner.name() + ": " + reason);
            }
            continue;
        }
        if (parsed.first.row > 1048576 || parsed.last.row > 1048576 ||
            parsed.first.column > 16384 || parsed.last.column > 16384) {
            if (report) {
                ++report->formulaReferencesSkipped;
                report->formulaDependencyDiagnostics.push_back(
                    "Skipped out-of-grid formula reference '" + it->str() + "' in " + owner.name());
            }
            continue;
        }
        if (report) ++report->formulaReferencesResolved;
        append(std::move(parsed));
    }

    // Resolve explicit defined-name tokens left after masking all A1 reference
    // spans. Function identifiers and sheet qualifiers therefore do not become
    // false named-range dependencies.
    for (std::sregex_iterator it(nameScan.begin(), nameScan.end(), namePattern), end; it != end; ++it) {
        const auto token = it->str();
        const auto* defined = findDefinedNameForSheet(workbook, owner, token);
        if (!defined) continue;

        std::size_t after = static_cast<std::size_t>(it->position() + it->length());
        while (after < nameScan.size() && std::isspace(static_cast<unsigned char>(nameScan[after]))) ++after;
        if (after < nameScan.size() && nameScan[after] == '(') continue;

        const bool dynamicName = isBoundedDynamicReference(defined->value());
        if (report) {
            ++report->definedNameDependenciesVisited;
            ++report->formulaReferencesVisited;
            if (dynamicName) ++report->dynamicDefinedNamesVisited;
        }
        ParsedChartReference parsed; std::string reason;
        if (!parseChartReference(workbook, owner, token, parsed, reason, true, currentCell)) {
            if (report) {
                ++report->definedNameDependenciesSkipped;
                ++report->formulaReferencesSkipped;
                if (dynamicName) ++report->dynamicDefinedNamesSkipped;
                report->formulaDependencyDiagnostics.push_back(
                    "Skipped defined-name dependency '" + token + "' in " + owner.name() + ": " + reason);
            }
            continue;
        }
        if (report) {
            ++report->definedNameDependenciesResolved;
            ++report->formulaReferencesResolved;
            if (dynamicName) ++report->dynamicDefinedNamesResolved;
        }
        append(std::move(parsed));
    }

    if (report && (referenceScan.find('[') != std::string::npos || referenceScan.find(']') != std::string::npos)) {
        report->formulaDependencyDiagnostics.push_back(
            "Formula in " + owner.name() + " contains unresolved external-workbook/structured-reference syntax; "
            "independently resolvable precedents were still followed");
    }
    return result;
}

enum class ChartReferenceTouchKind { None, Direct, FormulaPrecedent };

bool formulaRangeTouchesTrackedCells(const xlpp::Workbook& workbook, const ParsedChartReference& ref,
                                     std::size_t depth, std::size_t maxDepth,
                                     std::set<std::string>& visited,
                                     xlpp::ChartCacheSyncReport* report) {
    if (!ref.sheet || depth > maxDepth) return false;
    constexpr std::uint64_t columnMask = (std::uint64_t{1} << 20) - 1;
    for (const auto& entry : ref.sheet->cells()) {
        const auto row = static_cast<std::size_t>(entry.first >> 20);
        const auto column = static_cast<std::size_t>(entry.first & columnMask);
        if (row < ref.first.row || row > ref.last.row ||
            column < ref.first.column || column > ref.last.column) continue;
        const auto* cell = &entry.second;
        if (!cell->hasFormula()) continue;
        if (report) ++report->formulaDependenciesVisited;
        const auto key = ref.sheet->name() + "!" + cell->address();
        if (!visited.insert(key).second) {
            if (report) report->formulaDependencyDiagnostics.push_back(
                "Formula dependency cycle/revisit suppressed at " + key);
            continue;
        }
        const xlpp::CellReference formulaCell{row, column};
        for (const auto& precedent : simpleFormulaPrecedents(workbook, *ref.sheet, cell->formula(), report, &formulaCell)) {
            if (chartRangeDirectlyTouchesTrackedCells(precedent)) return true;
            if (depth < maxDepth &&
                formulaRangeTouchesTrackedCells(workbook, precedent, depth + 1, maxDepth, visited, report)) return true;
        }
    }
    return false;
}

ChartReferenceTouchKind chartReferenceTouchKind(const xlpp::Workbook& workbook, const ParsedChartReference& ref,
                                                const xlpp::ChartCacheSyncOptions& options,
                                                xlpp::ChartCacheSyncReport* report) {
    if (chartRangeDirectlyTouchesTrackedCells(ref)) return ChartReferenceTouchKind::Direct;
    if (!options.propagateFormulaDependencies || options.maxFormulaDependencyDepth == 0) return ChartReferenceTouchKind::None;
    std::set<std::string> visited;
    if (formulaRangeTouchesTrackedCells(workbook, ref, 0, options.maxFormulaDependencyDepth, visited, report)) {
        if (report) ++report->formulaDependenciesMatched;
        return ChartReferenceTouchKind::FormulaPrecedent;
    }
    return ChartReferenceTouchKind::None;
}


} // namespace

namespace xlpp {
std::vector<ChartCacheDependency> Workbook::chartCacheDependencies() const {
    std::vector<ChartCacheDependency> result;
    for (std::size_t sheetIndex = 0; sheetIndex < sheets_.size(); ++sheetIndex) {
        const auto& sheet = sheets_[sheetIndex];
        for (std::size_t chartIndex = 0; chartIndex < sheet.charts_.size(); ++chartIndex) {
            const auto& chart = sheet.charts_[chartIndex];
            for (std::size_t seriesIndex = 0; seriesIndex < chart.series().size(); ++seriesIndex) {
                const auto& series = chart.series()[seriesIndex];
                const auto append = [&](const std::string& reference, ChartCacheDependencyKind kind) {
                    if (reference.empty()) return;
                    ChartCacheDependency dependency;
                    dependency.ownerSheet = sheet.name();
                    dependency.chartStableId = chart.stableId();
                    dependency.chartIndex = chartIndex;
                    dependency.seriesIndex = seriesIndex;
                    dependency.kind = kind;
                    dependency.reference = reference;
                    ParsedChartReference parsed;
                    std::string reason;
                    if (parseChartReference(*this, sheet, reference, parsed, reason)) {
                        dependency.supported = true;
                        dependency.sourceSheet = parsed.sheet ? parsed.sheet->name() : std::string{};
                        dependency.first = parsed.first;
                        dependency.last = parsed.last;
                    } else {
                        dependency.issue = std::move(reason);
                    }
                    result.push_back(std::move(dependency));
                };
                append(series.titleReference(), ChartCacheDependencyKind::Title);
                append(series.categoriesReference(), ChartCacheDependencyKind::Category);
                append(series.valuesReference(), ChartCacheDependencyKind::Value);
            }
        }
    }
    return result;
}

void Workbook::clearChartCacheChangeTracking() noexcept {
    for (auto& sheet : sheets_) sheet.clearTrackedCellChanges();
}

ChartCacheSyncReport Workbook::synchronizeChangedChartCaches(ChartCacheSyncOptions options) {
    options.onlyChangedCells = true;
    return synchronizeChartCaches(options);
}

ChartCacheSyncReport Workbook::synchronizeChartCaches(const ChartCacheSyncOptions& options) {
    ChartCacheSyncReport report;
    const auto cacheEqual = [](const ChartSeriesCache& a, const ChartSeriesCache& b) {
        if (a.present != b.present || a.numeric != b.numeric || a.formatCode != b.formatCode || a.pointCount != b.pointCount || a.points.size() != b.points.size()) return false;
        for (std::size_t i = 0; i < a.points.size(); ++i)
            if (a.points[i].index != b.points[i].index || a.points[i].value != b.points[i].value) return false;
        return true;
    };
    for (auto& sheet : sheets_) {
        for (auto& chart : sheet.charts_) {
            ++report.chartsVisited;
            for (std::size_t seriesIndex = 0; seriesIndex < chart.series().size(); ++seriesIndex) {
                ++report.seriesVisited;
                auto& series = chart.series()[seriesIndex];
                bool unsupportedSelectedReference = false;
                auto synchronize = [&](const std::string& reference, ChartCacheKind kind, const ChartSeriesCache& existing,
                                       const char* label, bool enabled) {
                    if (!enabled || reference.empty()) return;
                    ++report.dependenciesVisited;
                    ParsedChartReference parsed; std::string reason;
                    if (!parseChartReference(*this, sheet, reference, parsed, reason)) {
                        ++report.referencesSkipped; unsupportedSelectedReference = true;
                        report.warnings.push_back(sheet.name() + ": chart " + (chart.stableId().empty() ? std::string("<generated>") : chart.stableId()) +
                                                  ", series " + std::to_string(seriesIndex) + ", " + label + " reference '" + reference + "': " + reason);
                        return;
                    }
                    auto touchKind = ChartReferenceTouchKind::Direct;
                    if (options.onlyChangedCells) {
                        touchKind = chartReferenceTouchKind(*this, parsed, options, &report);
                        if (touchKind == ChartReferenceTouchKind::None) {
                            ++report.dependenciesSkippedUnchanged;
                            return;
                        }
                    }
                    ++report.dependenciesMatched;
                    if (touchKind == ChartReferenceTouchKind::FormulaPrecedent) {
                        ++report.staleFormulaCachesPreserved;
                        if (options.requestHostRecalculationForFormulaDependencies) {
                            calcProps_.setCalcOnSave(true);
                            calcProps_.setFullCalcOnLoad(true);
                            report.hostRecalculationRequested = true;
                        }
                        return;
                    }
                    if (kind == ChartCacheKind::String && (parsed.first.row != parsed.last.row || parsed.first.column != parsed.last.column)) {
                        ++report.referencesSkipped; unsupportedSelectedReference = true;
                        report.warnings.push_back(sheet.name() + ": chart series title reference must resolve to one cell: " + reference);
                        return;
                    }
                    std::size_t formulaPointsReused = 0;
                    auto rebuilt = buildChartCache(parsed, kind, date1904_, existing,
                                                   options.preserveFormulaCachedValues, &formulaPointsReused, &report.warnings);
                    report.formulaCachePointsReused += formulaPointsReused;
                    if (!rebuilt.valid(true)) {
                        ++report.referencesSkipped; unsupportedSelectedReference = true;
                        report.warnings.push_back(sheet.name() + ": rebuilt " + label + " cache failed validation for " + reference);
                        return;
                    }
                    if (cacheEqual(rebuilt, existing)) return;
                    bool accepted = true;
                    if (chart.imported()) {
                        if (kind == ChartCacheKind::String) accepted = sheet.setChartSeriesTitleCache(chart.stableId(), seriesIndex, rebuilt);
                        else if (kind == ChartCacheKind::Numeric) accepted = sheet.setChartSeriesValueCache(chart.stableId(), seriesIndex, rebuilt);
                        else accepted = sheet.setChartSeriesCategoryCache(chart.stableId(), seriesIndex, rebuilt);
                    } else {
                        if (kind == ChartCacheKind::String) series.setTitleCache(rebuilt);
                        else if (kind == ChartCacheKind::Numeric) series.setValuesCache(rebuilt);
                        else series.setCategoriesCache(rebuilt);
                        sheet.dirty_ = true;
                        sheet.drawingAppendDirty_ = true;
                    }
                    if (accepted) ++report.cachesUpdated;
                    else report.warnings.push_back(sheet.name() + ": failed to apply rebuilt " + label + " cache for series " + std::to_string(seriesIndex));
                };
                synchronize(series.titleReference(), ChartCacheKind::String, series.titleCache(), "title", options.synchronizeTitles);
                synchronize(series.categoriesReference(), ChartCacheKind::Automatic, series.categoriesCache(), "category", options.synchronizeCategories);
                synchronize(series.valuesReference(), ChartCacheKind::Numeric, series.valuesCache(), "value", options.synchronizeValues);
                if (options.clearUnsupportedReferences && unsupportedSelectedReference) {
                    if (chart.imported()) {
                        if (sheet.clearChartSeriesCaches(chart.stableId(), seriesIndex)) ++report.cachesCleared;
                    } else {
                        const bool hadAny = series.titleCache().present || series.categoriesCache().present || series.valuesCache().present;
                        series.setTitleCache({}); series.setCategoriesCache({}); series.setValuesCache({});
                        if (hadAny) ++report.cachesCleared;
                        sheet.dirty_ = true; sheet.drawingAppendDirty_ = true;
                    }
                }
            }
        }
    }
    if (options.clearTrackedChangesAfterSync) clearChartCacheChangeTracking();
    return report;
}


} // namespace xlpp
