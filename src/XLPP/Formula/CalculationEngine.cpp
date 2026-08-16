#include <XLPP/Formula/CalculationEngine.h>
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Cell/CellReference.h>
#include <algorithm>
#include <cctype>
#include <charconv>
#include <cmath>
#include <cstring>
#include <ctime>
#include <functional>
#include <limits>
#include <optional>
#include <sstream>
#include <unordered_set>

namespace {

using xlpp::CellValue;
using xlpp::CellError;
using xlpp::DateTime;
using std::monostate;

constexpr std::string_view kKeySeparator = "\x1F";

// ---------------------------------------------------------------------------
// Value coercion helpers
// ---------------------------------------------------------------------------

bool isEmptyValue(const CellValue& v) noexcept { return std::holds_alternative<monostate>(v); }
bool isErrorValue(const CellValue& v) noexcept { return std::holds_alternative<CellError>(v); }
CellError errorOf(const CellValue& v) noexcept { return std::get<CellError>(v); }

std::string_view trimView(std::string_view s) {
    while (!s.empty() && std::isspace(static_cast<unsigned char>(s.front()))) s.remove_prefix(1);
    while (!s.empty() && std::isspace(static_cast<unsigned char>(s.back()))) s.remove_suffix(1);
    return s;
}

std::optional<double> parseNumberText(std::string_view s) {
    s = trimView(s);
    if (s.empty()) return std::nullopt;
    double value = 0.0;
    const auto [ptr, ec] = std::from_chars(s.data(), s.data() + s.size(), value);
    if (ec == std::errc{} && ptr == s.data() + s.size()) return value;
    // Percent sign and thousands separators are not handled by from_chars.
    return std::nullopt;
}

struct Numeric {
    double value{0.0};
    bool isError{false};
    CellError error{CellError::Value};
};

Numeric toNumeric(const CellValue& v) {
    if (const auto* d = std::get_if<double>(&v)) return {*d, false, CellError::Value};
    if (const auto* b = std::get_if<bool>(&v)) return {*b ? 1.0 : 0.0, false, CellError::Value};
    if (const auto* s = std::get_if<std::string>(&v)) {
        if (const auto parsed = parseNumberText(*s)) return {*parsed, false, CellError::Value};
        return {0.0, true, CellError::Value};
    }
    if (const auto* d = std::get_if<DateTime>(&v)) return {xlpp::toExcelSerial(*d, false), false, CellError::Value};
    if (isErrorValue(v)) return {0.0, true, errorOf(v)};
    return {0.0, false, CellError::Value};
}

std::string toTextValue(const CellValue& v) {
    if (const auto* s = std::get_if<std::string>(&v)) return *s;
    if (const auto* d = std::get_if<double>(&v)) {
        if (*d == static_cast<double>(static_cast<long long>(*d)) && std::abs(*d) < 1e15) {
            return std::to_string(static_cast<long long>(*d));
        }
        char buffer[32];
        const auto result = std::to_chars(buffer, buffer + sizeof(buffer), *d);
        return std::string(buffer, result.ptr);
    }
    if (const auto* b = std::get_if<bool>(&v)) return *b ? "TRUE" : "FALSE";
    if (std::holds_alternative<monostate>(v)) return "";
    if (const auto* d = std::get_if<DateTime>(&v)) {
        char buffer[32];
        const auto result = std::to_chars(buffer, buffer + sizeof(buffer), xlpp::toExcelSerial(*d, false));
        return std::string(buffer, result.ptr);
    }
    return "";
}

struct BooleanValue {
    bool value{false};
    bool isError{false};
    CellError error{CellError::Value};
};

BooleanValue toBooleanValue(const CellValue& v) {
    if (const auto* b = std::get_if<bool>(&v)) return {*b, false, CellError::Value};
    if (const auto* d = std::get_if<double>(&v)) return {*d != 0.0, false, CellError::Value};
    if (const auto* s = std::get_if<std::string>(&v)) {
        const auto upper = [&]{ std::string t = *s; std::transform(t.begin(), t.end(), t.begin(), [](unsigned char c){ return static_cast<char>(std::toupper(c)); }); return t; }();
        if (upper == "TRUE") return {true, false, CellError::Value};
        if (upper == "FALSE") return {false, false, CellError::Value};
        return {false, true, CellError::Value};
    }
    if (std::holds_alternative<monostate>(v)) return {false, false, CellError::Value};
    if (isErrorValue(v)) return {false, true, errorOf(v)};
    return {false, true, CellError::Value};
}

// ---------------------------------------------------------------------------
// Tokenizer
// ---------------------------------------------------------------------------

enum class TokenType { Number, String, Ident, Ref, SingleQuoted, Operator, LParen, RParen, Comma, Colon, LBracket, RBracket, Excl, End };

struct Token {
    TokenType type{TokenType::End};
    std::string_view text;
    double number{0.0};
    char op{0};
};

bool isRefStart(char c) { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '$'; }
bool isIdentChar(char c) {
    return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '.' || c == '_' || c == '$';
}

class Tokenizer {
public:
    explicit Tokenizer(std::string_view input) : input_(input) {}

    Token next() {
        skipWhitespace();
        if (pos_ >= input_.size()) return {TokenType::End, {}, 0.0, 0};
        const char c = input_[pos_];
        // Number literal.
        if (std::isdigit(static_cast<unsigned char>(c)) || (c == '.' && pos_ + 1 < input_.size()
            && std::isdigit(static_cast<unsigned char>(input_[pos_ + 1])))) {
            const std::size_t start = pos_;
            bool seenDigit = false, seenDot = false, seenExp = false;
            while (pos_ < input_.size()) {
                const char ch = input_[pos_];
                if (std::isdigit(static_cast<unsigned char>(ch))) { seenDigit = true; ++pos_; }
                else if (ch == '.' && !seenDot) { seenDot = true; ++pos_; }
                else if ((ch == 'e' || ch == 'E') && seenDigit && !seenExp && pos_ + 1 < input_.size()) {
                    const char nx = input_[pos_ + 1];
                    if (std::isdigit(static_cast<unsigned char>(nx)) || ((nx == '+' || nx == '-')
                        && pos_ + 2 < input_.size() && std::isdigit(static_cast<unsigned char>(input_[pos_ + 2])))) {
                        seenExp = true; ++pos_;
                        if (input_[pos_] == '+' || input_[pos_] == '-') ++pos_;
                    } else break;
                } else break;
            }
            double value = 0.0;
            std::from_chars(input_.data() + start, input_.data() + pos_, value);
            return {TokenType::Number, input_.substr(start, pos_ - start), value, 0};
        }
        // Quoted sheet name ('Sheet 1').
        if (c == '\'') {
            const std::size_t end = input_.find('\'', pos_ + 1);
            if (end == std::string_view::npos) { pos_ = input_.size(); return {TokenType::Ident, {}, 0.0, 0}; }
            Token token{TokenType::SingleQuoted, input_.substr(pos_ + 1, end - pos_ - 1), 0.0, 0};
            pos_ = end + 1;
            return token;
        }
        // Double-quoted string literal with "" escaping.
        if (c == '"') {
            std::size_t end = pos_ + 1;
            while (end < input_.size()) {
                if (input_[end] == '"') {
                    if (end + 1 < input_.size() && input_[end + 1] == '"') { end += 2; continue; }
                    break;
                }
                ++end;
            }
            if (end >= input_.size()) end = input_.size();
            Token token{TokenType::String, input_.substr(pos_ + 1, end > pos_ + 1 ? end - pos_ - 1 : 0), 0.0, 0};
            pos_ = end < input_.size() ? end + 1 : end;
            return token;
        }
        // Identifier / reference.
        if (isRefStart(c)) {
            const std::size_t start = pos_;
            while (pos_ < input_.size() && isIdentChar(input_[pos_])) ++pos_;
            return {TokenType::Ident, input_.substr(start, pos_ - start), 0.0, 0};
        }
        ++pos_;
        switch (c) {
            case '(': return {TokenType::LParen, {}, 0.0, 0};
            case ')': return {TokenType::RParen, {}, 0.0, 0};
            case ',': return {TokenType::Comma, {}, 0.0, 0};
            case ':': return {TokenType::Colon, {}, 0.0, 0};
            case '[': return {TokenType::LBracket, {}, 0.0, 0};
            case ']': return {TokenType::RBracket, {}, 0.0, 0};
            case '!': return {TokenType::Excl, {}, 0.0, 0};
            default: return {TokenType::Operator, {}, 0.0, c};
        }
    }

private:
    void skipWhitespace() {
        while (pos_ < input_.size() && std::isspace(static_cast<unsigned char>(input_[pos_]))) ++pos_;
    }
    std::string_view input_;
    std::size_t pos_{0};

public:
    std::size_t position() const noexcept { return pos_; }
};

// ---------------------------------------------------------------------------
// Parser / evaluator
// ---------------------------------------------------------------------------

struct ParsedRef {
    std::string sheet;
    std::string workbook;   // empty unless external reference
    xlpp::CellReference from;
    bool isRange{false};
    xlpp::CellReference to;
};

class Parser {
public:
    Parser(const xlpp::internal::CalculationEvaluator& ctx, std::string_view formula)
        : ctx_(ctx), formula_(formula) { advance(); }

    CellValue parse() {
        const auto result = parseComparison();
        if (tok_.type != TokenType::End) {
            consume(TokenType::End);
            return xlpp::CellValue{CellError::Value};
        }
        return result;
    }

private:
    const xlpp::internal::CalculationEvaluator& ctx_;
    std::string_view formula_;
    Token tok_;

    void advance() { tok_ = tokenizer_.next(); }

    // ---- low-level helpers ----
    bool at(TokenType type) const { return tok_.type == type; }
    void consume(TokenType type) {
        if (tok_.type != type) throw std::runtime_error("XL++ formula: unexpected token");
        advance();
    }
    bool isOperator(char op) const { return tok_.type == TokenType::Operator && tok_.op == op; }
    void consumeOperator(char op) {
        if (!isOperator(op)) throw std::runtime_error("XL++ formula: unexpected operator");
        advance();
    }
    CellValue errorValue(CellError error) const { return xlpp::CellValue{error}; }

    CellValue parseComparison() {
        auto left = parseConcat();
        while (true) {
            char op = 0;
            if (isOperator('=')) {
                op = '=';
                advance();
            } else if (isOperator('<')) {
                advance();
                if (isOperator('>')) { op = 'N'; advance(); }       // <>
                else if (isOperator('=')) { op = 'L'; advance(); }  // <=
                else op = '<';
            } else if (isOperator('>')) {
                advance();
                if (isOperator('=')) { op = 'G'; advance(); }       // >=
                else op = '>';
            } else {
                break;
            }
            auto right = parseConcat();
            left = compareValues(left, right, op);
        }
        return left;
    }

    CellValue compareValues(CellValue left, CellValue right, char op) {
        if (isErrorValue(left)) return left;
        if (isErrorValue(right)) return right;
        int result = 0;
        const auto lb = std::get_if<bool>(&left);
        const auto rb = std::get_if<bool>(&right);
        if ((lb && !rb) || (!lb && rb)) {
            // mixed bool/number: Excel orders booleans < numbers
            const bool lIsBool = lb != nullptr;
            const bool rIsBool = rb != nullptr;
            if (lIsBool != rIsBool) result = lIsBool ? -1 : 1;
            else result = 0;
        } else if (lb && rb) {
            result = (*lb == *rb) ? 0 : (*lb ? 1 : -1);
        } else {
            const auto ld = std::get_if<double>(&left);
            const auto rd = std::get_if<double>(&right);
            const auto ls = std::get_if<std::string>(&left);
            const auto rs = std::get_if<std::string>(&right);
            if (ld && rd) result = (*ld < *rd) ? -1 : (*ld > *rd ? 1 : 0);
            else if (ls && rs) {
                std::string a = *ls, b = *rs;
                std::transform(a.begin(), a.end(), a.begin(), [](unsigned char ch){ return static_cast<char>(std::toupper(ch)); });
                std::transform(b.begin(), b.end(), b.begin(), [](unsigned char ch){ return static_cast<char>(std::toupper(ch)); });
                result = (a < b) ? -1 : (a > b ? 1 : 0);
            } else if (isEmptyValue(left) || isEmptyValue(right)) {
                // Empty coerces to 0 or "" depending on the other side.
                const auto other = isEmptyValue(left) ? right : left;
                if (std::get_if<double>(&other) || std::get_if<DateTime>(&other)) {
                    const auto ln = toNumeric(left), rn = toNumeric(right);
                    result = (ln.value < rn.value) ? -1 : (ln.value > rn.value ? 1 : 0);
                } else if (std::get_if<bool>(&other)) {
                    const bool a = std::get_if<bool>(&left) ? *std::get_if<bool>(&left) : false;
                    const bool b = std::get_if<bool>(&right) ? *std::get_if<bool>(&right) : false;
                    result = (a == b) ? 0 : (a ? 1 : -1);
                } else {
                    const auto a = toTextValue(left), b = toTextValue(right);
                    result = (a < b) ? -1 : (a > b ? 1 : 0);
                }
            } else {
                const auto a = toTextValue(left), b = toTextValue(right);
                result = (a < b) ? -1 : (a > b ? 1 : 0);
            }
        }
        bool outcome = false;
        switch (op) {
            case '=': outcome = result == 0; break;
            case '<': outcome = result < 0; break;
            case '>': outcome = result > 0; break;
            case 'L': outcome = result <= 0; break;
            case 'G': outcome = result >= 0; break;
            case 'N': outcome = result != 0; break;
            default: return errorValue(CellError::Value);
        }
        return xlpp::CellValue{outcome};
    }

    // actual comparison operators handled in parseComparison via explicit checks
    CellValue parseConcat() {
        auto left = parseAddSub();
        while (isOperator('&')) {
            advance();
            auto right = parseAddSub();
            if (isErrorValue(left)) return left;
            if (isErrorValue(right)) return right;
            left = xlpp::CellValue{toTextValue(left) + toTextValue(right)};
        }
        return left;
    }

    CellValue parseAddSub() {
        auto left = parseMulDiv();
        while (isOperator('+') || isOperator('-')) {
            const char op = tok_.op;
            advance();
            auto right = parseMulDiv();
            const auto ln = toNumeric(left), rn = toNumeric(right);
            if (ln.isError) return errorValue(ln.error);
            if (rn.isError) return errorValue(rn.error);
            left = xlpp::CellValue{op == '+' ? ln.value + rn.value : ln.value - rn.value};
        }
        return left;
    }

    CellValue parseMulDiv() {
        auto left = parsePower();
        while (isOperator('*') || isOperator('/')) {
            const char op = tok_.op;
            advance();
            auto right = parsePower();
            const auto ln = toNumeric(left), rn = toNumeric(right);
            if (ln.isError) return errorValue(ln.error);
            if (rn.isError) return errorValue(rn.error);
            if (op == '*') left = xlpp::CellValue{ln.value * rn.value};
            else {
                if (rn.value == 0.0) return errorValue(CellError::DivisionByZero);
                left = xlpp::CellValue{ln.value / rn.value};
            }
        }
        return left;
    }

    CellValue parsePower() {
        auto left = parseUnary();
        if (isOperator('^')) {
            advance();
            auto right = parseUnary();
            const auto ln = toNumeric(left), rn = toNumeric(right);
            if (ln.isError) return errorValue(ln.error);
            if (rn.isError) return errorValue(rn.error);
            left = xlpp::CellValue{std::pow(ln.value, rn.value)};
        }
        return left;
    }

    CellValue parseUnary() {
        if (isOperator('-')) {
            advance();
            auto value = parseUnary();
            const auto n = toNumeric(value);
            if (n.isError) return errorValue(n.error);
            return xlpp::CellValue{-n.value};
        }
        if (isOperator('+')) {
            advance();
            return parseUnary();
        }
        return parsePostfix();
    }

    CellValue parsePostfix() {
        auto value = parsePrimary();
        while (isOperator('%')) {
            advance();
            const auto n = toNumeric(value);
            if (n.isError) return errorValue(n.error);
            value = xlpp::CellValue{n.value / 100.0};
        }
        return value;
    }

    // ---- references ----
    bool isCellReference(const std::string_view& text) const {
        std::size_t i = 0;
        if (i < text.size() && text[i] == '$') ++i;
        std::size_t letters = 0;
        while (i < text.size() && ((text[i] >= 'A' && text[i] <= 'Z') || (text[i] >= 'a' && text[i] <= 'z'))) { ++i; ++letters; }
        if (letters == 0 || letters > 3) return false;
        if (i < text.size() && text[i] == '$') ++i;
        std::size_t digits = 0;
        while (i < text.size() && text[i] >= '0' && text[i] <= '9') { ++i; ++digits; }
        return i == text.size() && digits > 0 && digits <= 7;
    }

    // Parses an already-consumed ref token text into a cell reference; returns
    // false when the token is not a valid A1 reference.
    bool parseCellReference(const std::string_view& text, xlpp::CellReference& out) {
        std::size_t i = 0;
        if (i < text.size() && text[i] == '$') ++i;
        std::size_t colLetters = 0;
        while (i < text.size() && ((text[i] >= 'A' && text[i] <= 'Z') || (text[i] >= 'a' && text[i] <= 'z'))) { ++i; ++colLetters; }
        std::size_t column = 0;
        for (std::size_t k = i - colLetters; k < i; ++k) {
            char ch = text[k];
            if (ch >= 'a' && ch <= 'z') ch = static_cast<char>(ch - 'a' + 'A');
            column = column * 26 + static_cast<std::size_t>(ch - 'A' + 1);
        }
        if (i < text.size() && text[i] == '$') ++i;
        std::size_t row = 0;
        while (i < text.size() && text[i] >= '0' && text[i] <= '9') { row = row * 10 + static_cast<std::size_t>(text[i] - '0'); ++i; }
        if (i != text.size() || row < 1 || column < 1 || column > 16384 || row > 1048576) return false;
        out = xlpp::CellReference{row, column};
        return true;
    }

    // Reads a fully-qualified reference: optional quoted sheet / [book] then
    // ! then an A1 cell or range. Requires tok_ to be at the sheet qualifier or
    // at the cell ref itself. Returns true when a reference was consumed.
    bool tryParseReference(ParsedRef& out) {
        std::string sheet;
        std::string workbook;
        if (at(TokenType::LBracket)) {
            advance();
            const auto start = tok_.type == TokenType::Ident || tok_.type == TokenType::SingleQuoted;
            std::string name;
            if (start) { name = std::string(tok_.text); advance(); }
            consume(TokenType::RBracket);
            if (!at(TokenType::Excl)) return false;
            workbook = name;
            advance();
            if (at(TokenType::SingleQuoted) || at(TokenType::Ident)) { sheet = std::string(tok_.text); advance(); }
            consume(TokenType::Excl);
        } else if (at(TokenType::SingleQuoted)) {
            sheet = std::string(tok_.text);
            advance();
            if (!at(TokenType::Excl)) return false;
            advance();
        } else if (at(TokenType::Ident)) {
            const std::string_view text = tok_.text;
            if (!isCellReference(text)) return false;
        }
        if (!at(TokenType::Ident)) return false;
        const std::string_view refText = tok_.text;
        if (!isCellReference(refText)) return false;
        xlpp::CellReference from;
        if (!parseCellReference(refText, from)) return false;
        advance();
        ParsedRef result;
        result.sheet = sheet.empty() ? std::string(ctx_.sheetName) : sheet;
        result.workbook = workbook;
        result.from = from;
        result.to = from;
        if (at(TokenType::Colon)) {
            advance();
            if (!at(TokenType::Ident)) return false;
            const std::string_view toText = tok_.text;
            xlpp::CellReference toRef;
            if (!isCellReference(toText) || !parseCellReference(toText, toRef)) return false;
            advance();
            result.to = toRef;
            result.isRange = true;
        }
        out = std::move(result);
        return true;
    }

    // Evaluate a range (or single cell) into an array of cell values.
    std::vector<CellValue> evaluateReferenceRange(const ParsedRef& ref) {
        std::vector<CellValue> result;
        if (!ref.workbook.empty()) {
            if (ctx_.options.externalReferenceResolver) {
                const auto value = ctx_.options.externalReferenceResolver(ref.workbook, ref.sheet, ref.from.address());
                if (ctx_.externalResolved) ++(*ctx_.externalResolved);
                if (value) result.push_back(*value);
                else {
                    if (ctx_.unresolvedExternal) ++(*ctx_.unresolvedExternal);
                    result.push_back(xlpp::CellValue{CellError::Reference});
                }
            } else {
                if (ctx_.unresolvedExternal) ++(*ctx_.unresolvedExternal);
                result.push_back(xlpp::CellValue{CellError::Reference});
            }
            return result;
        }
        const auto* sheet = ctx_.workbook.worksheet(ref.sheet);
        if (!sheet) {
            result.push_back(xlpp::CellValue{CellError::Reference});
            return result;
        }
        const std::size_t rowBegin = std::min(ref.from.row, ref.to.row);
        const std::size_t rowEnd = std::max(ref.from.row, ref.to.row);
        const std::size_t colBegin = std::min(ref.from.column, ref.to.column);
        const std::size_t colEnd = std::max(ref.from.column, ref.to.column);
        const std::size_t rowCount = rowEnd - rowBegin + 1;
        const std::size_t colCount = colEnd - colBegin + 1;
        result.reserve(rowCount * colCount);
        for (std::size_t r = rowBegin; r <= rowEnd; ++r) {
            for (std::size_t c = colBegin; c <= colEnd; ++c) {
                result.push_back(xlpp::internal::resolveReference(ctx_, ref.sheet, xlpp::CellReference{r, c}.address()));
            }
        }
        return result;
    }

    CellValue evaluateReferenceValue(const ParsedRef& ref) {
        if (ref.isRange) {
            const auto values = evaluateReferenceRange(ref);
            if (values.empty()) return xlpp::CellValue{};
            return values.front();
        }
        if (!ref.workbook.empty()) {
            if (ctx_.options.externalReferenceResolver) {
                const auto value = ctx_.options.externalReferenceResolver(ref.workbook, ref.sheet, ref.from.address());
                if (ctx_.externalResolved) ++(*ctx_.externalResolved);
                if (value) return *value;
                if (ctx_.unresolvedExternal) ++(*ctx_.unresolvedExternal);
            } else {
                if (ctx_.unresolvedExternal) ++(*ctx_.unresolvedExternal);
            }
            return xlpp::CellValue{CellError::Reference};
        }
        return xlpp::internal::resolveReference(ctx_, ref.sheet, ref.from.address());
    }

    // ---- primary expressions ----
    CellValue parsePrimary() {
        if (at(TokenType::Number)) {
            const double value = tok_.number;
            advance();
            return xlpp::CellValue{value};
        }
        if (at(TokenType::String)) {
            std::string text(tok_.text);
            std::string unescaped;
            unescaped.reserve(text.size());
            for (std::size_t i = 0; i < text.size(); ++i) {
                if (text[i] == '"' && i + 1 < text.size() && text[i + 1] == '"') { unescaped.push_back('"'); ++i; }
                else unescaped.push_back(text[i]);
            }
            advance();
            return xlpp::CellValue{unescaped};
        }
        if (at(TokenType::SingleQuoted) || at(TokenType::LBracket)) {
            ParsedRef ref;
            if (tryParseReference(ref)) return evaluateReferenceValue(ref);
            return errorValue(CellError::Reference);
        }
        if (at(TokenType::Ident)) {
            const std::string_view text = tok_.text;
            // Function call?
            if (text.size() >= 2 && text != "TRUE" && text != "FALSE" && text != "NA") {
                // Peek for '(' by re-tokenizing.
                if (peekIsLParen()) return parseFunctionCall();
            }
            if (text == "TRUE") { advance(); return xlpp::CellValue{true}; }
            if (text == "FALSE") { advance(); return xlpp::CellValue{false}; }
            if (isCellReference(text)) {
                ParsedRef ref;
                if (tryParseReference(ref)) return evaluateReferenceValue(ref);
                return errorValue(CellError::Reference);
            }
            // Defined name / unsupported symbol.
            advance();
            if (ctx_.unsupportedFormulas) ++(*ctx_.unsupportedFormulas);
            return errorValue(CellError::Name);
        }
        if (at(TokenType::LParen)) {
            advance();
            auto value = parseComparison();
            consume(TokenType::RParen);
            return value;
        }
        return errorValue(CellError::Value);
    }

    bool peekIsLParen() {
        // Re-tokenize from the current position to detect '('.
        Tokenizer peeker(std::string_view(formula_.data() + currentOffset(), formula_.size() - currentOffset()));
        const auto t = peeker.next();
        return t.type == TokenType::LParen;
    }

    std::size_t currentOffset() const { return tokenizer_.position(); }

    CellValue parseFunctionCall() {
        const std::string name(tok_.text);
        advance();
        consume(TokenType::LParen);
        std::vector<std::vector<CellValue>> args;
        if (!at(TokenType::RParen)) {
            while (true) {
                args.push_back(parseFunctionArgument());
                if (at(TokenType::Comma)) { advance(); continue; }
                break;
            }
        }
        consume(TokenType::RParen);
        return callFunction(name, args);
    }

    // A function argument is either a reference/range (evaluated as an array)
    // or an arbitrary expression (evaluated as a single value). A leading
    // reference is only treated as a bare argument when it is a range or a
    // standalone reference; "B1>10" and "A1/7" remain expressions.
    TokenType peekAfterLeadingReference() {
        Tokenizer peeker(formula_.substr(currentOffset()));
        Token t = peeker.next();
        while (t.type == TokenType::Ident || t.type == TokenType::SingleQuoted
               || t.type == TokenType::LBracket || t.type == TokenType::RBracket
               || t.type == TokenType::Excl) t = peeker.next();
        return t.type;
    }
    std::vector<CellValue> parseFunctionArgument() {
        if (at(TokenType::Ident) || at(TokenType::SingleQuoted) || at(TokenType::LBracket)) {
            const auto after = peekAfterLeadingReference();
            if (after == TokenType::Colon) {
                ParsedRef ref;
                if (tryParseReference(ref)) return evaluateReferenceRange(ref);
            } else if (after == TokenType::Comma || after == TokenType::RParen || after == TokenType::End) {
                ParsedRef ref;
                if (tryParseReference(ref)) return {evaluateReferenceValue(ref)};
            }
        }
        return {parseComparison()};
    }

    CellValue callFunction(const std::string& name, const std::vector<std::vector<CellValue>>& args);

    Tokenizer tokenizer_{formula_};
};

CellValue Parser::callFunction(const std::string& rawName, const std::vector<std::vector<CellValue>>& args) {
    std::string name = rawName;
    if (name.size() > 6 && name.substr(0, 6) == "_xlfn.") name = name.substr(6);
    std::transform(name.begin(), name.end(), name.begin(), [](unsigned char ch){ return static_cast<char>(std::toupper(ch)); });

    // Flatten a single scalar argument.
    auto scalar = [&](std::size_t index, CellValue fallback) -> CellValue {
        if (index >= args.size() || args[index].empty()) return fallback;
        return args[index].front();
    };
    // First numeric scalar argument.
    auto firstNumeric = [&](std::size_t index, Numeric& out) -> bool {
        if (index >= args.size() || args[index].empty()) { out = Numeric{0.0, false, CellError::Value}; return false; }
        out = toNumeric(args[index].front());
        return true;
    };

    const auto errorFromArgs = [&]() -> CellValue {
        for (const auto& group : args)
            for (const auto& value : group)
                if (isErrorValue(value)) return value;
        return xlpp::CellValue{CellError::NotAvailable};
    };

    if (name == "SUM" || name == "PRODUCT") {
        double total = 0.0;
        if (name == "PRODUCT") total = 1.0;
        for (const auto& group : args) {
            for (const auto& value : group) {
                if (isErrorValue(value)) return value;
                const auto n = toNumeric(value);
                if (n.isError) return errorValue(n.error);
                total = name == "SUM" ? total + n.value : total * n.value;
            }
        }
        return xlpp::CellValue{total};
    }
    if (name == "AVERAGE") {
        double total = 0.0;
        std::size_t count = 0;
        for (const auto& group : args) {
            for (const auto& value : group) {
                if (isErrorValue(value)) return value;
                const auto n = toNumeric(value);
                if (n.isError) return errorValue(n.error);
                total += n.value;
                ++count;
            }
        }
        if (count == 0) return errorValue(CellError::DivisionByZero);
        return xlpp::CellValue{total / static_cast<double>(count)};
    }
    if (name == "MIN" || name == "MAX") {
        bool any = false;
        double extreme = name == "MIN" ? std::numeric_limits<double>::max() : -std::numeric_limits<double>::max();
        for (const auto& group : args) {
            for (const auto& value : group) {
                if (isErrorValue(value)) return value;
                const auto n = toNumeric(value);
                if (n.isError) return errorValue(n.error);
                any = true;
                extreme = name == "MIN" ? std::min(extreme, n.value) : std::max(extreme, n.value);
            }
        }
        if (!any) return xlpp::CellValue{0.0};
        return xlpp::CellValue{extreme};
    }
    if (name == "COUNT" || name == "COUNTA") {
        std::size_t count = 0;
        for (const auto& group : args) {
            for (const auto& value : group) {
                if (isErrorValue(value)) return value;
                if (name == "COUNTA") {
                    if (!isEmptyValue(value)) ++count;
                } else if (std::get_if<double>(&value) || std::get_if<DateTime>(&value)) {
                    ++count;
                }
            }
        }
        return xlpp::CellValue{static_cast<double>(count)};
    }
    if (name == "IF") {
        const auto condition = toBooleanValue(scalar(0, xlpp::CellValue{false}));
        if (condition.isError) return errorValue(condition.error);
        if (args.size() < 3) return condition.value ? scalar(1, xlpp::CellValue{false}) : xlpp::CellValue{false};
        return condition.value ? scalar(1, xlpp::CellValue{false}) : scalar(2, xlpp::CellValue{false});
    }
    if (name == "IFERROR") {
        const auto candidate = scalar(0, xlpp::CellValue{});
        if (isErrorValue(candidate)) return scalar(1, xlpp::CellValue{});
        return candidate;
    }
    if (name == "IFNA") {
        const auto candidate = scalar(0, xlpp::CellValue{});
        if (isErrorValue(candidate) && errorOf(candidate) == CellError::NotAvailable) return scalar(1, xlpp::CellValue{});
        return candidate;
    }
    if (name == "AND" || name == "OR" || name == "XOR") {
        bool any = false;
        std::size_t trueCount = 0;
        std::size_t count = 0;
        for (const auto& group : args) {
            for (const auto& value : group) {
                if (isErrorValue(value)) return value;
                const auto b = toBooleanValue(value);
                if (b.isError) return errorValue(b.error);
                any = any || b.value;
                if (b.value) ++trueCount;
                ++count;
            }
        }
        if (name == "AND") return xlpp::CellValue{count > 0 && trueCount == count};
        if (name == "OR") return xlpp::CellValue{any};
        return xlpp::CellValue{(trueCount % 2) == 1};
    }
    if (name == "NOT") {
        const auto b = toBooleanValue(scalar(0, xlpp::CellValue{false}));
        if (b.isError) return errorValue(b.error);
        return xlpp::CellValue{!b.value};
    }
    if (name == "ABS" || name == "SIGN" || name == "INT" || name == "SQRT" || name == "EXP"
        || name == "LN" || name == "LOG10" || name == "CEILING" || name == "FLOOR") {
        Numeric n;
        firstNumeric(0, n);
        if (n.isError) return errorValue(n.error);
        double result = 0.0;
        if (name == "ABS") result = std::abs(n.value);
        else if (name == "SIGN") result = (n.value > 0) ? 1.0 : (n.value < 0 ? -1.0 : 0.0);
        else if (name == "INT") result = std::floor(n.value);
        else if (name == "SQRT") { if (n.value < 0) return errorValue(CellError::Number); result = std::sqrt(n.value); }
        else if (name == "EXP") result = std::exp(n.value);
        else if (name == "LN") { if (n.value <= 0) return errorValue(CellError::Number); result = std::log(n.value); }
        else if (name == "LOG10") { if (n.value <= 0) return errorValue(CellError::Number); result = std::log10(n.value); }
        else if (name == "CEILING") result = std::ceil(n.value);
        else if (name == "FLOOR") result = std::floor(n.value);
        return xlpp::CellValue{result};
    }
    if (name == "ROUND" || name == "ROUNDUP" || name == "ROUNDDOWN" || name == "MROUND") {
        Numeric n;
        firstNumeric(0, n);
        if (n.isError) return errorValue(n.error);
        double places = 0.0;
        if (args.size() > 1) { const auto p = toNumeric(args[1].front()); if (p.isError) return errorValue(p.error); places = p.value; }
        const double factor = std::pow(10.0, places);
        double scaled = n.value * factor;
        double rounded = 0.0;
        if (name == "ROUND") rounded = std::round(scaled);
        else if (name == "ROUNDUP") rounded = std::ceil(scaled);
        else rounded = std::floor(scaled);
        return xlpp::CellValue{rounded / factor};
    }
    if (name == "TRUNC") {
        Numeric n;
        firstNumeric(0, n);
        if (n.isError) return errorValue(n.error);
        double places = 0.0;
        if (args.size() > 1) { const auto p = toNumeric(args[1].front()); if (p.isError) return errorValue(p.error); places = p.value; }
        const double factor = std::pow(10.0, places);
        return xlpp::CellValue{std::trunc(n.value * factor) / factor};
    }
    if (name == "MOD") {
        const auto n = toNumeric(scalar(0, xlpp::CellValue{0.0}));
        const auto d = toNumeric(scalar(1, xlpp::CellValue{0.0}));
        if (n.isError) return errorValue(n.error);
        if (d.isError) return errorValue(d.error);
        if (d.value == 0.0) return errorValue(CellError::DivisionByZero);
        return xlpp::CellValue{std::fmod(n.value, d.value)};
    }
    if (name == "POWER") {
        const auto n = toNumeric(scalar(0, xlpp::CellValue{0.0}));
        const auto p = toNumeric(scalar(1, xlpp::CellValue{0.0}));
        if (n.isError) return errorValue(n.error);
        if (p.isError) return errorValue(p.error);
        return xlpp::CellValue{std::pow(n.value, p.value)};
    }
    if (name == "LOG") {
        const auto n = toNumeric(scalar(0, xlpp::CellValue{0.0}));
        if (n.isError) return errorValue(n.error);
        if (n.value <= 0) return errorValue(CellError::Number);
        if (args.size() > 1) {
            const auto b = toNumeric(scalar(1, xlpp::CellValue{10.0}));
            if (b.isError) return errorValue(b.error);
            if (b.value <= 0 || b.value == 1.0) return errorValue(CellError::DivisionByZero);
            return xlpp::CellValue{std::log(n.value) / std::log(b.value)};
        }
        return xlpp::CellValue{std::log10(n.value)};
    }
    if (name == "PI") return xlpp::CellValue{3.14159265358979323846};
    if (name == "RAND") return xlpp::CellValue{static_cast<double>(std::rand()) / RAND_MAX};

    if (name == "CONCATENATE" || name == "CONCAT") {
        std::string result;
        for (const auto& group : args)
            for (const auto& value : group) {
                if (isErrorValue(value)) return value;
                result += toTextValue(value);
            }
        return xlpp::CellValue{result};
    }
    if (name == "LEN") {
        const auto value = scalar(0, xlpp::CellValue{});
        if (isErrorValue(value)) return value;
        return xlpp::CellValue{static_cast<double>(toTextValue(value).size())};
    }
    if (name == "UPPER" || name == "LOWER" || name == "TRIM") {
        const auto value = scalar(0, xlpp::CellValue{});
        if (isErrorValue(value)) return value;
        std::string text = toTextValue(value);
        if (name == "TRIM") {
            std::string trimmed;
            bool space = true;
            for (char ch : text) {
                if (std::isspace(static_cast<unsigned char>(ch))) {
                    if (!space) trimmed.push_back(' ');
                    space = true;
                } else { trimmed.push_back(ch); space = false; }
            }
            if (!trimmed.empty() && trimmed.back() == ' ') trimmed.pop_back();
            return xlpp::CellValue{trimmed};
        }
        if (name == "UPPER") std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch){ return static_cast<char>(std::toupper(ch)); });
        else std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch){ return static_cast<char>(std::tolower(ch)); });
        return xlpp::CellValue{text};
    }
    if (name == "LEFT" || name == "RIGHT" || name == "MID") {
        const auto value = scalar(0, xlpp::CellValue{});
        if (isErrorValue(value)) return value;
        const std::string text = toTextValue(value);
        if (name == "LEFT" || name == "RIGHT") {
            double count = 1.0;
            if (args.size() > 1) { const auto c = toNumeric(args[1].front()); if (c.isError) return errorValue(c.error); count = c.value; }
            const std::size_t n = static_cast<std::size_t>(std::max(0.0, count));
            if (name == "LEFT") return xlpp::CellValue{text.substr(0, std::min(n, text.size()))};
            return xlpp::CellValue{n >= text.size() ? text : text.substr(text.size() - n)};
        }
        double startNum = 1.0, countNum = std::numeric_limits<double>::infinity();
        if (args.size() > 1) { const auto s = toNumeric(args[1].front()); if (s.isError) return errorValue(s.error); startNum = s.value; }
        if (args.size() > 2) { const auto c = toNumeric(args[2].front()); if (c.isError) return errorValue(c.error); countNum = c.value; }
        if (startNum < 1 || countNum < 0) return errorValue(CellError::Value);
        const std::size_t start = static_cast<std::size_t>(startNum) - 1;
        if (start >= text.size()) return xlpp::CellValue{std::string{}};
        const std::size_t count = static_cast<std::size_t>(std::min<double>(countNum, text.size() - start));
        return xlpp::CellValue{text.substr(start, count)};
    }
    if (name == "REPT") {
        const auto text = toTextValue(scalar(0, xlpp::CellValue{}));
        const auto n = toNumeric(scalar(1, xlpp::CellValue{0.0}));
        if (n.isError) return errorValue(n.error);
        if (n.value < 0) return errorValue(CellError::Value);
        std::string result;
        const std::size_t count = static_cast<std::size_t>(n.value);
        for (std::size_t i = 0; i < count; ++i) result += text;
        return xlpp::CellValue{result};
    }
    if (name == "EXACT") {
        const auto a = toTextValue(scalar(0, xlpp::CellValue{}));
        const auto b = toTextValue(scalar(1, xlpp::CellValue{}));
        return xlpp::CellValue{a == b};
    }
    if (name == "FIND" || name == "SEARCH") {
        const auto needle = toTextValue(scalar(0, xlpp::CellValue{}));
        const auto haystack = toTextValue(scalar(1, xlpp::CellValue{}));
        double startNum = 1.0;
        if (args.size() > 2) { const auto s = toNumeric(args[2].front()); if (s.isError) return errorValue(s.error); startNum = s.value; }
        if (startNum < 1) return errorValue(CellError::Value);
        std::string lowerNeedle = needle, lowerHaystack = haystack;
        if (name == "SEARCH") {
            std::transform(lowerNeedle.begin(), lowerNeedle.end(), lowerNeedle.begin(), [](unsigned char ch){ return static_cast<char>(std::tolower(ch)); });
            std::transform(lowerHaystack.begin(), lowerHaystack.end(), lowerHaystack.begin(), [](unsigned char ch){ return static_cast<char>(std::tolower(ch)); });
        }
        const std::size_t start = static_cast<std::size_t>(startNum) - 1;
        const auto pos = (name == "SEARCH" ? lowerHaystack : haystack).find(name == "SEARCH" ? lowerNeedle : needle, start);
        if (pos == std::string::npos) return errorValue(CellError::Value);
        return xlpp::CellValue{static_cast<double>(pos + 1)};
    }
    if (name == "SUBSTITUTE") {
        const auto text = toTextValue(scalar(0, xlpp::CellValue{}));
        const auto oldText = toTextValue(scalar(1, xlpp::CellValue{}));
        const auto newText = toTextValue(scalar(2, xlpp::CellValue{}));
        if (oldText.empty()) return xlpp::CellValue{text};
        if (args.size() > 3) {
            const auto n = toNumeric(args[3].front());
            if (n.isError) return errorValue(n.error);
            std::size_t occurrence = static_cast<std::size_t>(n.value);
            std::size_t pos = 0, current = 0;
            while (true) {
                const auto found = text.find(oldText, pos);
                if (found == std::string::npos) break;
                ++current;
                if (current == occurrence) {
                    std::string result = text.substr(0, found) + newText + text.substr(found + oldText.size());
                    return xlpp::CellValue{result};
                }
                pos = found + oldText.size();
            }
            return xlpp::CellValue{text};
        }
        std::string result = text;
        std::size_t pos = 0;
        while ((pos = result.find(oldText, pos)) != std::string::npos) {
            result.replace(pos, oldText.size(), newText);
            pos += newText.size();
        }
        return xlpp::CellValue{result};
    }
    if (name == "VALUE") {
        const auto value = scalar(0, xlpp::CellValue{});
        if (isErrorValue(value)) return value;
        if (const auto* d = std::get_if<double>(&value)) return *d;
        if (const auto* b = std::get_if<bool>(&value)) return xlpp::CellValue{*b ? 1.0 : 0.0};
        const auto parsed = parseNumberText(toTextValue(value));
        if (!parsed) return errorValue(CellError::Value);
        return xlpp::CellValue{*parsed};
    }
    if (name == "T") {
        const auto value = scalar(0, xlpp::CellValue{});
        if (std::get_if<std::string>(&value)) return value;
        return xlpp::CellValue{std::string{}};
    }

    // Lookup / reference
    if (name == "CHOOSE") {
        if (args.empty()) return errorValue(CellError::Value);
        const auto index = toNumeric(args[0].front());
        if (index.isError) return errorValue(index.error);
        const std::size_t i = static_cast<std::size_t>(index.value);
        if (i < 1 || i >= args.size()) return errorValue(CellError::Value);
        return args[i].empty() ? xlpp::CellValue{} : args[i].front();
    }
    if (name == "INDEX") {
        if (args.empty()) return errorValue(CellError::Value);
        const auto& array = args[0];
        if (array.empty()) return errorValue(CellError::Value);
        double rowNum = 0.0;
        if (args.size() > 1) { const auto n = toNumeric(args[1].front()); if (n.isError) return errorValue(n.error); rowNum = n.value; }
        double colNum = 0.0;
        if (args.size() > 2) { const auto n = toNumeric(args[2].front()); if (n.isError) return errorValue(n.error); colNum = n.value; }
        std::size_t index = 0;
        if (colNum > 0) index = static_cast<std::size_t>(colNum) - 1;
        else if (rowNum > 0) index = static_cast<std::size_t>(rowNum) - 1;
        if (index >= array.size()) return errorValue(CellError::Reference);
        return array[index];
    }
    if (name == "MATCH") {
        if (args.size() < 2) return errorValue(CellError::Value);
        const auto lookup = scalar(0, xlpp::CellValue{});
        const auto& array = args[1];
        double matchType = 1.0;
        if (args.size() > 2) { const auto m = toNumeric(args[2].front()); if (m.isError) return errorValue(m.error); matchType = m.value; }
        if (matchType == 0) {
            for (std::size_t i = 0; i < array.size(); ++i) {
                const auto a = toTextValue(array[i]);
                const auto b = toTextValue(lookup);
                if (a == b) return xlpp::CellValue{static_cast<double>(i + 1)};
            }
            return errorValue(CellError::NotAvailable);
        }
        for (std::size_t i = 0; i < array.size(); ++i) {
            const auto a = toNumeric(array[i]);
            const auto b = toNumeric(lookup);
            if (a.isError || b.isError) return errorValue(CellError::Value);
            const bool hit = matchType < 0 ? a.value <= b.value : a.value >= b.value;
            if (hit) return xlpp::CellValue{static_cast<double>(i + 1)};
        }
        return errorValue(CellError::NotAvailable);
    }
    if (name == "VLOOKUP" || name == "HLOOKUP") {
        if (args.size() < 3) return errorValue(CellError::Value);
        const auto lookup = scalar(0, xlpp::CellValue{});
        const auto& table = args[1];
        const auto colIndex = toNumeric(args[2].front());
        if (colIndex.isError) return errorValue(colIndex.error);
        const bool approximate = args.size() < 4 || toBooleanValue(args[3].front()).value;
        const std::size_t columns = name == "VLOOKUP" ? (table.empty() ? 0 : 1) : (table.size() == 1 ? 1 : 0);
        const std::size_t rows = name == "VLOOKUP" ? table.size() : 1;
        if (name == "VLOOKUP") {
            for (std::size_t r = 0; r < rows; ++r) {
                const auto candidate = table[r];
                const auto a = toNumeric(candidate);
                const auto b = toNumeric(lookup);
                if (a.isError || b.isError) {
                    if (toTextValue(candidate) == toTextValue(lookup)) {
                        const std::size_t col = static_cast<std::size_t>(colIndex.value);
                        if (col >= 1 && col <= columns) return table[r];
                        return errorValue(CellError::Reference);
                    }
                    continue;
                }
                const bool hit = approximate ? a.value <= b.value : a.value == b.value;
                if (hit) {
                    const std::size_t col = static_cast<std::size_t>(colIndex.value);
                    if (col >= 1 && col <= columns) return table[r];
                    return errorValue(CellError::Reference);
                }
            }
        }
        return errorValue(CellError::NotAvailable);
    }

    // Information
    if (name == "ISNUMBER") { const auto v = scalar(0, xlpp::CellValue{}); return xlpp::CellValue{std::get_if<double>(&v) != nullptr}; }
    if (name == "ISTEXT") { const auto v = scalar(0, xlpp::CellValue{}); return xlpp::CellValue{std::get_if<std::string>(&v) != nullptr}; }
    if (name == "ISBLANK") { const auto v = scalar(0, xlpp::CellValue{}); return xlpp::CellValue{isEmptyValue(v)}; }
    if (name == "ISLOGICAL") { const auto v = scalar(0, xlpp::CellValue{}); return xlpp::CellValue{std::get_if<bool>(&v) != nullptr}; }
    if (name == "ISERROR") { const auto v = scalar(0, xlpp::CellValue{}); return xlpp::CellValue{isErrorValue(v)}; }
    if (name == "ISNA") {
        const auto v = scalar(0, xlpp::CellValue{});
        return xlpp::CellValue{isErrorValue(v) && errorOf(v) == CellError::NotAvailable};
    }
    if (name == "N") {
        const auto value = scalar(0, xlpp::CellValue{});
        if (const auto* d = std::get_if<double>(&value)) return *d;
        if (const auto* b = std::get_if<bool>(&value)) return xlpp::CellValue{*b ? 1.0 : 0.0};
        if (const auto* d = std::get_if<DateTime>(&value)) return xlpp::CellValue{xlpp::toExcelSerial(*d, false)};
        if (isErrorValue(value)) return value;
        return xlpp::CellValue{0.0};
    }
    if (name == "NA") return errorValue(CellError::NotAvailable);
    if (name == "ERROR.TYPE") {
        const auto v = scalar(0, xlpp::CellValue{});
        if (!isErrorValue(v)) return errorValue(CellError::NotAvailable);
        switch (errorOf(v)) {
            case CellError::Null: return xlpp::CellValue{1.0};
            case CellError::DivisionByZero: return xlpp::CellValue{2.0};
            case CellError::Value: return xlpp::CellValue{3.0};
            case CellError::Reference: return xlpp::CellValue{4.0};
            case CellError::Name: return xlpp::CellValue{5.0};
            case CellError::Number: return xlpp::CellValue{6.0};
            case CellError::NotAvailable: return xlpp::CellValue{7.0};
            case CellError::GettingData: return xlpp::CellValue{8.0};
        }
        return errorValue(CellError::NotAvailable);
    }

    // Date / time
    if (name == "TODAY" || name == "NOW") {
        if (!ctx_.options.evaluateVolatileFunctions) return errorValue(CellError::NotAvailable);
        const std::time_t now = std::time(nullptr);
        std::tm local{};
#if defined(_WIN32)
        localtime_s(&local, &now);
#else
        localtime_r(&now, &local);
#endif
        const auto serial = xlpp::toExcelSerial(xlpp::DateTime{static_cast<std::int32_t>(local.tm_year + 1900),
            static_cast<std::int32_t>(local.tm_mon + 1), static_cast<std::int32_t>(local.tm_mday)}, false);
        if (name == "TODAY") return xlpp::CellValue{serial};
        const double timePart = (static_cast<double>(local.tm_hour) * 3600.0 + static_cast<double>(local.tm_min) * 60.0 + static_cast<double>(local.tm_sec)) / 86400.0;
        return xlpp::CellValue{serial + timePart};
    }
    if (name == "YEAR" || name == "MONTH" || name == "DAY" || name == "HOUR" || name == "MINUTE" || name == "SECOND") {
        const auto value = scalar(0, xlpp::CellValue{0.0});
        if (isErrorValue(value)) return value;
        const auto n = toNumeric(value);
        if (n.isError) return errorValue(n.error);
        const auto date = xlpp::fromExcelSerial(n.value, false);
        switch (name[0]) {
            case 'Y': return xlpp::CellValue{static_cast<double>(date.year)};
            case 'M': return xlpp::CellValue{static_cast<double>(date.month)};
            case 'D': return xlpp::CellValue{static_cast<double>(date.day)};
            case 'H': return xlpp::CellValue{static_cast<double>(date.hour)};
            case 'I': return xlpp::CellValue{static_cast<double>(date.minute)};
            case 'S': return xlpp::CellValue{static_cast<double>(date.second)};
        }
        return errorValue(CellError::Value);
    }
    if (name == "DATE") {
        const auto y = toNumeric(scalar(0, xlpp::CellValue{0.0}));
        const auto m = toNumeric(scalar(1, xlpp::CellValue{0.0}));
        const auto d = toNumeric(scalar(2, xlpp::CellValue{0.0}));
        if (y.isError || m.isError || d.isError) return errorValue(CellError::Value);
        return xlpp::CellValue{xlpp::toExcelSerial(xlpp::DateTime{static_cast<std::int32_t>(y.value),
            static_cast<std::int32_t>(m.value), static_cast<std::int32_t>(d.value)}, false)};
    }

    // Unsupported function.
    if (ctx_.unsupportedFormulas) ++(*ctx_.unsupportedFormulas);
    return errorValue(CellError::Name);
}

} // namespace

namespace xlpp {
namespace internal {

CellValue evaluateFormula(const CalculationEvaluator& ctx, const std::string& formula) {
    if (ctx.depth >= ctx.options.maxDepth) return CellValue{CellError::Number};
    std::string_view source(formula);
    if (!source.empty() && source.front() == '=') source.remove_prefix(1);
    if (source.empty()) return CellValue{0.0};
    Parser parser(ctx, source);
    try {
        return parser.parse();
    } catch (const std::exception&) {
        return CellValue{CellError::Value};
    }
}

CellValue resolveReference(const CalculationEvaluator& ctx,
                           const std::string& sheetName,
                           const std::string& cellAddress) {
    auto ref = CellReference::parse(cellAddress);
    const auto* sheet = ctx.workbook.worksheet(sheetName);
    if (!sheet) return CellValue{CellError::Reference};
    const auto* cell = sheet->tryCell(ref.row, ref.column);
    if (!cell) return CellValue{};
    const CellValue& value = cell->value();
    if (!cell->hasFormula()) return value;

    const std::string key = sheetName + std::string(kKeySeparator) + cellAddress;
    if (std::find(ctx.stack.begin(), ctx.stack.end(), key) != ctx.stack.end()) {
        // Circular reference: report once, fall back to the cached value.
        if (ctx.circularReferences) ++(*ctx.circularReferences);
        if (!isEmptyValue(value)) return value;
        return CellValue{0.0};
    }
    const auto memoized = ctx.memo.find(key);
    if (memoized != ctx.memo.end()) return memoized->second;

    ctx.stack.push_back(key);
    CalculationEvaluator nested = ctx;
    nested.depth = ctx.depth + 1;
    const auto result = evaluateFormula(nested, cell->formula());
    ctx.stack.pop_back();
    ctx.memo.emplace(key, result);
    return result;
}

} // namespace internal
} // namespace xlpp
