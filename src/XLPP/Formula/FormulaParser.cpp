#include "Formula/FormulaEvaluationContext.h"

#include <XLPP/Cell/CellReference.h>

#include <array>
#include <cctype>
#include <cmath>
#include <stdexcept>
#include <unordered_map>
#include <utility>

namespace xlpp::internal::formula {
namespace {

class Parser {
public:
    Parser(FormulaEvaluationContext& engine, Worksheet& sheet, std::string formula, std::size_t depth)
        : engine_(engine), sheet_(sheet), formula_(std::move(formula)), depth_(depth) {
        if (!formula_.empty() && formula_.front() == '=') formula_.erase(formula_.begin());
        next();
    }

    EvalValue parse();

private:
    enum class Kind {
        End, Number, String, Identifier, Error,
        Plus, Minus, Star, Slash, Caret, Amp, Percent,
        Equal, NotEqual, Less, LessEqual, Greater, GreaterEqual,
        LParen, RParen, LBrace, RBrace, Comma, Semicolon, Colon, Bang
    };
    struct Token { Kind kind{Kind::End}; std::string text; double number{0.0}; };

    FormulaEvaluationContext& engine_;
    Worksheet& sheet_;
    std::string formula_;
    std::size_t depth_{};
    std::size_t pos_{};
    Token token_;
    std::unordered_map<std::string, EvalValue> locals_;

    void next();
    EvalValue comparison();
    EvalValue concat();
    EvalValue additive();
    EvalValue multiplicative();
    EvalValue power();
    EvalValue unary();
    EvalValue postfix();
    EvalValue primary();
    EvalValue referenceOrName(std::string name);
    EvalValue functionCall(std::string name);
    [[noreturn]] void syntax(const std::string& message) const { throw std::runtime_error(message); }
};

void Parser::next() {
    while (pos_ < formula_.size() && std::isspace(static_cast<unsigned char>(formula_[pos_]))) ++pos_;
    if (pos_ >= formula_.size()) { token_ = {}; return; }
    const char c = formula_[pos_];
    auto one = [&](Kind kind) { token_ = {kind, std::string(1, formula_[pos_++]), 0.0}; };
    switch (c) {
        case '+': return one(Kind::Plus); case '-': return one(Kind::Minus); case '*': return one(Kind::Star);
        case '/': return one(Kind::Slash); case '^': return one(Kind::Caret); case '&': return one(Kind::Amp);
        case '%': return one(Kind::Percent); case '(': return one(Kind::LParen); case ')': return one(Kind::RParen);
        case '{': return one(Kind::LBrace); case '}': return one(Kind::RBrace);
        case ',': return one(Kind::Comma); case ';': return one(Kind::Semicolon); case ':': return one(Kind::Colon);
        case '!': return one(Kind::Bang); case '=': return one(Kind::Equal);
        case '<':
            ++pos_; if (pos_ < formula_.size() && formula_[pos_] == '=') { ++pos_; token_ = {Kind::LessEqual, "<=", 0}; }
            else if (pos_ < formula_.size() && formula_[pos_] == '>') { ++pos_; token_ = {Kind::NotEqual, "<>", 0}; }
            else token_ = {Kind::Less, "<", 0};
            return;
        case '>':
            ++pos_; if (pos_ < formula_.size() && formula_[pos_] == '=') { ++pos_; token_ = {Kind::GreaterEqual, ">=", 0}; }
            else token_ = {Kind::Greater, ">", 0};
            return;
        case '"': {
            ++pos_; std::string text;
            while (pos_ < formula_.size()) {
                if (formula_[pos_] == '"') {
                    if (pos_ + 1 < formula_.size() && formula_[pos_ + 1] == '"') { text.push_back('"'); pos_ += 2; continue; }
                    ++pos_; token_ = {Kind::String, std::move(text), 0}; return;
                }
                text.push_back(formula_[pos_++]);
            }
            syntax("Unterminated formula string");
        }
        case '\'': {
            ++pos_; std::string text;
            while (pos_ < formula_.size()) {
                if (formula_[pos_] == '\'') {
                    if (pos_ + 1 < formula_.size() && formula_[pos_ + 1] == '\'') { text.push_back('\''); pos_ += 2; continue; }
                    ++pos_; token_ = {Kind::Identifier, std::move(text), 0}; return;
                }
                text.push_back(formula_[pos_++]);
            }
            syntax("Unterminated quoted worksheet name");
        }
        case '#': {
            static constexpr std::array<std::string_view,10> errors{"#GETTING_DATA","#DIV/0!","#VALUE!","#NULL!","#REF!","#NAME?","#NUM!","#N/A","#SPILL!","#CALC!"};
            for(auto error:errors)if(formula_.compare(pos_,error.size(),error)==0){pos_+=error.size();token_={Kind::Error,std::string(error),0};return;}
            syntax("Unknown Excel error literal");
        }
        default: break;
    }
    if (c == '[') {
        const auto start = pos_;
        const auto close = formula_.find(']', pos_ + 1);
        if (close == std::string::npos) syntax("Unterminated external workbook reference");
        pos_ = close + 1;
        while (pos_ < formula_.size()) {
            const char ch = formula_[pos_];
            if (ch == '!' || ch == ':' || ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '^' ||
                ch == '&' || ch == '=' || ch == '<' || ch == '>' || ch == '(' || ch == ')' || ch == ',' || ch == ';' ||
                std::isspace(static_cast<unsigned char>(ch))) break;
            ++pos_;
        }
        token_ = {Kind::Identifier, formula_.substr(start, pos_ - start), 0};
        return;
    }
    if (std::isdigit(static_cast<unsigned char>(c)) || (c == '.' && pos_ + 1 < formula_.size() && std::isdigit(static_cast<unsigned char>(formula_[pos_ + 1])))) {
        const char* begin = formula_.c_str() + pos_; char* end = nullptr; const double value = std::strtod(begin, &end);
        if (end == begin) syntax("Invalid number");
        const auto consumed = static_cast<std::size_t>(end - begin); token_ = {Kind::Number, formula_.substr(pos_, consumed), value}; pos_ += consumed; return;
    }
    if (std::isalpha(static_cast<unsigned char>(c)) || c == '_' || c == '$' || c == '\\') {
        const auto start = pos_++;
        while (pos_ < formula_.size()) {
            const unsigned char ch = static_cast<unsigned char>(formula_[pos_]);
            if (!(std::isalnum(ch) || formula_[pos_] == '_' || formula_[pos_] == '.' || formula_[pos_] == '$' || formula_[pos_] == '\\')) break;
            ++pos_;
        }
        if (pos_ < formula_.size() && formula_[pos_] == '[') {
            int brackets = 0;
            do {
                if (formula_[pos_] == '[') ++brackets;
                else if (formula_[pos_] == ']') --brackets;
                ++pos_;
            } while (pos_ < formula_.size() && brackets > 0);
            if (brackets != 0) syntax("Unterminated structured table reference");
        }
        token_ = {Kind::Identifier, formula_.substr(start, pos_ - start), 0}; return;
    }
    syntax(std::string("Unexpected character in formula: ") + c);
}

EvalValue Parser::parse() {
    auto value = comparison();
    if (token_.kind != Kind::End) syntax("Unexpected trailing formula token: " + token_.text);
    return value;
}

EvalValue Parser::comparison() {
    auto lhs = concat();
    while (token_.kind == Kind::Equal || token_.kind == Kind::NotEqual || token_.kind == Kind::Less || token_.kind == Kind::LessEqual || token_.kind == Kind::Greater || token_.kind == Kind::GreaterEqual) {
        const auto op = token_.text; next(); auto rhs = concat();
        auto l = firstScalar(lhs), r = firstScalar(rhs); if (isError(l)) return EvalValue::fromScalar(l);
        if (isError(r)) return EvalValue::fromScalar(r);
        lhs = EvalValue::fromScalar(engine_.compare(l, r, op));
    }
    return lhs;
}

EvalValue Parser::concat() {
    auto lhs = additive();
    while (token_.kind == Kind::Amp) {
        next(); auto rhs = additive(); auto l = firstScalar(lhs), r = firstScalar(rhs);
        if (isError(l)) return EvalValue::fromScalar(l);
        if (isError(r)) return EvalValue::fromScalar(r);
        lhs = EvalValue::fromScalar(scalarText(l, engine_.date1904()) + scalarText(r, engine_.date1904()));
    }
    return lhs;
}

EvalValue Parser::additive() {
    auto lhs = multiplicative();
    while (token_.kind == Kind::Plus || token_.kind == Kind::Minus) {
        const auto op = token_.kind; next(); auto rhs = multiplicative(); auto l = firstScalar(lhs), r = firstScalar(rhs);
        if (isError(l)) return EvalValue::fromScalar(l);
        if (isError(r)) return EvalValue::fromScalar(r);
        const auto a = numberValue(l, engine_.date1904()), b = numberValue(r, engine_.date1904()); if (!a || !b) return EvalValue::fromScalar(CellError::Value);
        lhs = EvalValue::fromScalar(op == Kind::Plus ? *a + *b : *a - *b);
    }
    return lhs;
}

EvalValue Parser::multiplicative() {
    auto lhs = power();
    while (token_.kind == Kind::Star || token_.kind == Kind::Slash) {
        const auto op = token_.kind; next(); auto rhs = power(); auto l = firstScalar(lhs), r = firstScalar(rhs);
        if (isError(l)) return EvalValue::fromScalar(l);
        if (isError(r)) return EvalValue::fromScalar(r);
        const auto a = numberValue(l, engine_.date1904()), b = numberValue(r, engine_.date1904()); if (!a || !b) return EvalValue::fromScalar(CellError::Value);
        if (op == Kind::Slash && *b == 0.0) lhs = EvalValue::fromScalar(CellError::DivisionByZero);
        else lhs = EvalValue::fromScalar(op == Kind::Star ? *a * *b : *a / *b);
    }
    return lhs;
}

EvalValue Parser::power() {
    auto lhs = unary();
    if (token_.kind == Kind::Caret) {
        next(); auto rhs = power(); auto l = firstScalar(lhs), r = firstScalar(rhs);
        if (isError(l)) return EvalValue::fromScalar(l);
        if (isError(r)) return EvalValue::fromScalar(r);
        const auto a = numberValue(l, engine_.date1904()), b = numberValue(r, engine_.date1904()); if (!a || !b) return EvalValue::fromScalar(CellError::Value);
        const auto value = std::pow(*a, *b); return EvalValue::fromScalar(std::isfinite(value) ? Scalar{value} : Scalar{CellError::Number});
    }
    return lhs;
}

EvalValue Parser::unary() {
    if (token_.kind == Kind::Plus) { next(); return unary(); }
    if (token_.kind == Kind::Minus) {
        next(); auto v = firstScalar(unary()); if (isError(v)) return EvalValue::fromScalar(v);
        const auto n = numberValue(v, engine_.date1904()); return EvalValue::fromScalar(n ? Scalar{-*n} : Scalar{CellError::Value});
    }
    return postfix();
}

EvalValue Parser::postfix() {
    auto value = primary();
    while (token_.kind == Kind::Percent) {
        next(); auto scalar = firstScalar(value); if (isError(scalar)) return EvalValue::fromScalar(scalar);
        const auto n = numberValue(scalar, engine_.date1904()); value = EvalValue::fromScalar(n ? Scalar{*n / 100.0} : Scalar{CellError::Value});
    }
    return value;
}

EvalValue Parser::primary() {
    if (token_.kind == Kind::Number) { const double value = token_.number; next(); return EvalValue::fromScalar(value); }
    if (token_.kind == Kind::String) { auto value = token_.text; next(); return EvalValue::fromScalar(std::move(value)); }
    if (token_.kind == Kind::Error) { auto text = token_.text; next(); return EvalValue::fromScalar(cellErrorFromString(text)); }
    if (token_.kind == Kind::Identifier) { auto name = token_.text; next(); return referenceOrName(std::move(name)); }
    if (token_.kind == Kind::LBrace) {
        next();
        std::vector<Scalar> values;
        std::size_t rows = 1, columns = 0, currentColumns = 0;
        if (token_.kind == Kind::RBrace) syntax("Empty array constant");
        for (;;) {
            auto item = comparison();
            if (item.isRange) syntax("Array constant elements must be scalar");
            values.push_back(item.scalar); ++currentColumns;
            if (token_.kind == Kind::Comma) { next(); continue; }
            if (token_.kind == Kind::Semicolon) {
                if (columns == 0) columns = currentColumns;
                else if (columns != currentColumns) syntax("Array constant rows must have equal width");
                currentColumns = 0; ++rows; next(); continue;
            }
            break;
        }
        if (token_.kind != Kind::RBrace) syntax("Expected '}' after array constant");
        if (columns == 0) columns = currentColumns;
        else if (columns != currentColumns) syntax("Array constant rows must have equal width");
        next();
        return EvalValue::fromRange(std::move(values), rows, columns);
    }
    if (token_.kind == Kind::LParen) { next(); auto value = comparison(); if (token_.kind != Kind::RParen) syntax("Expected ')' in formula"); next(); return value; }
    syntax("Expected formula value");
}

EvalValue Parser::referenceOrName(std::string name) {
    if (token_.kind == Kind::LParen) return functionCall(std::move(name));
    const auto upper = upperAscii(name);
    if (const auto local = locals_.find(upper); local != locals_.end()) return local->second;
    if ((!name.empty() && name.front() != '[' && name.find('[') != std::string::npos) ||
        (!name.empty() && name.front() == '[' && token_.kind != Kind::Bang))
        return engine_.resolveStructuredReference(sheet_, name, depth_ + 1);
    if (upper == "TRUE") return EvalValue::fromScalar(true);
    if (upper == "FALSE") return EvalValue::fromScalar(false);

    std::string sheetName;
    std::string first = name;
    if (token_.kind == Kind::Bang) {
        sheetName = std::move(first); next();
        if (token_.kind != Kind::Identifier) return EvalValue::fromScalar(CellError::Reference);
        first = token_.text; next();
    }
    auto isCell = [](const std::string& value) {
        try { (void)CellReference::parse(value); return true; } catch (...) { return false; }
    };
    if (!sheetName.empty() && sheetName.front() == '[') {
        const auto close = sheetName.find(']');
        if (close == std::string::npos || close == 1 || close + 1 >= sheetName.size() || !isCell(first))
            return EvalValue::fromScalar(CellError::Reference);
        const auto workbookToken = sheetName.substr(1, close - 1);
        const auto externalSheet = sheetName.substr(close + 1);
        std::optional<std::string> last;
        if (token_.kind == Kind::Colon) {
            next();
            if (token_.kind != Kind::Identifier || !isCell(token_.text)) return EvalValue::fromScalar(CellError::Reference);
            last = token_.text; next();
        }
        return engine_.resolveExternalReference(workbookToken, externalSheet, first, last);
    }
    if (isCell(first)) {
        std::optional<std::string> last;
        if (token_.kind == Kind::Colon) {
            next(); if (token_.kind != Kind::Identifier || !isCell(token_.text)) return EvalValue::fromScalar(CellError::Reference);
            last = token_.text; next();
        }
        return engine_.resolveReference(sheet_, sheetName, first, last, depth_ + 1);
    }
    if (!sheetName.empty()) return EvalValue::fromScalar(CellError::Reference);
    return engine_.resolveDefinedName(sheet_, name, depth_ + 1);
}

EvalValue Parser::functionCall(std::string name) {
    next(); // '('
    std::string normalized = name;
    while (normalized.rfind("_xlfn.", 0) == 0) normalized.erase(0, 6);
    normalized = upperAscii(normalized);
    if (normalized == "LET") {
        const auto savedLocals = locals_;
        auto restore = [&]() { locals_ = savedLocals; };
        if (token_.kind != Kind::Identifier) { restore(); syntax("LET requires a variable name"); }
        for (;;) {
            if (token_.kind != Kind::Identifier) { restore(); syntax("LET requires a variable name"); }
            const auto variable = upperAscii(token_.text);
            next();
            if (token_.kind != Kind::Comma && token_.kind != Kind::Semicolon) { restore(); syntax("LET variable must be followed by a value"); }
            next();
            auto value = comparison();
            locals_[variable] = std::move(value);
            if (token_.kind != Kind::Comma && token_.kind != Kind::Semicolon) { restore(); syntax("LET requires a final calculation"); }
            next();

            // A subsequent binding has the lexical shape name,value.  If the
            // identifier is followed by any other operator/token it starts the
            // final calculation, so restore that token and parse normally.
            if (token_.kind == Kind::Identifier) {
                const auto candidate = token_; const auto candidatePos = pos_;
                next();
                if (token_.kind == Kind::Comma || token_.kind == Kind::Semicolon) {
                    token_ = candidate; pos_ = candidatePos;
                    continue;
                }
                token_ = candidate; pos_ = candidatePos;
            }
            auto result = comparison();
            if (token_.kind != Kind::RParen) { restore(); syntax("Expected ')' after LET calculation"); }
            next();
            restore();
            return result;
        }
    }

    std::vector<EvalValue> args;
    if (token_.kind != Kind::RParen) {
        for (;;) {
            args.push_back(comparison());
            if (token_.kind == Kind::Comma || token_.kind == Kind::Semicolon) { next(); continue; }
            break;
        }
    }
    if (token_.kind != Kind::RParen) syntax("Expected ')' after function arguments");
    next();
    return engine_.callFunction(std::move(name), args, sheet_, depth_ + 1);
}


} // namespace

EvalValue parseFormula(FormulaEvaluationContext& context, Worksheet& sheet,
                       std::string formula, std::size_t depth) {
    Parser parser(context, sheet, std::move(formula), depth);
    return parser.parse();
}

} // namespace xlpp::internal::formula
