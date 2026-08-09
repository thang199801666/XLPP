#include "Formula/FormulaFunctionSupport.h"

#include <cctype>
#include <string_view>

namespace xlpp::internal::formula {

std::vector<Scalar> flattened(const std::vector<EvalValue>& args) {
    std::vector<Scalar> out;
    for (const auto& arg : args) {
        if (arg.isRange) out.insert(out.end(), arg.range.begin(), arg.range.end());
        else out.push_back(arg.scalar);
    }
    return out;
}

bool wildcardMatchInsensitive(std::string_view text, std::string_view pattern) {
    auto fold=[](unsigned char c){return static_cast<char>(std::toupper(c));};
    std::size_t ti=0,pi=0,star=std::string_view::npos,match=0;
    while(ti<text.size()){
        if(pi<pattern.size()&&pattern[pi]=='~'&&pi+1<pattern.size()){
            if(fold(static_cast<unsigned char>(text[ti]))==fold(static_cast<unsigned char>(pattern[pi+1]))){++ti;pi+=2;continue;}
        } else if(pi<pattern.size()&&(pattern[pi]=='?'||fold(static_cast<unsigned char>(pattern[pi]))==fold(static_cast<unsigned char>(text[ti])))) {++ti;++pi;continue;}
        else if(pi<pattern.size()&&pattern[pi]=='*'){star=pi++;match=ti;continue;}
        if(star!=std::string_view::npos){pi=star+1;ti=++match;continue;}
        return false;
    }
    while(pi<pattern.size()&&pattern[pi]=='*')++pi;
    return pi==pattern.size();
}

std::optional<CellError> FormulaFunctionCall::error() const {
    for (const auto& value : flattened(args))
        if (const auto* error = std::get_if<CellError>(&value)) return *error;
    return std::nullopt;
}

Scalar FormulaFunctionCall::scalarArg(std::size_t index) const {
    return index < args.size() ? firstScalar(args[index]) : Scalar{std::monostate{}};
}

std::optional<double> FormulaFunctionCall::numArg(std::size_t index) const {
    return numberValue(scalarArg(index), engine.date1904());
}

std::optional<bool> FormulaFunctionCall::boolArg(std::size_t index) const {
    return boolValue(scalarArg(index), engine.date1904());
}

std::string FormulaFunctionCall::textArg(std::size_t index) const {
    return scalarText(scalarArg(index), engine.date1904());
}

std::vector<double> FormulaFunctionCall::numericValues() const {
    std::vector<double> out;
    for (const auto& value : flattened(args))
        if (auto number = numberValue(value, engine.date1904())) out.push_back(*number);
    return out;
}

bool FormulaFunctionCall::matchesCriterion(const Scalar& candidate, const Scalar& criterionValue) const {
    if (isError(candidate)) return false;
    if (!std::holds_alternative<std::string>(criterionValue)) return engine.compare(candidate, criterionValue, "=");
    std::string criterion=std::get<std::string>(criterionValue),op="=";
    for(const char* candidateOp:{">=","<=","<>",">","<","="}){
        if(criterion.rfind(candidateOp,0)==0){op=candidateOp;criterion.erase(0,std::char_traits<char>::length(candidateOp));break;}
    }
    if((op=="="||op=="<>")&&(criterion.find('*')!=std::string::npos||criterion.find('?')!=std::string::npos||criterion.find('~')!=std::string::npos)){
        const bool matched=wildcardMatchInsensitive(scalarText(candidate,engine.date1904()),criterion);return op=="="?matched:!matched;
    }
    Scalar right=criterion;if(auto number=parseNumberText(criterion))right=*number;else {auto upper=upperAscii(trimAscii(criterion));if(upper=="TRUE")right=true;else if(upper=="FALSE")right=false;}
    return engine.compare(candidate,right,op);
}

} // namespace xlpp::internal::formula
