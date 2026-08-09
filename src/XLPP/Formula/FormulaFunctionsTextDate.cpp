#include "Formula/FormulaFunctionSupport.h"
#include <algorithm>
#include <chrono>
#include <cmath>
namespace xlpp::internal::formula {
    FormulaFunctionResult evaluateLogicalTextDateFunctions(FormulaFunctionCall& call, const std::string& name) {
        auto& engine = call.engine;
        const auto& args = call.args;
        auto scalarArg = [&](std::size_t index) {
            return call.scalarArg(index);
        };
        auto numArg = [&](std::size_t index) {
            return call.numArg(index);
        };
        auto boolArg = [&](std::size_t index) {
            return call.boolArg(index);
        };
        auto textArg = [&](std::size_t index) {
            return call.textArg(index);
        };
        if (name == "LEN") return EvalValue::fromScalar(static_cast<double>(textArg(0).size()));
        if (name == "LOWER" || name == "UPPER") {
            auto text = textArg(0);
            std::transform(text.begin(), text.end(), text.begin(), [&](unsigned char c) { return static_cast<char>(name == "LOWER" ? std::tolower(c) : std::toupper(c)); });
            return EvalValue::fromScalar(std::move(text));
        }
        if (name == "TRIM") {
            std::istringstream in(textArg(0));
            std::string word, out;
            while (in >> word) {
                if (!out.empty()) out += ' ';
                out += word;
            }
            return EvalValue::fromScalar(std::move(out));
        }
        if (name == "LEFT" || name == "RIGHT") {
            const auto text = textArg(0);
            const auto n = args.size() >= 2 ? numArg(1) : std::optional<double> {
                1.0
            };
            if (!n || *n < 0) return EvalValue::fromScalar(CellError::Value);
            const auto count = std::min<std::size_t>(text.size(), static_cast<std::size_t>(*n));
            return EvalValue::fromScalar(name == "LEFT" ? text.substr(0, count) : text.substr(text.size() - count));
        }
        if (name == "MID") {
            const auto text = textArg(0);
            const auto start = numArg(1), countValue = numArg(2);
            if (!start || !countValue || *start < 1 || *countValue < 0) return EvalValue::fromScalar(CellError::Value);
            const auto begin = static_cast<std::size_t>(*start - 1);
            if (begin >= text.size()) return EvalValue::fromScalar(std::string{});
            return EvalValue::fromScalar(text.substr(begin, static_cast<std::size_t>(*countValue)));
        }
        if (name == "CONCAT" || name == "CONCATENATE") {
            std::string out;
            for (const auto& v : flattened(args)) {
                if (isError(v)) return EvalValue::fromScalar(errorOr(v));
                out += scalarText(v, engine.date1904());
            }
            return EvalValue::fromScalar(std::move(out));
        }
        if (name == "FIND" || name == "SEARCH") {
            if(args.size()<2)return EvalValue::fromScalar(CellError::Value);
            auto needleText=textArg(0),hay=textArg(1);
            auto start=args.size()>=3?numArg(2):std::optional<double> {
                1.0
            };
            if(!start||*start<1)return EvalValue::fromScalar(CellError::Value);
            std::size_t begin=static_cast<std::size_t>(*start-1);
            if(begin>hay.size())return EvalValue::fromScalar(CellError::Value);
            if(name=="SEARCH") {
                needleText=upperAscii(needleText);
                hay=upperAscii(hay);
            }
            auto pos=hay.find(needleText,begin);
            return EvalValue::fromScalar(pos==std::string::npos?Scalar{CellError::Value}:Scalar{static_cast<double>(pos+1)});
        }
        if(name=="SUBSTITUTE") {
            if(args.size()<3)return EvalValue::fromScalar(CellError::Value);
            auto text=textArg(0),oldText=textArg(1),newText=textArg(2);
            if(oldText.empty())return EvalValue::fromScalar(text);
            std::optional<std::size_t> instance;
            if(args.size()>=4) {
                auto n=numArg(3);
                if(!n||*n<1)return EvalValue::fromScalar(CellError::Value);
                instance=static_cast<std::size_t>(*n);
            }
            std::string out;
            std::size_t pos=0,seen=0;
            while(true) {
                auto found=text.find(oldText,pos);
                if(found==std::string::npos) {
                    out+=text.substr(pos);
                    break;
                }
                out+=text.substr(pos,found-pos);
                ++seen;
                if(!instance||seen==*instance)out+=newText;
                else out+=oldText;
                pos=found+oldText.size();
            }
            return EvalValue::fromScalar(std::move(out));
        }
        if(name=="EXACT")return EvalValue::fromScalar(textArg(0)==textArg(1));
        if(name=="VALUE") {
            auto n=parseNumberText(trimAscii(textArg(0)));
            return EvalValue::fromScalar(n?Scalar{*n}:Scalar{CellError::Value});
        }
        if(name=="REPT") {
            auto n=numArg(1);
            if(!n||*n<0||*n>32767)return EvalValue::fromScalar(CellError::Value);
            std::string out;
            auto text=textArg(0);
            for(std::size_t i=0;i<static_cast<std::size_t>(*n);++i)out+=text;
            return EvalValue::fromScalar(std::move(out));
        }
        if(name=="TEXTJOIN") {
            if(args.size()<3)return EvalValue::fromScalar(CellError::Value);
            auto delimiter=textArg(0);
            auto ignore=boolArg(1);
            if(!ignore)return EvalValue::fromScalar(CellError::Value);
            std::string out;
            bool first=true;
            for(std::size_t a=2;a<args.size();++a) {
                std::vector<Scalar> values=args[a].isRange?args[a].range:std::vector<Scalar> {
                    args[a].scalar
                };
                for(const auto&v:values) {
                    if(isError(v))return EvalValue::fromScalar(errorOr(v));
                    auto text=scalarText(v,engine.date1904());
                    if(*ignore&&text.empty())continue;
                    if(!first)out+=delimiter;
                    out+=text;
                    first=false;
                }
            }
            return EvalValue::fromScalar(std::move(out));
        }
        if(name=="NA")return EvalValue::fromScalar(CellError::NotAvailable);
        if (name == "ISNUMBER") return EvalValue::fromScalar(std::holds_alternative<double>(scalarArg(0)) || std::holds_alternative<DateTime>(scalarArg(0)));
        if (name == "ISTEXT") return EvalValue::fromScalar(std::holds_alternative<std::string>(scalarArg(0)));
        if (name == "ISBLANK") return EvalValue::fromScalar(std::holds_alternative<std::monostate>(scalarArg(0)));
        if (name == "ISERROR") return EvalValue::fromScalar(isError(scalarArg(0)));
        if (name == "ISNA") {
            const auto value=scalarArg(0);
            const auto* e=std::get_if<CellError>(&value);
            return EvalValue::fromScalar(e&&*e==CellError::NotAvailable);
        }
        if (name == "ISERR") {
            const auto value=scalarArg(0);
            const auto* e=std::get_if<CellError>(&value);
            return EvalValue::fromScalar(e&&*e!=CellError::NotAvailable);
        }
        if (name == "DATE") {
            const auto y = numArg(0), m = numArg(1), d = numArg(2);
            if (!y || !m || !d) return EvalValue::fromScalar(CellError::Value);
            int year = static_cast<int>(*y);
            if (year >= 0 && year < 1900) year += 1900;
            int month = static_cast<int>(*m);
            int day = static_cast<int>(*d);
            // Normalize through serial arithmetic so month/day overflow matches Excel reasonably closely.
            const int zeroMonth = month - 1;
            year += zeroMonth / 12;
            month = zeroMonth % 12 + 1;
            if (month <= 0) {
                month += 12;
                --year;
            }
            DateTime base {
                year, month, 1
            };
            const auto serial = toExcelSerial(base, engine.date1904()) + day - 1;
            return EvalValue::fromScalar(fromExcelSerial(serial, engine.date1904()));
        }
        if (name == "YEAR" || name == "MONTH" || name == "DAY") {
            const auto value = scalarArg(0);
            DateTime date;
            if (const auto* d = std::get_if<DateTime>(&value)) date = *d;
            else if (const auto n = numberValue(value, engine.date1904())) date = fromExcelSerial(*n, engine.date1904());
            else return EvalValue::fromScalar(CellError::Value);
            if (name == "YEAR") return EvalValue::fromScalar(static_cast<double>(date.year));
            if (name == "MONTH") return EvalValue::fromScalar(static_cast<double>(date.month));
            return EvalValue::fromScalar(static_cast<double>(date.day));
        }
        if(name=="DAYS") {
            auto endDate=numArg(0),startDate=numArg(1);
            if(!endDate||!startDate)return EvalValue::fromScalar(CellError::Value);
            return EvalValue::fromScalar(*endDate-*startDate);
        }
        if(name=="TIME") {
            auto h=numArg(0),m=numArg(1),sec=numArg(2);
            if(!h||!m||!sec)return EvalValue::fromScalar(CellError::Value);
            double seconds=*h*3600.0+*m*60.0+*sec;
            if(seconds<0)return EvalValue::fromScalar(CellError::Number);
            seconds=std::fmod(seconds,86400.0);
            return EvalValue::fromScalar(seconds/86400.0);
        }
        if(name=="HOUR"||name=="MINUTE"||name=="SECOND") {
            auto n=numArg(0);
            if(!n)return EvalValue::fromScalar(CellError::Value);
            double fraction=*n-std::floor(*n);
            if(fraction<0)fraction+=1.0;
            double seconds=std::round(fraction*86400.0);
            if(seconds>=86400)seconds=0;
            int h=static_cast<int>(seconds/3600),m=static_cast<int>(seconds/60)%60,sec=static_cast<int>(seconds)%60;
            return EvalValue::fromScalar(static_cast<double>(name=="HOUR"?h:(name=="MINUTE"?m:sec)));
        }
        if (name == "TODAY" || name == "NOW") {
            if (!engine.evaluateVolatileFunctions()) return engine.unsupportedFunction(name);
            const auto now = std::chrono::system_clock::to_time_t(std::chrono::system_clock::now());
            std::tm local {
            };
            #if defined(_WIN32)
            localtime_s(&local, &now);
            #else
            localtime_r(&now, &local);
            #endif
            DateTime date {
                local.tm_year + 1900, local.tm_mon + 1, local.tm_mday, name == "NOW" ? local.tm_hour : 0, name == "NOW" ? local.tm_min : 0, name == "NOW" ? static_cast<double>(local.tm_sec) : 0.0
            };
            return EvalValue::fromScalar(date);
        }
        return std::nullopt;
    }
}
// namespace xlpp::internal::formula
