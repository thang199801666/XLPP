#include "Formula/FormulaFunctionSupport.h"
#include <algorithm>
#include <array>
#include <chrono>
#include <cctype>
#include <cmath>
#include <unordered_set>
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
        if (name == "CLEAN") {
            auto text = textArg(0);
            text.erase(std::remove_if(text.begin(), text.end(), [](unsigned char c) { return c < 32; }), text.end());
            return EvalValue::fromScalar(std::move(text));
        }
        if (name == "PROPER") {
            auto text = textArg(0);
            bool capitalize = true;
            for (char& value : text) {
                const auto c = static_cast<unsigned char>(value);
                if (std::isalpha(c)) {
                    value = static_cast<char>(capitalize ? std::toupper(c) : std::tolower(c));
                    capitalize = false;
                } else {
                    capitalize = true;
                }
            }
            return EvalValue::fromScalar(std::move(text));
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
        if(name=="REPLACE") {
            if(args.size()<4)return EvalValue::fromScalar(CellError::Value);
            auto text=textArg(0),replacement=textArg(3);
            auto start=numArg(1),count=numArg(2);
            if(!start||!count||*start<1||*count<0)return EvalValue::fromScalar(CellError::Value);
            const auto begin=static_cast<std::size_t>(*start-1);
            if(begin>text.size())text.append(begin-text.size(),' ');
            text.replace(begin,std::min<std::size_t>(static_cast<std::size_t>(*count),text.size()-begin),replacement);
            return EvalValue::fromScalar(std::move(text));
        }
        if(name=="EXACT")return EvalValue::fromScalar(textArg(0)==textArg(1));
        if(name=="T") {
            const auto value=scalarArg(0);
            if(isError(value))return EvalValue::fromScalar(errorOr(value));
            if(const auto* text=std::get_if<std::string>(&value))return EvalValue::fromScalar(*text);
            return EvalValue::fromScalar(std::string{});
        }
        if(name=="N") {
            const auto value=scalarArg(0);
            if(isError(value))return EvalValue::fromScalar(errorOr(value));
            if(const auto* number=std::get_if<double>(&value))return EvalValue::fromScalar(*number);
            if(const auto* boolean=std::get_if<bool>(&value))return EvalValue::fromScalar(*boolean?1.0:0.0);
            if(const auto* date=std::get_if<DateTime>(&value))return EvalValue::fromScalar(toExcelSerial(*date,engine.date1904()));
            return EvalValue::fromScalar(0.0);
        }
        if(name=="CHAR") {
            auto code=numArg(0);
            if(!code||*code<1||*code>255)return EvalValue::fromScalar(CellError::Value);
            return EvalValue::fromScalar(std::string(1,static_cast<char>(static_cast<unsigned char>(static_cast<int>(*code)))));
        }
        if(name=="CODE") {
            const auto text=textArg(0);
            if(text.empty())return EvalValue::fromScalar(CellError::Value);
            return EvalValue::fromScalar(static_cast<double>(static_cast<unsigned char>(text.front())));
        }
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
        if(name=="EDATE"||name=="EOMONTH") {
            auto serial=numArg(0),months=numArg(1);
            if(!serial||!months)return EvalValue::fromScalar(CellError::Value);
            auto date=fromExcelSerial(*serial,engine.date1904());
            const int offset=static_cast<int>(std::trunc(*months));
            const int monthIndex=date.year*12+(date.month-1)+offset;
            int year=monthIndex/12,month=monthIndex%12+1;
            if(month<=0){month+=12;--year;}
            auto leap=[](int y){return y%4==0&&(y%100!=0||y%400==0);};
            auto daysInMonth=[&](int y,int m){
                static constexpr int days[]{31,28,31,30,31,30,31,31,30,31,30,31};
                return m==2&&leap(y)?29:days[m-1];
            };
            const int day=name=="EOMONTH"?daysInMonth(year,month):std::min(date.day,daysInMonth(year,month));
            return EvalValue::fromScalar(DateTime{year,month,day,date.hour,date.minute,date.second});
        }
        if(name=="WEEKDAY") {
            auto serial=numArg(0);
            auto returnType=args.size()>=2?numArg(1):std::optional<double>{1.0};
            if(!serial||!returnType)return EvalValue::fromScalar(CellError::Value);
            const auto date=fromExcelSerial(*serial,engine.date1904());
            int dow=static_cast<int>((daysFromCivil(date.year,date.month,date.day)+4)%7);
            if(dow<0)dow+=7; // Sunday=0
            const int type=static_cast<int>(*returnType);
            int start=0;
            bool zeroBased=false;
            switch(type){
                case 1: case 17: start=0; break;
                case 2: case 11: start=1; break;
                case 3: start=1; zeroBased=true; break;
                case 12: start=2; break; case 13: start=3; break; case 14: start=4; break;
                case 15: start=5; break; case 16: start=6; break;
                default: return EvalValue::fromScalar(CellError::Number);
            }
            return EvalValue::fromScalar(static_cast<double>((dow-start+7)%7+(zeroBased?0:1)));
        }
        if(name=="WEEKNUM"||name=="ISOWEEKNUM") {
            auto serial=numArg(0);
            if(!serial)return EvalValue::fromScalar(CellError::Value);
            const auto date=fromExcelSerial(*serial,engine.date1904());
            const auto weekday=[](int year,int month,int day) {
                int value=static_cast<int>((daysFromCivil(year,month,day)+4)%7);
                return value<0?value+7:value; // Sunday=0
            };
            const bool iso=name=="ISOWEEKNUM"||
                (args.size()>=2&&numArg(1)&&static_cast<int>(*numArg(1))==21);
            if(iso) {
                const auto current=daysFromCivil(date.year,date.month,date.day);
                const int mondayBased=(weekday(date.year,date.month,date.day)+6)%7;
                const auto thursday=current+(3-mondayBased);
                int isoYear=0,isoMonth=0,isoDay=0;
                civilFromDays(thursday,isoYear,isoMonth,isoDay);
                const auto january4=daysFromCivil(isoYear,1,4);
                const int january4Weekday=(weekday(isoYear,1,4)+6)%7;
                const auto week1Monday=january4-january4Weekday;
                return EvalValue::fromScalar(static_cast<double>((thursday-week1Monday)/7+1));
            }
            auto returnType=args.size()>=2?numArg(1):std::optional<double>{1.0};
            if(!returnType)return EvalValue::fromScalar(CellError::Value);
            int start=0;
            switch(static_cast<int>(*returnType)) {
                case 1: case 17: start=0; break;
                case 2: case 11: start=1; break;
                case 12: start=2; break; case 13: start=3; break; case 14: start=4; break;
                case 15: start=5; break; case 16: start=6; break;
                default: return EvalValue::fromScalar(CellError::Number);
            }
            const auto dayOfYear=daysFromCivil(date.year,date.month,date.day)-daysFromCivil(date.year,1,1)+1;
            const int offset=(weekday(date.year,1,1)-start+7)%7;
            return EvalValue::fromScalar(static_cast<double>((dayOfYear+offset-1)/7+1));
        }
        if(name=="NETWORKDAYS"||name=="NETWORKDAYS.INTL"||name=="WORKDAY"||name=="WORKDAY.INTL") {
            const bool network=name.rfind("NETWORKDAYS",0)==0;
            const bool international=name.ends_with(".INTL");
            const std::size_t weekendIndex=2;
            const std::size_t holidaysIndex=international?3:2;
            std::array<bool,7> weekend{}; // Sunday through Saturday
            weekend[0]=true;
            weekend[6]=true;
            auto applyWeekendCode=[&](int code) {
                weekend.fill(false);
                if(code>=1&&code<=7) {
                    const int first=(code==1?6:code-2); // 1=Sat/Sun, 2=Sun/Mon ... 7=Fri/Sat
                    weekend[static_cast<std::size_t>(first)]=true;
                    weekend[static_cast<std::size_t>((first+1)%7)]=true;
                    return true;
                }
                if(code>=11&&code<=17) {
                    weekend[static_cast<std::size_t>(code-11)]=true; // 11=Sun ... 17=Sat
                    return true;
                }
                return false;
            };
            if(international&&args.size()>weekendIndex) {
                const auto weekendValue=scalarArg(weekendIndex);
                if(const auto* text=std::get_if<std::string>(&weekendValue)) {
                    if(text->size()!=7||std::any_of(text->begin(),text->end(),[](char c){return c!='0'&&c!='1';})||
                       std::all_of(text->begin(),text->end(),[](char c){return c=='1';}))
                        return EvalValue::fromScalar(CellError::Value);
                    weekend.fill(false);
                    for(std::size_t index=0;index<7;++index) weekend[(index+1)%7]=(*text)[index]=='1';
                } else {
                    auto code=numArg(weekendIndex);
                    if(!code||std::trunc(*code)!=*code||!applyWeekendCode(static_cast<int>(*code)))
                        return EvalValue::fromScalar(CellError::Number);
                }
            }
            std::unordered_set<long long> holidays;
            if(args.size()>holidaysIndex) {
                const auto values=args[holidaysIndex].isRange?args[holidaysIndex].range:std::vector<Scalar>{args[holidaysIndex].scalar};
                for(const auto& value:values) {
                    if(isError(value))return EvalValue::fromScalar(errorOr(value));
                    if(auto serial=numberValue(value,engine.date1904())) {
                        const auto date=fromExcelSerial(*serial,engine.date1904());
                        holidays.insert(daysFromCivil(date.year,date.month,date.day));
                    }
                }
            }
            auto civilDay=[&](double serial) {
                const auto date=fromExcelSerial(serial,engine.date1904());
                return daysFromCivil(date.year,date.month,date.day);
            };
            auto workingDay=[&](long long day) {
                int year=0,month=0,date=0;
                civilFromDays(day,year,month,date);
                int dow=static_cast<int>((day+4)%7);
                if(dow<0)dow+=7;
                return !weekend[static_cast<std::size_t>(dow)]&&holidays.find(day)==holidays.end();
            };
            if(network) {
                auto start=numArg(0),end=numArg(1);
                if(!start||!end)return EvalValue::fromScalar(CellError::Value);
                auto first=civilDay(*start),last=civilDay(*end);
                const int direction=first<=last?1:-1;
                std::size_t count=0;
                for(auto day=first;;day+=direction) {
                    if(workingDay(day))++count;
                    if(day==last)break;
                }
                return EvalValue::fromScalar(static_cast<double>(direction)*static_cast<double>(count));
            }
            auto start=numArg(0),days=numArg(1);
            if(!start||!days)return EvalValue::fromScalar(CellError::Value);
            if(!std::isfinite(*days)||std::fabs(*days)>10000000.0)return EvalValue::fromScalar(CellError::Number);
            const auto requested=static_cast<long long>(std::trunc(*days));
            const int direction=requested<0?-1:1;
            auto current=civilDay(*start);
            auto remaining=static_cast<unsigned long long>(requested<0?-requested:requested);
            while(remaining>0) {
                current+=direction;
                if(workingDay(current))--remaining;
            }
            DateTime result;
            civilFromDays(current,result.year,result.month,result.day);
            return EvalValue::fromScalar(result);
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
