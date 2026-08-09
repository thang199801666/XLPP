#include "Formula/FormulaFunctionSupport.h"
#include <algorithm>
#include <cmath>
#include <limits>
namespace xlpp::internal::formula {
    FormulaFunctionResult evaluateMathStatFinancialFunctions(FormulaFunctionCall& call, const std::string& name) {
        auto& engine = call.engine;
        const auto& args = call.args;
        auto err = [&]() {
            return call.error();
        };
        auto scalarArg = [&](std::size_t index) {
            return call.scalarArg(index);
        };
        auto numArg = [&](std::size_t index) {
            return call.numArg(index);
        };
        auto boolArg = [&](std::size_t index) {
            return call.boolArg(index);
        };
        auto numericValues = [&]() {
            return call.numericValues();
        };
        if (name == "SUM" || name == "AVERAGE" || name == "MIN" || name == "MAX" || name == "PRODUCT") {
            if (auto e = err()) return EvalValue::fromScalar(*e);
            const auto values = numericValues();
            if (name == "SUM") {
                double total = 0.0;
                for (double v : values) total += v;
                return EvalValue::fromScalar(total);
            }
            if (name == "AVERAGE") {
                if (values.empty()) return EvalValue::fromScalar(CellError::DivisionByZero);
                double total = 0.0;
                for (double v : values) total += v;
                return EvalValue::fromScalar(total / static_cast<double>(values.size()));
            }
            if (name == "MIN") return EvalValue::fromScalar(values.empty() ? 0.0 : *std::min_element(values.begin(), values.end()));
            if (name == "MAX") return EvalValue::fromScalar(values.empty() ? 0.0 : *std::max_element(values.begin(), values.end()));
            double product = 1.0;
            for (double v : values) product *= v;
            return EvalValue::fromScalar(product);
        }
        if (name == "COUNT") {
            std::size_t count = 0;
            for (const auto& v : flattened(args)) if (std::holds_alternative<double>(v) || std::holds_alternative<DateTime>(v)) ++count;
            return EvalValue::fromScalar(static_cast<double>(count));
        }
        if (name == "COUNTA") {
            std::size_t count = 0;
            for (const auto& v : flattened(args)) if (!std::holds_alternative<std::monostate>(v)) ++count;
            return EvalValue::fromScalar(static_cast<double>(count));
        }
        if (name == "COUNTBLANK") {
            std::size_t count=0;
            for(const auto& v:flattened(args)) if(std::holds_alternative<std::monostate>(v)||(std::holds_alternative<std::string>(v)&&std::get<std::string>(v).empty())) ++count;
            return EvalValue::fromScalar(static_cast<double>(count));
        }
        if (name == "IF") {
            if (args.size() < 2) return EvalValue::fromScalar(CellError::Value);
            const auto condition = boolArg(0);
            if (!condition) return EvalValue::fromScalar(CellError::Value);
            return *condition ? EvalValue::fromScalar(scalarArg(1)) : EvalValue::fromScalar(args.size() >= 3 ? scalarArg(2) : Scalar{false});
        }
        if (name == "IFERROR") {
            if (args.size() < 2) return EvalValue::fromScalar(CellError::Value);
            const auto value = scalarArg(0);
            return EvalValue::fromScalar(isError(value) ? scalarArg(1) : value);
        }
        if (name == "IFNA") {
            if (args.size() < 2) return EvalValue::fromScalar(CellError::Value);
            const auto value = scalarArg(0);
            return EvalValue::fromScalar(std::get_if<CellError>(&value) && std::get<CellError>(value)==CellError::NotAvailable ? scalarArg(1) : value);
        }
        if (name == "AND" || name == "OR") {
            bool result = name == "AND";
            for (const auto& v : flattened(args)) {
                if (isError(v)) return EvalValue::fromScalar(errorOr(v));
                const auto b = boolValue(v, engine.date1904());
                if (!b) continue;
                if (name == "AND") result = result && *b;
                else result = result || *b;
            }
            return EvalValue::fromScalar(result);
        }
        if(name=="XOR") {
            bool result=false;
            bool any=false;
            for(const auto&v:flattened(args)) {
                if(isError(v))return EvalValue::fromScalar(errorOr(v));
                auto b=boolValue(v,engine.date1904());
                if(b) {
                    result^=*b;
                    any=true;
                }
            }
            return EvalValue::fromScalar(any?Scalar{result}:Scalar{CellError::Value});
        }
        if(name=="IFS") {
            if(args.size()<2||args.size()%2)return EvalValue::fromScalar(CellError::Value);
            for(std::size_t i=0;i<args.size();i+=2) {
                auto b=boolValue(firstScalar(args[i]),engine.date1904());
                if(!b)return EvalValue::fromScalar(CellError::Value);
                if(*b)return EvalValue::fromScalar(firstScalar(args[i+1]));
            }
            return EvalValue::fromScalar(CellError::NotAvailable);
        }
        if(name=="SWITCH") {
            if(args.size()<3)return EvalValue::fromScalar(CellError::Value);
            auto expression=scalarArg(0);
            const bool hasDefault=(args.size()%2)==0;
            const std::size_t pairEnd=hasDefault?args.size()-1:args.size();
            for(std::size_t i=1;i+1<pairEnd;i+=2)if(engine.compare(expression,firstScalar(args[i]),"="))return EvalValue::fromScalar(firstScalar(args[i+1]));
            return EvalValue::fromScalar(hasDefault?firstScalar(args.back()):Scalar{CellError::NotAvailable});
        }
        if (name == "NOT") {
            const auto b = boolArg(0);
            return EvalValue::fromScalar(b ? Scalar{!*b} : Scalar{CellError::Value});
        }
        if (name == "ABS" || name == "SQRT" || name == "INT") {
            const auto n = numArg(0);
            if (!n) return EvalValue::fromScalar(CellError::Value);
            if (name == "ABS") return EvalValue::fromScalar(std::fabs(*n));
            if (name == "SQRT") return EvalValue::fromScalar(*n < 0 ? Scalar{CellError::Number} : Scalar{std::sqrt(*n)});
            return EvalValue::fromScalar(std::floor(*n));
        }
        if (name == "POWER" || name == "MOD") {
            const auto a = numArg(0), b = numArg(1);
            if (!a || !b) return EvalValue::fromScalar(CellError::Value);
            if (name == "MOD") return EvalValue::fromScalar(*b == 0.0 ? Scalar{CellError::DivisionByZero} : Scalar{*a - *b * std::floor(*a / *b)});
            const auto value = std::pow(*a, *b);
            return EvalValue::fromScalar(std::isfinite(value) ? Scalar{value} : Scalar{CellError::Number});
        }
        if (name == "ROUND" || name == "ROUNDUP" || name == "ROUNDDOWN") {
            const auto n = numArg(0), digitsValue = numArg(1);
            if (!n || !digitsValue) return EvalValue::fromScalar(CellError::Value);
            const int digits = static_cast<int>(*digitsValue);
            const double scale = std::pow(10.0, digits);
            if (!std::isfinite(scale) || scale == 0.0) return EvalValue::fromScalar(CellError::Number);
            const double scaled = *n * scale;
            double rounded = 0.0;
            if (name == "ROUND") rounded = scaled >= 0 ? std::floor(scaled + 0.5) : std::ceil(scaled - 0.5);
            else if (name == "ROUNDUP") rounded = scaled >= 0 ? std::ceil(scaled) : std::floor(scaled);
            else rounded = scaled >= 0 ? std::floor(scaled) : std::ceil(scaled);
            return EvalValue::fromScalar(rounded / scale);
        }
        if(name=="SIGN"||name=="EXP"||name=="LN"||name=="LOG10") {
            auto n=numArg(0);
            if(!n)return EvalValue::fromScalar(CellError::Value);
            if(name=="SIGN")return EvalValue::fromScalar(*n>0?1.0:(*n<0?-1.0:0.0));
            if((name=="LN"||name=="LOG10")&&*n<=0)return EvalValue::fromScalar(CellError::Number);
            double v=name=="EXP"?std::exp(*n):(name=="LN"?std::log(*n):std::log10(*n));
            return EvalValue::fromScalar(std::isfinite(v)?Scalar{v}:Scalar{CellError::Number});
        }
        if(name=="LOG") {
            auto n=numArg(0),base=args.size()>=2?numArg(1):std::optional<double> {
                10.0
            };
            if(!n||!base||*n<=0||*base<=0||*base==1)return EvalValue::fromScalar(CellError::Number);
            return EvalValue::fromScalar(std::log(*n)/std::log(*base));
        }
        if(name=="SIN"||name=="COS"||name=="TAN"||name=="ASIN"||name=="ACOS"||name=="ATAN"||name=="RADIANS"||name=="DEGREES") {
            auto n=numArg(0);
            if(!n)return EvalValue::fromScalar(CellError::Value);
            double v=0.0;
            if(name=="SIN")v=std::sin(*n);
            else if(name=="COS")v=std::cos(*n);
            else if(name=="TAN")v=std::tan(*n);
            else if(name=="ASIN") {
                if(*n < -1.0 || *n > 1.0)return EvalValue::fromScalar(CellError::Number);
                v=std::asin(*n);
            } else if(name=="ACOS") {
                if(*n < -1.0 || *n > 1.0)return EvalValue::fromScalar(CellError::Number);
                v=std::acos(*n);
            } else if(name=="ATAN")v=std::atan(*n);
            else if(name=="RADIANS")v=*n*3.141592653589793238462643383279502884/180.0;
            else v=*n*180.0/3.141592653589793238462643383279502884;
            return EvalValue::fromScalar(std::isfinite(v)?Scalar{v}:Scalar{CellError::Number});
        }
        if(name=="ATAN2") {
            auto x=numArg(0),y=numArg(1);
            if(!x||!y)return EvalValue::fromScalar(CellError::Value);
            if(*x==0.0&&*y==0.0)return EvalValue::fromScalar(CellError::DivisionByZero);
            return EvalValue::fromScalar(std::atan2(*y,*x));
        }
        if(name=="SUMSQ") {
            if(auto e=err())return EvalValue::fromScalar(*e);
            double total=0.0;
            for(double v:numericValues())total+=v*v;
            return EvalValue::fromScalar(total);
        }
        if(name=="TRUE")return EvalValue::fromScalar(true);
        if(name=="FALSE")return EvalValue::fromScalar(false);
        if(name=="NPV") {
            if(args.size()<2)return EvalValue::fromScalar(CellError::Value);
            auto rate=numArg(0);
            if(!rate||*rate<=-1.0)return EvalValue::fromScalar(CellError::Number);
            double total=0.0;
            std::size_t period=1;
            for(std::size_t a=1;a<args.size();++a) {
                const auto values=args[a].isRange?args[a].range:std::vector<Scalar> {
                    args[a].scalar
                };
                for(const auto&v:values)if(auto n=numberValue(v,engine.date1904())) {
                    total+=*n/std::pow(1.0+*rate,static_cast<double>(period));
                    ++period;
                }
            }
            return EvalValue::fromScalar(std::isfinite(total)?Scalar{total}:Scalar{CellError::Number});
        }
        if(name=="FV"||name=="PV"||name=="PMT") {
            if(args.size()<3)return EvalValue::fromScalar(CellError::Value);
            auto rate=numArg(0),nper=numArg(1),third=numArg(2);
            if(!rate||!nper||!third||*nper==0.0)return EvalValue::fromScalar(CellError::Value);
            const double optional4=args.size()>=4&&numArg(3)?*numArg(3):0.0;
            const double type=args.size()>=5&&numArg(4)?*numArg(4):0.0;
            if(type!=0.0&&type!=1.0)return EvalValue::fromScalar(CellError::Number);
            const double r=*rate,n=*nper;
            double result=0.0;
            if(name=="FV") {
                const double pmt=*third,pv=optional4;
                if(r==0.0)result=-(pv+pmt*n);
                else {
                    const double factor=std::pow(1.0+r,n);
                    result=-(pv*factor+pmt*(1.0+r*type)*(factor-1.0)/r);
                }
            } else if(name=="PV") {
                const double pmt=*third,fv=optional4;
                if(r==0.0)result=-(fv+pmt*n);
                else {
                    const double factor=std::pow(1.0+r,n);
                    result=-(fv+pmt*(1.0+r*type)*(factor-1.0)/r)/factor;
                }
            } else {
                const double pv=*third,fv=optional4;
                if(r==0.0)result=-(fv+pv)/n;
                else {
                    const double factor=std::pow(1.0+r,n),denom=(1.0+r*type)*(factor-1.0);
                    if(denom==0.0)return EvalValue::fromScalar(CellError::DivisionByZero);
                    result=-(fv+pv*factor)*r/denom;
                }
            }
            return EvalValue::fromScalar(std::isfinite(result)?Scalar{result}:Scalar{CellError::Number});
        }
        if(name=="IRR") {
            if(args.empty())return EvalValue::fromScalar(CellError::Value);
            std::vector<double>cash;
            for(const auto&v:flattened(args.size()>1?std::vector<EvalValue>{args[0]}:args))if(auto n=numberValue(v,engine.date1904()))cash.push_back(*n);
            if(cash.size()<2)return EvalValue::fromScalar(CellError::Number);
            bool positive=false,negative=false;
            for(double v:cash) {
                positive|=v>0;
                negative|=v<0;
            }
            if(!positive||!negative)return EvalValue::fromScalar(CellError::Number);
            double rate=args.size()>=2&&numArg(1)?*numArg(1):0.1;
            auto npv=[&](double r) {
                double value=0.0;
                for(std::size_t i=0;i<cash.size();++i)value+=cash[i]/std::pow(1.0+r,static_cast<double>(i));
                return value;
            };
            for(int iteration=0;iteration<80;++iteration) {
                if(rate<=-0.999999999)rate=-0.9;
                double value=0.0,derivative=0.0;
                for(std::size_t i=0;i<cash.size();++i) {
                    const double base=std::pow(1.0+rate,static_cast<double>(i));
                    value+=cash[i]/base;
                    if(i)derivative-=static_cast<double>(i)*cash[i]/(base*(1.0+rate));
                }
                if(std::fabs(value)<1e-10)return EvalValue::fromScalar(rate);
                if(std::fabs(derivative)<1e-14)break;
                const double next=rate-value/derivative;
                if(!std::isfinite(next)||next<=-1.0)break;
                if(std::fabs(next-rate)<1e-12)return EvalValue::fromScalar(next);
                rate=next;
            }
            double lo=-0.9999,hi=10.0,flo=npv(lo),fhi=npv(hi);
            if(!std::isfinite(flo)||!std::isfinite(fhi)||flo*fhi>0)return EvalValue::fromScalar(CellError::Number);
            for(int i=0;i<160;++i) {
                double mid=(lo+hi)/2.0,fmid=npv(mid);
                if(std::fabs(fmid)<1e-10)return EvalValue::fromScalar(mid);
                if(flo*fmid<=0) {
                    hi=mid;
                    fhi=fmid;
                } else {
                    lo=mid;
                    flo=fmid;
                }
            }
            return EvalValue::fromScalar((lo+hi)/2.0);
        }
        if(name=="SUMPRODUCT") {
            if(args.empty())return EvalValue::fromScalar(0.0);
            std::size_t count=args[0].isRange?args[0].range.size():1;
            for(const auto&a:args)if((a.isRange?a.range.size():1)!=count)return EvalValue::fromScalar(CellError::Value);
            double total=0;
            for(std::size_t i=0;i<count;++i) {
                double product=1;
                for(const auto&a:args) {
                    auto v=a.isRange?a.range[i]:a.scalar;
                    auto n=numberValue(v,engine.date1904());
                    product*=n?*n:0.0;
                }
                total+=product;
            }
            return EvalValue::fromScalar(total);
        }
        if(name=="MEDIAN") {
            if(auto e=err())return EvalValue::fromScalar(*e);
            auto values=numericValues();
            if(values.empty())return EvalValue::fromScalar(0.0);
            std::sort(values.begin(),values.end());
            const auto n=values.size();
            return EvalValue::fromScalar(n%2?values[n/2]:(values[n/2-1]+values[n/2])/2.0);
        }
        if(name=="SMALL"||name=="LARGE") {
            if(args.size()<2)return EvalValue::fromScalar(CellError::Value);
            std::vector<double>values;
            const auto source=args[0].isRange?args[0].range:std::vector<Scalar> {
                args[0].scalar
            };
            for(const auto&v:source)if(auto n=numberValue(v,engine.date1904()))values.push_back(*n);
            auto k=numArg(1);
            if(!k||*k<1||static_cast<std::size_t>(*k)>values.size())return EvalValue::fromScalar(CellError::Number);
            std::sort(values.begin(),values.end());
            auto index=static_cast<std::size_t>(*k)-1;
            return EvalValue::fromScalar(name=="SMALL"?values[index]:values[values.size()-1-index]);
        }
        if(name=="VAR.S"||name=="VAR"||name=="VAR.P"||name=="STDEV.S"||name=="STDEV"||name=="STDEV.P") {
            if(auto e=err())return EvalValue::fromScalar(*e);
            auto values=numericValues();
            const bool sample=name=="VAR.S"||name=="VAR"||name=="STDEV.S"||name=="STDEV";
            if(values.size()<(sample?2u:1u))return EvalValue::fromScalar(CellError::DivisionByZero);
            double mean=0;
            for(double v:values)mean+=v;
            mean/=static_cast<double>(values.size());
            double sum=0;
            for(double v:values) {
                auto d=v-mean;
                sum+=d*d;
            }
            double variance=sum/static_cast<double>(values.size()-(sample?1:0));
            return EvalValue::fromScalar((name.rfind("STDEV",0)==0)?std::sqrt(variance):variance);
        }
        if(name=="PI")return EvalValue::fromScalar(3.141592653589793238462643383279502884);
        if(name=="CEILING"||name=="CEILING.MATH"||name=="FLOOR"||name=="FLOOR.MATH") {
            auto n=numArg(0);
            auto sig=args.size()>=2?numArg(1):std::optional<double> {
                1.0
            };
            if(!n||!sig||*sig==0)return EvalValue::fromScalar(CellError::Value);
            const auto step=std::fabs(*sig);
            double q=*n/step;
            double v=(name.rfind("CEILING",0)==0?std::ceil(q):std::floor(q))*step;
            return EvalValue::fromScalar(v);
        }
        return std::nullopt;
    }
}
// namespace xlpp::internal::formula
