#include "Formula/FormulaFunctionSupport.h"
#include <algorithm>
#include <limits>
namespace xlpp::internal::formula {
    FormulaFunctionResult evaluateCriteriaLookupFunctions(FormulaFunctionCall& call, const std::string& name) {
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
        auto matchesCriterion = [&](const Scalar& candidate, const Scalar& criterion) {
            return call.matchesCriterion(candidate, criterion);
        };
        if(name=="COUNTIF"||name=="SUMIF"||name=="AVERAGEIF") {
            if(args.size()<2||!args[0].isRange)return EvalValue::fromScalar(CellError::Value);
            const auto&criteriaRange=args[0].range;
            const std::vector<Scalar>* valueRange=&criteriaRange;
            if((name=="SUMIF"||name=="AVERAGEIF")&&args.size()>=3) {
                if(!args[2].isRange)return EvalValue::fromScalar(CellError::Value);
                valueRange=&args[2].range;
            }
            double total=0;
            std::size_t count=0;
            for(std::size_t i=0;i<criteriaRange.size();++i)if(matchesCriterion(criteriaRange[i],scalarArg(1))) {
                ++count;
                if(i<valueRange->size())if(auto n=numberValue((*valueRange)[i],engine.date1904()))total+=*n;
            }
            if(name=="COUNTIF")return EvalValue::fromScalar(static_cast<double>(count));
            if(name=="AVERAGEIF") {
                std::size_t numericCount=0;
                for(std::size_t i=0;i<criteriaRange.size();++i)if(matchesCriterion(criteriaRange[i],scalarArg(1))&&i<valueRange->size()&&numberValue((*valueRange)[i],engine.date1904()))++numericCount;
                return EvalValue::fromScalar(numericCount?Scalar{total/static_cast<double>(numericCount)}:Scalar{CellError::DivisionByZero});
            }
            return EvalValue::fromScalar(total);
        }
        if(name=="COUNTIFS"||name=="SUMIFS"||name=="AVERAGEIFS"||name=="MINIFS"||name=="MAXIFS") {
            const bool countOnly=name=="COUNTIFS";
            std::size_t criteriaStart=countOnly?0:1;
            if((countOnly&&(args.size()<2||args.size()%2))||(!countOnly&&(args.size()<3||(args.size()-1)%2)))return EvalValue::fromScalar(CellError::Value);
            const EvalValue* values=countOnly?nullptr:&args[0];
            const auto&firstRange=args[criteriaStart];
            if(!firstRange.isRange)return EvalValue::fromScalar(CellError::Value);
            const auto n=firstRange.range.size();
            if(values&&(!values->isRange||values->range.size()!=n))return EvalValue::fromScalar(CellError::Value);
            for(std::size_t a=criteriaStart;a<args.size();a+=2)if(!args[a].isRange||args[a].range.size()!=n)return EvalValue::fromScalar(CellError::Value);
            double total=0;
            std::size_t count=0;
            for(std::size_t i=0;i<n;++i) {
                bool ok=true;
                for(std::size_t a=criteriaStart;a<args.size();a+=2)if(!matchesCriterion(args[a].range[i],firstScalar(args[a+1]))) {
                    ok=false;
                    break;
                }
                if(ok) {
                    ++count;
                    if(values)if(auto v=numberValue(values->range[i],engine.date1904()))total+=*v;
                }
            }
            if(countOnly)return EvalValue::fromScalar(static_cast<double>(count));
            if(name=="MINIFS"||name=="MAXIFS") {
                std::optional<double>extreme;
                for(std::size_t i=0;i<n;++i) {
                    bool ok=true;
                    for(std::size_t a=criteriaStart;a<args.size();a+=2)if(!matchesCriterion(args[a].range[i],firstScalar(args[a+1]))) {
                        ok=false;
                        break;
                    }
                    if(ok)if(auto v=numberValue(values->range[i],engine.date1904()))extreme=!extreme?*v:(name=="MINIFS"?std::min(*extreme,*v):std::max(*extreme,*v));
                }
                return EvalValue::fromScalar(extreme?Scalar{*extreme}:Scalar{0.0});
            }
            if(name=="AVERAGEIFS") {
                std::size_t numericCount=0;
                for(std::size_t i=0;i<n;++i) {
                    bool ok=true;
                    for(std::size_t a=criteriaStart;a<args.size();a+=2)if(!matchesCriterion(args[a].range[i],firstScalar(args[a+1]))) {
                        ok=false;
                        break;
                    }
                    if(ok&&numberValue(values->range[i],engine.date1904()))++numericCount;
                }
                return EvalValue::fromScalar(numericCount?Scalar{total/static_cast<double>(numericCount)}:Scalar{CellError::DivisionByZero});
            }
            return EvalValue::fromScalar(total);
        }
        if(name=="INDEX") {
            if(args.size()<2||!args[0].isRange)return EvalValue::fromScalar(CellError::Value);
            auto rowValue=numArg(1),colValue=args.size()>=3?numArg(2):std::optional<double> {
                1.0
            };
            if(!rowValue||!colValue)return EvalValue::fromScalar(CellError::Value);
            auto row=static_cast<std::size_t>(*rowValue),col=static_cast<std::size_t>(*colValue);
            if(row<1||col<1||row>args[0].rows||col>args[0].columns)return EvalValue::fromScalar(CellError::Reference);
            return EvalValue::fromScalar(args[0].range[(row-1)*args[0].columns+(col-1)]);
        }
        if(name=="MATCH") {
            if(args.size()<2||!args[1].isRange)return EvalValue::fromScalar(CellError::Value);
            auto lookup=scalarArg(0);
            int mode=0;
            if(args.size()>=3) {
                auto m=numArg(2);
                if(!m)return EvalValue::fromScalar(CellError::Value);
                mode=static_cast<int>(*m);
            }
            const auto&range=args[1].range;
            if(mode==0) {
                for(std::size_t i=0;i<range.size();++i)if(engine.compare(range[i],lookup,"="))return EvalValue::fromScalar(static_cast<double>(i+1));
                return EvalValue::fromScalar(CellError::NotAvailable);
            }
            std::optional<std::size_t> best;
            for(std::size_t i=0;i<range.size();++i) {
                auto n=numberValue(range[i],engine.date1904()),target=numberValue(lookup,engine.date1904());
                if(n&&target) {
                    if((mode>0&&*n<=*target)||(mode<0&&*n>=*target))best=i;
                } else if(engine.compare(range[i],lookup,"="))return EvalValue::fromScalar(static_cast<double>(i+1));
            }
            return best?EvalValue::fromScalar(static_cast<double>(*best+1)):EvalValue::fromScalar(CellError::NotAvailable);
        }
        if(name=="VLOOKUP"||name=="HLOOKUP") {
            if(args.size()<3||!args[1].isRange)return EvalValue::fromScalar(CellError::Value);
            auto indexValue=numArg(2);
            if(!indexValue)return EvalValue::fromScalar(CellError::Value);
            const auto index=static_cast<std::size_t>(*indexValue);
            bool approximate=true;
            if(args.size()>=4) {
                auto b=boolArg(3);
                if(!b)return EvalValue::fromScalar(CellError::Value);
                approximate=*b;
            }
            const auto&table=args[1];
            const bool vertical=name=="VLOOKUP";
            const auto lookupCount=vertical?table.rows:table.columns;
            const auto resultLimit=vertical?table.columns:table.rows;
            if(index<1||index>resultLimit)return EvalValue::fromScalar(CellError::Reference);
            std::optional<std::size_t> found;
            for(std::size_t i=0;i<lookupCount;++i) {
                const auto pos=vertical?i*table.columns:i;
                const auto&candidate=table.range[pos];
                if(engine.compare(candidate,scalarArg(0),"=")) {
                    found=i;
                    break;
                }
                if(approximate) {
                    auto c=numberValue(candidate,engine.date1904()),l=numberValue(scalarArg(0),engine.date1904());
                    if(c&&l&&*c<=*l)found=i;
                }
            }
            if(!found)return EvalValue::fromScalar(CellError::NotAvailable);
            const auto pos=vertical?(*found)*table.columns+(index-1):(index-1)*table.columns+(*found);
            return EvalValue::fromScalar(table.range[pos]);
        }
        if(name=="XLOOKUP") {
            if(args.size()<3||!args[1].isRange||!args[2].isRange||args[1].range.size()!=args[2].range.size())return EvalValue::fromScalar(CellError::Value);
            int matchMode=0,searchMode=1;
            if(args.size()>=5) {
                auto m=numArg(4);
                if(!m)return EvalValue::fromScalar(CellError::Value);
                matchMode=static_cast<int>(*m);
            }
            if(args.size()>=6) {
                auto m=numArg(5);
                if(!m)return EvalValue::fromScalar(CellError::Value);
                searchMode=static_cast<int>(*m);
            }
            const auto&look=args[1].range;
            std::optional<std::size_t> found;
            auto exact=[&](const Scalar&v) {
                if(matchMode==2)return wildcardMatchInsensitive(scalarText(v,engine.date1904()),scalarText(scalarArg(0),engine.date1904()));
                return engine.compare(v,scalarArg(0),"=");
            };
            if(searchMode<0) {
                for(std::size_t i=look.size();i-->0;)if(exact(look[i])) {
                    found=i;
                    break;
                }
            } else for(std::size_t i=0;i<look.size();++i)if(exact(look[i])) {
                found=i;
                break;
            }
            if(!found&&(matchMode==-1||matchMode==1)) {
                auto target=numberValue(scalarArg(0),engine.date1904());
                if(target) {
                    double best=matchMode==-1?-std::numeric_limits<double>::infinity():std::numeric_limits<double>::infinity();
                    for(std::size_t i=0;i<look.size();++i)if(auto n=numberValue(look[i],engine.date1904())) {
                        if(matchMode==-1&&*n<=*target&&*n>best) {
                            best=*n;
                            found=i;
                        }
                        if(matchMode==1&&*n>=*target&&*n<best) {
                            best=*n;
                            found=i;
                        }
                    }
                }
            }
            if(!found)return EvalValue::fromScalar(args.size()>=4?scalarArg(3):Scalar{CellError::NotAvailable});
            return EvalValue::fromScalar(args[2].range[*found]);
        }
        return std::nullopt;
    }
}
// namespace xlpp::internal::formula
