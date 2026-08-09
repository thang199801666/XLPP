#include "Formula/FormulaFunctionSupport.h"
#include <XLPP/Cell/CellReference.h>
#include <algorithm>
#include <cmath>
#include <limits>
#include <unordered_map>
#include <unordered_set>
namespace xlpp::internal::formula {
    FormulaFunctionResult evaluateDynamicReferenceFunctions(FormulaFunctionCall& call, const std::string& name) {
        auto& engine = call.engine;
        const auto& args = call.args;
        auto& worksheet = call.worksheet;
        const auto depth = call.depth;
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
        if(name=="SEQUENCE") {
            auto rv=numArg(0);
            if(!rv) return EvalValue::fromScalar(CellError::Value);
            auto cv=args.size()>=2?numArg(1):std::optional<double> {
                1.0
            };
            auto sv=args.size()>=3?numArg(2):std::optional<double> {
                1.0
            };
            auto step=args.size()>=4?numArg(3):std::optional<double> {
                1.0
            };
            if(!cv||!sv||!step||*rv<1||*cv<1) return EvalValue::fromScalar(CellError::Value);
            const auto rows=static_cast<std::size_t>(*rv), cols=static_cast<std::size_t>(*cv);
            if(rows>1048576||cols>16384||rows>1'000'000/cols) return EvalValue::fromScalar(CellError::Number);
            std::vector<Scalar> out;
            out.reserve(rows*cols);
            double value=*sv;
            for(std::size_t i=0;i<rows*cols;++i,value+=*step) out.emplace_back(value);
            return EvalValue::fromRange(std::move(out),rows,cols);
        }
        if(name=="TRANSPOSE") {
            if(args.empty()||!args[0].isRange) return args.empty()?EvalValue::fromScalar(CellError::Value):args[0];
            const auto& a=args[0];
            std::vector<Scalar> out(a.range.size());
            for(std::size_t r=0;r<a.rows;++r) for(std::size_t c=0;c<a.columns;++c) out[c*a.rows+r]=a.range[r*a.columns+c];
            return EvalValue::fromRange(std::move(out),a.columns,a.rows);
        }
        if(name=="SORT") {
            if(args.empty()||!args[0].isRange) return args.empty()?EvalValue::fromScalar(CellError::Value):args[0];
            const auto& a=args[0];
            auto iv=args.size()>=2?numArg(1):std::optional<double> {
                1.0
            };
            auto ov=args.size()>=3?numArg(2):std::optional<double> {
                1.0
            };
            auto bv=args.size()>=4?boolArg(3):std::optional<bool> {
                false
            };
            if(!iv||!ov||!bv||(*ov!=1.0&&*ov!=-1.0)) return EvalValue::fromScalar(CellError::Value);
            const auto key=static_cast<std::size_t>(*iv);
            std::vector<Scalar> out;
            out.reserve(a.range.size());
            if(!*bv) {
                if(key<1||key>a.columns) return EvalValue::fromScalar(CellError::Value);
                std::vector<std::size_t> order(a.rows);
                for(std::size_t i=0;i<a.rows;++i)order[i]=i;
                std::stable_sort(order.begin(),order.end(),[&](auto x,auto y){return engine.compare(a.range[x*a.columns+key-1],a.range[y*a.columns+key-1],*ov>0?"<":">");});
                for(auto r:order) for(std::size_t c=0;c<a.columns;++c) out.push_back(a.range[r*a.columns+c]);
                return EvalValue::fromRange(std::move(out),a.rows,a.columns);
            }
            if(key<1||key>a.rows) return EvalValue::fromScalar(CellError::Value);
            std::vector<std::size_t> order(a.columns);
            for(std::size_t i=0;i<a.columns;++i)order[i]=i;
            std::stable_sort(order.begin(),order.end(),[&](auto x,auto y){return engine.compare(a.range[(key-1)*a.columns+x],a.range[(key-1)*a.columns+y],*ov>0?"<":">");});
            for(std::size_t r=0;r<a.rows;++r) for(auto c:order) out.push_back(a.range[r*a.columns+c]);
            return EvalValue::fromRange(std::move(out),a.rows,a.columns);
        }
        if(name=="UNIQUE") {
            if(args.empty()||!args[0].isRange) return args.empty()?EvalValue::fromScalar(CellError::Value):args[0];
            const auto& a=args[0];
            auto bv=args.size()>=2?boolArg(1):std::optional<bool> {
                false
            };
            auto ev=args.size()>=3?boolArg(2):std::optional<bool> {
                false
            };
            if(!bv||!ev)return EvalValue::fromScalar(CellError::Value);
            const std::size_t items=*bv?a.columns:a.rows, width=*bv?a.rows:a.columns;
            auto keyOf=[&](std::size_t item) {
                std::string k;
                for(std::size_t j=0;j<width;++j) {
                    const auto&v=*bv?a.range[j*a.columns+item]:a.range[item*a.columns+j];
                    k+=std::to_string(v.index())+":"+scalarText(v,engine.date1904())+"\x1f";
                }
                return k;
            };
            std::unordered_map<std::string,std::size_t> counts;
            for(std::size_t i=0;i<items;++i)++counts[keyOf(i)];
            std::unordered_set<std::string> emitted;
            std::vector<std::size_t> kept;
            for(std::size_t i=0;i<items;++i) {
                auto k=keyOf(i);
                if((*ev&&counts[k]!=1)||!emitted.insert(k).second)continue;
                kept.push_back(i);
            }
            std::vector<Scalar> out;
            if(!*bv) {
                for(auto r:kept)for(std::size_t c=0;c<a.columns;++c)out.push_back(a.range[r*a.columns+c]);
                return EvalValue::fromRange(std::move(out),kept.size(),a.columns);
            }
            for(std::size_t r=0;r<a.rows;++r)for(auto c:kept)out.push_back(a.range[r*a.columns+c]);
            return EvalValue::fromRange(std::move(out),a.rows,kept.size());
        }
        if(name=="FILTER") {
            if(args.size()<2||!args[0].isRange||!args[1].isRange) return EvalValue::fromScalar(CellError::Value);
            const auto&a=args[0];
            const auto&inc=args[1];
            if(inc.range.size()==a.rows) {
                std::vector<Scalar>out;
                std::size_t kept=0;
                for(std::size_t r=0;r<a.rows;++r) {
                    auto b=boolValue(inc.range[r],engine.date1904());
                    if(b&&*b) {
                        ++kept;
                        for(std::size_t c=0;c<a.columns;++c)out.push_back(a.range[r*a.columns+c]);
                    }
                }
                if(!kept)return EvalValue::fromScalar(args.size()>=3?scalarArg(2):Scalar{CellError::Calculation});
                return EvalValue::fromRange(std::move(out),kept,a.columns);
            }
            if(inc.range.size()==a.columns) {
                std::vector<std::size_t>cols;
                for(std::size_t c=0;c<a.columns;++c) {
                    auto b=boolValue(inc.range[c],engine.date1904());
                    if(b&&*b)cols.push_back(c);
                }
                if(cols.empty())return EvalValue::fromScalar(args.size()>=3?scalarArg(2):Scalar{CellError::Calculation});
                std::vector<Scalar>out;
                for(std::size_t r=0;r<a.rows;++r)for(auto c:cols)out.push_back(a.range[r*a.columns+c]);
                return EvalValue::fromRange(std::move(out),a.rows,cols.size());
            }
            return EvalValue::fromScalar(CellError::Value);
        }
        if(name=="TAKE"||name=="DROP") {
            if(args.size()<2||!args[0].isRange)return EvalValue::fromScalar(CellError::Value);
            const auto&a=args[0];
            auto rv=numArg(1);
            auto cv=args.size()>=3?numArg(2):std::optional<double> {
                static_cast<double>(a.columns)
            };
            if(!rv||!cv)return EvalValue::fromScalar(CellError::Value);
            long rows=static_cast<long>(*rv),cols=static_cast<long>(*cv);
            std::size_t r0=0,r1=a.rows,c0=0,c1=a.columns;
            auto boundTake=[](long n,std::size_t size,std::size_t&lo,std::size_t&hi) {
                auto count=(std::min)(static_cast<std::size_t>(std::labs(n)),size);
                if(n>=0)hi=lo+count;
                else lo=hi-count;
            };
            auto boundDrop=[](long n,std::size_t size,std::size_t&lo,std::size_t&hi) {
                auto count=(std::min)(static_cast<std::size_t>(std::labs(n)),size);
                if(n>=0)lo+=count;
                else hi-=count;
            };
            if(name=="TAKE") {
                boundTake(rows,a.rows,r0,r1);
                boundTake(cols,a.columns,c0,c1);
            } else {
                boundDrop(rows,a.rows,r0,r1);
                boundDrop(cols,a.columns,c0,c1);
            }
            if(r0>=r1||c0>=c1)return EvalValue::fromScalar(CellError::Calculation);
            std::vector<Scalar>out;
            for(std::size_t r=r0;r<r1;++r)for(std::size_t c=c0;c<c1;++c)out.push_back(a.range[r*a.columns+c]);
            return EvalValue::fromRange(std::move(out),r1-r0,c1-c0);
        }
        if(name=="CHOOSECOLS"||name=="CHOOSEROWS") {
            if(args.size()<2||!args[0].isRange)return EvalValue::fromScalar(CellError::Value);
            const auto&a=args[0];
            const bool chooseCols=name=="CHOOSECOLS";
            const auto limit=chooseCols?a.columns:a.rows;
            std::vector<std::size_t> selected;
            selected.reserve(args.size()-1);
            for(std::size_t i=1;i<args.size();++i) {
                auto n=numArg(i);
                if(!n||*n==0)return EvalValue::fromScalar(CellError::Value);
                long idx=static_cast<long>(*n);
                if(idx<0)idx=static_cast<long>(limit)+idx+1;
                if(idx<1||static_cast<std::size_t>(idx)>limit)return EvalValue::fromScalar(CellError::Value);
                selected.push_back(static_cast<std::size_t>(idx-1));
            }
            std::vector<Scalar>out;
            if(chooseCols) {
                out.reserve(a.rows*selected.size());
                for(std::size_t r=0;r<a.rows;++r)for(auto c:selected)out.push_back(a.range[r*a.columns+c]);
                return EvalValue::fromRange(std::move(out),a.rows,selected.size());
            }
            out.reserve(selected.size()*a.columns);
            for(auto r:selected)for(std::size_t c=0;c<a.columns;++c)out.push_back(a.range[r*a.columns+c]);
            return EvalValue::fromRange(std::move(out),selected.size(),a.columns);
        }
        if(name=="HSTACK"||name=="VSTACK") {
            if(args.empty())return EvalValue::fromScalar(CellError::Value);
            auto shape=[](const EvalValue&v) {
                return std::pair<std::size_t,std::size_t> {
                    v.isRange?v.rows:1,v.isRange?v.columns:1
                };
            };
            auto at=[](const EvalValue&v,std::size_t r,std::size_t c)->Scalar {
                if(!v.isRange)return (r==0&&c==0)?v.scalar:Scalar {
                    CellError::NotAvailable
                };
                if(r>=v.rows||c>=v.columns)return CellError::NotAvailable;
                return v.range[r*v.columns+c];
            };
            if(name=="HSTACK") {
                std::size_t rows=0,cols=0;
                for(const auto&v:args) {
                    auto [r,c]=shape(v);
                    rows=std::max(rows,r);
                    cols+=c;
                }
                std::vector<Scalar>out;
                out.reserve(rows*cols);
                for(std::size_t r=0;r<rows;++r)for(const auto&v:args) {
                    auto [vr,vc]=shape(v);
                    for(std::size_t c=0;c<vc;++c)out.push_back(at(v,r,c));
                }
                return EvalValue::fromRange(std::move(out),rows,cols);
            }
            std::size_t rows=0,cols=0;
            for(const auto&v:args) {
                auto [r,c]=shape(v);
                rows+=r;
                cols=std::max(cols,c);
            }
            std::vector<Scalar>out;
            out.reserve(rows*cols);
            for(const auto&v:args) {
                auto [vr,vc]=shape(v);
                for(std::size_t r=0;r<vr;++r)for(std::size_t c=0;c<cols;++c)out.push_back(at(v,r,c));
            }
            return EvalValue::fromRange(std::move(out),rows,cols);
        }
        if(name=="TOROW"||name=="TOCOL") {
            if(args.empty())return EvalValue::fromScalar(CellError::Value);
            const auto&a=args[0];
            std::vector<Scalar>src=a.isRange?a.range:std::vector<Scalar> {
                a.scalar
            };
            int ignore=0;
            if(args.size()>=2) {
                auto n=numArg(1);
                if(!n)return EvalValue::fromScalar(CellError::Value);
                ignore=static_cast<int>(*n);
            }
            std::vector<Scalar>out;
            out.reserve(src.size());
            for(const auto&v:src) {
                const bool blank=std::holds_alternative<std::monostate>(v);
                const bool error=isError(v);
                if((ignore==1&&blank)||(ignore==2&&error)||(ignore==3&&(blank||error)))continue;
                out.push_back(v);
            }
            if(out.empty())return EvalValue::fromScalar(CellError::Calculation);
            const auto count=out.size();
            return name=="TOROW"?EvalValue::fromRange(std::move(out),1,count):EvalValue::fromRange(std::move(out),count,1);
        }
        if(name=="ROWS"||name=="COLUMNS") {
            if(args.empty())return EvalValue::fromScalar(CellError::Value);
            return EvalValue::fromScalar(static_cast<double>(name=="ROWS"?(args[0].isRange?args[0].rows:1):(args[0].isRange?args[0].columns:1)));
        }
        if(name=="ROW"||name=="COLUMN") {
            if(args.empty())return EvalValue::fromScalar(CellError::Value);
            if(!args[0].hasReference())return EvalValue::fromScalar(CellError::Value);
            return EvalValue::fromScalar(static_cast<double>(name=="ROW"?args[0].referenceRow:args[0].referenceColumn));
        }
        if(name=="ADDRESS") {
            auto rv=numArg(0),cv=numArg(1);
            if(!rv||!cv||*rv<1||*rv>1048576||*cv<1||*cv>16384)return EvalValue::fromScalar(CellError::Value);
            int absNum=1;
            if(args.size()>=3) {
                auto av=numArg(2);
                if(!av)return EvalValue::fromScalar(CellError::Value);
                absNum=static_cast<int>(*av);
                if(absNum<1||absNum>4)return EvalValue::fromScalar(CellError::Value);
            }
            if(args.size()>=4) {
                auto av=boolArg(3);
                if(!av||!*av)return engine.unsupportedFunction("ADDRESS(R1C1)");
            }
            const auto row=static_cast<std::size_t>(*rv),col=static_cast<std::size_t>(*cv);
            std::string out;
            const bool absRow=absNum==1||absNum==2,absCol=absNum==1||absNum==3;
            if(absCol)out+='$';
            out+=CellReference::columnName(col);
            if(absRow)out+='$';
            out+=std::to_string(row);
            if(args.size()>=5&&!textArg(4).empty())out="'"+textArg(4)+"'!"+out;
            return EvalValue::fromScalar(std::move(out));
        }
        if(name=="INDIRECT") {
            if(args.empty())return EvalValue::fromScalar(CellError::Value);
            if(args.size()>=2) {
                auto a1=boolArg(1);
                if(!a1||!*a1)return engine.unsupportedFunction("INDIRECT(R1C1)");
            }
            const auto reference=textArg(0);
            if(reference.empty())return EvalValue::fromScalar(CellError::Reference);
            try {
                return parseFormula(engine, worksheet, reference, depth + 1);
            } catch (...) {
                return EvalValue::fromScalar(CellError::Reference);
            }
        }
        if(name=="OFFSET") {
            if(args.size()<3||!args[0].hasReference())return EvalValue::fromScalar(CellError::Reference);
            auto ro=numArg(1),co=numArg(2);
            if(!ro||!co)return EvalValue::fromScalar(CellError::Value);
            const auto baseRows=args[0].isRange?args[0].rows:1,baseCols=args[0].isRange?args[0].columns:1;
            auto hv=args.size()>=4?numArg(3):std::optional<double> {
                static_cast<double>(baseRows)
            };
            auto wv=args.size()>=5?numArg(4):std::optional<double> {
                static_cast<double>(baseCols)
            };
            if(!hv||!wv||*hv<1||*wv<1)return EvalValue::fromScalar(CellError::Reference);
            const long r=static_cast<long>(args[0].referenceRow)+static_cast<long>(*ro),c=static_cast<long>(args[0].referenceColumn)+static_cast<long>(*co),h=static_cast<long>(*hv),w=static_cast<long>(*wv);
            if(r<1||c<1||r+h-1>1048576||c+w-1>16384)return EvalValue::fromScalar(CellError::Reference);
            const auto first=CellReference(static_cast<std::size_t>(r),static_cast<std::size_t>(c)).address();
            const auto last=CellReference(static_cast<std::size_t>(r+h-1),static_cast<std::size_t>(c+w-1)).address();
            return engine.resolveReference(*args[0].referenceSheet,args[0].referenceSheet->name(),first,last,depth+1);
        }
        if(name=="CHOOSE") {
            if(args.size()<2)return EvalValue::fromScalar(CellError::Value);
            auto n=numArg(0);
            if(!n)return EvalValue::fromScalar(CellError::Value);
            auto index=static_cast<std::size_t>(*n);
            if(index<1||index>=args.size())return EvalValue::fromScalar(CellError::Value);
            return args[index];
        }
        return std::nullopt;
    }
}
// namespace xlpp::internal::formula
