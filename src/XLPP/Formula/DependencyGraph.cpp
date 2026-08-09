#include <XLPP/Formula/DependencyGraph.h>
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Cell/CellReference.h>

#include <algorithm>
#include <cctype>
#include <optional>
#include <string_view>
#include <unordered_set>

namespace xlpp {
namespace {

std::string upper(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](unsigned char c){ return static_cast<char>(std::toupper(c)); });
    return value;
}
bool ieq(std::string_view a, std::string_view b) {
    if (a.size()!=b.size()) return false;
    for (std::size_t i=0;i<a.size();++i) if (std::toupper(static_cast<unsigned char>(a[i]))!=std::toupper(static_cast<unsigned char>(b[i]))) return false;
    return true;
}
std::string unquoteSheet(std::string value) {
    if (value.size()>=2 && value.front()=='\'' && value.back()=='\'') {
        value=value.substr(1,value.size()-2);
        std::size_t p=0; while((p=value.find("''",p))!=std::string::npos){value.replace(p,2,"'");++p;}
    }
    return value;
}

struct ParsedRef { std::string sheet; std::string ref; std::size_t end{}; bool external{false}; };

std::optional<ParsedRef> parseA1(std::string_view f, std::size_t start, std::string_view contextSheet) {
    std::size_t p=start; std::string sheet;
    if (p<f.size() && f[p]=='\'') {
        const auto begin=p++; std::string decoded;
        while(p<f.size()) {
            if(f[p]=='\'' && p+1<f.size() && f[p+1]=='\''){decoded.push_back('\'');p+=2;continue;}
            if(f[p]=='\'' && p+1<f.size() && f[p+1]=='!'){sheet=decoded;p+=2;break;}
            decoded.push_back(f[p++]);
        }
        if(sheet.empty()) p=begin;
    } else {
        std::size_t q=p;
        while(q<f.size() && (std::isalnum(static_cast<unsigned char>(f[q]))||f[q]=='_'||f[q]=='.'||f[q]==' '||f[q]=='-'||f[q]=='['||f[q]==']')) ++q;
        if(q<f.size()&&f[q]=='!'&&q>p){sheet=std::string(f.substr(p,q-p));p=q+1;}
    }
    const auto cellBegin=p;
    if(p<f.size()&&f[p]=='$')++p;
    const auto colBegin=p; while(p<f.size()&&std::isalpha(static_cast<unsigned char>(f[p])))++p;
    if(p==colBegin)return std::nullopt;
    if(p<f.size()&&f[p]=='$')++p;
    const auto rowBegin=p;while(p<f.size()&&std::isdigit(static_cast<unsigned char>(f[p])))++p;
    if(p==rowBegin)return std::nullopt;
    std::string first(f.substr(cellBegin,p-cellBegin));
    try{(void)CellReference::parse(first);}catch(...){return std::nullopt;}
    if(sheet.empty() && p<f.size() && f[p]=='(') return std::nullopt; // LOG10(...)
    std::string ref=first;
    if(p<f.size()&&f[p]==':'){
        const auto colon=p++;const auto secondBegin=p;if(p<f.size()&&f[p]=='$')++p;const auto cb=p;while(p<f.size()&&std::isalpha(static_cast<unsigned char>(f[p])))++p;if(p==cb){p=colon;}
        else{if(p<f.size()&&f[p]=='$')++p;const auto rb=p;while(p<f.size()&&std::isdigit(static_cast<unsigned char>(f[p])))++p;if(p==rb)p=colon;else{std::string second(f.substr(secondBegin,p-secondBegin));try{(void)CellReference::parse(second);ref+=':'+second;}catch(...){p=colon;}}}
    }
    if(p<f.size()&&(std::isalnum(static_cast<unsigned char>(f[p]))||f[p]=='_'||f[p]=='.'))return std::nullopt;
    ParsedRef out;out.sheet=sheet.empty()?std::string(contextSheet):unquoteSheet(sheet);out.ref=ref;out.end=p;out.external=out.sheet.find('[')!=std::string::npos;return out;
}

bool containsCell(std::string_view range, std::string_view cell) {
    try {
        const auto colon=range.find(':');
        auto a=CellReference::parse(std::string(colon==std::string_view::npos?range:range.substr(0,colon)));
        auto b=CellReference::parse(std::string(colon==std::string_view::npos?range:range.substr(colon+1)));
        auto c=CellReference::parse(std::string(cell));
        const auto r1=std::min(a.row,b.row),r2=std::max(a.row,b.row),c1=std::min(a.column,b.column),c2=std::max(a.column,b.column);
        return c.row>=r1&&c.row<=r2&&c.column>=c1&&c.column<=c2;
    } catch (...) { return false; }
}

const DefinedName* definedNameFor(const Workbook& wb, const Worksheet& sheet, std::string_view name) {
    const DefinedName* global=nullptr; const auto sheetIndex=wb.index(sheet);
    for(const auto& d:wb.definedNames()){
        if(!ieq(d.name(),name))continue;
        if(d.localSheetId()&&*d.localSheetId()==sheetIndex)return &d;
        if(!d.localSheetId())global=&d;
    }
    return global;
}


} // namespace

std::vector<FormulaDependency> FormulaDependencyGraph::precedentsOf(const std::string& sheet, const std::string& cell) const {
    std::vector<FormulaDependency> out;for(const auto&e:edges_)if(ieq(e.dependentSheet,sheet)&&ieq(e.dependentCell,cell))out.push_back(e);return out;
}
std::vector<FormulaDependency> FormulaDependencyGraph::dependentsOf(const std::string& sheet, const std::string& cell) const {
    std::vector<FormulaDependency> out;for(const auto&e:edges_)if(e.kind==FormulaDependencyKind::CellOrRange&&ieq(e.precedentSheet,sheet)&&containsCell(e.precedentReference,cell))out.push_back(e);return out;
}
bool FormulaDependencyGraph::dependsOn(const std::string& ds,const std::string& dc,const std::string& ps,const std::string& pc) const {
    const auto values=dependentsOf(ps,pc);return std::any_of(values.begin(),values.end(),[&](const auto&e){return ieq(e.dependentSheet,ds)&&ieq(e.dependentCell,dc);});
}

FormulaDependencyGraph buildFormulaDependencyGraph(const Workbook& workbook) {
    FormulaDependencyGraph graph;
    std::unordered_set<std::string> dedupe;
    auto emit=[&](FormulaDependency edge){
        const auto key=edge.dependentSheet+"\x1f"+edge.dependentCell+"\x1f"+std::to_string(static_cast<int>(edge.kind))+"\x1f"+edge.precedentSheet+"\x1f"+edge.precedentReference+"\x1f"+edge.symbol;
        if(!dedupe.insert(key).second)return;
        switch(edge.kind){
            case FormulaDependencyKind::CellOrRange:++graph.report_.cellOrRangeEdges;break;
            case FormulaDependencyKind::DefinedName:++graph.report_.definedNameEdges;break;
            case FormulaDependencyKind::Table:++graph.report_.tableEdges;break;
            case FormulaDependencyKind::ExternalReference:++graph.report_.externalEdges;break;
            case FormulaDependencyKind::VolatileReference:++graph.report_.volatileReferences;break;
        }
        graph.edges_.push_back(std::move(edge));
    };
    static const std::unordered_set<std::string> volatileFns{"INDIRECT","OFFSET","TODAY","NOW","RAND","RANDBETWEEN"};
    static const std::unordered_set<std::string> literals{"TRUE","FALSE"};
    for (const auto& sheet : workbook.worksheets()) for (const auto& [_, cell] : sheet.cells()) {
        if (!cell.hasFormula()) continue;
        ++graph.report_.formulaCells;
        const auto dependent=CellReference(cell.row(),cell.column()).address();
        std::string_view f=cell.formula();if(!f.empty()&&f.front()=='=')f.remove_prefix(1);
        for(std::size_t i=0;i<f.size();){
            if(f[i]=='"'){++i;while(i<f.size()){if(f[i]=='"'){if(i+1<f.size()&&f[i+1]=='"'){i+=2;continue;}++i;break;}++i;}continue;}
            if(auto ref=parseA1(f,i,sheet.name())){
                emit({sheet.name(),dependent,ref->external?FormulaDependencyKind::ExternalReference:FormulaDependencyKind::CellOrRange,ref->sheet,ref->ref,{}});i=ref->end;continue;
            }
            if(std::isalpha(static_cast<unsigned char>(f[i]))||f[i]=='_'||f[i]=='\\'){
                const auto begin=i++;while(i<f.size()&&(std::isalnum(static_cast<unsigned char>(f[i]))||f[i]=='_'||f[i]=='.'||f[i]=='\\'))++i;
                const auto token=std::string(f.substr(begin,i-begin));const auto u=upper(token);
                if (i < f.size() && f[i] == '[') {
                    int depth = 0;
                    const auto b = i;
                    do {
                        if (f[i] == '[') ++depth;
                        else if (f[i] == ']') --depth;
                        ++i;
                    } while (i < f.size() && depth > 0);
                    std::string expr = token + std::string(f.substr(b, i - b));
                    bool resolvedTable = false;
                    for (const auto& ts : workbook.worksheets()) {
                        for (const auto& t : ts.tables()) {
                            if (ieq(t.name(), token) || ieq(t.displayName(), token)) {
                                emit({sheet.name(), dependent, FormulaDependencyKind::Table,
                                      ts.name(), t.reference(), expr});
                                emit({sheet.name(), dependent, FormulaDependencyKind::CellOrRange,
                                      ts.name(), t.reference(), expr});
                                resolvedTable = true;
                                break;
                            }
                        }
                        if (resolvedTable) break;
                    }
                    if (!resolvedTable) {
                        ++graph.report_.unresolvedSymbols;
                        graph.report_.warnings.push_back("Unresolved table reference: " + expr);
                    }
                    continue;
                }
                if(i<f.size()&&f[i]=='('){if(volatileFns.count(u))emit({sheet.name(),dependent,FormulaDependencyKind::VolatileReference,sheet.name(),{},u});continue;}
                if(literals.count(u))continue;
                if(const auto*d=definedNameFor(workbook,sheet,token)){
                    emit({sheet.name(),dependent,FormulaDependencyKind::DefinedName,sheet.name(),d->value(),d->name()});
                    std::string_view named=d->value();if(!named.empty()&&named.front()=='=')named.remove_prefix(1);
                    for(std::size_t np=0;np<named.size();){if(auto nr=parseA1(named,np,sheet.name())){emit({sheet.name(),dependent,nr->external?FormulaDependencyKind::ExternalReference:FormulaDependencyKind::CellOrRange,nr->sheet,nr->ref,d->name()});np=nr->end;}else ++np;}
                    continue;
                }
                // Operators and Excel error-name fragments are handled elsewhere;
                // record only plausible user symbols so diagnostics stay actionable.
                if(token.find('.')==std::string::npos){++graph.report_.unresolvedSymbols;graph.report_.warnings.push_back("Unresolved formula symbol: "+token+" in "+sheet.name()+"!"+dependent);}
                continue;
            }
            ++i;
        }
    }
    graph.report_.edges=graph.edges_.size();return graph;
}

} // namespace xlpp
