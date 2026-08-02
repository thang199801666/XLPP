#pragma once
#include <string>
namespace xlpp {
class DocumentProperties {
public:
#define XLPP_PROP(Name,name) const std::string& name() const noexcept{return name##_;} void set##Name(std::string v){name##_=std::move(v);}
    XLPP_PROP(Title,title) XLPP_PROP(Subject,subject) XLPP_PROP(Creator,creator) XLPP_PROP(Description,description)
    XLPP_PROP(Keywords,keywords) XLPP_PROP(Category,category) XLPP_PROP(LastModifiedBy,lastModifiedBy)
#undef XLPP_PROP
private: std::string title_,subject_,creator_,description_,keywords_,category_,lastModifiedBy_;
};}
