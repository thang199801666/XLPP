#pragma once
#include <string>
namespace xlpp {
class Comment {
public:
    Comment()=default; Comment(std::string text,std::string author):text_(std::move(text)),author_(std::move(author)){}
    const std::string& text() const noexcept{return text_;} void setText(std::string v){text_=std::move(v);} 
    const std::string& author() const noexcept{return author_;} void setAuthor(std::string v){author_=std::move(v);} 
private: std::string text_,author_;
};}
