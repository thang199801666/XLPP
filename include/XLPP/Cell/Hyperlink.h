#pragma once
#include <string>
namespace xlpp {
class Hyperlink {
public:
    Hyperlink()=default; explicit Hyperlink(std::string target):target_(std::move(target)){}
    const std::string& target() const noexcept{return target_;} void setTarget(std::string v){target_=std::move(v);} 
    const std::string& display() const noexcept{return display_;} void setDisplay(std::string v){display_=std::move(v);} 
    const std::string& tooltip() const noexcept{return tooltip_;} void setTooltip(std::string v){tooltip_=std::move(v);} 
    bool external() const noexcept{return external_;} void setExternal(bool v) noexcept{external_=v;}
private: std::string target_,display_,tooltip_; bool external_{true};
};}
