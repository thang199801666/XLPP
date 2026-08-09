#pragma once
#include <string>

namespace xlpp {

enum class CalculationMode { Automatic, AutomaticExceptDataTables, Manual };

class CalcProperties {
public:
    int calcId() const noexcept { return calcId_; } void setCalcId(int v) noexcept { calcId_ = v; }
    const std::string& calcMode() const noexcept { return calcMode_; }
    void setCalcMode(std::string v) { calcMode_ = std::move(v); }
    CalculationMode calculationMode() const noexcept {
        if (calcMode_ == "manual") return CalculationMode::Manual;
        if (calcMode_ == "autoNoTable") return CalculationMode::AutomaticExceptDataTables;
        return CalculationMode::Automatic;
    }
    void setCalculationMode(CalculationMode v) {
        switch (v) {
            case CalculationMode::Manual: calcMode_ = "manual"; break;
            case CalculationMode::AutomaticExceptDataTables: calcMode_ = "autoNoTable"; break;
            case CalculationMode::Automatic: calcMode_ = "auto"; break;
        }
    }
    bool calcOnSave() const noexcept { return calcOnSave_; }
    void setCalcOnSave(bool v) noexcept { calcOnSave_ = v; }
    bool fullCalcOnLoad() const noexcept { return fullCalcOnLoad_; }
    void setFullCalcOnLoad(bool v) noexcept { fullCalcOnLoad_ = v; }
    bool fullPrecision() const noexcept { return fullPrecision_; }
    void setFullPrecision(bool v) noexcept { fullPrecision_ = v; }
    bool iterate() const noexcept { return iterate_; } void setIterate(bool v) noexcept { iterate_ = v; }
    int iterateCount() const noexcept { return iterateCount_; } void setIterateCount(int v) noexcept { iterateCount_ = v; }
    double iterateDelta() const noexcept { return iterateDelta_; } void setIterateDelta(double v) noexcept { iterateDelta_ = v; }
private:
    int calcId_{191029}, iterateCount_{100};
    double iterateDelta_{0.001};
    std::string calcMode_{"auto"};
    bool calcOnSave_{false}, fullCalcOnLoad_{false}, fullPrecision_{true}, iterate_{false};
};

}
