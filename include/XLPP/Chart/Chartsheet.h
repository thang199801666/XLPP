#pragma once

#include <XLPP/Chart/Chart.h>
#include <XLPP/Worksheet/PageSetup.h>
#include <XLPP/Protection/LegacyPassword.h>

#include <optional>
#include <stdexcept>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace xlpp::internal { struct WorkbookChartsheetIOAccess; struct WorkbookChartsheetPackageAccess; }

namespace xlpp {

class ChartsheetProperties {
public:
    const std::string& codeName() const noexcept { return codeName_; }
    void setCodeName(std::string value) { codeName_ = std::move(value); }
    const std::optional<std::string>& tabColor() const noexcept { return tabColor_; }
    void setTabColor(std::string rgb) { tabColor_ = std::move(rgb); }
    void clearTabColor() noexcept { tabColor_.reset(); }
    const std::optional<bool>& published() const noexcept { return published_; }
    void setPublished(bool value) noexcept { published_ = value; }
    void clearPublished() noexcept { published_.reset(); }
    bool empty() const noexcept { return codeName_.empty() && !tabColor_ && !published_; }
private:
    std::string codeName_;
    std::optional<std::string> tabColor_;
    std::optional<bool> published_;
};

class ChartsheetView {
public:
    int workbookViewId() const noexcept { return workbookViewId_; }
    void setWorkbookViewId(int value) noexcept { workbookViewId_ = value; }
    bool tabSelected() const noexcept { return tabSelected_; }
    void setTabSelected(bool value) noexcept { tabSelected_ = value; }
    const std::optional<unsigned>& zoomScale() const noexcept { return zoomScale_; }
    void setZoomScale(unsigned value) { if (value < 10 || value > 400) throw std::out_of_range("Chartsheet zoomScale must be between 10 and 400"); zoomScale_ = value; }
    void clearZoomScale() noexcept { zoomScale_.reset(); }
    bool zoomToFit() const noexcept { return zoomToFit_; }
    void setZoomToFit(bool value) noexcept { zoomToFit_ = value; }
private:
    int workbookViewId_{0};
    bool tabSelected_{false};
    std::optional<unsigned> zoomScale_;
    bool zoomToFit_{true};
};

class ChartsheetProtection {
public:
    bool enabled() const noexcept { return enabled_; }
    void setEnabled(bool value) noexcept { enabled_ = value; }
    bool content() const noexcept { return content_; }
    void setContent(bool value) noexcept { content_ = value; enabled_ = true; }
    bool objects() const noexcept { return objects_; }
    void setObjects(bool value) noexcept { objects_ = value; enabled_ = true; }
    const std::string& passwordHash() const noexcept { return passwordHash_; }
    void setPasswordHash(std::string value) { passwordHash_ = std::move(value); enabled_ = true; }
    void setPassword(std::string_view password) { passwordHash_ = legacyProtectionPasswordHash(password); enabled_ = true; }
    void clearPassword() noexcept { passwordHash_.clear(); }
    const std::optional<std::string>& algorithmName() const noexcept { return algorithmName_; }
    void setAlgorithmName(std::string value) { algorithmName_ = std::move(value); enabled_ = true; }
    void clearAlgorithmName() noexcept { algorithmName_.reset(); }
    const std::optional<std::string>& hashValue() const noexcept { return hashValue_; }
    void setHashValue(std::string value) { hashValue_ = std::move(value); enabled_ = true; }
    void clearHashValue() noexcept { hashValue_.reset(); }
    const std::optional<std::string>& saltValue() const noexcept { return saltValue_; }
    void setSaltValue(std::string value) { saltValue_ = std::move(value); enabled_ = true; }
    void clearSaltValue() noexcept { saltValue_.reset(); }
    const std::optional<unsigned>& spinCount() const noexcept { return spinCount_; }
    void setSpinCount(unsigned value) noexcept { spinCount_ = value; enabled_ = true; }
    void clearSpinCount() noexcept { spinCount_.reset(); }
private:
    bool enabled_{false};
    bool content_{true};
    bool objects_{true};
    std::string passwordHash_;
    std::optional<std::string> algorithmName_, hashValue_, saltValue_;
    std::optional<unsigned> spinCount_;
};

enum class CustomChartsheetViewState { Visible, Hidden, VeryHidden };

class CustomChartsheetView {
public:
    const std::string& guid() const noexcept { return guid_; }
    void setGuid(std::string value) { guid_ = std::move(value); }
    const std::optional<unsigned>& scale() const noexcept { return scale_; }
    void setScale(unsigned value) { if (value < 10 || value > 400) throw std::out_of_range("Custom Chartsheet view scale must be between 10 and 400"); scale_ = value; }
    void clearScale() noexcept { scale_.reset(); }
    CustomChartsheetViewState state() const noexcept { return state_; }
    void setState(CustomChartsheetViewState value) noexcept { state_ = value; }
    const std::optional<bool>& zoomToFit() const noexcept { return zoomToFit_; }
    void setZoomToFit(bool value) noexcept { zoomToFit_ = value; }
    void clearZoomToFit() noexcept { zoomToFit_.reset(); }
    PageMargins& pageMargins() noexcept { if (!pageMargins_) pageMargins_.emplace(); return *pageMargins_; }
    const std::optional<PageMargins>& pageMarginsValue() const noexcept { return pageMargins_; }
    PageSetup& pageSetup() noexcept { if (!pageSetup_) pageSetup_.emplace(); return *pageSetup_; }
    const std::optional<PageSetup>& pageSetupValue() const noexcept { return pageSetup_; }
    HeaderFooter& headerFooter() noexcept { if (!headerFooter_) headerFooter_.emplace(); return *headerFooter_; }
    const std::optional<HeaderFooter>& headerFooterValue() const noexcept { return headerFooter_; }
private:
    std::string guid_;
    std::optional<unsigned> scale_;
    CustomChartsheetViewState state_{CustomChartsheetViewState::Visible};
    std::optional<bool> zoomToFit_;
    std::optional<PageMargins> pageMargins_;
    std::optional<PageSetup> pageSetup_;
    std::optional<HeaderFooter> headerFooter_;
};

// First-class chart-only workbook sheet. A Chartsheet participates in the
// workbook <sheets> order but, unlike Worksheet, owns a chart drawing instead
// of cells. Imported chartsheets keep their original package subtree until the
// chart is mutated or replaced. Sheet-level properties are independently dirty
// so view/page/protection edits can patch only the chartsheet part while the
// imported drawing/chart subtree remains byte-preserved.
class Chartsheet {
    friend class Workbook;
    friend struct internal::WorkbookChartsheetIOAccess;
    friend struct internal::WorkbookChartsheetPackageAccess;
public:
    Chartsheet() = default;
    explicit Chartsheet(std::string name) : name_(std::move(name)) {}
    Chartsheet(std::string name, Chart chart) : name_(std::move(name)), chart_(std::move(chart)) {}

    const std::string& name() const noexcept { return name_; }
    bool hasChart() const noexcept { return chart_.has_value(); }

    const Chart& chart() const {
        if (!chart_) throw std::logic_error("Chartsheet does not contain a materialized chart");
        return *chart_;
    }
    Chart& chart() {
        if (!chart_) chart_.emplace();
        chartDirty_ = true;
        return *chart_;
    }
    void setChart(Chart chart) {
        chart_ = std::move(chart);
        chartDirty_ = true;
    }

    ChartsheetProperties& properties() noexcept { sheetDirty_ = true; return properties_; }
    const ChartsheetProperties& properties() const noexcept { return properties_; }
    ChartsheetView& view() noexcept { sheetDirty_ = true; return view_; }
    const ChartsheetView& view() const noexcept { return view_; }
    ChartsheetProtection& protection() noexcept { sheetDirty_ = true; return protection_; }
    const ChartsheetProtection& protection() const noexcept { return protection_; }
    PageMargins& pageMargins() noexcept { sheetDirty_ = true; pageMarginsPresent_ = true; return pageMargins_; }
    const PageMargins& pageMargins() const noexcept { return pageMargins_; }
    bool hasPageMargins() const noexcept { return pageMarginsPresent_; }
    PageSetup& pageSetup() noexcept { sheetDirty_ = true; pageSetupPresent_ = true; return pageSetup_; }
    const PageSetup& pageSetup() const noexcept { return pageSetup_; }
    bool hasPageSetup() const noexcept { return pageSetupPresent_; }

    // Raw DevMode printer-settings payload owned by this Chartsheet. OOXML
    // stores the platform/printer-specific bytes in a separate
    // xl/printerSettings/printerSettingsN.bin part referenced by pageSetup@r:id.
    // XL++ intentionally treats the payload as opaque bytes: callers may attach
    // a producer-specific DevMode blob without the core pretending to decode
    // platform printer structures.
    bool hasPrinterSettings() const noexcept { return printerSettingsData_.has_value(); }
    const std::optional<std::string>& printerSettingsData() const noexcept { return printerSettingsData_; }
    void setPrinterSettingsData(std::string bytes) {
        printerSettingsData_ = std::move(bytes);
        printerSettingsDirty_ = true;
        sheetDirty_ = true;
        pageSetupPresent_ = true;
    }
    void clearPrinterSettings() noexcept {
        printerSettingsData_.reset();
        printerSettingsDirty_ = true;
        sheetDirty_ = true;
        pageSetup_.clearRelationshipId();
        printerSettingsRelationshipId_.clear();
    }
    HeaderFooter& headerFooter() noexcept { sheetDirty_ = true; headerFooterPresent_ = true; return headerFooter_; }
    const HeaderFooter& headerFooter() const noexcept { return headerFooter_; }
    bool hasHeaderFooter() const noexcept { return headerFooterPresent_; }
    std::vector<CustomChartsheetView>& customViews() noexcept { sheetDirty_ = true; customViewsDirty_ = true; return customViews_; }
    const std::vector<CustomChartsheetView>& customViews() const noexcept { return customViews_; }
    bool hasCustomViews() const noexcept { return !customViews_.empty() || !customViewsRawXml_.empty(); }

    // Imported chartsheets are preservation-backed. A const read does not
    // force regeneration; mutable chart access opts into full chart regeneration.
    bool imported() const noexcept { return imported_; }
    bool chartDirty() const noexcept { return chartDirty_; }
    bool sheetDirty() const noexcept { return sheetDirty_; }
    const std::string& sourcePart() const noexcept { return sourcePart_; }

private:
    void rename(std::string value) { name_ = std::move(value); }
    void setImportedSource(std::string sourcePart, std::string sourceXml, std::string drawingRelationshipId) {
        sourcePart_ = std::move(sourcePart);
        sourceXml_ = std::move(sourceXml);
        drawingRelationshipId_ = std::move(drawingRelationshipId);
        imported_ = true;
        chartDirty_ = false;
        sheetDirty_ = false;
        customViewsDirty_ = false;
        printerSettingsDirty_ = false;
    }
    void setImportedPrinterSettings(std::string sourcePart, std::string relationshipId, std::string data) {
        printerSettingsSourcePart_ = std::move(sourcePart);
        printerSettingsRelationshipId_ = std::move(relationshipId);
        printerSettingsData_ = std::move(data);
        printerSettingsDirty_ = false;
    }
    void clearDirty() const noexcept { chartDirty_ = false; sheetDirty_ = false; printerSettingsDirty_ = false; }

    std::string name_;
    std::optional<Chart> chart_;
    ChartsheetProperties properties_;
    ChartsheetView view_;
    ChartsheetProtection protection_;
    PageMargins pageMargins_;
    PageSetup pageSetup_;
    HeaderFooter headerFooter_;
    std::vector<CustomChartsheetView> customViews_;
    std::string customViewsRawXml_;
    std::string sourcePart_;
    mutable std::string sourceXml_;
    std::string drawingRelationshipId_{"rId1"};
    std::optional<std::string> printerSettingsData_;
    std::string printerSettingsSourcePart_;
    std::string printerSettingsRelationshipId_;
    bool pageMarginsPresent_{false};
    bool pageSetupPresent_{false};
    bool headerFooterPresent_{false};
    bool customViewsDirty_{true};
    bool imported_{false};
    mutable bool chartDirty_{true};
    mutable bool sheetDirty_{true};
    mutable bool printerSettingsDirty_{false};
};

} // namespace xlpp
