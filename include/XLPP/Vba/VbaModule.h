#pragma once

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

namespace xlpp {

enum class VbaModuleType {
    Standard,
    Document,
    Class,
    // Designer modules back Office Forms/UserForms and other registered
    // designers. Their source lives in VBA/<module>, while designer-specific
    // binary state lives in a sibling root storage with the same module name.
    Designer
};

struct VbaModule {
    std::string name;
    std::string source;
    VbaModuleType type{VbaModuleType::Standard};
    std::string docString;
    bool readOnly{false};
    bool isPrivate{false};
    int helpContextId{0};
    // CLSID written by PROJECT Package= for a Designer module. The default is
    // the Microsoft Forms 2.0 UserForm designer CLSID used by common Excel
    // UserForms; callers may override it for another registered designer.
    std::string designerClassId{"{AC9F2F90-E877-11CE-9F68-00AA00574A4F}"};
    // Attribute VB_Base value for a designer module. Imported UserForms keep
    // this exact value; new forms use the MS-OVBA example-compatible default.
    std::string designerBaseClass{"0{842E9C5E-88B5-439A-912E-4C2D9AA0EC27}{2DC3C962-DA1C-47BA-AB63-E9D578FC2637}"};
};

enum class VbaReferenceKind {
    Registered,
    Project,
    Control
};

struct VbaReference {
    std::string name;
    // Registered type-library LibId, e.g. *\G{GUID}#major.minor#lcid#path#description.
    std::string libid;
    VbaReferenceKind kind{VbaReferenceKind::Registered};
    // REFERENCEPROJECT additionally carries a relative LibId and project
    // version. These remain empty/zero for a registered type-library reference.
    std::string relativeLibid;
    std::uint32_t majorVersion{0};
    std::uint16_t minorVersion{0};
    // REFERENCECONTROL metadata used by ActiveX/UserForm designer libraries.
    // `libid` is the extended type-library LibId for this kind.
    std::string twiddledLibid;
    std::string extendedName;
    std::string originalTypeLib;
    std::uint32_t controlCookie{0};
};

struct VbaDesignerStream {
    // Path relative to the designer storage root. Nested storage components
    // are separated by '/'. Bytes are preserved exactly.
    std::string path;
    std::vector<unsigned char> data;
};

struct VbaDesignerStorage {
    // Root-storage name; for a UserForm this normally equals MODULESTREAMNAME.
    std::string name;
    std::vector<VbaDesignerStream> streams;

    const VbaDesignerStream* findStream(const std::string& relativePath) const noexcept {
        for (const auto& stream : streams) if (stream.path == relativePath) return &stream;
        return nullptr;
    }
};


struct VbaUserFormProperties {
    std::uint8_t minorVersion{0};
    std::uint8_t majorVersion{0};
    std::uint32_t propertyMask{0};
    std::optional<std::uint32_t> backColor;
    std::optional<std::uint32_t> foreColor;
    std::optional<std::uint32_t> nextAvailableId;
    std::optional<std::uint32_t> booleanProperties;
    std::optional<std::uint8_t> borderStyle;
    std::optional<std::uint8_t> mousePointer;
    std::optional<std::uint8_t> scrollBars;
    std::optional<std::int32_t> groupCount;
    std::optional<std::uint8_t> cycle;
    std::optional<std::uint8_t> specialEffect;
    std::optional<std::uint32_t> borderColor;
    std::optional<std::string> caption;
    std::optional<std::uint32_t> zoom;
    std::optional<std::uint8_t> pictureAlignment;
    std::optional<std::uint8_t> pictureSizeMode;
    std::optional<std::uint32_t> shapeCookie;
    std::optional<std::uint32_t> drawBuffer;
    std::optional<std::int32_t> displayedWidth;
    std::optional<std::int32_t> displayedHeight;
    std::optional<std::int32_t> logicalWidth;
    std::optional<std::int32_t> logicalHeight;
    std::optional<std::int32_t> scrollTop;
    std::optional<std::int32_t> scrollLeft;
};

struct VbaUserFormPropertiesPatch {
    std::optional<std::uint32_t> backColor;
    std::optional<std::uint32_t> foreColor;
    std::optional<std::uint32_t> nextAvailableId;
    std::optional<std::uint32_t> booleanProperties;
    std::optional<std::uint8_t> borderStyle;
    std::optional<std::uint8_t> mousePointer;
    std::optional<std::uint8_t> scrollBars;
    std::optional<std::int32_t> groupCount;
    std::optional<std::uint8_t> cycle;
    std::optional<std::uint8_t> specialEffect;
    std::optional<std::uint32_t> borderColor;
    std::optional<std::string> caption;
    std::optional<std::uint32_t> zoom;
    std::optional<std::uint8_t> pictureAlignment;
    std::optional<std::uint8_t> pictureSizeMode;
    std::optional<std::uint32_t> shapeCookie;
    std::optional<std::uint32_t> drawBuffer;
    std::optional<std::int32_t> displayedWidth;
    std::optional<std::int32_t> displayedHeight;
    std::optional<std::int32_t> logicalWidth;
    std::optional<std::int32_t> logicalHeight;
    std::optional<std::int32_t> scrollTop;
    std::optional<std::int32_t> scrollLeft;
};

struct VbaUserFormInspection {
    bool valid{false};
    std::string error;
    VbaUserFormProperties properties;
    std::size_t trailingBytes{0};
};

// Site-level metadata for one embedded UserForm control. This comes from the
// MS-OFORMS FormSiteData/OleSiteConcreteControl structure in stream "f"; the
// control-specific payload remains in stream "o" and is not reinterpreted here.
enum class VbaUserFormControlKind {
    Unknown,
    Form,
    Image,
    Frame,
    MorphData,
    SpinButton,
    CommandButton,
    TabStrip,
    Label,
    TextBox,
    ListBox,
    ComboBox,
    CheckBox,
    OptionButton,
    ToggleButton,
    ScrollBar,
    MultiPage,
    CustomClass
};

struct VbaUserFormControlSite {
    std::uint8_t depth{0};
    std::uint8_t siteType{1};
    std::uint16_t version{0};
    std::uint32_t propertyMask{0};
    std::optional<std::string> name;
    std::optional<std::string> tag;
    std::optional<std::int32_t> id;
    std::optional<std::int32_t> helpContextId;
    std::optional<std::uint32_t> bitFlags;
    std::optional<std::uint32_t> objectStreamSize;
    std::optional<std::int16_t> tabIndex;
    std::optional<std::uint16_t> clsidCacheIndex;
    VbaUserFormControlKind kind{VbaUserFormControlKind::Unknown};
    std::optional<std::uint16_t> groupId;
    std::optional<std::string> controlTipText;
    std::optional<std::string> runtimeLicenseKey;
    std::optional<std::string> controlSource;
    std::optional<std::string> rowSource;
    std::optional<std::int32_t> top;
    std::optional<std::int32_t> left;
    // Slice of the Designer Storage "o" stream owned by this site according
    // to ObjectStreamSize. P1E exposes the bytes losslessly before semantic
    // parsing of individual control classes is attempted.
    std::size_t objectStreamOffset{0};
    std::vector<unsigned char> objectData;
};

struct VbaUserFormControlSitePatch {
    std::optional<std::string> name;
    std::optional<std::string> tag;
    std::optional<std::int32_t> helpContextId;
    std::optional<std::uint32_t> bitFlags;
    std::optional<std::int16_t> tabIndex;
    std::optional<std::uint16_t> groupId;
    std::optional<std::string> controlTipText;
    std::optional<std::string> controlSource;
    std::optional<std::string> rowSource;
    std::optional<std::int32_t> top;
    std::optional<std::int32_t> left;
};

struct VbaUserFormControlObjectProperties {
    VbaUserFormControlKind kind{VbaUserFormControlKind::Unknown};
    std::uint8_t minorVersion{0};
    std::uint8_t majorVersion{0};
    std::uint16_t cbControl{0};
    std::uint32_t propertyMask{0};
    bool semanticPropertiesSupported{false};
    std::optional<std::uint32_t> foreColor;
    std::optional<std::uint32_t> backColor;
    std::optional<std::uint32_t> variousPropertyBits;
    std::optional<std::string> caption;
    std::optional<std::uint32_t> picturePosition;
    std::optional<std::uint8_t> mousePointer;
    std::optional<std::uint32_t> borderColor;
    std::optional<std::uint16_t> borderStyle;
    std::optional<std::uint16_t> specialEffect;
    std::optional<std::uint16_t> accelerator;
    std::optional<std::int32_t> width;
    std::optional<std::int32_t> height;
};

struct VbaUserFormControlObjectInspection {
    bool valid{false};
    std::string error;
    VbaUserFormControlObjectProperties properties;
    std::size_t objectBytes{0};
    std::size_t semanticSectionBytes{0};
    std::size_t trailingBytes{0};
};

struct VbaUserFormControlObjectPatch {
    std::optional<std::uint32_t> foreColor;
    std::optional<std::uint32_t> backColor;
    std::optional<std::uint32_t> variousPropertyBits;
    std::optional<std::string> caption;
    std::optional<std::uint32_t> picturePosition;
    std::optional<std::uint8_t> mousePointer;
    std::optional<std::uint32_t> borderColor;
    std::optional<std::uint16_t> borderStyle;
    std::optional<std::uint16_t> specialEffect;
    std::optional<std::uint16_t> accelerator;
    std::optional<std::int32_t> width;
    std::optional<std::int32_t> height;
};

struct VbaUserFormControlsInspection {
    bool valid{false};
    std::string error;
    std::size_t siteDataOffset{0};
    std::size_t classInfoCount{0};
    std::vector<VbaUserFormControlSite> controls;
    std::size_t totalObjectStreamBytes{0};
    std::size_t objectStreamBytes{0};
    std::size_t unassignedObjectStreamBytes{0};
};

struct VbaDesignerValidationIssue {
    std::string designerName;
    std::string message;
};

struct VbaDesignerValidationReport {
    std::size_t designerModules{0};
    std::size_t designerStorages{0};
    std::size_t validUserFormStreams{0};
    std::vector<VbaDesignerValidationIssue> issues;
    bool ok() const noexcept { return issues.empty(); }
};

struct VbaProjectInfo {
    std::string name{"VBAProject"};
    std::string description;
    std::string helpFile;
    int helpContextId{0};
    std::string constants;
    // Locale/code-page metadata from the OVBA dir stream. Defaults match the
    // source-only Windows VBA project emitted by previous XL++ releases.
    std::uint32_t systemKind{3};
    std::uint32_t lcid{0x0409};
    std::uint32_t lcidInvoke{0x0409};
    std::uint16_t codePage{1252};
    // Registered references found in / written to the VBA dir stream. XL++
    // always retains its minimal stdole/Office baseline and merges these by
    // LibId, so callers can add project-specific type libraries safely.
    std::vector<VbaReference> references;
    // Raw designer storages (UserForms/ActiveX designers). XL++ preserves
    // their recursive CFB stream tree byte-for-byte while exposing the module
    // source separately through VbaModuleType::Designer. P1D/P1E add safe
    // semantic Form and control-site layers without discarding unknown streams.
    std::vector<VbaDesignerStorage> designerStorages;
    // Project ID written to the PROJECT stream. XL++ keeps a stable default
    // for source-generated projects but callers may supply their own GUID text.
    std::string projectId{"{9E394C0B-697E-4AEE-9FA6-446F51FB30DC}"};
};

} // namespace xlpp
