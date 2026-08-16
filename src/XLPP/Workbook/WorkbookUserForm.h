#pragma once
#include <XLPP/Vba/VbaModule.h>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

namespace xlpp {
namespace internal {

// MS-OFORMS UserForm stream parsing helpers shared by the VBA modules. These
// decode the opaque frm/control binary layouts into inspection records.

struct UserFormStreamLayout {
    VbaUserFormInspection inspection;
    std::size_t sectionEnd{0};
    std::size_t dataEnd{0};
    std::size_t extraStart{0};
    std::optional<std::size_t> backColor, foreColor, nextId, booleanProperties;
    std::optional<std::size_t> borderStyle, mousePointer, scrollBars, groupCount;
    std::optional<std::size_t> cycle, specialEffect, borderColor, captionCount;
    std::optional<std::size_t> zoom, pictureAlignment, pictureSizeMode, shapeCookie, drawBuffer;
    std::optional<std::size_t> displayedSize, logicalSize, scrollPosition, captionData;
    std::size_t captionBytes{0};
    bool captionCompressed{false};
};

struct UserFormSiteStringLayout {
    std::optional<std::size_t> countOffset;
    std::optional<std::size_t> dataOffset;
    std::size_t byteCount{0};
    bool compressed{false};
};

struct UserFormControlSiteLayout {
    VbaUserFormControlSite site;
    std::size_t begin{0};
    std::size_t end{0};
    std::optional<std::size_t> helpContextId, bitFlags, objectStreamSize, tabIndex, clsidCacheIndex, groupId;
    std::optional<std::size_t> position;
    UserFormSiteStringLayout name, tag, controlTipText, runtimeLicenseKey, controlSource, rowSource;
};

struct UserFormSiteDataLayout {
    VbaUserFormControlsInspection inspection;
    std::size_t countOfBytesOffset{0};
    std::vector<UserFormControlSiteLayout> sites;
};

struct UserFormControlObjectLayout {
    VbaUserFormControlObjectInspection inspection;
    std::optional<std::size_t> foreColor, backColor, variousPropertyBits, captionCount,
        picturePosition, mousePointer, borderColor, borderStyle, specialEffect,
        accelerator, size;
    std::optional<std::size_t> captionData;
    std::size_t captionBytes{0};
    bool captionCompressed{false};
    std::size_t semanticEnd{0};

    std::optional<std::size_t> scrollBars, displayStyle, enterKeyBehavior, tabKeyBehavior,
        maxLength, wordWrap, textCount, autoWordSelect, integralHeight, passwordChar,
        valueCount, lineCount, multiLine, multiSelect, hideSelection, dataEntry, dragBehavior,
        listRows;
    std::optional<std::size_t> textData, valueData;
    std::size_t textBytes{0}, valueBytes{0};
    bool textCompressed{false}, valueCompressed{false};
    std::optional<std::size_t> groupNameCount, groupNumber, tripleState;
    std::optional<std::size_t> groupNameData;
    std::size_t groupNameBytes{0};
    bool groupNameCompressed{false};
    std::optional<std::size_t> min, max, smallChange, largeChange, orientation;
    std::optional<std::size_t> listWidth, boundColumn, textColumn, columnCount, columnWidthsCount;
    std::optional<std::size_t> style, listStyle, matchEntry, showDropButtonWhen, dropButtonStyle,
        matchFound, iMEMode, iMEStatus;
    std::optional<std::size_t> columnWidthsData;
    std::size_t columnWidthsBytes{0};
    bool columnWidthsCompressed{false};
    bool maskIs64Bit{false};
};

// Serializes a high-level VbaUserFormDesign into a complete Designer Storage
// ("f" FormControl + FormSiteData, "o" control-object streams, "f3"
// FormControlEx and "Designer" Designer_CompData). The result round-trips
// through the inspection parsers in this module and is installable through
// Workbook::addUserForm / setVbaDesignerStorage.
VbaDesignerStorage buildUserFormDesign(const VbaUserFormDesign& design);

// Byte-level MS-OFORMS primitives shared by UserForm readers and writers.
std::size_t userFormAlign(std::size_t offset, std::size_t alignment);
std::uint16_t userFormU16(const std::vector<unsigned char>& bytes, std::size_t offset);
std::uint32_t userFormU32(const std::vector<unsigned char>& bytes, std::size_t offset);
std::int32_t userFormI32(const std::vector<unsigned char>& bytes, std::size_t offset);
std::int16_t userFormI16(const std::vector<unsigned char>& bytes, std::size_t offset);
void userFormPutU16(std::vector<unsigned char>& bytes, std::size_t offset, std::uint16_t value);
void userFormPutU32(std::vector<unsigned char>& bytes, std::size_t offset, std::uint32_t value);
void userFormPutI32(std::vector<unsigned char>& bytes, std::size_t offset, std::int32_t value);
void userFormPutI16(std::vector<unsigned char>& bytes, std::size_t offset, std::int16_t value);
std::vector<std::uint32_t> userFormUtf8Codepoints(const std::string& text);
std::string userFormAppendUtf8(std::string out, std::uint32_t cp);
std::string decodeUserFormString(const std::vector<unsigned char>& bytes,
                                 std::size_t offset, std::size_t byteCount, bool compressed);
// Encodes text as a LengthAndCompression field payload: returns (bytes,
// compressed). The length header is written by the caller via userFormPutU32.
std::pair<std::vector<unsigned char>, bool> encodeUserFormString(const std::string& text);

UserFormStreamLayout parseUserFormStream(const std::vector<unsigned char>& bytes);
UserFormSiteDataLayout parseUserFormSiteData(const std::vector<unsigned char>& bytes, std::size_t minimumOffset);
UserFormControlObjectLayout parseSimpleUserFormControlObject(const std::vector<unsigned char>& bytes,
                                                            VbaUserFormControlKind kind);
void requireUserFormField(bool requested, const std::optional<std::size_t>& offset, const char* name);
void requireUserFormObjectField(bool requested, const std::optional<std::size_t>& offset, const char* name);
void requireUserFormSiteField(bool requested, const std::optional<std::size_t>& offset, const char* name);

} // namespace internal
} // namespace xlpp
