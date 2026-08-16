#include <XLPP/Workbook/Workbook.h>
#include "../Vba/VbaProjectBinary.h"
#include "WorkbookUserForm.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <optional>
#include <set>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

using xlpp::internal::UserFormStreamLayout;
using xlpp::internal::UserFormSiteStringLayout;
using xlpp::internal::UserFormControlSiteLayout;
using xlpp::internal::UserFormSiteDataLayout;
using xlpp::internal::UserFormControlObjectLayout;
using xlpp::internal::userFormAlign;
using xlpp::internal::userFormU16;
using xlpp::internal::userFormU32;
using xlpp::internal::userFormI32;
using xlpp::internal::userFormI16;
using xlpp::internal::userFormPutU16;
using xlpp::internal::userFormPutU32;
using xlpp::internal::userFormPutI32;
using xlpp::internal::userFormPutI16;
using xlpp::internal::userFormUtf8Codepoints;
using xlpp::internal::userFormAppendUtf8;
using xlpp::internal::decodeUserFormString;
using xlpp::internal::encodeUserFormString;
using xlpp::internal::parseUserFormStream;
using xlpp::internal::parseUserFormSiteData;
using xlpp::internal::parseSimpleUserFormControlObject;
using xlpp::internal::requireUserFormField;
using xlpp::internal::requireUserFormObjectField;
using xlpp::internal::requireUserFormSiteField;

namespace xlpp {
namespace {

std::vector<std::string> ensureWorksheetVbaCodeNames(std::deque<Worksheet>& sheets) {
    auto lower = [](std::string value) {
        std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
        return value;
    };
    std::set<std::string> used;
    std::vector<std::string> result(sheets.size());
    for (std::size_t i = 0; i < sheets.size(); ++i) {
        if (sheets[i].vbaCodeName().empty()) continue;
        internal::validateVbaModuleName(sheets[i].vbaCodeName());
        const auto key = lower(sheets[i].vbaCodeName());
        if (!used.insert(key).second)
            throw std::invalid_argument("Duplicate worksheet VBA code name: " + sheets[i].vbaCodeName());
        result[i] = sheets[i].vbaCodeName();
    }
    std::size_t candidate = 1;
    for (std::size_t i = 0; i < result.size(); ++i) {
        if (!result[i].empty()) continue;
        for (;;) {
            auto name = "Sheet" + std::to_string(candidate++);
            if (used.insert(lower(name)).second) {
                result[i] = std::move(name);
                break;
            }
        }
    }
    for (std::size_t i = 0; i < sheets.size(); ++i)
        if (sheets[i].vbaCodeName().empty()) sheets[i].setVbaCodeName(result[i]);
    return result;
}

} // namespace


VbaUserFormInspection Workbook::inspectVbaUserForm(const std::string& storageName) const {
    const auto storages = vbaDesignerStorages();
    const auto it = std::find_if(storages.begin(), storages.end(), [&](const auto& storage) { return storage.name == storageName; });
    if (it == storages.end()) return {false, "Designer storage not found: " + storageName, {}, 0};
    const auto* stream = it->findStream("f");
    if (!stream) return {false, "Designer storage does not contain an MS-OFORMS Form stream named f", {}, 0};
    return parseUserFormStream(stream->data).inspection;
}

bool Workbook::updateVbaUserFormProperties(const std::string& storageName,
                                           const VbaUserFormPropertiesPatch& patch) {
    auto info = vbaProjectInfo();
    auto storageIt = std::find_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
        return storage.name == storageName;
    });
    if (storageIt == info.designerStorages.end()) return false;
    auto streamIt = std::find_if(storageIt->streams.begin(), storageIt->streams.end(), [](const auto& stream) { return stream.path == "f"; });
    if (streamIt == storageIt->streams.end()) return false;
    auto layout = parseUserFormStream(streamIt->data);
    if (!layout.inspection.valid) throw std::runtime_error("Malformed UserForm Form stream: " + layout.inspection.error);
    auto& bytes = streamIt->data;

    requireUserFormField(patch.backColor.has_value(), layout.backColor, "BackColor");
    requireUserFormField(patch.foreColor.has_value(), layout.foreColor, "ForeColor");
    requireUserFormField(patch.nextAvailableId.has_value(), layout.nextId, "NextAvailableID");
    requireUserFormField(patch.booleanProperties.has_value(), layout.booleanProperties, "BooleanProperties");
    requireUserFormField(patch.borderStyle.has_value(), layout.borderStyle, "BorderStyle");
    requireUserFormField(patch.mousePointer.has_value(), layout.mousePointer, "MousePointer");
    requireUserFormField(patch.scrollBars.has_value(), layout.scrollBars, "ScrollBars");
    requireUserFormField(patch.groupCount.has_value(), layout.groupCount, "GroupCount");
    requireUserFormField(patch.cycle.has_value(), layout.cycle, "Cycle");
    requireUserFormField(patch.specialEffect.has_value(), layout.specialEffect, "SpecialEffect");
    requireUserFormField(patch.borderColor.has_value(), layout.borderColor, "BorderColor");
    requireUserFormField(patch.caption.has_value(), layout.captionCount, "Caption");
    requireUserFormField(patch.zoom.has_value(), layout.zoom, "Zoom");
    requireUserFormField(patch.pictureAlignment.has_value(), layout.pictureAlignment, "PictureAlignment");
    requireUserFormField(patch.pictureSizeMode.has_value(), layout.pictureSizeMode, "PictureSizeMode");
    requireUserFormField(patch.shapeCookie.has_value(), layout.shapeCookie, "ShapeCookie");
    requireUserFormField(patch.drawBuffer.has_value(), layout.drawBuffer, "DrawBuffer");
    requireUserFormField(patch.displayedWidth.has_value() || patch.displayedHeight.has_value(), layout.displayedSize, "DisplayedSize");
    requireUserFormField(patch.logicalWidth.has_value() || patch.logicalHeight.has_value(), layout.logicalSize, "LogicalSize");
    requireUserFormField(patch.scrollLeft.has_value() || patch.scrollTop.has_value(), layout.scrollPosition, "ScrollPosition");

    if (patch.backColor) userFormPutU32(bytes, *layout.backColor, *patch.backColor);
    if (patch.foreColor) userFormPutU32(bytes, *layout.foreColor, *patch.foreColor);
    if (patch.nextAvailableId) userFormPutU32(bytes, *layout.nextId, *patch.nextAvailableId);
    if (patch.booleanProperties) userFormPutU32(bytes, *layout.booleanProperties, *patch.booleanProperties);
    if (patch.borderStyle) bytes[*layout.borderStyle] = *patch.borderStyle;
    if (patch.mousePointer) bytes[*layout.mousePointer] = *patch.mousePointer;
    if (patch.scrollBars) bytes[*layout.scrollBars] = *patch.scrollBars;
    if (patch.groupCount) userFormPutI32(bytes, *layout.groupCount, *patch.groupCount);
    if (patch.cycle) bytes[*layout.cycle] = *patch.cycle;
    if (patch.specialEffect) bytes[*layout.specialEffect] = *patch.specialEffect;
    if (patch.borderColor) userFormPutU32(bytes, *layout.borderColor, *patch.borderColor);
    if (patch.zoom) userFormPutU32(bytes, *layout.zoom, *patch.zoom);
    if (patch.pictureAlignment) bytes[*layout.pictureAlignment] = *patch.pictureAlignment;
    if (patch.pictureSizeMode) bytes[*layout.pictureSizeMode] = *patch.pictureSizeMode;
    if (patch.shapeCookie) userFormPutU32(bytes, *layout.shapeCookie, *patch.shapeCookie);
    if (patch.drawBuffer) userFormPutU32(bytes, *layout.drawBuffer, *patch.drawBuffer);
    if (patch.displayedWidth) userFormPutI32(bytes, *layout.displayedSize, *patch.displayedWidth);
    if (patch.displayedHeight) userFormPutI32(bytes, *layout.displayedSize + 4, *patch.displayedHeight);
    if (patch.logicalWidth) userFormPutI32(bytes, *layout.logicalSize, *patch.logicalWidth);
    if (patch.logicalHeight) userFormPutI32(bytes, *layout.logicalSize + 4, *patch.logicalHeight);
    if (patch.scrollLeft) userFormPutI32(bytes, *layout.scrollPosition, *patch.scrollLeft);
    if (patch.scrollTop) userFormPutI32(bytes, *layout.scrollPosition + 4, *patch.scrollTop);

    if (patch.caption) {
        auto [encoded, compressed] = encodeUserFormString(*patch.caption);
        const auto oldBegin = *layout.captionData;
        const auto oldEnd = oldBegin + layout.captionBytes;
        std::vector<unsigned char> rebuilt;
        rebuilt.reserve(bytes.size() - layout.captionBytes + encoded.size());
        rebuilt.insert(rebuilt.end(), bytes.begin(), bytes.begin() + static_cast<std::ptrdiff_t>(oldBegin));
        rebuilt.insert(rebuilt.end(), encoded.begin(), encoded.end());
        rebuilt.insert(rebuilt.end(), bytes.begin() + static_cast<std::ptrdiff_t>(oldEnd), bytes.end());
        const auto newCount = static_cast<std::uint32_t>(encoded.size()) | (compressed ? 0x80000000u : 0u);
        userFormPutU32(rebuilt, *layout.captionCount, newCount);
        const auto delta = static_cast<std::ptrdiff_t>(encoded.size()) - static_cast<std::ptrdiff_t>(layout.captionBytes);
        const auto newCb = static_cast<std::ptrdiff_t>(userFormU16(rebuilt, 2)) + delta;
        if (newCb < 4 || newCb > 0xFFFF) throw std::overflow_error("UserForm cbForm exceeds 16-bit storage after caption edit");
        userFormPutU16(rebuilt, 2, static_cast<std::uint16_t>(newCb));
        bytes = std::move(rebuilt);
    }

    setVbaProjectInfo(std::move(info));
    return true;
}

VbaUserFormControlsInspection Workbook::inspectVbaUserFormControls(const std::string& storageName) const {
    const auto storages = vbaDesignerStorages();
    const auto it = std::find_if(storages.begin(), storages.end(), [&](const auto& storage) { return storage.name == storageName; });
    if (it == storages.end()) return {false, "Designer storage not found: " + storageName};
    const auto* stream = it->findStream("f");
    if (!stream) return {false, "Designer storage does not contain an MS-OFORMS Form stream named f"};
    const auto form = parseUserFormStream(stream->data);
    if (!form.inspection.valid) return {false, form.inspection.error};
    auto inspection = parseUserFormSiteData(stream->data, form.sectionEnd).inspection;
    if (!inspection.valid) return inspection;

    const auto* objectStream = it->findStream("o");
    if (inspection.totalObjectStreamBytes != 0 && !objectStream) {
        inspection.valid = false;
        inspection.error = "UserForm control sites reference object-stream bytes but Designer Storage has no stream named o";
        return inspection;
    }
    if (objectStream) {
        inspection.objectStreamBytes = objectStream->data.size();
        if (objectStream->data.size() < inspection.totalObjectStreamBytes) {
            inspection.valid = false;
            inspection.error = "UserForm object stream is shorter than the sum of OleSite ObjectStreamSize values";
            return inspection;
        }
        std::size_t offset = 0;
        for (auto& control : inspection.controls) {
            const auto size = static_cast<std::size_t>(control.objectStreamSize.value_or(0));
            control.objectStreamOffset = offset;
            control.objectData.assign(objectStream->data.begin() + static_cast<std::ptrdiff_t>(offset),
                                      objectStream->data.begin() + static_cast<std::ptrdiff_t>(offset + size));
            offset += size;
        }
        inspection.unassignedObjectStreamBytes = objectStream->data.size() - offset;
    }
    return inspection;
}

VbaUserFormControlObjectInspection Workbook::inspectVbaUserFormControlObject(const std::string& storageName,
                                                                                 std::size_t controlIndex) const {
    const auto controls = inspectVbaUserFormControls(storageName);
    if (!controls.valid) return {false, controls.error};
    if (controlIndex >= controls.controls.size()) return {false, "UserForm control-site index is out of range"};
    const auto& control = controls.controls[controlIndex];
    return parseSimpleUserFormControlObject(control.objectData, control.kind).inspection;
}

bool Workbook::updateVbaUserFormControlObject(const std::string& storageName,
                                              std::size_t controlIndex,
                                              const VbaUserFormControlObjectPatch& patch) {
    auto info = vbaProjectInfo();
    auto storageIt = std::find_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
        return storage.name == storageName;
    });
    if (storageIt == info.designerStorages.end()) return false;
    auto fIt = std::find_if(storageIt->streams.begin(), storageIt->streams.end(), [](const auto& stream) { return stream.path == "f"; });
    auto oIt = std::find_if(storageIt->streams.begin(), storageIt->streams.end(), [](const auto& stream) { return stream.path == "o"; });
    if (fIt == storageIt->streams.end() || oIt == storageIt->streams.end()) return false;

    auto form = parseUserFormStream(fIt->data);
    if (!form.inspection.valid) throw std::runtime_error("Cannot edit malformed UserForm Form stream: " + form.inspection.error);
    auto sites = parseUserFormSiteData(fIt->data, form.sectionEnd);
    if (!sites.inspection.valid) throw std::runtime_error("Cannot edit UserForm control object: " + sites.inspection.error);
    if (controlIndex >= sites.sites.size()) throw std::out_of_range("UserForm control-site index is out of range");
    auto& site = sites.sites[controlIndex];
    const auto objectSize = static_cast<std::size_t>(site.site.objectStreamSize.value_or(0));
    std::size_t objectOffset = 0;
    for (std::size_t i = 0; i < controlIndex; ++i)
        objectOffset += static_cast<std::size_t>(sites.sites[i].site.objectStreamSize.value_or(0));
    if (objectOffset + objectSize > oIt->data.size())
        throw std::runtime_error("UserForm object stream is shorter than the control's ObjectStreamSize slice");
    std::vector<unsigned char> object(oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset),
                                      oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset + objectSize));
    auto layout = parseSimpleUserFormControlObject(object, site.site.kind);
    if (!layout.inspection.valid) throw std::runtime_error("Cannot edit malformed UserForm control object: " + layout.inspection.error);
    if (!layout.inspection.properties.semanticPropertiesSupported)
        throw std::invalid_argument("Semantic object editing is currently supported only for built-in CommandButton and Label controls");

    requireUserFormObjectField(patch.foreColor.has_value(), layout.foreColor, "ForeColor");
    requireUserFormObjectField(patch.backColor.has_value(), layout.backColor, "BackColor");
    requireUserFormObjectField(patch.variousPropertyBits.has_value(), layout.variousPropertyBits, "VariousPropertyBits");
    requireUserFormObjectField(patch.caption.has_value(), layout.captionCount, "Caption");
    requireUserFormObjectField(patch.picturePosition.has_value(), layout.picturePosition, "PicturePosition");
    requireUserFormObjectField(patch.mousePointer.has_value(), layout.mousePointer, "MousePointer");
    requireUserFormObjectField(patch.borderColor.has_value(), layout.borderColor, "BorderColor");
    requireUserFormObjectField(patch.borderStyle.has_value(), layout.borderStyle, "BorderStyle");
    requireUserFormObjectField(patch.specialEffect.has_value(), layout.specialEffect, "SpecialEffect");
    requireUserFormObjectField(patch.accelerator.has_value(), layout.accelerator, "Accelerator");
    requireUserFormObjectField(patch.width.has_value() || patch.height.has_value(), layout.size, "Size");

    // P1Y-A extended control fields. Each is optional in the patch; a field is
    // written only when the target control actually materializes it.
    requireUserFormObjectField(patch.scrollBars.has_value(), layout.scrollBars, "ScrollBars");
    requireUserFormObjectField(patch.displayStyle.has_value(), layout.displayStyle, "DisplayStyle");
    requireUserFormObjectField(patch.enterKeyBehavior.has_value(), layout.enterKeyBehavior, "EnterKeyBehavior");
    requireUserFormObjectField(patch.tabKeyBehavior.has_value(), layout.tabKeyBehavior, "TabKeyBehavior");
    requireUserFormObjectField(patch.maxLength.has_value(), layout.maxLength, "MaxLength");
    requireUserFormObjectField(patch.wordWrap.has_value(), layout.wordWrap, "WordWrap");
    requireUserFormObjectField(patch.text.has_value(), layout.textCount, "Text");
    requireUserFormObjectField(patch.autoWordSelect.has_value(), layout.autoWordSelect, "AutoWordSelect");
    requireUserFormObjectField(patch.integralHeight.has_value(), layout.integralHeight, "IntegralHeight");
    requireUserFormObjectField(patch.passwordChar.has_value(), layout.passwordChar, "PasswordChar");
    requireUserFormObjectField(patch.value.has_value(), layout.valueCount, "Value");
    requireUserFormObjectField(patch.multiLine.has_value(), layout.multiLine, "MultiLine");
    requireUserFormObjectField(patch.multiSelect.has_value(), layout.multiSelect, "MultiSelect");
    requireUserFormObjectField(patch.hideSelection.has_value(), layout.hideSelection, "HideSelection");
    requireUserFormObjectField(patch.dataEntry.has_value(), layout.dataEntry, "DataEntry");
    requireUserFormObjectField(patch.dragBehavior.has_value(), layout.dragBehavior, "DragBehavior");
    requireUserFormObjectField(patch.listRows.has_value(), layout.listRows, "ListRows");
    requireUserFormObjectField(patch.groupName.has_value(), layout.groupNameCount, "GroupName");
    requireUserFormObjectField(patch.groupNumber.has_value(), layout.groupNumber, "GroupNumber");
    requireUserFormObjectField(patch.tripleState.has_value(), layout.tripleState, "TripleState");
    requireUserFormObjectField(patch.min.has_value(), layout.min, "Min");
    requireUserFormObjectField(patch.max.has_value(), layout.max, "Max");
    requireUserFormObjectField(patch.smallChange.has_value(), layout.smallChange, "SmallChange");
    requireUserFormObjectField(patch.largeChange.has_value(), layout.largeChange, "LargeChange");
    requireUserFormObjectField(patch.orientation.has_value(), layout.orientation, "Orientation");
    requireUserFormObjectField(patch.listWidth.has_value(), layout.listWidth, "ListWidth");
    requireUserFormObjectField(patch.boundColumn.has_value(), layout.boundColumn, "BoundColumn");
    requireUserFormObjectField(patch.textColumn.has_value(), layout.textColumn, "TextColumn");
    requireUserFormObjectField(patch.columnCount.has_value(), layout.columnCount, "ColumnCount");
    requireUserFormObjectField(patch.columnWidths.has_value(), layout.columnWidthsCount, "ColumnWidths");
    requireUserFormObjectField(patch.style.has_value(), layout.style, "Style");
    requireUserFormObjectField(patch.listStyle.has_value(), layout.listStyle, "ListStyle");
    requireUserFormObjectField(patch.matchEntry.has_value(), layout.matchEntry, "MatchEntry");
    requireUserFormObjectField(patch.showDropButtonWhen.has_value(), layout.showDropButtonWhen, "ShowDropButtonWhen");
    requireUserFormObjectField(patch.dropButtonStyle.has_value(), layout.dropButtonStyle, "DropButtonStyle");
    requireUserFormObjectField(patch.matchFound.has_value(), layout.matchFound, "MatchFound");
    requireUserFormObjectField(patch.iMEMode.has_value(), layout.iMEMode, "IMEMode");
    requireUserFormObjectField(patch.iMEStatus.has_value(), layout.iMEStatus, "IMEStatus");

    if (patch.caption) {
        requireUserFormObjectField(true, layout.captionData, "Caption");
        auto [encoded, compressed] = encodeUserFormString(*patch.caption);
        const auto oldBegin = *layout.captionData;
        const auto oldEnd = oldBegin + layout.captionBytes;
        std::vector<unsigned char> rebuilt;
        rebuilt.reserve(object.size() - layout.captionBytes + encoded.size());
        rebuilt.insert(rebuilt.end(), object.begin(), object.begin() + static_cast<std::ptrdiff_t>(oldBegin));
        rebuilt.insert(rebuilt.end(), encoded.begin(), encoded.end());
        rebuilt.insert(rebuilt.end(), object.begin() + static_cast<std::ptrdiff_t>(oldEnd), object.end());
        userFormPutU32(rebuilt, *layout.captionCount,
                       static_cast<std::uint32_t>(encoded.size()) | (compressed ? 0x80000000u : 0u));
        const auto delta = static_cast<std::ptrdiff_t>(encoded.size()) - static_cast<std::ptrdiff_t>(layout.captionBytes);
        const auto newCb = static_cast<std::ptrdiff_t>(userFormU16(rebuilt, 2)) + delta;
        if (newCb < 4 || newCb > 0xFFFF) throw std::overflow_error("UserForm control cbControl exceeds 16-bit storage after caption edit");
        userFormPutU16(rebuilt, 2, static_cast<std::uint16_t>(newCb));
        object = std::move(rebuilt);
        layout = parseSimpleUserFormControlObject(object, site.site.kind);
        if (!layout.inspection.valid) throw std::runtime_error("UserForm control object became malformed after caption edit: " + layout.inspection.error);
    }

    if (patch.foreColor) userFormPutU32(object, *layout.foreColor, *patch.foreColor);
    if (patch.backColor) userFormPutU32(object, *layout.backColor, *patch.backColor);
    if (patch.variousPropertyBits) userFormPutU32(object, *layout.variousPropertyBits, *patch.variousPropertyBits);
    if (patch.picturePosition) userFormPutU32(object, *layout.picturePosition, *patch.picturePosition);
    if (patch.mousePointer) object[*layout.mousePointer] = *patch.mousePointer;
    if (patch.borderColor) userFormPutU32(object, *layout.borderColor, *patch.borderColor);
    if (patch.borderStyle) userFormPutU16(object, *layout.borderStyle, *patch.borderStyle);
    if (patch.specialEffect) userFormPutU16(object, *layout.specialEffect, *patch.specialEffect);
    if (patch.accelerator) userFormPutU16(object, *layout.accelerator, *patch.accelerator);
    if (patch.width) userFormPutI32(object, *layout.size, *patch.width);
    if (patch.height) userFormPutI32(object, *layout.size + 4, *patch.height);

    // P1Y-A extended scalar field writes.
    if (patch.scrollBars) userFormPutU16(object, *layout.scrollBars, *patch.scrollBars);
    if (patch.displayStyle) object[*layout.displayStyle] = *patch.displayStyle;
    if (patch.enterKeyBehavior) object[*layout.enterKeyBehavior] = *patch.enterKeyBehavior;
    if (patch.tabKeyBehavior) object[*layout.tabKeyBehavior] = *patch.tabKeyBehavior;
    if (patch.maxLength) userFormPutU32(object, *layout.maxLength, *patch.maxLength);
    if (patch.wordWrap) userFormPutU16(object, *layout.wordWrap, *patch.wordWrap);
    if (patch.autoWordSelect) object[*layout.autoWordSelect] = *patch.autoWordSelect;
    if (patch.integralHeight) object[*layout.integralHeight] = *patch.integralHeight;
    if (patch.passwordChar) userFormPutU16(object, *layout.passwordChar, *patch.passwordChar);
    if (patch.multiLine) userFormPutU16(object, *layout.multiLine, *patch.multiLine);
    if (patch.multiSelect) userFormPutU16(object, *layout.multiSelect, *patch.multiSelect);
    if (patch.hideSelection) userFormPutU16(object, *layout.hideSelection, *patch.hideSelection);
    if (patch.dataEntry) object[*layout.dataEntry] = *patch.dataEntry;
    if (patch.dragBehavior) object[*layout.dragBehavior] = *patch.dragBehavior;
    if (patch.listRows) userFormPutU32(object, *layout.listRows, *patch.listRows);
    if (patch.groupNumber) userFormPutU16(object, *layout.groupNumber, *patch.groupNumber);
    if (patch.tripleState) userFormPutU16(object, *layout.tripleState, *patch.tripleState);
    if (patch.min) userFormPutU32(object, *layout.min, *patch.min);
    if (patch.max) userFormPutU32(object, *layout.max, *patch.max);
    if (patch.smallChange) userFormPutU32(object, *layout.smallChange, *patch.smallChange);
    if (patch.largeChange) userFormPutU32(object, *layout.largeChange, *patch.largeChange);
    if (patch.orientation) object[*layout.orientation] = *patch.orientation;
    if (patch.listWidth) userFormPutU32(object, *layout.listWidth, *patch.listWidth);
    if (patch.boundColumn) userFormPutU32(object, *layout.boundColumn, *patch.boundColumn);
    if (patch.textColumn) userFormPutU32(object, *layout.textColumn, *patch.textColumn);
    if (patch.columnCount) userFormPutU32(object, *layout.columnCount, *patch.columnCount);
    if (patch.style) object[*layout.style] = *patch.style;
    if (patch.listStyle) object[*layout.listStyle] = *patch.listStyle;
    if (patch.matchEntry) object[*layout.matchEntry] = *patch.matchEntry;
    if (patch.showDropButtonWhen) object[*layout.showDropButtonWhen] = *patch.showDropButtonWhen;
    if (patch.dropButtonStyle) object[*layout.dropButtonStyle] = *patch.dropButtonStyle;
    if (patch.matchFound) object[*layout.matchFound] = *patch.matchFound;
    if (patch.iMEMode) object[*layout.iMEMode] = *patch.iMEMode;
    if (patch.iMEStatus) object[*layout.iMEStatus] = *patch.iMEStatus;

    // Variable-length string replacements for Text, Value and GroupName. These
    // follow the same length-and-compression rebuild used for Caption.
    const auto patchObjectString = [&](const std::optional<std::string>& requested,
                                       const std::optional<std::size_t>& countOffset,
                                       const std::optional<std::size_t>& dataOffset,
                                       std::size_t oldBytes) {
        if (!requested) return;
        if (!countOffset || !dataOffset)
            throw std::invalid_argument("UserForm control string property is not materialized in PropMask");
        auto [encoded, compressed] = encodeUserFormString(*requested);
        const auto oldBegin = *dataOffset;
        const auto oldEnd = oldBegin + oldBytes;
        // MS-OFORMS aligns every ExtraDataBlock string payload to a 4-byte
        // boundary and emits alignment bytes after the payload. When a string
        // length changes, the following string/Size fields must stay aligned.
        // Rebuild the prefix up to the old payload start, insert the new
        // payload, align to 4, then copy the tail from the old aligned Size
        // boundary so Size/StreamData keep their 4-byte phase.
        const auto oldTailStart = userFormAlign(oldEnd, 4);
        std::vector<unsigned char> rebuilt;
        rebuilt.reserve(object.size() - oldBytes + encoded.size() + 3);
        rebuilt.insert(rebuilt.end(), object.begin(), object.begin() + static_cast<std::ptrdiff_t>(oldBegin));
        rebuilt.insert(rebuilt.end(), encoded.begin(), encoded.end());
        while ((rebuilt.size() % 4) != 0) rebuilt.push_back(0);
        rebuilt.insert(rebuilt.end(), object.begin() + static_cast<std::ptrdiff_t>(oldTailStart), object.end());
        userFormPutU32(rebuilt, *countOffset,
                       static_cast<std::uint32_t>(encoded.size()) | (compressed ? 0x80000000u : 0u));
        const auto newCb = static_cast<std::ptrdiff_t>(rebuilt.size()) - static_cast<std::ptrdiff_t>(4);
        if (newCb < 4 || newCb > 0xFFFF)
            throw std::overflow_error("UserForm control cbControl exceeds 16-bit storage after string edit");
        userFormPutU16(rebuilt, 2, static_cast<std::uint16_t>(newCb));
        object = std::move(rebuilt);
        layout = parseSimpleUserFormControlObject(object, site.site.kind);
        if (!layout.inspection.valid)
            throw std::runtime_error("UserForm control object became malformed after string edit: " + layout.inspection.error);
    };
    patchObjectString(patch.text, layout.textCount, layout.textData, layout.textBytes);
    patchObjectString(patch.value, layout.valueCount, layout.valueData, layout.valueBytes);
    patchObjectString(patch.groupName, layout.groupNameCount, layout.groupNameData, layout.groupNameBytes);
    patchObjectString(patch.columnWidths, layout.columnWidthsCount, layout.columnWidthsData, layout.columnWidthsBytes);

    if (object.size() != objectSize) {
        requireUserFormSiteField(true, site.objectStreamSize, "ObjectStreamSize");
        if (object.size() > 0xFFFFFFFFu) throw std::overflow_error("UserForm control object exceeds 32-bit ObjectStreamSize");
        userFormPutU32(fIt->data, *site.objectStreamSize, static_cast<std::uint32_t>(object.size()));
    }
    oIt->data.erase(oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset),
                    oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset + objectSize));
    oIt->data.insert(oIt->data.begin() + static_cast<std::ptrdiff_t>(objectOffset), object.begin(), object.end());
    setVbaProjectInfo(std::move(info));
    return true;
}

bool Workbook::updateVbaUserFormControlSite(const std::string& storageName,
                                            std::size_t controlIndex,
                                            const VbaUserFormControlSitePatch& patch) {
    auto info = vbaProjectInfo();
    auto storageIt = std::find_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
        return storage.name == storageName;
    });
    if (storageIt == info.designerStorages.end()) return false;
    auto streamIt = std::find_if(storageIt->streams.begin(), storageIt->streams.end(), [](const auto& stream) { return stream.path == "f"; });
    if (streamIt == storageIt->streams.end()) return false;
    auto& bytes = streamIt->data;
    const auto form = parseUserFormStream(bytes);
    if (!form.inspection.valid) throw std::runtime_error("Cannot edit malformed UserForm Form stream: " + form.inspection.error);

    auto parseSites = [&]() {
        auto parsed = parseUserFormSiteData(bytes, parseUserFormStream(bytes).sectionEnd);
        if (!parsed.inspection.valid) throw std::runtime_error("Cannot edit UserForm controls: " + parsed.inspection.error);
        if (controlIndex >= parsed.sites.size()) throw std::out_of_range("UserForm control-site index is out of range");
        return parsed;
    };

    auto patchString = [&](const std::optional<std::string>& requested, auto member, const char* name) {
        if (!requested) return;
        auto parsed = parseSites();
        auto& site = parsed.sites[controlIndex];
        const auto& layout = site.*member;
        requireUserFormSiteField(true, layout.countOffset, name);
        requireUserFormSiteField(true, layout.dataOffset, name);
        auto [encoded, compressed] = encodeUserFormString(*requested);
        const auto oldBytes = layout.byteCount;
        const auto dataOffset = *layout.dataOffset;
        bytes.erase(bytes.begin() + static_cast<std::ptrdiff_t>(dataOffset),
                    bytes.begin() + static_cast<std::ptrdiff_t>(dataOffset + oldBytes));
        bytes.insert(bytes.begin() + static_cast<std::ptrdiff_t>(dataOffset), encoded.begin(), encoded.end());
        const auto newCount = static_cast<std::uint32_t>(encoded.size()) | (compressed ? 0x80000000u : 0u);
        userFormPutU32(bytes, *layout.countOffset, newCount);
        const auto delta = static_cast<std::ptrdiff_t>(encoded.size()) - static_cast<std::ptrdiff_t>(oldBytes);
        const auto oldCbSite = static_cast<std::ptrdiff_t>(userFormU16(bytes, site.begin + 2));
        const auto newCbSite = oldCbSite + delta;
        if (newCbSite < 4 || newCbSite > 0xFFFF) throw std::overflow_error("UserForm control cbSite exceeds 16-bit storage after string edit");
        userFormPutU16(bytes, site.begin + 2, static_cast<std::uint16_t>(newCbSite));
        const auto oldCountBytes = static_cast<std::int64_t>(userFormU32(bytes, parsed.countOfBytesOffset));
        const auto newCountBytes = oldCountBytes + delta;
        if (newCountBytes < 0 || newCountBytes > 0xFFFFFFFFll) throw std::overflow_error("UserForm FormSiteData CountOfBytes overflow");
        userFormPutU32(bytes, parsed.countOfBytesOffset, static_cast<std::uint32_t>(newCountBytes));
    };

    patchString(patch.name, &UserFormControlSiteLayout::name, "Name");
    patchString(patch.tag, &UserFormControlSiteLayout::tag, "Tag");
    patchString(patch.controlTipText, &UserFormControlSiteLayout::controlTipText, "ControlTipText");
    patchString(patch.controlSource, &UserFormControlSiteLayout::controlSource, "ControlSource");
    patchString(patch.rowSource, &UserFormControlSiteLayout::rowSource, "RowSource");

    auto parsed = parseSites();
    auto& site = parsed.sites[controlIndex];
    requireUserFormSiteField(patch.helpContextId.has_value(), site.helpContextId, "HelpContextID");
    requireUserFormSiteField(patch.bitFlags.has_value(), site.bitFlags, "BitFlags");
    requireUserFormSiteField(patch.tabIndex.has_value(), site.tabIndex, "TabIndex");
    requireUserFormSiteField(patch.groupId.has_value(), site.groupId, "GroupID");
    requireUserFormSiteField(patch.top.has_value() || patch.left.has_value(), site.position, "Position");
    if (patch.helpContextId) userFormPutI32(bytes, *site.helpContextId, *patch.helpContextId);
    if (patch.bitFlags) userFormPutU32(bytes, *site.bitFlags, *patch.bitFlags);
    if (patch.tabIndex) userFormPutI16(bytes, *site.tabIndex, *patch.tabIndex);
    if (patch.groupId) userFormPutU16(bytes, *site.groupId, *patch.groupId);
    if (patch.top) userFormPutI32(bytes, *site.position, *patch.top);
    if (patch.left) userFormPutI32(bytes, *site.position + 4, *patch.left);

    setVbaProjectInfo(std::move(info));
    return true;
}

VbaDesignerValidationReport Workbook::validateVbaDesignerProject() const {
    VbaDesignerValidationReport report;
    const auto modules = vbaModules();
    const auto storages = vbaDesignerStorages();
    std::set<std::string> moduleNames;
    std::set<std::string> storageNames;
    for (const auto& module : modules) {
        if (module.type != VbaModuleType::Designer) continue;
        ++report.designerModules;
        if (!moduleNames.insert(module.name).second) {
            report.issues.push_back({module.name, "Duplicate VBA Designer module name"});
            continue;
        }
        const auto storage = std::find_if(storages.begin(), storages.end(), [&](const auto& item) { return item.name == module.name; });
        if (storage == storages.end()) {
            report.issues.push_back({module.name, "Designer module has no matching root Designer Storage"});
            continue;
        }
        const auto* form = storage->findStream("f");
        const bool looksLikeUserForm = module.designerClassId.find("AC9F2F90-E877-11CE-9F68-00AA00574A4F") != std::string::npos;
        if (!form) {
            if (looksLikeUserForm) report.issues.push_back({module.name, "UserForm Designer Storage has no Form stream named f"});
            continue;
        }
        const auto parsed = parseUserFormStream(form->data);
        if (!parsed.inspection.valid) report.issues.push_back({module.name, "Malformed MS-OFORMS Form stream: " + parsed.inspection.error});
        else ++report.validUserFormStreams;
    }
    for (const auto& storage : storages) {
        ++report.designerStorages;
        if (!storageNames.insert(storage.name).second) {
            report.issues.push_back({storage.name, "Duplicate root Designer Storage name"});
            continue;
        }
        const auto module = std::find_if(modules.begin(), modules.end(), [&](const auto& item) {
            return item.type == VbaModuleType::Designer && item.name == storage.name;
        });
        if (module == modules.end()) report.issues.push_back({storage.name, "Designer Storage is orphaned from VBA Designer modules"});
    }
    return report;
}

void Workbook::addVbaProject(const std::filesystem::path& path) {
    std::ifstream stream(path, std::ios::binary);
    if (!stream) throw std::runtime_error("Cannot open VBA project file: " + path.string());
    std::vector<unsigned char> bytes((std::istreambuf_iterator<char>(stream)), std::istreambuf_iterator<char>());
    if (bytes.empty()) throw std::invalid_argument("VBA project file is empty");
    setVbaProject(std::move(bytes));
}

void Workbook::setVbaProject(std::vector<unsigned char> bytes) {
    if (bytes.empty()) throw std::invalid_argument("VBA project bytes cannot be empty");
    removeVbaProject();
    generatedVbaProject_ = false;
    PreservedPart part;
    part.name = "xl/vbaProject.bin";
    part.data.assign(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    part.overrideType = "application/vnd.ms-office.vbaProject";
    part.extension = "bin";
    part.defaultType = "application/vnd.ms-office.vbaProject";
    part.compress = false;
    preservedParts_.push_back(std::move(part));
}

bool Workbook::hasVbaProject() const noexcept {
    return std::any_of(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
}

bool Workbook::removeVbaProject() noexcept {
    const auto oldSize = preservedParts_.size();
    preservedParts_.erase(std::remove_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin"
            || part.name == "xl/vbaProjectSignature.bin"
            || part.name == "xl/_rels/vbaProject.bin.rels";
    }), preservedParts_.end());
    const bool removed = preservedParts_.size() != oldSize;
    if (removed) generatedVbaProject_ = false;
    return removed;
}

std::vector<VbaModule> Workbook::vbaModules() const {
    const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
    if (it == preservedParts_.end()) return {};
    const std::vector<unsigned char> bytes(it->data.begin(), it->data.end());
    return internal::readVbaProjectBinary(bytes);
}

std::optional<std::string> Workbook::vbaModuleText(const std::string& moduleName) const {
    const auto target = [&] {
        std::string value = moduleName;
        std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
        return value;
    }();
    for (const auto& module : vbaModules()) {
        std::string name = module.name;
        std::transform(name.begin(), name.end(), name.begin(), [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
        if (name == target) return module.source;
    }
    return std::nullopt;
}

void Workbook::setVbaModule(VbaModule module) {
    if (module.type == VbaModuleType::Document) {
        const auto lower = [](std::string value) {
            std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) { return static_cast<char>(std::tolower(ch)); });
            return value;
        };
        const auto name = lower(module.name);
        bool validDocument = name == "thisworkbook";
        const auto codeNames = ensureWorksheetVbaCodeNames(sheets_);
        if (!validDocument) validDocument = std::any_of(codeNames.begin(), codeNames.end(), [&](const auto& codeName) {
            return lower(codeName) == name;
        });
        if (!validDocument) throw std::invalid_argument("Unknown VBA document module: " + module.name);
    } else {
        internal::validateVbaModuleName(module.name);
    }
    module.source = internal::normalizeVbaSource(std::move(module.source));

    std::vector<VbaModule> modules;
    VbaProjectInfo info;
    if (hasVbaProject()) {
        try {
            modules = vbaModules();
            info = vbaProjectInfo();
        } catch (const std::exception&) {
            throw std::runtime_error("Cannot edit VBA source in an unsupported existing vbaProject.bin; attach a generated project or replace it first");
        }
    }
    const auto sameName = [&](const VbaModule& existing) {
        if (existing.name.size() != module.name.size()) return false;
        return std::equal(existing.name.begin(), existing.name.end(), module.name.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
    };
    const auto it = std::find_if(modules.begin(), modules.end(), sameName);
    if (it == modules.end()) modules.push_back(std::move(module));
    else {
        if (it->type == VbaModuleType::Document && module.type != VbaModuleType::Document)
            throw std::invalid_argument("Cannot replace a VBA document module with a standard/class module: " + module.name);
        if (it->type != VbaModuleType::Document && module.type == VbaModuleType::Document)
            throw std::invalid_argument("Cannot replace a VBA standard/class module with a document module: " + module.name);
        *it = std::move(module);
    }
    setVbaProject(internal::buildVbaProjectBinary(modules, ensureWorksheetVbaCodeNames(sheets_), info));
    generatedVbaProject_ = true;
}

void Workbook::setVbaModuleText(std::string moduleName, std::string source) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Standard});
}

void Workbook::setVbaClassModuleText(std::string moduleName, std::string source) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Class});
}

void Workbook::addUserForm(const VbaUserFormDesign& design) {
    auto storage = internal::buildUserFormDesign(design);
    std::string source = design.vbaSource;
    source += "\nPrivate Sub UserForm_Initialize()\n";
    source += "Me.Caption = \"" + design.caption + "\"\n";
    source += "End Sub\n";
    setVbaDesignerModule(design.name, source, std::move(storage));
}

void Workbook::setVbaDesignerModule(std::string moduleName, std::string source, VbaDesignerStorage storage,
                                    std::string designerClassId) {
    if (storage.name.empty()) storage.name = moduleName;
    if (storage.name != moduleName)
        throw std::invalid_argument("VBA designer storage name must match the designer module name");
    VbaModule module;
    module.name = std::move(moduleName);
    module.source = std::move(source);
    module.type = VbaModuleType::Designer;
    module.designerClassId = std::move(designerClassId);
    setVbaModule(std::move(module));
    setVbaDesignerStorage(std::move(storage));
}

std::vector<VbaDesignerStorage> Workbook::vbaDesignerStorages() const {
    if (!hasVbaProject()) return {};
    return vbaProjectInfo().designerStorages;
}

void Workbook::setVbaDesignerStorage(VbaDesignerStorage storage) {
    if (storage.name.empty()) throw std::invalid_argument("VBA designer storage name cannot be empty");
    if (!hasVbaProject()) throw std::logic_error("A VBA project is required before adding designer storage");
    auto info = vbaProjectInfo();
    const auto same = [&](const VbaDesignerStorage& existing) {
        if (existing.name.size() != storage.name.size()) return false;
        return std::equal(existing.name.begin(), existing.name.end(), storage.name.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
    };
    const auto it = std::find_if(info.designerStorages.begin(), info.designerStorages.end(), same);
    if (it == info.designerStorages.end()) info.designerStorages.push_back(std::move(storage));
    else *it = std::move(storage);
    setVbaProjectInfo(std::move(info));
}

bool Workbook::removeVbaDesignerStorage(const std::string& storageName) {
    if (!hasVbaProject()) return false;
    auto info = vbaProjectInfo();
    const auto oldSize = info.designerStorages.size();
    info.designerStorages.erase(std::remove_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
        if (storage.name.size() != storageName.size()) return false;
        return std::equal(storage.name.begin(), storage.name.end(), storageName.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
    }), info.designerStorages.end());
    if (info.designerStorages.size() == oldSize) return false;
    setVbaProjectInfo(std::move(info));
    return true;
}

void Workbook::setVbaDocumentModuleText(std::string moduleName, std::string source) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Document});
}

VbaProjectInfo Workbook::vbaProjectInfo() const {
    const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
    if (it == preservedParts_.end()) return {};
    const std::vector<unsigned char> bytes(it->data.begin(), it->data.end());
    return internal::readVbaProjectInfoBinary(bytes);
}

void Workbook::setVbaProjectInfo(VbaProjectInfo info) {
    if (info.name.empty()) throw std::invalid_argument("VBA project name cannot be empty");
    if (info.name.size() > 31) throw std::invalid_argument("VBA project name must contain at most 31 characters");
    auto modules = hasVbaProject() ? vbaModules() : std::vector<VbaModule>{};
    setVbaProject(internal::buildVbaProjectBinary(modules, ensureWorksheetVbaCodeNames(sheets_), info));
    generatedVbaProject_ = true;
}

bool Workbook::removeVbaModule(const std::string& moduleName) {
    if (!hasVbaProject()) return false;
    auto modules = vbaModules();
    bool removedDesigner = false;
    const auto oldSize = modules.size();
    modules.erase(std::remove_if(modules.begin(), modules.end(), [&](const VbaModule& module) {
        if (module.type == VbaModuleType::Document || module.name.size() != moduleName.size()) return false;
        const bool same = std::equal(module.name.begin(), module.name.end(), moduleName.begin(), [](unsigned char a, unsigned char b) {
            return std::tolower(a) == std::tolower(b);
        });
        if (same && module.type == VbaModuleType::Designer) removedDesigner = true;
        return same;
    }), modules.end());
    if (modules.size() == oldSize) return false;
    const bool hasUserSource = std::any_of(modules.begin(), modules.end(), [](const VbaModule& module) {
        return module.type != VbaModuleType::Document || !module.source.empty();
    });
    if (!hasUserSource) {
        removeVbaProject();
    } else {
        VbaProjectInfo info;
        try { info = vbaProjectInfo(); } catch (...) {}
        if (removedDesigner) {
            info.designerStorages.erase(std::remove_if(info.designerStorages.begin(), info.designerStorages.end(), [&](const auto& storage) {
                if (storage.name.size() != moduleName.size()) return false;
                return std::equal(storage.name.begin(), storage.name.end(), moduleName.begin(), [](unsigned char a, unsigned char b) {
                    return std::tolower(a) == std::tolower(b);
                });
            }), info.designerStorages.end());
        }
        setVbaProject(internal::buildVbaProjectBinary(modules, ensureWorksheetVbaCodeNames(sheets_), info));
        generatedVbaProject_ = true;
    }
    return true;
}

} // namespace xlpp
