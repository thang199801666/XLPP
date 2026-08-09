#include <XLPP/Worksheet/Worksheet.h>
#include <algorithm>
#include <cmath>
#include <filesystem>
#include <utility>

namespace xlpp {
const Image* Worksheet::imageByStableId(const std::string& stableId) const noexcept {
    const auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    return it == images_.end() ? nullptr : &*it;
}

Worksheet::ImportedImageEdit& Worksheet::ensureImportedImageEdit(const Image& image) {
    const auto existing = std::find_if(importedImageEdits_.begin(), importedImageEdits_.end(), [&](const ImportedImageEdit& edit) {
        return edit.stableId == image.stableId_;
    });
    if (existing != importedImageEdits_.end()) return *existing;
    ImportedImageEdit edit;
    edit.stableId = image.stableId_;
    edit.sourceDrawingPart = image.sourceDrawingPart_;
    edit.sourceMediaPart = image.sourceMediaPart_;
    edit.sourceRelationshipId = image.sourceRelationshipId_;
    edit.originalAnchor = image.anchorInfo_;
    edit.anchor = image.anchorInfo_;
    importedImageEdits_.push_back(std::move(edit));
    return importedImageEdits_.back();
}

bool Worksheet::moveImage(const std::string& stableId, const std::string& anchor) {
    auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end() || it->anchorInfo_.type == DrawingAnchorType::Absolute) return false;
    const auto ref = CellReference::parse(anchor);
    const auto oldRow = static_cast<long long>(it->anchorInfo_.from.row);
    const auto oldColumn = static_cast<long long>(it->anchorInfo_.from.column);
    const auto rowDelta = static_cast<long long>(ref.row) - oldRow;
    const auto columnDelta = static_cast<long long>(ref.column) - oldColumn;
    auto updated = it->anchorInfo_;
    updated.from.row = ref.row;
    updated.from.column = ref.column;
    if (updated.type == DrawingAnchorType::TwoCell) {
        const auto newToRow = static_cast<long long>(updated.to.row) + rowDelta;
        const auto newToColumn = static_cast<long long>(updated.to.column) + columnDelta;
        if (newToRow < 1 || newToColumn < 1) return false;
        updated.to.row = static_cast<std::size_t>(newToRow);
        updated.to.column = static_cast<std::size_t>(newToColumn);
    }
    auto& edit = ensureImportedImageEdit(*it);
    edit.anchor = updated;
    edit.moved = true;
    it->anchor_ = CellReference{ref.row, ref.column}.address();
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::moveImageAbsolute(const std::string& stableId, long long xEmu, long long yEmu) {
    auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end() || it->anchorInfo_.type != DrawingAnchorType::Absolute || xEmu < 0 || yEmu < 0) return false;
    auto updated = it->anchorInfo_;
    updated.xEmu = xEmu;
    updated.yEmu = yEmu;
    auto& edit = ensureImportedImageEdit(*it);
    edit.anchor = updated;
    edit.moved = true;
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::resizeImage(const std::string& stableId, double widthPixels, double heightPixels) {
    if (!(widthPixels > 0.0) || !(heightPixels > 0.0)) return false;
    auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end()) return false;
    auto updated = it->anchorInfo_;
    updated.widthEmu = static_cast<long long>(std::llround(widthPixels * 9525.0));
    updated.heightEmu = static_cast<long long>(std::llround(heightPixels * 9525.0));
    if (updated.type == DrawingAnchorType::TwoCell) {
        const auto resizeTerminalMarker = [](std::size_t fromIndex, long long fromOffset,
                                             std::size_t oldToIndex, long long oldToOffset,
                                             long long oldExtent, long long newExtent,
                                             std::size_t& newToIndex, long long& newToOffset) {
            if (oldExtent <= 0 || newExtent <= 0 || oldToIndex < fromIndex) return false;
            const auto span = oldToIndex - fromIndex;
            if (span == 0) {
                // The original image already ends in the starting cell. A
                // smaller/equal resize can safely retain that cell without
                // needing workbook font/column metrics.
                if (newExtent > oldExtent) return false;
                newToIndex = fromIndex;
                newToOffset = fromOffset + newExtent;
                return true;
            }
            const long double averageCellExtent =
                static_cast<long double>(oldExtent + fromOffset - oldToOffset) / static_cast<long double>(span);
            if (!(averageCellExtent > 0.0L)) return false;
            const long double terminal = static_cast<long double>(fromOffset + newExtent);
            auto cells = static_cast<long long>(std::floor(terminal / averageCellExtent));
            if (cells < 0) cells = 0;
            auto offset = static_cast<long long>(std::llround(terminal - static_cast<long double>(cells) * averageCellExtent));
            const auto roundedCell = static_cast<long long>(std::llround(averageCellExtent));
            if (roundedCell > 0 && offset >= roundedCell) { ++cells; offset = 0; }
            newToIndex = fromIndex + static_cast<std::size_t>(cells);
            newToOffset = std::max<long long>(0, offset);
            return true;
        };
        std::size_t toColumn = updated.to.column;
        std::size_t toRow = updated.to.row;
        long long toColumnOffset = updated.to.columnOffsetEmu;
        long long toRowOffset = updated.to.rowOffsetEmu;
        const bool columnOk = resizeTerminalMarker(updated.from.column, updated.from.columnOffsetEmu,
                                                   updated.to.column, updated.to.columnOffsetEmu,
                                                   it->anchorInfo_.widthEmu, updated.widthEmu,
                                                   toColumn, toColumnOffset);
        const bool rowOk = resizeTerminalMarker(updated.from.row, updated.from.rowOffsetEmu,
                                                updated.to.row, updated.to.rowOffsetEmu,
                                                it->anchorInfo_.heightEmu, updated.heightEmu,
                                                toRow, toRowOffset);
        if (!columnOk || !rowOk) return false;
        updated.to.column = toColumn;
        updated.to.columnOffsetEmu = toColumnOffset;
        updated.to.row = toRow;
        updated.to.rowOffsetEmu = toRowOffset;
    }
    auto& edit = ensureImportedImageEdit(*it);
    edit.anchor = updated;
    edit.resized = true;
    it->widthPixels_ = widthPixels;
    it->heightPixels_ = heightPixels;
    it->anchorInfo_ = updated;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::replaceImage(const std::string& stableId, Image replacement) {
    if (replacement.extension_ != "png" && replacement.extension_ != "jpg" && replacement.extension_ != "jpeg") return false;
    auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end() || replacement.bytes_.empty()) return false;
    if (replacement.extension_ == "jpeg") replacement.extension_ = "jpg";
    auto& edit = ensureImportedImageEdit(*it);
    edit.replaced = true;
    edit.replacementBytes = replacement.bytes_;
    edit.replacementExtension = replacement.extension_;
    it->bytes_ = replacement.bytes_;
    it->extension_ = replacement.extension_;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

bool Worksheet::replaceImage(const std::string& stableId, const std::filesystem::path& path) {
    const auto* current = imageByStableId(stableId);
    if (!current) return false;
    return replaceImage(stableId, Image::fromFile(path, current->anchor()));
}

bool Worksheet::removeImage(const std::string& stableId) {
    const auto it = std::find_if(images_.begin(), images_.end(), [&](const Image& image) {
        return image.imported_ && image.stableId_ == stableId;
    });
    if (it == images_.end()) return false;
    auto& edit = ensureImportedImageEdit(*it);
    edit.removed = true;
    const auto index = static_cast<std::size_t>(std::distance(images_.begin(), it));
    images_.erase(it);
    if (index < loadedImageCount_ && loadedImageCount_ > 0) --loadedImageCount_;
    dirty_ = true;
    drawingAppendDirty_ = true;
    return true;
}

} // namespace xlpp
