#include <XLPP/XLPP.h>
#include "Package/Zip/ZipArchive.h"
#include "Package/Zip/ZipArchiveReader.h"
#include "Package/Opc/RelationshipGraph.h"
#include "Platform/MappedFile.h"
#include "Streaming/SharedStringsReader.h"
#include "Core/Threading/ThreadPool.h"
#include "Package/Xml/SimdScan.h"
#include "Package/Xml/XmlScanner.h"
#include "Package/Xml/XmlUtilities.h"
#include "VBA/VbaProjectBinary.h"
#include "Encryption/OfficeEncryption.h"
#include <array>
#include <chrono>
#include <cmath>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iostream>
#include <set>
#include <sstream>
#include <stdexcept>
#include <string>
#include <variant>
#include <vector>

#include "../TestFramework.h"
#include "RegressionTests.h"

namespace {
std::vector<unsigned char> onePixelPngBytes() {
    return {
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x04, 0x00, 0x00, 0x00, 0xb5, 0x1c, 0x0c, 0x02, 0x00, 0x00, 0x00, 0x0b,
        0x49, 0x44, 0x41, 0x54, 0x78, 0xda, 0x63, 0x64, 0xf8, 0x0f, 0x00, 0x01, 0x05,
        0x01, 0x01, 0x27, 0x18, 0xe3, 0x66, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
        0x44, 0xae, 0x42, 0x60, 0x82
    };
}

} // namespace

void testImagePackageRegression(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_image_package.xlsx";
    const std::vector<unsigned char> png{137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,137,0,0,0,13,73,68,65,84,8,215,99,248,207,192,240,31,0,5,0,1,255,137,153,61,29,0,0,0,0,73,69,78,68,174,66,96,130};
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Image");
    auto image = xlpp::Image("D2", png, "png");
    image.setWidthPixels(64);
    image.setHeightPixels(64);
    sheet.addImage(std::move(image));
    workbook.save(path);
    const auto zip = xlpp::internal::ZipArchive::open(path);
    test.checkTrue(zip.contains("xl/media/image1.png"), "Image binary is packaged");
    test.checkEqual(zip.get("xl/media/image1.png").size(), png.size(), "Image bytes are preserved");
    test.checkTrue(zip.get("xl/worksheets/sheet1.xml").find("<drawing r:id=\"rIdDrawing\"/>") != std::string::npos,
                   "Worksheet references drawing part");
    const auto drawing = zip.get("xl/drawings/drawing1.xml");
    test.checkTrue(drawing.find("<xdr:col>3</xdr:col>") != std::string::npos, "Image D2 anchor column is serialized");
    test.checkTrue(drawing.find("<xdr:row>1</xdr:row>") != std::string::npos, "Image D2 anchor row is serialized");
    test.checkTrue(zip.get("xl/drawings/_rels/drawing1.xml.rels").find("../media/image1.png") != std::string::npos,
                   "Drawing relationship targets packaged image");
    std::filesystem::remove(path);
}

void testDrawingImageReaderMetadata(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};

    for (const auto& [producer, path] : fixtures) {
        xlpp::Workbook workbook;
        workbook.load(path);
        const auto& constWorkbook = workbook;
        const auto* sheet = constWorkbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(producer) + " drawing-reader worksheet loads");
        test.checkEqual(sheet->images().size(), std::size_t{1}, std::string(producer) + " image is exposed through the reader model");
        const auto& image = sheet->images().front();
        test.checkTrue(image.imported(), std::string(producer) + " image is marked as package-imported");
        test.checkEqual(image.anchor(), std::string("D2"), std::string(producer) + " image top-left marker is converted to A1 notation");
        test.checkTrue(!image.stableId().empty(), std::string(producer) + " image has a stable drawing-object ID");
        test.checkEqual(image.sourceMediaPart(), std::string("xl/media/image1.png"), std::string(producer) + " media package target is resolved");
        test.checkTrue(!image.sourceRelationshipId().empty(), std::string(producer) + " drawing relationship ID is retained");
        test.checkEqual(image.anchorInfo().from.column, std::size_t{4}, std::string(producer) + " anchor column is parsed as 1-based");
        test.checkEqual(image.anchorInfo().from.row, std::size_t{2}, std::string(producer) + " anchor row is parsed as 1-based");
        test.checkTrue(image.anchorInfo().widthEmu > 0 && image.anchorInfo().heightEmu > 0,
                       std::string(producer) + " image extents are available in EMU");
        if (std::string(producer) == "OpenPyXL") {
            test.checkEqual(static_cast<unsigned>(image.anchorInfo().type), static_cast<unsigned>(xlpp::DrawingAnchorType::OneCell),
                            "OpenPyXL one-cell image anchor is classified correctly");
            test.checkEqual(image.anchorInfo().widthEmu, 1143000LL, "OpenPyXL image width EMU is parsed exactly");
            test.checkEqual(image.anchorInfo().heightEmu, 762000LL, "OpenPyXL image height EMU is parsed exactly");
        } else {
            test.checkEqual(static_cast<unsigned>(image.anchorInfo().type), static_cast<unsigned>(xlpp::DrawingAnchorType::TwoCell),
                            "LibreOffice two-cell image anchor is classified correctly");
            test.checkEqual(image.anchorInfo().to.column, std::size_t{5}, "LibreOffice image end column is parsed");
            test.checkEqual(image.anchorInfo().to.row, std::size_t{5}, "LibreOffice image end row is parsed");
            test.checkEqual(image.anchorInfo().to.columnOffsetEmu, 531000LL, "LibreOffice end-column offset is parsed");
            test.checkEqual(image.anchorInfo().to.rowOffsetEmu, 190080LL, "LibreOffice end-row offset is parsed");
        }
    }
}

void testAbsoluteImageAnchorReader(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto fixturePath = std::filesystem::temp_directory_path() / "xlpp_absolute_anchor_fixture.xlsx";
    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto drawing = package.get("xl/drawings/drawing1.xml");
    const auto anchors = xlpp::internal::tags(drawing, "oneCellAnchor");
    test.checkEqual(anchors.size(), std::size_t{2}, "Absolute-anchor fixture starts from two OpenPyXL one-cell anchors");
    if (anchors.size() >= 2) {
        const auto& imageAnchor = anchors[1];
        const auto pictures = xlpp::internal::tags(imageAnchor, "pic");
        test.checkTrue(!pictures.empty(), "Absolute-anchor fixture locates the image payload");
        if (!pictures.empty()) {
            const std::string absoluteAnchor =
                "<absoluteAnchor><pos x=\"123456\" y=\"654321\"/><ext cx=\"1143000\" cy=\"762000\"/>" +
                pictures.front() + "<clientData/></absoluteAnchor>";
            const auto position = drawing.find(imageAnchor);
            if (position != std::string::npos) drawing.replace(position, imageAnchor.size(), absoluteAnchor);
            package.replace("xl/drawings/drawing1.xml", drawing);
            package.save(fixturePath);
        }
    }

    xlpp::Workbook workbook;
    workbook.load(fixturePath);
    const auto& constWorkbook = workbook;
    const auto* sheet = constWorkbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "Absolute-anchor worksheet loads");
    test.checkEqual(sheet->images().size(), std::size_t{1}, "Absolute-anchor image is discovered");
    const auto& image = sheet->images().front();
    test.checkEqual(static_cast<unsigned>(image.anchorInfo().type), static_cast<unsigned>(xlpp::DrawingAnchorType::Absolute),
                    "Absolute image anchor is classified correctly");
    test.checkEqual(image.anchorInfo().xEmu, 123456LL, "Absolute anchor X position is parsed");
    test.checkEqual(image.anchorInfo().yEmu, 654321LL, "Absolute anchor Y position is parsed");
    test.checkEqual(image.anchorInfo().widthEmu, 1143000LL, "Absolute anchor width is parsed");
    test.checkEqual(image.anchorInfo().heightEmu, 762000LL, "Absolute anchor height is parsed");
    std::filesystem::remove(fixturePath);
}

void testAppendImageToPreservedDrawing(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};

    for (const auto& [producer, sourcePath] : fixtures) {
        const auto outputPath = std::filesystem::temp_directory_path() / (std::string("xlpp_append_image_") + producer + ".xlsx");
        const auto secondPath = std::filesystem::temp_directory_path() / (std::string("xlpp_append_image_second_") + producer + ".xlsx");
        const auto before = xlpp::internal::ZipArchive::open(sourcePath);
        const auto media = before.get("xl/media/image1.png");

        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        auto* sheet = workbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(producer) + " append-image worksheet loads");
        xlpp::Image added("J3", std::vector<unsigned char>(media.begin(), media.end()), "png");
        added.setName("Added by XLPP");
        added.setWidthPixels(80.0);
        added.setHeightPixels(50.0);
        sheet->addImage(std::move(added));
        test.checkEqual(sheet->loadedImageCount(), std::size_t{1}, std::string(producer) + " imported image baseline remains one");
        test.checkEqual(sheet->appendedImageCount(), std::size_t{1}, std::string(producer) + " one image is tracked as an additive drawing mutation");
        workbook.save(outputPath);

        const auto after = xlpp::internal::ZipArchive::open(outputPath);
        const auto validation = xlpp::internal::RelationshipGraph::fromArchive(after).validate();
        const auto inventory = xlpp::internal::RelationshipGraph::fromArchive(after).objectInventory();
        test.checkTrue(validation.ok(), std::string(producer) + " drawing remains graph-valid after appending an image");
        test.checkEqual(inventory.drawings, std::size_t{1}, std::string(producer) + " append reuses the existing drawing part");
        test.checkEqual(inventory.images, std::size_t{2}, std::string(producer) + " original and appended images are both visible");
        test.checkEqual(inventory.charts, std::size_t{1}, std::string(producer) + " existing chart remains visible");
        test.checkTrue(after.contains("xl/media/image2.png"), std::string(producer) + " appended image gets a collision-free media part");
        test.checkEqual(after.get("xl/media/image1.png"), before.get("xl/media/image1.png"), std::string(producer) + " original media bytes remain untouched");
        test.checkEqual(after.get("xl/charts/chart1.xml"), before.get("xl/charts/chart1.xml"), std::string(producer) + " existing chart XML remains byte-identical");
        test.checkTrue(after.get("xl/drawings/drawing1.xml").find("Added by XLPP") != std::string::npos,
                       std::string(producer) + " appended anchor is injected into the preserved drawing XML");

        xlpp::Workbook reopened;
        reopened.load(outputPath);
        const auto& constReopened = reopened;
        const auto* reopenedSheet = constReopened.worksheet("Objects");
        test.checkTrue(reopenedSheet != nullptr, std::string(producer) + " appended workbook reloads");
        test.checkEqual(reopenedSheet->images().size(), std::size_t{2}, std::string(producer) + " reader sees both images after append");
        test.checkTrue(std::any_of(reopenedSheet->images().begin(), reopenedSheet->images().end(), [](const auto& image) {
            return image.anchor() == "J3" && image.name() == "Added by XLPP";
        }), std::string(producer) + " appended image anchor round-trips");

        // A repeated save of the same in-memory workbook must not duplicate or
        // lose the appended object even though its source package remains the
        // original external workbook.
        workbook.save(secondPath);
        const auto second = xlpp::internal::ZipArchive::open(secondPath);
        const auto secondInventory = xlpp::internal::RelationshipGraph::fromArchive(second).objectInventory();
        test.checkEqual(secondInventory.images, std::size_t{2}, std::string(producer) + " repeated save keeps exactly two visible images");
        test.checkEqual(secondInventory.charts, std::size_t{1}, std::string(producer) + " repeated save keeps the original chart");

        std::filesystem::remove(outputPath);
        std::filesystem::remove(secondPath);
    }
}

void testSelectiveImportedImageMoveResize(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};

    for (const auto& [producer, sourcePath] : fixtures) {
        const auto outputPath = std::filesystem::temp_directory_path() /
            (std::string("xlpp_selective_move_resize_") + producer + ".xlsx");
        const auto secondPath = std::filesystem::temp_directory_path() /
            (std::string("xlpp_selective_move_resize_second_") + producer + ".xlsx");
        const auto before = xlpp::internal::ZipArchive::open(sourcePath);

        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        auto* sheet = workbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(producer) + " selective-mutation worksheet loads");
        const auto& constSheet = *sheet;
        test.checkEqual(constSheet.images().size(), std::size_t{1}, std::string(producer) + " selective-mutation fixture has one image");
        const auto stableId = constSheet.images().front().stableId();
        test.checkTrue(sheet->moveImage(stableId, "H6"), std::string(producer) + " imported image moves by stable ID");
        test.checkTrue(sheet->resizeImage(stableId, 64.0, 48.0), std::string(producer) + " imported image resizes by stable ID");
        test.checkTrue(!sheet->moveImage("missing-image", "A1"), std::string(producer) + " unknown stable ID is rejected");
        test.checkTrue(!sheet->resizeImage(stableId, 0.0, 48.0), std::string(producer) + " invalid resize is rejected");
        workbook.save(outputPath);

        const auto after = xlpp::internal::ZipArchive::open(outputPath);
        const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
        test.checkTrue(graph.validate().ok(), std::string(producer) + " move/resize keeps package graph valid");
        test.checkEqual(graph.objectInventory().images, std::size_t{1}, std::string(producer) + " move/resize keeps one visible image");
        test.checkEqual(graph.objectInventory().charts, std::size_t{1}, std::string(producer) + " move/resize preserves sibling chart");
        test.checkEqual(after.get("xl/charts/chart1.xml"), before.get("xl/charts/chart1.xml"),
                        std::string(producer) + " sibling chart remains byte-identical");
        test.checkEqual(after.get("xl/media/image1.png"), before.get("xl/media/image1.png"),
                        std::string(producer) + " move/resize does not rewrite media bytes");

        xlpp::Workbook reopened;
        reopened.load(outputPath);
        const auto& reopenedConst = reopened;
        const auto* reopenedSheet = reopenedConst.worksheet("Objects");
        test.checkTrue(reopenedSheet != nullptr, std::string(producer) + " move/resize result reloads");
        test.checkEqual(reopenedSheet->images().size(), std::size_t{1}, std::string(producer) + " moved image reload count");
        const auto& image = reopenedSheet->images().front();
        test.checkEqual(image.anchor(), std::string("H6"), std::string(producer) + " moved top-left anchor round-trips");
        test.checkEqual(image.anchorInfo().widthEmu, 609600LL, std::string(producer) + " resized width round-trips exactly");
        test.checkEqual(image.anchorInfo().heightEmu, 457200LL, std::string(producer) + " resized height round-trips exactly");
        if (std::string(producer) == "LibreOffice") {
            test.checkEqual(image.anchorInfo().to.column, std::size_t{8},
                            "LibreOffice two-cell resize updates the terminal column");
            test.checkEqual(image.anchorInfo().to.row, std::size_t{8},
                            "LibreOffice two-cell resize updates the terminal row");
            test.checkEqual(image.anchorInfo().to.columnOffsetEmu, 609600LL,
                            "LibreOffice two-cell resize computes the terminal column offset");
            test.checkEqual(image.anchorInfo().to.rowOffsetEmu, 76080LL,
                            "LibreOffice two-cell resize computes the terminal row offset");
        }

        workbook.save(secondPath);
        const auto second = xlpp::internal::ZipArchive::open(secondPath);
        const auto secondGraph = xlpp::internal::RelationshipGraph::fromArchive(second);
        test.checkTrue(secondGraph.validate().ok(), std::string(producer) + " repeated selective save remains graph-valid");
        test.checkEqual(secondGraph.objectInventory().images, std::size_t{1}, std::string(producer) + " repeated selective save does not duplicate image");
        std::filesystem::remove(outputPath);
        std::filesystem::remove(secondPath);
    }
}

void testSelectiveAbsoluteImageMoveResize(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto fixturePath = std::filesystem::temp_directory_path() / "xlpp_absolute_mutation_fixture.xlsx";
    const auto outputPath = std::filesystem::temp_directory_path() / "xlpp_absolute_mutation_result.xlsx";
    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto drawing = package.get("xl/drawings/drawing1.xml");
    const auto anchors = xlpp::internal::tags(drawing, "oneCellAnchor");
    test.checkTrue(anchors.size() >= 2, "Absolute mutation fixture locates image anchor");
    if (anchors.size() >= 2) {
        const auto pictures = xlpp::internal::tags(anchors[1], "pic");
        test.checkTrue(!pictures.empty(), "Absolute mutation fixture locates picture payload");
        if (!pictures.empty()) {
            const std::string absoluteAnchor =
                "<absoluteAnchor><pos x=\"123456\" y=\"654321\"/><ext cx=\"1143000\" cy=\"762000\"/>" +
                pictures.front() + "<clientData/></absoluteAnchor>";
            const auto position = drawing.find(anchors[1]);
            drawing.replace(position, anchors[1].size(), absoluteAnchor);
            package.replace("xl/drawings/drawing1.xml", drawing);
            package.save(fixturePath);
        }
    }

    xlpp::Workbook workbook;
    workbook.load(fixturePath);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "Absolute mutation worksheet loads");
    const auto& constSheet = *sheet;
    const auto stableId = constSheet.images().front().stableId();
    test.checkTrue(!sheet->moveImage(stableId, "A1"), "Cell-based move rejects absolute anchors");
    test.checkTrue(sheet->moveImageAbsolute(stableId, 222222, 333333), "Absolute image position can be changed in EMU");
    test.checkTrue(sheet->resizeImage(stableId, 50.0, 30.0), "Absolute image can be resized selectively");
    workbook.save(outputPath);

    xlpp::Workbook reopened;
    reopened.load(outputPath);
    const auto& reopenedConst = reopened;
    const auto* reopenedSheet = reopenedConst.worksheet("Objects");
    const auto& image = reopenedSheet->images().front();
    test.checkEqual(image.anchorInfo().xEmu, 222222LL, "Absolute image X mutation round-trips");
    test.checkEqual(image.anchorInfo().yEmu, 333333LL, "Absolute image Y mutation round-trips");
    test.checkEqual(image.anchorInfo().widthEmu, 476250LL, "Absolute image width mutation round-trips");
    test.checkEqual(image.anchorInfo().heightEmu, 285750LL, "Absolute image height mutation round-trips");
    const auto after = xlpp::internal::ZipArchive::open(outputPath);
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "Absolute selective mutation keeps package graph valid");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Absolute selective mutation preserves sibling chart");
    std::filesystem::remove(fixturePath);
    std::filesystem::remove(outputPath);
}

void testSelectiveImportedImageRemove(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};
    for (const auto& [producer, sourcePath] : fixtures) {
        const auto outputPath = std::filesystem::temp_directory_path() /
            (std::string("xlpp_selective_remove_") + producer + ".xlsx");
        const auto before = xlpp::internal::ZipArchive::open(sourcePath);
        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        auto* sheet = workbook.worksheet("Objects");
        const auto stableId = static_cast<const xlpp::Worksheet&>(*sheet).images().front().stableId();
        test.checkTrue(sheet->removeImage(stableId), std::string(producer) + " imported image removes by stable ID");
        test.checkTrue(sheet->imageByStableId(stableId) == nullptr, std::string(producer) + " removed image disappears from the in-memory model");
        test.checkTrue(!sheet->removeImage(stableId), std::string(producer) + " removing the same imported image twice is rejected");
        workbook.save(outputPath);
        const auto after = xlpp::internal::ZipArchive::open(outputPath);
        const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
        test.checkTrue(graph.validate().ok(), std::string(producer) + " image removal keeps package graph valid");
        test.checkEqual(graph.objectInventory().images, std::size_t{0}, std::string(producer) + " image removal removes the visible image");
        test.checkEqual(graph.objectInventory().charts, std::size_t{1}, std::string(producer) + " image removal preserves sibling chart");
        test.checkTrue(!after.contains("xl/media/image1.png"), std::string(producer) + " unreferenced media part is cleaned up");
        test.checkEqual(after.get("xl/charts/chart1.xml"), before.get("xl/charts/chart1.xml"),
                        std::string(producer) + " sibling chart XML remains byte-identical after image removal");
        xlpp::Workbook reopened;
        reopened.load(outputPath);
        const auto* reopenedSheet = static_cast<const xlpp::Workbook&>(reopened).worksheet("Objects");
        test.checkEqual(reopenedSheet->images().size(), std::size_t{0}, std::string(producer) + " removed image stays removed after reload");
        std::filesystem::remove(outputPath);
    }
}

void testSelectiveImageReplaceWithSharedMedia(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto fixturePath = std::filesystem::temp_directory_path() / "xlpp_shared_media_fixture.xlsx";
    const auto replacePath = std::filesystem::temp_directory_path() / "xlpp_shared_media_replace.xlsx";
    const auto removePath = std::filesystem::temp_directory_path() / "xlpp_shared_media_remove.xlsx";
    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto drawing = package.get("xl/drawings/drawing1.xml");
    const auto anchors = xlpp::internal::tags(drawing, "oneCellAnchor");
    test.checkTrue(anchors.size() >= 2, "Shared-media fixture locates image anchor");
    if (anchors.size() >= 2) {
        auto duplicate = anchors[1];
        auto replaceOnce = [](std::string& value, const std::string& from, const std::string& to) {
            const auto pos = value.find(from);
            if (pos != std::string::npos) value.replace(pos, from.size(), to);
        };
        replaceOnce(duplicate, "id=\"2\"", "id=\"3\"");
        replaceOnce(duplicate, "name=\"Image 2\"", "name=\"Image 3\"");
        replaceOnce(duplicate, "r:embed=\"rId2\"", "r:embed=\"rId3\"");
        replaceOnce(duplicate, "<col>3</col>", "<col>5</col>");
        const auto close = drawing.rfind("</wsDr>");
        drawing.insert(close, duplicate);
        package.replace("xl/drawings/drawing1.xml", drawing);

        auto rels = package.get("xl/drawings/_rels/drawing1.xml.rels");
        const std::string rel =
            "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
            "Target=\"/xl/media/image1.png\" Id=\"rId3\"/>";
        rels.insert(rels.rfind("</Relationships>"), rel);
        package.replace("xl/drawings/_rels/drawing1.xml.rels", rels);
        package.save(fixturePath);
    }

    const auto fixture = xlpp::internal::ZipArchive::open(fixturePath);
    const auto fixtureGraph = xlpp::internal::RelationshipGraph::fromArchive(fixture);
    test.checkTrue(fixtureGraph.validate().ok(), "Shared-media fixture graph is valid");
    test.checkEqual(fixtureGraph.objectInventory().images, std::size_t{2}, "Shared-media fixture exposes two images");

    xlpp::Workbook replaceWorkbook;
    replaceWorkbook.load(fixturePath);
    auto* replaceSheet = replaceWorkbook.worksheet("Objects");
    const auto& constReplaceSheet = *replaceSheet;
    const auto target = std::find_if(constReplaceSheet.images().begin(), constReplaceSheet.images().end(), [](const auto& image) {
        return image.name() == "Image 3";
    });
    test.checkTrue(target != constReplaceSheet.images().end(), "Shared-media replacement target is found by stable object metadata");
    const auto targetId = target->stableId();
    xlpp::Image replacement("A1", onePixelPngBytes(), "png");
    replacement.setName("Replacement");
    test.checkTrue(replaceSheet->replaceImage(targetId, std::move(replacement)), "Shared-media image can be replaced selectively");
    replaceWorkbook.save(replacePath);

    const auto replaced = xlpp::internal::ZipArchive::open(replacePath);
    const auto replacedGraph = xlpp::internal::RelationshipGraph::fromArchive(replaced);
    test.checkTrue(replacedGraph.validate().ok(), "Shared-media replacement keeps graph valid");
    test.checkEqual(replacedGraph.objectInventory().images, std::size_t{2}, "Shared-media replacement preserves both visible images");
    test.checkEqual(replacedGraph.objectInventory().charts, std::size_t{1}, "Shared-media replacement preserves chart");
    test.checkEqual(replaced.get("xl/media/image1.png"), fixture.get("xl/media/image1.png"),
                    "Replacing one shared-media image leaves the original media bytes untouched");
    test.checkTrue(replaced.contains("xl/media/image2.png"), "Replacing one shared-media image allocates a private media part");
    const auto replacementBytes = onePixelPngBytes();
    const std::string replacementText(reinterpret_cast<const char*>(replacementBytes.data()), replacementBytes.size());
    test.checkEqual(replaced.get("xl/media/image2.png"), replacementText, "Replacement media part contains the requested bytes");
    const auto relationships = replacedGraph.relationshipsFrom("xl/drawings/drawing1.xml");
    const auto rId2 = std::find_if(relationships.begin(), relationships.end(), [](const auto& rel) { return rel.id == "rId2"; });
    const auto rId3 = std::find_if(relationships.begin(), relationships.end(), [](const auto& rel) { return rel.id == "rId3"; });
    test.checkTrue(rId2 != relationships.end() && xlpp::internal::RelationshipGraph::resolveTarget(rId2->sourcePart, rId2->target) == "xl/media/image1.png",
                   "First image keeps the shared original media target");
    test.checkTrue(rId3 != relationships.end() && xlpp::internal::RelationshipGraph::resolveTarget(rId3->sourcePart, rId3->target) == "xl/media/image2.png",
                   "Replaced image relationship is retargeted without renumbering its relationship ID");
    test.checkEqual(replaced.get("xl/charts/chart1.xml"), fixture.get("xl/charts/chart1.xml"),
                    "Shared-media replacement preserves sibling chart XML byte-for-byte");

    xlpp::Workbook removeWorkbook;
    removeWorkbook.load(fixturePath);
    auto* removeSheet = removeWorkbook.worksheet("Objects");
    const auto& constRemoveSheet = *removeSheet;
    const auto removeTarget = std::find_if(constRemoveSheet.images().begin(), constRemoveSheet.images().end(), [](const auto& image) {
        return image.name() == "Image 3";
    });
    test.checkTrue(removeTarget != constRemoveSheet.images().end(), "Shared-media remove target is found");
    test.checkTrue(removeSheet->removeImage(removeTarget->stableId()), "One of two shared-media images can be removed selectively");
    removeWorkbook.save(removePath);
    const auto removed = xlpp::internal::ZipArchive::open(removePath);
    const auto removedGraph = xlpp::internal::RelationshipGraph::fromArchive(removed);
    test.checkTrue(removedGraph.validate().ok(), "Shared-media removal keeps graph valid");
    test.checkEqual(removedGraph.objectInventory().images, std::size_t{1}, "Shared-media removal leaves the sibling image visible");
    test.checkTrue(removed.contains("xl/media/image1.png"), "Shared media part is retained while another image still references it");
    test.checkEqual(removed.get("xl/media/image1.png"), fixture.get("xl/media/image1.png"), "Shared media bytes remain unchanged after sibling removal");

    std::filesystem::remove(fixturePath);
    std::filesystem::remove(replacePath);
    std::filesystem::remove(removePath);
}

void testMultiDrawingSelectiveMutation(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    const auto fixturePath = std::filesystem::temp_directory_path() / "xlpp_multi_drawing_fixture.xlsx";
    const auto outputPath = std::filesystem::temp_directory_path() / "xlpp_multi_drawing_mutated.xlsx";

    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto sheetXml = package.get("xl/worksheets/sheet1.xml");
    const std::string secondDrawingOwner =
        "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId2\"/>";
    sheetXml.insert(sheetXml.rfind("</worksheet>"), secondDrawingOwner);
    package.replace("xl/worksheets/sheet1.xml", sheetXml);

    auto sheetRels = package.get("xl/worksheets/_rels/sheet1.xml.rels");
    sheetRels.insert(sheetRels.rfind("</Relationships>"),
        "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
        "Target=\"/xl/drawings/drawing2.xml\" Id=\"rId2\"/>");
    package.replace("xl/worksheets/_rels/sheet1.xml.rels", sheetRels);

    auto contentTypes = package.get("[Content_Types].xml");
    contentTypes.insert(contentTypes.rfind("</Types>"),
        "<Override PartName=\"/xl/drawings/drawing2.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
    package.replace("[Content_Types].xml", contentTypes);

    const std::string drawing2 =
        "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<xdr:oneCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        "<xdr:ext cx=\"1143000\" cy=\"762000\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"21\" name=\"Second Drawing Image\"/>"
        "<xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rIdImage\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
        "<xdr:spPr><a:prstGeom prst=\"rect\"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"
        "<xdr:oneCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        "<xdr:ext cx=\"1500000\" cy=\"500000\"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"22\" name=\"Preserved Text Box\">"
        "<a:hlinkClick r:id=\"rIdShapeLink\"/></xdr:cNvPr><xdr:cNvSpPr txBox=\"1\"/></xdr:nvSpPr><xdr:spPr/>"
        "<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>keep me</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>"
        "</xdr:wsDr>";
    package.add("xl/drawings/drawing2.xml", drawing2);
    package.add("xl/drawings/_rels/drawing2.xml.rels",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"/xl/media/image1.png\"/>"
        "<Relationship Id=\"rIdShapeLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
        "Target=\"https://example.com/preserved-shape\" TargetMode=\"External\"/>"
        "</Relationships>");
    package.save(fixturePath);

    const auto fixture = xlpp::internal::ZipArchive::open(fixturePath);
    const auto fixtureGraph = xlpp::internal::RelationshipGraph::fromArchive(fixture);
    test.checkTrue(fixtureGraph.validate().ok(), "Multi-drawing fixture is graph-valid including the shape hyperlink");
    test.checkEqual(fixtureGraph.objectInventory().drawings, std::size_t{2}, "Validator sees both worksheet drawing parts");
    test.checkEqual(fixtureGraph.objectInventory().images, std::size_t{2}, "Validator sees images in both drawing parts");
    test.checkEqual(fixtureGraph.objectInventory().charts, std::size_t{1}, "Chart in the first drawing remains visible");
    test.checkEqual(fixtureGraph.objectInventory().shapes, std::size_t{1}, "Validator inventories the unsupported shape");
    test.checkEqual(fixtureGraph.objectInventory().textBoxes, std::size_t{1}, "Validator inventories the text box shape");

    xlpp::Workbook workbook;
    workbook.load(fixturePath);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "Multi-drawing worksheet loads");
    const auto& constSheet = *sheet;
    test.checkEqual(constSheet.images().size(), std::size_t{2}, "Reader exposes images from both preserved drawings");
    const auto secondImage = std::find_if(constSheet.images().begin(), constSheet.images().end(), [](const auto& image) {
        return image.sourceDrawingPart() == "xl/drawings/drawing2.xml";
    });
    test.checkTrue(secondImage != constSheet.images().end(), "Second drawing image retains its source drawing part");
    const auto secondStableId = secondImage->stableId();
    test.checkTrue(sheet->moveImage(secondStableId, "K6"), "Selective move can target an image in the second preserved drawing");
    workbook.save(outputPath);

    const auto after = xlpp::internal::ZipArchive::open(outputPath);
    const auto afterGraph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(afterGraph.validate().ok(), "Multi-drawing selective mutation keeps the package graph valid");
    test.checkEqual(afterGraph.objectInventory().drawings, std::size_t{2}, "Both drawing parts survive selective mutation");
    test.checkEqual(afterGraph.objectInventory().images, std::size_t{2}, "Both images survive selective mutation");
    test.checkEqual(afterGraph.objectInventory().charts, std::size_t{1}, "Sibling chart survives second-drawing mutation");
    test.checkEqual(afterGraph.objectInventory().shapes, std::size_t{1}, "Unsupported shape survives selective mutation");
    test.checkEqual(afterGraph.objectInventory().textBoxes, std::size_t{1}, "Unsupported text box survives selective mutation");
    test.checkEqual(after.get("xl/drawings/drawing1.xml"), fixture.get("xl/drawings/drawing1.xml"),
                    "Untouched first drawing stays byte-identical");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), fixture.get("xl/drawings/_rels/drawing1.xml.rels"),
                    "Untouched first drawing relationships stay byte-identical");
    test.checkTrue(after.get("xl/drawings/drawing2.xml").find("<xdr:col>10</xdr:col>") != std::string::npos &&
                   after.get("xl/drawings/drawing2.xml").find("<xdr:row>5</xdr:row>") != std::string::npos,
                   "Second drawing image anchor moves to K6 without rebuilding sibling objects");
    test.checkTrue(after.get("xl/drawings/drawing2.xml").find("Preserved Text Box") != std::string::npos,
                   "Unknown DrawingML text-box XML remains in the selectively patched drawing");
    test.checkTrue(after.get("xl/drawings/_rels/drawing2.xml.rels").find("rIdShapeLink") != std::string::npos,
                   "Unknown drawing hyperlink relationship remains connected");

    xlpp::Workbook reopened;
    reopened.load(outputPath);
    const auto* reopenedSheet = static_cast<const xlpp::Workbook&>(reopened).worksheet("Objects");
    test.checkTrue(reopenedSheet != nullptr, "Mutated multi-drawing workbook reloads");
    test.checkTrue(std::any_of(reopenedSheet->images().begin(), reopenedSheet->images().end(), [](const auto& image) {
        return image.sourceDrawingPart() == "xl/drawings/drawing2.xml" && image.anchor() == "K6";
    }), "Reader observes the moved image in drawing2 after round-trip");

    std::filesystem::remove(fixturePath);
    std::filesystem::remove(outputPath);
}

void testUnknownDrawingRelationshipValidation(TestContext& test) {
    const auto sourcePath = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl" / "image_chart.xlsx";
    auto package = xlpp::internal::ZipArchive::open(sourcePath);
    auto drawing = package.get("xl/drawings/drawing1.xml");
    auto rels = package.get("xl/drawings/_rels/drawing1.xml.rels");

    // Inject an unsupported DrawingML shape carrying an external hyperlink.
    const std::string shape =
        "<oneCellAnchor><from><col>8</col><colOff>0</colOff><row>10</row><rowOff>0</rowOff></from><ext cx=\"1000000\" cy=\"400000\"/>"
        "<sp><nvSpPr><cNvPr id=\"77\" name=\"Unknown shape\"><a:hlinkClick xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdUnknown\"/></cNvPr><cNvSpPr/></nvSpPr>"
        "<spPr/><txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"
        "<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>unknown</a:t></a:r></a:p></txBody></sp><clientData/></oneCellAnchor>";
    drawing.insert(drawing.rfind("</wsDr>"), shape);
    rels.insert(rels.rfind("</Relationships>"),
        "<Relationship Id=\"rIdUnknown\" Type=\"urn:xlpp:test:unsupportedDrawingRelationship\" "
        "Target=\"https://example.com/unknown\" TargetMode=\"External\"/>");
    package.replace("xl/drawings/drawing1.xml", drawing);
    package.replace("xl/drawings/_rels/drawing1.xml.rels", rels);

    const auto validGraph = xlpp::internal::RelationshipGraph::fromArchive(package);
    test.checkTrue(validGraph.validate().ok(), "Referenced unknown DrawingML relationship is preserved as a valid owner edge");
    test.checkEqual(validGraph.objectInventory().shapes, std::size_t{1}, "Unknown shape is included in drawing inventory");
    test.checkEqual(validGraph.objectInventory().textBoxes, std::size_t{1}, "Unknown shape text body is inventoried as a text box");

    auto orphanedRelationship = package;
    auto brokenDrawing = orphanedRelationship.get("xl/drawings/drawing1.xml");
    const auto attribute = std::string(" r:id=\"rIdUnknown\"");
    const auto position = brokenDrawing.find(attribute);
    test.checkTrue(position != std::string::npos, "Negative drawing fixture locates the relationship-bearing attribute");
    if (position != std::string::npos) brokenDrawing.erase(position, attribute.size());
    orphanedRelationship.replace("xl/drawings/drawing1.xml", brokenDrawing);
    const auto orphanReport = xlpp::internal::RelationshipGraph::fromArchive(orphanedRelationship).validate();
    test.checkTrue(std::any_of(orphanReport.ownerReferenceErrors.begin(), orphanReport.ownerReferenceErrors.end(), [](const auto& issue) {
        return issue.find("rIdUnknown") != std::string::npos && issue.find("not referenced") != std::string::npos;
    }), "Validator rejects an unknown drawing relationship that is no longer referenced by DrawingML");

    auto missingRelationship = package;
    auto brokenRels = missingRelationship.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto relationBegin = brokenRels.find("<Relationship Id=\"rIdUnknown\"");
    test.checkTrue(relationBegin != std::string::npos, "Negative drawing fixture locates the unknown relationship node");
    if (relationBegin != std::string::npos) {
        const auto relationEnd = brokenRels.find("/>", relationBegin);
        if (relationEnd != std::string::npos) brokenRels.erase(relationBegin, relationEnd + 2 - relationBegin);
    }
    missingRelationship.replace("xl/drawings/_rels/drawing1.xml.rels", brokenRels);
    const auto missingReport = xlpp::internal::RelationshipGraph::fromArchive(missingRelationship).validate();
    test.checkTrue(std::any_of(missingReport.ownerReferenceErrors.begin(), missingReport.ownerReferenceErrors.end(), [](const auto& issue) {
        return issue.find("rIdUnknown") != std::string::npos && issue.find("missing relationship") != std::string::npos;
    }), "Validator rejects DrawingML that references an unknown missing relationship");
}

void testImageFileApi(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto pngPath = dir / "xlpp_image_api.PNG";
    const auto gifPath = dir / "xlpp_image_api.gif";
    const auto output = dir / "xlpp_image_api.xlsx";
    const std::vector<unsigned char> png{137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,137,0,0,0,13,73,68,65,84,8,215,99,248,207,192,240,31,0,5,0,1,255,137,153,61,29,0,0,0,0,73,69,78,68,174,66,96,130};
    { std::ofstream stream(pngPath, std::ios::binary); stream.write(reinterpret_cast<const char*>(png.data()), static_cast<std::streamsize>(png.size())); }
    { std::ofstream stream(gifPath, std::ios::binary); stream << "GIF89a"; }

    auto image = xlpp::Image::fromFile(pngPath, "C3");
    test.checkEqual(image.anchor(), std::string("C3"), "Image file anchor");
    test.checkEqual(image.extension(), std::string("png"), "Upper-case PNG extension normalizes");
    test.checkEqual(image.name(), std::string("xlpp_image_api"), "Image name comes from file stem");
    image.setAnchor("D4"); image.setName("Front View"); image.setWidthPixels(64); image.setHeightPixels(32);
    test.checkEqual(image.anchor(), std::string("D4"), "Image anchor setter");
    test.checkEqual(image.name(), std::string("Front View"), "Image name setter");
    test.checkNear(image.widthPixels(), 64.0, 1e-12, "Image width setter");
    test.checkNear(image.heightPixels(), 32.0, 1e-12, "Image height setter");

    bool unsupportedThrown = false;
    try { (void)xlpp::Image::fromFile(gifPath, "A1"); } catch (const std::invalid_argument&) { unsupportedThrown = true; }
    test.checkTrue(unsupportedThrown, "Unsupported image extension throws");
    bool missingThrown = false;
    try { (void)xlpp::Image::fromFile(dir / "missing-xlpp-image.png", "A1"); } catch (const std::runtime_error&) { missingThrown = true; }
    test.checkTrue(missingThrown, "Missing image file throws");

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Image");
    sheet.addImage(pngPath, "B2");
    workbook.save(output);
    test.checkTrue(xlpp::internal::ZipArchive::open(output).contains("xl/media/image1.png"), "Worksheet path overload packages image");
    std::filesystem::remove(pngPath); std::filesystem::remove(gifPath); std::filesystem::remove(output);
}
