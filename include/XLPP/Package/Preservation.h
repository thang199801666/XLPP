#pragma once
#include <string>

namespace xlpp {

// OPC relationship retained from an existing package so unsupported or
// selectively edited package topology can round-trip without being severed.
struct PreservedRelationship {
    std::string sourcePart;   // empty means package root (_rels/.rels)
    std::string id;
    std::string type;
    std::string target;
    std::string targetMode;   // empty/Internal or External
};

// Package part retained verbatim when XLPP does not semantically own the whole
// part. Preservation metadata is package-level state and intentionally lives
// outside the Workbook model header to keep OPC/package code model-neutral.
struct PreservedPart {
    std::string name;          // package path, e.g. "customXml/item1.xml"
    std::string data;          // raw part bytes
    std::string overrideType;  // content type emitted as <Override>
    std::string extension;     // file extension (empty if none)
    std::string defaultType;   // content type of the <Default> rule
    bool compress{true};
};

} // namespace xlpp
