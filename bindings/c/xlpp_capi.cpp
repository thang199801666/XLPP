// XLPP C API Implementation
#include "xlpp_capi.h"
#include <XLPP/XLPP.h>
#include <XLPP/Cell/RichText.h>
#include <string>
#include <vector>
#include <cstring>
#include <memory>
#include <cstdio>
#include <algorithm>
#include <limits>

// Simple handle wrapping: cast between opaque pointer and C++ pointer
#define WB(h)   reinterpret_cast<xlpp::Workbook*>(h)
#define WS(h)   reinterpret_cast<xlpp::Worksheet*>(h)
#define CELL(h) reinterpret_cast<xlpp::Cell*>(h)
#define FONT(h) reinterpret_cast<xlpp::Font*>(h)
#define FILL(h) reinterpret_cast<xlpp::Fill*>(h)
#define BDR(h)  reinterpret_cast<xlpp::Border*>(h)
#define BS(h)   reinterpret_cast<xlpp::BorderSide*>(h)
#define ALN(h)  reinterpret_cast<xlpp::Alignment*>(h)
#define STY(h)  reinterpret_cast<xlpp::Style*>(h)
#define PROP(h) reinterpret_cast<xlpp::DocumentProperties*>(h)

namespace {
thread_local std::string g_lastError;
void clearError() noexcept { g_lastError.clear(); }
void setError(const char* message) noexcept { g_lastError = message ? message : "XLPP C API error"; }
void setError(const std::exception& error) noexcept { g_lastError = error.what(); }
}

// String copy helper into caller buffer.
static void copyStr(const std::string& s, char* out, int outSize) {
    if (!out || outSize <= 0) return;
    const auto n = (std::min)(s.size(), static_cast<std::size_t>(outSize - 1));
    std::memcpy(out, s.c_str(), n);
    out[n] = '\0';
}

// ============================================================
// Workbook
// ============================================================
extern "C" {

XLPP_API const char* xlpp_version(void) {
    return "1.1.2";
}

XLPP_API xlpp_workbook xlpp_workbook_create(void) {
    return reinterpret_cast<xlpp_workbook>(new xlpp::Workbook());
}

XLPP_API void xlpp_workbook_destroy(xlpp_workbook wb) {
    delete WB(wb);
}

XLPP_API xlpp_worksheet xlpp_workbook_add_sheet(xlpp_workbook wb, const char* name) {
    if (!wb) { setError("Workbook handle is null"); return nullptr; }
    if (!name || !name[0]) { setError("Workbook and sheet name are required"); return nullptr; }
    try {
        return reinterpret_cast<xlpp_worksheet>(&WB(wb)->addWorksheet(name));
    } catch (const std::exception& e) { setError(e); return nullptr; }
}

XLPP_API int xlpp_workbook_sheet_count(xlpp_workbook wb) {
    return static_cast<int>(WB(wb)->sheetCount());
}

XLPP_API xlpp_worksheet xlpp_workbook_get_sheet(xlpp_workbook wb, int index) {
    try {
        return reinterpret_cast<xlpp_worksheet>(&(*WB(wb))[static_cast<std::size_t>(index)]);
    } catch (...) { return nullptr; }
}

XLPP_API xlpp_worksheet xlpp_workbook_sheet_by_name(xlpp_workbook wb, const char* name) {
    auto* ws = WB(wb)->worksheet(name);
    return reinterpret_cast<xlpp_worksheet>(ws);
}

XLPP_API int xlpp_workbook_remove_sheet(xlpp_workbook wb, const char* name) {
    if (!wb) { setError("Workbook handle is null"); return 0; }
    if (!name || !name[0]) { setError("Workbook and sheet name are required"); return 0; }
    try {
        clearError();
        return WB(wb)->removeWorksheet(name) ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_rename_sheet(xlpp_workbook wb, const char* old_name, const char* new_name) {
    if (!wb) { setError("Workbook handle is null"); return 0; }
    if (!old_name || !old_name[0] || !new_name || !new_name[0]) {
        setError("Workbook, old sheet name, and new sheet name are required");
        return 0;
    }
    try {
        clearError();
        return WB(wb)->renameWorksheet(old_name, new_name) ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API xlpp_worksheet xlpp_workbook_copy_sheet(xlpp_workbook wb, xlpp_worksheet src, const char* new_name) {
    try {
        return reinterpret_cast<xlpp_worksheet>(&WB(wb)->copyWorksheet(*WS(src), new_name));
    } catch (...) { return nullptr; }
}

XLPP_API int xlpp_workbook_sheet_index(xlpp_workbook wb, xlpp_worksheet ws) {
    try {
        return static_cast<int>(WB(wb)->index(*WS(ws)));
    } catch (...) { return -1; }
}

XLPP_API const char* xlpp_workbook_sheet_name(xlpp_workbook wb, int index, char* out, int outSize) {
    try {
        copyStr((*WB(wb))[static_cast<std::size_t>(index)].name(), out, outSize);
    } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
    return out;
}

XLPP_API int xlpp_workbook_sheet_names_count(xlpp_workbook wb) {
    return static_cast<int>(WB(wb)->sheetNames().size());
}

#define CSHEET(h) reinterpret_cast<xlpp::Chartsheet*>(h)
XLPP_API int xlpp_workbook_tab_count(xlpp_workbook wb) {
    return wb ? static_cast<int>(WB(wb)->workbookSheetCount()) : 0;
}
XLPP_API const char* xlpp_workbook_tab_name(xlpp_workbook wb, int index, char* out, int outSize) {
    try {
        if (!wb || index < 0) throw std::out_of_range("Workbook tab index is out of range");
        const auto names = WB(wb)->workbookSheetNames();
        copyStr(names.at(static_cast<std::size_t>(index)), out, outSize);
    } catch (const std::exception& e) {
        setError(e);
        if (out && outSize > 0) out[0] = '\0';
    }
    return out;
}
XLPP_API int xlpp_workbook_tab_kind(xlpp_workbook wb, int index) {
    try {
        if (!wb || index < 0) return -1;
        const auto tabs = WB(wb)->workbookSheets();
        return tabs.at(static_cast<std::size_t>(index)).kind == xlpp::WorkbookSheetKind::Chartsheet ? 1 : 0;
    } catch (...) { return -1; }
}
XLPP_API int xlpp_workbook_tab_visibility(xlpp_workbook wb, int index) {
    try {
        if (!wb || index < 0) return -1;
        return static_cast<int>(WB(wb)->workbookSheetVisibility(static_cast<std::size_t>(index)));
    } catch (...) { return -1; }
}
XLPP_API int xlpp_workbook_set_tab_visibility(xlpp_workbook wb, int index, int visibility) {
    try {
        if (!wb || index < 0 || visibility < 0 || visibility > 2) return 0;
        WB(wb)->setWorkbookSheetVisibility(static_cast<std::size_t>(index), static_cast<xlpp::WorkbookSheetVisibility>(visibility));
        clearError(); return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}
XLPP_API int xlpp_workbook_active_tab(xlpp_workbook wb) {
    return wb ? static_cast<int>(WB(wb)->activeWorkbookSheetIndex()) : -1;
}
XLPP_API int xlpp_workbook_set_active_tab(xlpp_workbook wb, int index) {
    try {
        if (!wb || index < 0) return 0;
        WB(wb)->setActiveWorkbookSheetIndex(static_cast<std::size_t>(index));
        clearError(); return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}
XLPP_API int xlpp_workbook_move_tab(xlpp_workbook wb, int from_index, int to_index) {
    try {
        if (!wb || from_index < 0 || to_index < 0) return 0;
        WB(wb)->moveWorkbookSheet(static_cast<std::size_t>(from_index), static_cast<std::size_t>(to_index));
        clearError();
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}
XLPP_API xlpp_chartsheet xlpp_workbook_add_chartsheet(xlpp_workbook wb, const char* name, int chart_type) {
    try {
        if (!wb || !name || !name[0]) throw std::invalid_argument("Workbook and chartsheet name are required");
        clearError();
        return reinterpret_cast<xlpp_chartsheet>(&WB(wb)->addChartsheet(name, xlpp::Chart(static_cast<xlpp::Chart::Type>(chart_type))));
    } catch (const std::exception& e) { setError(e); return nullptr; }
}
XLPP_API int xlpp_workbook_chartsheet_count(xlpp_workbook wb) {
    return wb ? static_cast<int>(WB(wb)->chartsheetCount()) : 0;
}
XLPP_API xlpp_chartsheet xlpp_workbook_chartsheet_at(xlpp_workbook wb, int index) {
    try {
        if (!wb || index < 0) return nullptr;
        return reinterpret_cast<xlpp_chartsheet>(&WB(wb)->chartsheets().at(static_cast<std::size_t>(index)));
    } catch (...) { return nullptr; }
}
XLPP_API xlpp_chartsheet xlpp_workbook_chartsheet_by_name(xlpp_workbook wb, const char* name) {
    if (!wb || !name) return nullptr;
    return reinterpret_cast<xlpp_chartsheet>(WB(wb)->chartsheet(name));
}
XLPP_API int xlpp_workbook_rename_chartsheet(xlpp_workbook wb, const char* old_name, const char* new_name) {
    try {
        if (!wb || !old_name || !new_name) return 0;
        return WB(wb)->renameChartsheet(old_name, new_name) ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}
XLPP_API int xlpp_workbook_remove_chartsheet(xlpp_workbook wb, const char* name) {
    try {
        if (!wb || !name) return 0;
        return WB(wb)->removeChartsheet(name) ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}
XLPP_API const char* xlpp_chartsheet_name(xlpp_chartsheet cs) {
    return cs ? CSHEET(cs)->name().c_str() : "";
}
XLPP_API xlpp_chart xlpp_chartsheet_chart(xlpp_chartsheet cs) {
    try {
        if (!cs) return nullptr;
        return reinterpret_cast<xlpp_chart>(&CSHEET(cs)->chart());
    } catch (const std::exception& e) { setError(e); return nullptr; }
}
XLPP_API int xlpp_chartsheet_set_printer_settings(xlpp_chartsheet cs, const unsigned char* data, uint64_t size) {
    try {
        if (!cs || (size != 0 && !data)) return 0;
        CSHEET(cs)->setPrinterSettingsData(size == 0 ? std::string{} : std::string(reinterpret_cast<const char*>(data), static_cast<std::size_t>(size)));
        clearError();
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}
XLPP_API uint64_t xlpp_chartsheet_printer_settings_size(xlpp_chartsheet cs) {
    if (!cs || !CSHEET(cs)->printerSettingsData()) return 0;
    return static_cast<uint64_t>(CSHEET(cs)->printerSettingsData()->size());
}
XLPP_API uint64_t xlpp_chartsheet_copy_printer_settings(xlpp_chartsheet cs, unsigned char* out, uint64_t capacity) {
    if (!cs || !CSHEET(cs)->printerSettingsData()) return 0;
    const auto& data = *CSHEET(cs)->printerSettingsData();
    const auto needed = static_cast<uint64_t>(data.size());
    if (!out || capacity < needed) return needed;
    std::memcpy(out, data.data(), data.size());
    return needed;
}
XLPP_API void xlpp_chartsheet_clear_printer_settings(xlpp_chartsheet cs) {
    if (cs) CSHEET(cs)->clearPrinterSettings();
}

XLPP_API void xlpp_workbook_set_template(xlpp_workbook wb, int enabled) { if (wb) WB(wb)->setTemplate(enabled != 0); }
XLPP_API int xlpp_workbook_is_template(xlpp_workbook wb) { return wb && WB(wb)->isTemplate() ? 1 : 0; }

XLPP_API int xlpp_workbook_load(xlpp_workbook wb, const char* path) {
    try {
        WB(wb)->load(std::filesystem::path(path));
        return 1;
    } catch (...) { return 0; }
}

XLPP_API int xlpp_workbook_save(xlpp_workbook wb, const char* path) {
    try {
        WB(wb)->save(std::filesystem::path(path));
        return 1;
    } catch (...) { return 0; }
}

XLPP_API int xlpp_workbook_load_password(xlpp_workbook wb, const char* path, const char* password_utf8) {
    if (!wb || !path || !password_utf8) { setError("Workbook, path, and password are required"); return 0; }
    try {
        clearError();
        xlpp::LoadOptions options;
        options.passwordToOpen = password_utf8;
        WB(wb)->load(std::filesystem::path(path), options);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_load_password_ex(xlpp_workbook wb, const char* path, const char* password_utf8,
                                                    uint64_t max_spin_count, uint64_t max_decrypted_package_bytes,
                                                    int allow_standard_encryption, int require_agile_data_integrity,
                                                    uint64_t max_encryption_info_bytes) {
    if (!wb || !path || !password_utf8) { setError("Workbook, path, and password are required"); return 0; }
    if (max_spin_count > 0xffffffffull) { setError("max_spin_count exceeds the ECMA-376 32-bit range"); return 0; }
    if (max_decrypted_package_bytes > static_cast<uint64_t>((std::numeric_limits<std::size_t>::max)()) ||
        max_encryption_info_bytes > static_cast<uint64_t>((std::numeric_limits<std::size_t>::max)())) {
        setError("Encryption size policy exceeds this platform's size_t range"); return 0;
    }
    try {
        clearError();
        xlpp::LoadOptions options;
        options.passwordToOpen = password_utf8;
        options.maxEncryptionSpinCount = static_cast<std::uint32_t>(max_spin_count);
        options.maxDecryptedPackageBytes = static_cast<std::size_t>(max_decrypted_package_bytes);
        options.allowStandardEncryption = allow_standard_encryption != 0;
        options.requireAgileDataIntegrity = require_agile_data_integrity != 0;
        options.maxEncryptionInfoBytes = static_cast<std::size_t>(max_encryption_info_bytes);
        WB(wb)->load(std::filesystem::path(path), options);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_save_password(xlpp_workbook wb, const char* path, const char* password_utf8, uint64_t spin_count) {
    if (!wb || !path || !password_utf8) { setError("Workbook, path, and password are required"); return 0; }
    if (spin_count > 0xffffffffull) { setError("spin_count exceeds the ECMA-376 32-bit range"); return 0; }
    try {
        clearError();
        xlpp::SaveOptions options;
        options.encryption.enabled = true;
        options.encryption.password = password_utf8;
        options.encryption.spinCount = spin_count ? static_cast<std::uint32_t>(spin_count) : 100000u;
        WB(wb)->save(std::filesystem::path(path), options);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_save_password_ex(xlpp_workbook wb, const char* path, const char* password_utf8, int mode, unsigned key_bits, int hash_algorithm, uint64_t spin_count) {
    if (!wb || !path || !password_utf8) { setError("Workbook, path, and password are required"); return 0; }
    if (spin_count > 0xffffffffull) { setError("spin_count exceeds the ECMA-376 32-bit range"); return 0; }
    if (mode < XLPP_ENCRYPTION_AGILE || mode > XLPP_ENCRYPTION_STANDARD) { setError("Invalid encryption mode"); return 0; }
    if (hash_algorithm < XLPP_ENCRYPTION_HASH_SHA1 || hash_algorithm > XLPP_ENCRYPTION_HASH_SHA512) { setError("Invalid encryption hash algorithm"); return 0; }
    try {
        clearError();
        xlpp::SaveOptions options;
        options.encryption.enabled = true;
        options.encryption.password = password_utf8;
        options.encryption.mode = mode == XLPP_ENCRYPTION_STANDARD ? xlpp::PackageEncryptionMode::Standard : xlpp::PackageEncryptionMode::Agile;
        options.encryption.keyBits = key_bits;
        options.encryption.hashAlgorithm = static_cast<xlpp::PackageEncryptionHash>(hash_algorithm);
        options.encryption.spinCount = spin_count ? static_cast<std::uint32_t>(spin_count) : 100000u;
        WB(wb)->save(std::filesystem::path(path), options);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_is_password_encrypted_file(const char* path) {
    if (!path) { setError("Path is required"); return 0; }
    try {
        clearError();
        return xlpp::Workbook::isPasswordEncryptedFile(std::filesystem::path(path)) ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_encryption_profile(const char* path, int* format, unsigned* key_bits, int* hash_algorithm, uint64_t* spin_count, int* has_data_integrity) {
    if (!path) { setError("Path is required"); return 0; }
    try {
        clearError();
        const auto info = xlpp::Workbook::inspectPasswordEncryptionFile(std::filesystem::path(path));
        if (format) *format = static_cast<int>(info.format);
        if (key_bits) *key_bits = info.keyBits;
        if (hash_algorithm) *hash_algorithm = static_cast<int>(info.hashAlgorithm);
        if (spin_count) *spin_count = info.spinCount;
        if (has_data_integrity) *has_data_integrity = info.hasDataIntegrity ? 1 : 0;
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_encryption_key_encryptor_counts(const char* path, uint64_t* total_key_encryptors,
                                                           uint64_t* password_key_encryptors,
                                                           uint64_t* certificate_key_encryptors) {
    if (!path) { setError("Path is required"); return 0; }
    try {
        clearError();
        const auto info = xlpp::Workbook::inspectPasswordEncryptionFile(std::filesystem::path(path));
        if (total_key_encryptors) *total_key_encryptors = static_cast<uint64_t>(info.keyEncryptorCount);
        if (password_key_encryptors) *password_key_encryptors = static_cast<uint64_t>(info.passwordKeyEncryptorCount);
        if (certificate_key_encryptors) *certificate_key_encryptors = static_cast<uint64_t>(info.certificateKeyEncryptors.size());
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API void xlpp_workbook_set_date1904(xlpp_workbook wb, int v) { WB(wb)->setDate1904(v != 0); }
XLPP_API int  xlpp_workbook_date1904(xlpp_workbook wb) { return WB(wb)->date1904() ? 1 : 0; }
XLPP_API void xlpp_workbook_clear(xlpp_workbook wb) { WB(wb)->clear(); }
XLPP_API int  xlpp_workbook_strict_namespaces(xlpp_workbook wb) { return WB(wb)->strictNamespaces() ? 1 : 0; }

XLPP_API xlpp_properties xlpp_workbook_properties(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_properties>(&WB(wb)->properties());
}
XLPP_API xlpp_wbprotection xlpp_workbook_protection(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_wbprotection>(&WB(wb)->protection());
}
XLPP_API xlpp_calcprops xlpp_workbook_calc_properties(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_calcprops>(&WB(wb)->calcProperties());
}
XLPP_API xlpp_customprops xlpp_workbook_custom_properties(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_customprops>(&WB(wb)->customProperties());
}

XLPP_API xlpp_namedstyle xlpp_workbook_add_named_style(xlpp_workbook wb, const char* name, int* ok) {
    try {
        auto& s = WB(wb)->addNamedStyle(xlpp::NamedStyle(name));
        if (ok) *ok = 1;
        return reinterpret_cast<xlpp_namedstyle>(&s);
    } catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API xlpp_namedstyle xlpp_workbook_named_style(xlpp_workbook wb, const char* name) {
    return reinterpret_cast<xlpp_namedstyle>(WB(wb)->namedStyle(name));
}
XLPP_API int xlpp_workbook_named_styles_count(xlpp_workbook wb) {
    return static_cast<int>(WB(wb)->namedStyles().size());
}
XLPP_API xlpp_namedstyle xlpp_workbook_named_style_at(xlpp_workbook wb, int index) {
    try {
        return reinterpret_cast<xlpp_namedstyle>(const_cast<xlpp::NamedStyle*>(&WB(wb)->namedStyles()[static_cast<std::size_t>(index)]));
    } catch (...) { return nullptr; }
}
XLPP_API void xlpp_workbook_apply_named_style(xlpp_workbook wb, xlpp_cell c, const char* name) {
    try { WB(wb)->applyNamedStyle(*CELL(c), name); } catch (...) {}
}

XLPP_API xlpp_definedname xlpp_workbook_add_defined_name(xlpp_workbook wb, const char* name, const char* value, int* ok) {
    try {
        auto& d = WB(wb)->addDefinedName(xlpp::DefinedName(name, value));
        if (ok) *ok = 1;
        return reinterpret_cast<xlpp_definedname>(&d);
    } catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API xlpp_definedname xlpp_workbook_defined_name(xlpp_workbook wb, const char* name) {
    return reinterpret_cast<xlpp_definedname>(WB(wb)->definedName(name));
}
XLPP_API int xlpp_workbook_defined_names_count(xlpp_workbook wb) {
    return static_cast<int>(WB(wb)->definedNames().size());
}
XLPP_API xlpp_definedname xlpp_workbook_defined_name_at(xlpp_workbook wb, int index) {
    try {
        return reinterpret_cast<xlpp_definedname>(&WB(wb)->definedNames()[static_cast<std::size_t>(index)]);
    } catch (...) { return nullptr; }
}

// ============================================================
// Properties
// ============================================================
XLPP_API void xlpp_properties_set_title(xlpp_properties p, const char* v)    { PROP(p)->setTitle(v); }
XLPP_API void xlpp_properties_set_creator(xlpp_properties p, const char* v)   { PROP(p)->setCreator(v); }
XLPP_API void xlpp_properties_set_subject(xlpp_properties p, const char* v)   { PROP(p)->setSubject(v); }
XLPP_API void xlpp_properties_set_description(xlpp_properties p, const char* v) { PROP(p)->setDescription(v); }
XLPP_API void xlpp_properties_set_keywords(xlpp_properties p, const char* v)  { PROP(p)->setKeywords(v); }
XLPP_API void xlpp_properties_set_category(xlpp_properties p, const char* v)  { PROP(p)->setCategory(v); }
XLPP_API void xlpp_properties_set_last_modified_by(xlpp_properties p, const char* v) { PROP(p)->setLastModifiedBy(v); }
XLPP_API const char* xlpp_properties_get_title(xlpp_properties p)             { return PROP(p)->title().c_str(); }
XLPP_API const char* xlpp_properties_get_creator(xlpp_properties p)           { return PROP(p)->creator().c_str(); }
XLPP_API const char* xlpp_properties_get_subject(xlpp_properties p)           { return PROP(p)->subject().c_str(); }
XLPP_API const char* xlpp_properties_get_description(xlpp_properties p)       { return PROP(p)->description().c_str(); }
XLPP_API const char* xlpp_properties_get_keywords(xlpp_properties p)          { return PROP(p)->keywords().c_str(); }
XLPP_API const char* xlpp_properties_get_category(xlpp_properties p)          { return PROP(p)->category().c_str(); }
XLPP_API const char* xlpp_properties_get_last_modified_by(xlpp_properties p)  { return PROP(p)->lastModifiedBy().c_str(); }

// ============================================================
// Workbook protection
// ============================================================
#define WBP(h) reinterpret_cast<xlpp::WorkbookProtection*>(h)
XLPP_API void xlpp_wbprotection_set_lock_structure(xlpp_wbprotection p, int v) { WBP(p)->setLockStructure(v != 0); }
XLPP_API int  xlpp_wbprotection_lock_structure(xlpp_wbprotection p) { return WBP(p)->lockStructure() ? 1 : 0; }
XLPP_API void xlpp_wbprotection_set_lock_windows(xlpp_wbprotection p, int v) { WBP(p)->setLockWindows(v != 0); }
XLPP_API int  xlpp_wbprotection_lock_windows(xlpp_wbprotection p) { return WBP(p)->lockWindows() ? 1 : 0; }
XLPP_API void xlpp_wbprotection_set_lock_revision(xlpp_wbprotection p, int v) { WBP(p)->setLockRevision(v != 0); }
XLPP_API int  xlpp_wbprotection_lock_revision(xlpp_wbprotection p) { return WBP(p)->lockRevision() ? 1 : 0; }
XLPP_API void xlpp_wbprotection_set_password_hash(xlpp_wbprotection p, const char* v) { WBP(p)->setWorkbookPasswordHash(v); }
XLPP_API const char* xlpp_wbprotection_password_hash(xlpp_wbprotection p) { return WBP(p)->workbookPasswordHash().c_str(); }

// ============================================================
// CalcProperties
// ============================================================
#define CP(h) reinterpret_cast<xlpp::CalcProperties*>(h)
XLPP_API void xlpp_calcprops_set_calc_id(xlpp_calcprops p, int v) { CP(p)->setCalcId(v); }
XLPP_API int  xlpp_calcprops_calc_id(xlpp_calcprops p) { return CP(p)->calcId(); }
XLPP_API void xlpp_calcprops_set_calc_mode(xlpp_calcprops p, const char* v) { CP(p)->setCalcMode(v); }
XLPP_API const char* xlpp_calcprops_calc_mode(xlpp_calcprops p) { return CP(p)->calcMode().c_str(); }
XLPP_API void xlpp_calcprops_set_calc_on_save(xlpp_calcprops p, int v) { CP(p)->setCalcOnSave(v != 0); }
XLPP_API int  xlpp_calcprops_calc_on_save(xlpp_calcprops p) { return CP(p)->calcOnSave() ? 1 : 0; }
XLPP_API void xlpp_calcprops_set_full_calc_on_load(xlpp_calcprops p, int v) { CP(p)->setFullCalcOnLoad(v != 0); }
XLPP_API int  xlpp_calcprops_full_calc_on_load(xlpp_calcprops p) { return CP(p)->fullCalcOnLoad() ? 1 : 0; }
XLPP_API void xlpp_calcprops_set_full_precision(xlpp_calcprops p, int v) { CP(p)->setFullPrecision(v != 0); }
XLPP_API int  xlpp_calcprops_full_precision(xlpp_calcprops p) { return CP(p)->fullPrecision() ? 1 : 0; }
XLPP_API void xlpp_calcprops_set_iterate(xlpp_calcprops p, int v) { CP(p)->setIterate(v != 0); }
XLPP_API int  xlpp_calcprops_iterate(xlpp_calcprops p) { return CP(p)->iterate() ? 1 : 0; }
XLPP_API void xlpp_calcprops_set_iterate_count(xlpp_calcprops p, int v) { CP(p)->setIterateCount(v); }
XLPP_API int  xlpp_calcprops_iterate_count(xlpp_calcprops p) { return CP(p)->iterateCount(); }
XLPP_API void xlpp_calcprops_set_iterate_delta(xlpp_calcprops p, double v) { CP(p)->setIterateDelta(v); }
XLPP_API double xlpp_calcprops_iterate_delta(xlpp_calcprops p) { return CP(p)->iterateDelta(); }

// ============================================================
// Custom properties
// ============================================================
#define CUST(h) reinterpret_cast<xlpp::CustomProperties*>(h)
#define CUSTP(h) reinterpret_cast<xlpp::CustomProperty*>(h)
XLPP_API xlpp_customprop xlpp_customprops_add(xlpp_customprops c, const char* name, const char* value, int type) {
    try {
        xlpp::CustomProperty p;
        const std::string vstr = value ? value : "";
        switch (type) {
        case 0: p = xlpp::CustomProperty(name, vstr); break;
        case 1: p = xlpp::CustomProperty(name, vstr.empty() ? 0 : std::atoi(vstr.c_str())); break;
        case 2: p = xlpp::CustomProperty(name, vstr.empty() ? 0.0 : std::atof(vstr.c_str())); break;
        case 3: p = xlpp::CustomProperty(name, vstr == "true"); break;
        default: p = xlpp::CustomProperty(name, vstr); break;
        }
        CUST(c)->add(std::move(p));
        return reinterpret_cast<xlpp_customprop>(&CUST(c)->items().back());
    } catch (...) { return nullptr; }
}
XLPP_API int xlpp_customprops_count(xlpp_customprops c) { return static_cast<int>(CUST(c)->items().size()); }
XLPP_API xlpp_customprop xlpp_customprops_at(xlpp_customprops c, int index) {
    try { return reinterpret_cast<xlpp_customprop>(&CUST(c)->items()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API const char* xlpp_customprop_name(xlpp_customprop p) { return CUSTP(p)->name().c_str(); }
XLPP_API const char* xlpp_customprop_value(xlpp_customprop p) { return CUSTP(p)->value().c_str(); }
XLPP_API const char* xlpp_customprop_type(xlpp_customprop p) { return CUSTP(p)->type().c_str(); }

// ============================================================
// Worksheet
// ============================================================
XLPP_API const char* xlpp_sheet_name(xlpp_worksheet ws) { return WS(ws)->name().c_str(); }
XLPP_API void xlpp_sheet_rename(xlpp_worksheet ws, const char* name) { WS(ws)->rename(name); }

XLPP_API xlpp_cell xlpp_sheet_cell(xlpp_worksheet ws, const char* address) {
    try { return reinterpret_cast<xlpp_cell>(&WS(ws)->cell(address)); } catch (...) { return nullptr; }
}
XLPP_API xlpp_cell xlpp_sheet_cell_rc(xlpp_worksheet ws, uint64_t row, uint64_t col) {
    try { return reinterpret_cast<xlpp_cell>(&WS(ws)->cell(static_cast<std::size_t>(row), static_cast<std::size_t>(col))); } catch (...) { return nullptr; }
}
XLPP_API int xlpp_sheet_has_cell(xlpp_worksheet ws, const char* address) {
    return WS(ws)->tryCell(address) != nullptr ? 1 : 0;
}
XLPP_API int xlpp_sheet_has_cell_rc(xlpp_worksheet ws, uint64_t row, uint64_t col) {
    return WS(ws)->tryCell(static_cast<std::size_t>(row), static_cast<std::size_t>(col)) != nullptr ? 1 : 0;
}

XLPP_API void xlpp_sheet_append_row(xlpp_worksheet ws, const char** values, int count) {
    std::vector<xlpp::CellValue> cv;
    cv.reserve(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) {
        if (values[i] && values[i][0]) cv.push_back(std::string(values[i]));
        else cv.push_back(std::monostate{});
    }
    WS(ws)->append(cv);
}

XLPP_API void xlpp_sheet_append_doubles(xlpp_worksheet ws, const double* values, int count) {
    std::vector<xlpp::CellValue> cv;
    cv.reserve(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) cv.push_back(values[i]);
    WS(ws)->append(cv);
}

XLPP_API void xlpp_sheet_append_values(xlpp_worksheet ws, const double* nums, const int* types, int count) {
    std::vector<xlpp::CellValue> cv;
    cv.reserve(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) {
        switch (types[i]) {
        case XLPP_VALUE_NUMBER: cv.push_back(nums[i]); break;
        case XLPP_VALUE_BOOL: cv.push_back(nums[i] != 0.0); break;
        case XLPP_VALUE_STRING: cv.push_back(std::string(reinterpret_cast<const char*>(static_cast<uintptr_t>(static_cast<uint64_t>(nums[i]))))); break;
        default: cv.push_back(std::monostate{}); break;
        }
    }
    WS(ws)->append(cv);
}

XLPP_API xlpp_cellrange xlpp_sheet_range(xlpp_worksheet ws, const char* address) {
    try {
        auto r = WS(ws)->range(address);
        return reinterpret_cast<xlpp_cellrange>(new xlpp::CellRange(std::move(r)));
    } catch (...) { return nullptr; }
}
XLPP_API xlpp_cellrange xlpp_sheet_range_rc(xlpp_worksheet ws, uint64_t minRow, uint64_t minCol, uint64_t maxRow, uint64_t maxCol) {
    try {
        auto r = WS(ws)->range(static_cast<std::size_t>(minRow), static_cast<std::size_t>(minCol),
                               static_cast<std::size_t>(maxRow), static_cast<std::size_t>(maxCol));
        return reinterpret_cast<xlpp_cellrange>(new xlpp::CellRange(std::move(r)));
    } catch (...) { return nullptr; }
}

XLPP_API void xlpp_sheet_merge_cells(xlpp_worksheet ws, const char* range)     { try { WS(ws)->mergeCells(range); } catch (...) {} }
XLPP_API void xlpp_sheet_unmerge_cells(xlpp_worksheet ws, const char* range)   { try { WS(ws)->unmergeCells(range); } catch (...) {} }
XLPP_API int xlpp_sheet_is_merged(xlpp_worksheet ws, const char* cell)         { try { return WS(ws)->isMerged(cell) ? 1 : 0; } catch (...) { return 0; } }
XLPP_API int xlpp_sheet_merged_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->mergedRanges().size()); }
XLPP_API void xlpp_sheet_merged_at(xlpp_worksheet ws, int index, char* out, int outSize) {
    try { copyStr(WS(ws)->mergedRanges()[static_cast<std::size_t>(index)], out, outSize); } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}

XLPP_API void xlpp_sheet_freeze_panes(xlpp_worksheet ws, const char* cell)     { try { WS(ws)->freezePanes(cell); } catch (...) {} }
XLPP_API void xlpp_sheet_clear_freeze_panes(xlpp_worksheet ws) { WS(ws)->clearFreezePanes(); }
XLPP_API void xlpp_sheet_frozen_pane(xlpp_worksheet ws, char* out, int outSize) {
    const auto& fp = WS(ws)->frozenPane();
    if (fp) copyStr(*fp, out, outSize);
    else if (out && outSize > 0) out[0] = '\0';
}

XLPP_API xlpp_rowdim xlpp_sheet_row_dimension(xlpp_worksheet ws, uint64_t row) {
    return reinterpret_cast<xlpp_rowdim>(&WS(ws)->rowDimension(static_cast<std::size_t>(row)));
}
XLPP_API xlpp_coldim xlpp_sheet_col_dimension(xlpp_worksheet ws, uint64_t col) {
    return reinterpret_cast<xlpp_coldim>(&WS(ws)->columnDimension(static_cast<std::size_t>(col)));
}

XLPP_API uint64_t xlpp_sheet_max_row(xlpp_worksheet ws) { return WS(ws)->maxRow(); }
XLPP_API uint64_t xlpp_sheet_max_col(xlpp_worksheet ws) { return WS(ws)->maxColumn(); }
XLPP_API void xlpp_sheet_dimensions(xlpp_worksheet ws, char* out, int outSize) { copyStr(WS(ws)->dimensions(), out, outSize); }
XLPP_API int xlpp_sheet_empty(xlpp_worksheet ws) { return WS(ws)->empty() ? 1 : 0; }
XLPP_API uint64_t xlpp_sheet_row_count(xlpp_worksheet ws) { return WS(ws)->rowCount(); }
XLPP_API uint64_t xlpp_sheet_col_count(xlpp_worksheet ws) { return WS(ws)->columnCount(); }

XLPP_API void xlpp_sheet_insert_rows(xlpp_worksheet ws, uint64_t index, uint64_t amount) {
    try { WS(ws)->insertRows(static_cast<std::size_t>(index), static_cast<std::size_t>(amount)); } catch (...) {}
}
XLPP_API void xlpp_sheet_delete_rows(xlpp_worksheet ws, uint64_t index, uint64_t amount) {
    try { WS(ws)->deleteRows(static_cast<std::size_t>(index), static_cast<std::size_t>(amount)); } catch (...) {}
}
XLPP_API void xlpp_sheet_insert_cols(xlpp_worksheet ws, uint64_t index, uint64_t amount) {
    try { WS(ws)->insertColumns(static_cast<std::size_t>(index), static_cast<std::size_t>(amount)); } catch (...) {}
}
XLPP_API void xlpp_sheet_delete_cols(xlpp_worksheet ws, uint64_t index, uint64_t amount) {
    try { WS(ws)->deleteColumns(static_cast<std::size_t>(index), static_cast<std::size_t>(amount)); } catch (...) {}
}

XLPP_API void xlpp_sheet_set_print_area(xlpp_worksheet ws, const char* v) { WS(ws)->setPrintArea(v); }
XLPP_API void xlpp_sheet_print_area(xlpp_worksheet ws, char* out, int outSize) { copyStr(WS(ws)->printArea(), out, outSize); }
XLPP_API void xlpp_sheet_set_print_titles_rows(xlpp_worksheet ws, const char* v) { WS(ws)->setPrintTitlesRows(v); }
XLPP_API void xlpp_sheet_print_titles_rows(xlpp_worksheet ws, char* out, int outSize) { copyStr(WS(ws)->printTitlesRows(), out, outSize); }
XLPP_API void xlpp_sheet_set_print_titles_cols(xlpp_worksheet ws, const char* v) { WS(ws)->setPrintTitlesCols(v); }
XLPP_API void xlpp_sheet_print_titles_cols(xlpp_worksheet ws, char* out, int outSize) { copyStr(WS(ws)->printTitlesCols(), out, outSize); }

XLPP_API xlpp_autofilter xlpp_sheet_auto_filter(xlpp_worksheet ws) { return reinterpret_cast<xlpp_autofilter>(&WS(ws)->autoFilter()); }
XLPP_API xlpp_cfcollection xlpp_sheet_conditional_formatting(xlpp_worksheet ws) { return reinterpret_cast<xlpp_cfcollection>(&WS(ws)->conditionalFormatting()); }
XLPP_API xlpp_dvcollection xlpp_sheet_data_validations(xlpp_worksheet ws) { return reinterpret_cast<xlpp_dvcollection>(&WS(ws)->dataValidations()); }
XLPP_API xlpp_pagesetup xlpp_sheet_page_setup(xlpp_worksheet ws) { return reinterpret_cast<xlpp_pagesetup>(&WS(ws)->pageSetup()); }
XLPP_API xlpp_pagemargins xlpp_sheet_page_margins(xlpp_worksheet ws) { return reinterpret_cast<xlpp_pagemargins>(&WS(ws)->pageMargins()); }
XLPP_API xlpp_printopts xlpp_sheet_print_options(xlpp_worksheet ws) { return reinterpret_cast<xlpp_printopts>(&WS(ws)->printOptions()); }
XLPP_API xlpp_headerfooter xlpp_sheet_header_footer(xlpp_worksheet ws) { return reinterpret_cast<xlpp_headerfooter>(&WS(ws)->headerFooter()); }
XLPP_API xlpp_wssprotection xlpp_sheet_protection(xlpp_worksheet ws) { return reinterpret_cast<xlpp_wssprotection>(&WS(ws)->protection()); }
XLPP_API xlpp_sheetview xlpp_sheet_view(xlpp_worksheet ws) { return reinterpret_cast<xlpp_sheetview>(&WS(ws)->sheetView()); }

XLPP_API xlpp_table xlpp_sheet_add_table(xlpp_worksheet ws, const char* name, const char* reference, int* ok) {
    try { auto& t = WS(ws)->addTable(name, reference); if (ok) *ok = 1; return reinterpret_cast<xlpp_table>(&t); }
    catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API xlpp_table xlpp_sheet_table(xlpp_worksheet ws, const char* name) { return reinterpret_cast<xlpp_table>(WS(ws)->table(name)); }
XLPP_API int xlpp_sheet_table_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->tables().size()); }
XLPP_API xlpp_table xlpp_sheet_table_at(xlpp_worksheet ws, int index) {
    try { return reinterpret_cast<xlpp_table>(&WS(ws)->tables()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}

XLPP_API xlpp_image xlpp_sheet_add_image(xlpp_worksheet ws, const char* path, const char* anchor, int* ok) {
    try { auto& img = WS(ws)->addImage(std::filesystem::path(path), anchor); if (ok) *ok = 1; return reinterpret_cast<xlpp_image>(&img); }
    catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API int xlpp_sheet_image_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->images().size()); }
XLPP_API xlpp_image xlpp_sheet_image_at(xlpp_worksheet ws, int index) {
    try { return reinterpret_cast<xlpp_image>(&WS(ws)->images()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}

XLPP_API void xlpp_sheet_add_chart(xlpp_worksheet ws, int type) {
    WS(ws)->addChart(xlpp::Chart(static_cast<xlpp::Chart::Type>(type)));
}
XLPP_API int xlpp_sheet_chart_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->chartCount()); }
XLPP_API xlpp_chart xlpp_sheet_chart_at(xlpp_worksheet ws, int index) {
    try { return reinterpret_cast<xlpp_chart>(&WS(ws)->chart(static_cast<std::size_t>(index))); } catch (...) { return nullptr; }
}

XLPP_API void xlpp_sheet_add_pivot(xlpp_worksheet ws, const char* name, const char* location) {
    xlpp::PivotTable pt(name);
    pt.setLocation(location);
    WS(ws)->addPivotTable(std::move(pt));
}
XLPP_API int xlpp_sheet_pivot_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->pivotTables().size()); }
XLPP_API xlpp_pivottable xlpp_sheet_pivot_at(xlpp_worksheet ws, int index) {
    try { return reinterpret_cast<xlpp_pivottable>(&WS(ws)->pivotTables()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}

// ============================================================
// Row / Column dimensions
// ============================================================
#define RD(h) reinterpret_cast<xlpp::RowDimension*>(h)
#define CD(h) reinterpret_cast<xlpp::ColumnDimension*>(h)
XLPP_API void xlpp_rowdim_set_height(xlpp_rowdim d, double v) { RD(d)->height = v; }
XLPP_API int  xlpp_rowdim_has_height(xlpp_rowdim d) { return RD(d)->height.has_value() ? 1 : 0; }
XLPP_API double xlpp_rowdim_height(xlpp_rowdim d) { return RD(d)->height.value_or(0.0); }
XLPP_API void xlpp_rowdim_set_hidden(xlpp_rowdim d, int v) { RD(d)->hidden = v != 0; }
XLPP_API int  xlpp_rowdim_hidden(xlpp_rowdim d) { return RD(d)->hidden ? 1 : 0; }
XLPP_API void xlpp_rowdim_set_outline_level(xlpp_rowdim d, int v) { RD(d)->outlineLevel = v; }
XLPP_API int  xlpp_rowdim_outline_level(xlpp_rowdim d) { return RD(d)->outlineLevel; }
XLPP_API void xlpp_rowdim_set_collapsed(xlpp_rowdim d, int v) { RD(d)->collapsed = v != 0; }
XLPP_API int  xlpp_rowdim_collapsed(xlpp_rowdim d) { return RD(d)->collapsed ? 1 : 0; }

XLPP_API void xlpp_coldim_set_width(xlpp_coldim d, double v) { CD(d)->width = v; }
XLPP_API int  xlpp_coldim_has_width(xlpp_coldim d) { return CD(d)->width.has_value() ? 1 : 0; }
XLPP_API double xlpp_coldim_width(xlpp_coldim d) { return CD(d)->width.value_or(0.0); }
XLPP_API void xlpp_coldim_set_hidden(xlpp_coldim d, int v) { CD(d)->hidden = v != 0; }
XLPP_API int  xlpp_coldim_hidden(xlpp_coldim d) { return CD(d)->hidden ? 1 : 0; }
XLPP_API void xlpp_coldim_set_best_fit(xlpp_coldim d, int v) { CD(d)->bestFit = v != 0; }
XLPP_API int  xlpp_coldim_best_fit(xlpp_coldim d) { return CD(d)->bestFit ? 1 : 0; }
XLPP_API void xlpp_coldim_set_outline_level(xlpp_coldim d, int v) { CD(d)->outlineLevel = v; }
XLPP_API int  xlpp_coldim_outline_level(xlpp_coldim d) { return CD(d)->outlineLevel; }
XLPP_API void xlpp_coldim_set_collapsed(xlpp_coldim d, int v) { CD(d)->collapsed = v != 0; }
XLPP_API int  xlpp_coldim_collapsed(xlpp_coldim d) { return CD(d)->collapsed ? 1 : 0; }

// ============================================================
// CellRange
// ============================================================
#define RNG(h) reinterpret_cast<xlpp::CellRange*>(h)
XLPP_API uint64_t xlpp_range_min_row(xlpp_cellrange r) { return RNG(r)->minRow(); }
XLPP_API uint64_t xlpp_range_min_col(xlpp_cellrange r) { return RNG(r)->minColumn(); }
XLPP_API uint64_t xlpp_range_max_row(xlpp_cellrange r) { return RNG(r)->maxRow(); }
XLPP_API uint64_t xlpp_range_max_col(xlpp_cellrange r) { return RNG(r)->maxColumn(); }
XLPP_API uint64_t xlpp_range_row_count(xlpp_cellrange r) { return RNG(r)->rowCount(); }
XLPP_API uint64_t xlpp_range_col_count(xlpp_cellrange r) { return RNG(r)->columnCount(); }
XLPP_API void xlpp_range_address(xlpp_cellrange r, char* out, int outSize) { copyStr(RNG(r)->address(), out, outSize); }
XLPP_API xlpp_cell xlpp_range_cell(xlpp_cellrange r, uint64_t relRow, uint64_t relCol) {
    return reinterpret_cast<xlpp_cell>(&RNG(r)->cell(static_cast<std::size_t>(relRow), static_cast<std::size_t>(relCol)));
}
XLPP_API void xlpp_range_set_value(xlpp_cellrange r, double v) { RNG(r)->setValue(v); }
XLPP_API void xlpp_range_set_string(xlpp_cellrange r, const char* v) { RNG(r)->setValue(std::string(v)); }
XLPP_API void xlpp_range_clear(xlpp_cellrange r) { RNG(r)->clear(); }
XLPP_API void xlpp_range_values(xlpp_cellrange r, double* out, int* outCount) {
    auto values = RNG(r)->values();
    if (outCount) *outCount = static_cast<int>(values.size());
    if (out) {
        for (std::size_t i = 0; i < values.size() && out; ++i)
            out[i] = std::holds_alternative<double>(values[i]) ? std::get<double>(values[i]) : 0.0;
    }
}

// ============================================================
// Cell
// ============================================================
XLPP_API const char* xlpp_cell_address(xlpp_cell c)  { return CELL(c)->address().c_str(); }
XLPP_API uint64_t xlpp_cell_row(xlpp_cell c)          { return CELL(c)->row(); }
XLPP_API uint64_t xlpp_cell_column(xlpp_cell c)       { return CELL(c)->column(); }

XLPP_API int xlpp_cell_value_type(xlpp_cell c) {
    const auto& v = CELL(c)->value();
    if (std::holds_alternative<std::monostate>(v)) return XLPP_VALUE_EMPTY;
    if (std::holds_alternative<bool>(v))           return XLPP_VALUE_BOOL;
    if (std::holds_alternative<double>(v))           return XLPP_VALUE_NUMBER;
    if (std::holds_alternative<std::string>(v))      return XLPP_VALUE_STRING;
    if (std::holds_alternative<xlpp::CellError>(v))  return XLPP_VALUE_ERROR;
    if (std::holds_alternative<xlpp::DateTime>(v))   return XLPP_VALUE_DATE;
    return XLPP_VALUE_EMPTY;
}

XLPP_API int xlpp_cell_get_bool(xlpp_cell c) {
    try { return std::get<bool>(CELL(c)->value()) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API double xlpp_cell_get_number(xlpp_cell c) {
    if (auto* v = std::get_if<double>(&CELL(c)->value())) return *v;
    if (auto* v = std::get_if<xlpp::DateTime>(&CELL(c)->value())) return xlpp::toExcelSerial(*v, false);
    return 0.0;
}
XLPP_API const char* xlpp_cell_get_string(xlpp_cell c) {
    if (auto* v = std::get_if<std::string>(&CELL(c)->value())) return v->c_str();
    return "";
}
XLPP_API int xlpp_cell_is_empty(xlpp_cell c) { return CELL(c)->empty() ? 1 : 0; }
XLPP_API int xlpp_cell_has_value(xlpp_cell c) { return CELL(c)->hasValue() ? 1 : 0; }
XLPP_API int xlpp_cell_is_numeric(xlpp_cell c) { return CELL(c)->isNumeric() ? 1 : 0; }
XLPP_API int xlpp_cell_is_string(xlpp_cell c) { return CELL(c)->isString() ? 1 : 0; }
XLPP_API int xlpp_cell_is_bool(xlpp_cell c) { return CELL(c)->isBoolean() ? 1 : 0; }
XLPP_API int xlpp_cell_is_date(xlpp_cell c) { return CELL(c)->isDate() ? 1 : 0; }
XLPP_API int xlpp_cell_is_error(xlpp_cell c) { return CELL(c)->isError() ? 1 : 0; }
XLPP_API int xlpp_cell_error_code(xlpp_cell c) {
    auto e = CELL(c)->error();
    if (!e) return -1;
    switch (*e) {
    case xlpp::CellError::Null: return XLPP_ERROR_NULL;
    case xlpp::CellError::DivisionByZero: return XLPP_ERROR_DIV0;
    case xlpp::CellError::Value: return XLPP_ERROR_VALUE;
    case xlpp::CellError::Reference: return XLPP_ERROR_REF;
    case xlpp::CellError::Name: return XLPP_ERROR_NAME;
    case xlpp::CellError::Number: return XLPP_ERROR_NUM;
    case xlpp::CellError::NotAvailable: return XLPP_ERROR_NA;
    case xlpp::CellError::GettingData: return XLPP_ERROR_GETTING_DATA;
    }
    return XLPP_ERROR_VALUE;
}
XLPP_API int xlpp_cell_date(xlpp_cell c, int* year, int* month, int* day, int* hour, int* minute, double* second) {
    auto d = CELL(c)->date();
    if (!d) return 0;
    if (year) *year = d->year;
    if (month) *month = d->month;
    if (day) *day = d->day;
    if (hour) *hour = d->hour;
    if (minute) *minute = d->minute;
    if (second) *second = d->second;
    return 1;
}

XLPP_API void xlpp_cell_set_string(xlpp_cell c, const char* v)   { CELL(c)->setStringValue(v ? v : ""); }
XLPP_API void xlpp_cell_set_number(xlpp_cell c, double v)        { CELL(c)->setNumericValue(v); }
XLPP_API void xlpp_cell_set_bool(xlpp_cell c, int v)             { CELL(c)->setBoolValue(v != 0); }
XLPP_API void xlpp_cell_set_empty(xlpp_cell c)                   { CELL(c)->setValue(std::monostate{}); }
XLPP_API void xlpp_cell_set_error(xlpp_cell c, int errorCode) {
    switch (errorCode) {
    case XLPP_ERROR_NULL: CELL(c)->setError(xlpp::CellError::Null); break;
    case XLPP_ERROR_DIV0: CELL(c)->setError(xlpp::CellError::DivisionByZero); break;
    case XLPP_ERROR_REF: CELL(c)->setError(xlpp::CellError::Reference); break;
    case XLPP_ERROR_NAME: CELL(c)->setError(xlpp::CellError::Name); break;
    case XLPP_ERROR_NUM: CELL(c)->setError(xlpp::CellError::Number); break;
    case XLPP_ERROR_NA: CELL(c)->setError(xlpp::CellError::NotAvailable); break;
    case XLPP_ERROR_GETTING_DATA: CELL(c)->setError(xlpp::CellError::GettingData); break;
    default: CELL(c)->setError(xlpp::CellError::Value); break;
    }
}
XLPP_API void xlpp_cell_set_date(xlpp_cell c, int year, int month, int day, int hour, int minute, double second, int hasTime) {
    xlpp::DateTime d{year, month, day, hour, minute, second};
    if (hasTime) CELL(c)->setDateTime(d);
    else CELL(c)->setDate(d);
}
XLPP_API void xlpp_cell_clear(xlpp_cell c)                       { CELL(c)->clear(); }

XLPP_API const char* xlpp_cell_get_formula(xlpp_cell c)          { return CELL(c)->formula().c_str(); }
XLPP_API void xlpp_cell_set_formula(xlpp_cell c, const char* f)   { CELL(c)->setFormula(f ? f : ""); }
XLPP_API void xlpp_cell_set_shared_formula(xlpp_cell c, const char* f, unsigned sharedIndex, const char* reference) {
    CELL(c)->setSharedFormula(f ? f : "", sharedIndex, reference ? reference : "");
}
XLPP_API void xlpp_cell_set_array_formula(xlpp_cell c, const char* f, const char* reference) {
    CELL(c)->setArrayFormula(f ? f : "", reference ? reference : "");
}
XLPP_API void xlpp_cell_set_dynamic_array_formula(xlpp_cell c, const char* f, const char* reference) {
    CELL(c)->setDynamicArrayFormula(f ? f : "", reference ? reference : "");
}
XLPP_API int xlpp_cell_has_formula(xlpp_cell c)                  { return CELL(c)->hasFormula() ? 1 : 0; }
XLPP_API void xlpp_cell_clear_formula(xlpp_cell c)               { CELL(c)->clearFormula(); }

XLPP_API xlpp_style xlpp_cell_style(xlpp_cell c)         { return reinterpret_cast<xlpp_style>(&CELL(c)->style()); }
XLPP_API xlpp_font xlpp_cell_font(xlpp_cell c)           { return reinterpret_cast<xlpp_font>(&CELL(c)->font()); }
XLPP_API xlpp_fill xlpp_cell_fill(xlpp_cell c)           { return reinterpret_cast<xlpp_fill>(&CELL(c)->fill()); }
XLPP_API xlpp_border xlpp_cell_border(xlpp_cell c)       { return reinterpret_cast<xlpp_border>(&CELL(c)->border()); }
XLPP_API xlpp_alignment xlpp_cell_alignment(xlpp_cell c)  { return reinterpret_cast<xlpp_alignment>(&CELL(c)->alignment()); }
XLPP_API void xlpp_cell_set_number_format(xlpp_cell c, const char* v) { CELL(c)->setNumberFormat(v); }
XLPP_API void xlpp_cell_number_format(xlpp_cell c, char* out, int outSize) { copyStr(CELL(c)->numberFormat(), out, outSize); }
XLPP_API void xlpp_cell_set_named_style(xlpp_cell c, const char* name) { CELL(c)->setNamedStyle(name); }
XLPP_API void xlpp_cell_named_style(xlpp_cell c, char* out, int outSize) {
    const auto& ns = CELL(c)->namedStyle();
    if (ns) copyStr(*ns, out, outSize);
    else if (out && outSize > 0) out[0] = '\0';
}

XLPP_API int xlpp_cell_has_hyperlink(xlpp_cell c) { return CELL(c)->hasHyperlink() ? 1 : 0; }
XLPP_API xlpp_hyperlink xlpp_cell_hyperlink(xlpp_cell c) { return reinterpret_cast<xlpp_hyperlink>(&CELL(c)->hyperlink()); }
XLPP_API void xlpp_cell_set_hyperlink(xlpp_cell c, const char* url) {
    CELL(c)->setHyperlink(xlpp::Hyperlink(url ? url : ""));
}
XLPP_API void xlpp_cell_set_hyperlink_full(xlpp_cell c, const char* url, const char* display, const char* tooltip, int external) {
    xlpp::Hyperlink h(url ? url : "");
    if (display) h.setDisplay(display);
    if (tooltip) h.setTooltip(tooltip);
    h.setExternal(external != 0);
    CELL(c)->setHyperlink(std::move(h));
}
XLPP_API void xlpp_cell_clear_hyperlink(xlpp_cell c) { CELL(c)->clearHyperlink(); }

XLPP_API int xlpp_cell_has_comment(xlpp_cell c) { return CELL(c)->hasComment() ? 1 : 0; }
XLPP_API xlpp_comment xlpp_cell_comment(xlpp_cell c) { return reinterpret_cast<xlpp_comment>(&CELL(c)->comment()); }
XLPP_API void xlpp_cell_set_comment(xlpp_cell c, const char* text, const char* author) {
    CELL(c)->setComment(xlpp::Comment(text ? text : "", author ? author : ""));
}
XLPP_API void xlpp_cell_clear_comment(xlpp_cell c) { CELL(c)->clearComment(); }

// ============================================================
// Hyperlink / Comment
// ============================================================
#define HL(h) reinterpret_cast<xlpp::Hyperlink*>(h)
#define CM(h) reinterpret_cast<xlpp::Comment*>(h)
XLPP_API void xlpp_hyperlink_set_target(xlpp_hyperlink h, const char* v) { HL(h)->setTarget(v); }
XLPP_API void xlpp_hyperlink_set_display(xlpp_hyperlink h, const char* v) { HL(h)->setDisplay(v); }
XLPP_API void xlpp_hyperlink_set_tooltip(xlpp_hyperlink h, const char* v) { HL(h)->setTooltip(v); }
XLPP_API void xlpp_hyperlink_set_external(xlpp_hyperlink h, int v) { HL(h)->setExternal(v != 0); }
XLPP_API const char* xlpp_hyperlink_target(xlpp_hyperlink h) { return HL(h)->target().c_str(); }
XLPP_API const char* xlpp_hyperlink_display(xlpp_hyperlink h) { return HL(h)->display().c_str(); }
XLPP_API const char* xlpp_hyperlink_tooltip(xlpp_hyperlink h) { return HL(h)->tooltip().c_str(); }
XLPP_API int xlpp_hyperlink_external(xlpp_hyperlink h) { return HL(h)->external() ? 1 : 0; }

XLPP_API void xlpp_comment_set_text(xlpp_comment c, const char* v) { CM(c)->setText(v); }
XLPP_API void xlpp_comment_set_author(xlpp_comment c, const char* v) { CM(c)->setAuthor(v); }
XLPP_API const char* xlpp_comment_text(xlpp_comment c) { return CM(c)->text().c_str(); }
XLPP_API const char* xlpp_comment_author(xlpp_comment c) { return CM(c)->author().c_str(); }

// ============================================================
// RichText
// ============================================================
#define RT(h) reinterpret_cast<xlpp::RichText*>(h)
#define RTR(h) reinterpret_cast<xlpp::RichTextRun*>(h)
XLPP_API xlpp_richtext xlpp_richtext_create(void) {
    return reinterpret_cast<xlpp_richtext>(new xlpp::RichText());
}
XLPP_API void xlpp_richtext_destroy(xlpp_richtext rt) { delete RT(rt); }
XLPP_API int xlpp_richtext_run_count(xlpp_richtext rt) { return static_cast<int>(RT(rt)->runs().size()); }
XLPP_API xlpp_richtextrun xlpp_richtext_add_run(xlpp_richtext rt, const char* text) {
    RT(rt)->addRun(xlpp::RichTextRun(text ? text : ""));
    return reinterpret_cast<xlpp_richtextrun>(&RT(rt)->runs().back());
}
XLPP_API xlpp_richtextrun xlpp_richtext_run_at(xlpp_richtext rt, int index) {
    try { return reinterpret_cast<xlpp_richtextrun>(&RT(rt)->runs()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API int xlpp_richtext_empty(xlpp_richtext rt) { return RT(rt)->empty() ? 1 : 0; }
XLPP_API void xlpp_richtext_plain_text(xlpp_richtext rt, char* out, int outSize) { copyStr(RT(rt)->plainText(), out, outSize); }

XLPP_API void xlpp_richtextrun_set_text(xlpp_richtextrun r, const char* v) { RTR(r)->setText(v); }
XLPP_API void xlpp_richtextrun_set_bold(xlpp_richtextrun r, int v) { RTR(r)->setBold(v != 0); }
XLPP_API void xlpp_richtextrun_set_italic(xlpp_richtextrun r, int v) { RTR(r)->setItalic(v != 0); }
XLPP_API void xlpp_richtextrun_set_underline(xlpp_richtextrun r, int v) { RTR(r)->setUnderline(v != 0); }
XLPP_API void xlpp_richtextrun_set_strike(xlpp_richtextrun r, int v) { RTR(r)->setStrike(v != 0); }
XLPP_API void xlpp_richtextrun_set_color(xlpp_richtextrun r, const char* v) { RTR(r)->setColor(v); }
XLPP_API void xlpp_richtextrun_set_size(xlpp_richtextrun r, double v) { RTR(r)->setSize(v); }
XLPP_API void xlpp_richtextrun_set_font(xlpp_richtextrun r, const char* v) { RTR(r)->setFontName(v); }
XLPP_API const char* xlpp_richtextrun_text(xlpp_richtextrun r) { return RTR(r)->text().c_str(); }

// ============================================================
// Named style / Defined name
// ============================================================
#define NS(h) reinterpret_cast<xlpp::NamedStyle*>(h)
#define DN(h) reinterpret_cast<xlpp::DefinedName*>(h)
XLPP_API void xlpp_namedstyle_set_name(xlpp_namedstyle s, const char* v) { NS(s)->setName(v); }
XLPP_API const char* xlpp_namedstyle_name(xlpp_namedstyle s) { return NS(s)->name().c_str(); }
XLPP_API xlpp_style xlpp_namedstyle_style(xlpp_namedstyle s) { return reinterpret_cast<xlpp_style>(&NS(s)->style()); }

XLPP_API void xlpp_definedname_set_value(xlpp_definedname d, const char* v) { try { DN(d)->setValue(v); } catch (...) {} }
XLPP_API const char* xlpp_definedname_name(xlpp_definedname d) { return DN(d)->name().c_str(); }
XLPP_API const char* xlpp_definedname_value(xlpp_definedname d) { return DN(d)->value().c_str(); }
XLPP_API void xlpp_definedname_set_local_sheet_id(xlpp_definedname d, uint64_t v) { DN(d)->setLocalSheetId(static_cast<std::size_t>(v)); }
XLPP_API void xlpp_definedname_clear_local_sheet_id(xlpp_definedname d) { DN(d)->clearLocalSheetId(); }
XLPP_API int xlpp_definedname_has_local_sheet_id(xlpp_definedname d) { return DN(d)->localSheetId().has_value() ? 1 : 0; }
XLPP_API void xlpp_definedname_set_hidden(xlpp_definedname d, int v) { DN(d)->setHidden(v != 0); }
XLPP_API int xlpp_definedname_hidden(xlpp_definedname d) { return DN(d)->hidden() ? 1 : 0; }
XLPP_API void xlpp_definedname_set_comment(xlpp_definedname d, const char* v) { DN(d)->setComment(v); }
XLPP_API const char* xlpp_definedname_comment(xlpp_definedname d) { return DN(d)->comment().c_str(); }

// ============================================================
// AutoFilter / Sort
// ============================================================
#define AF(h) reinterpret_cast<xlpp::AutoFilter*>(h)
#define FC(h) reinterpret_cast<xlpp::FilterColumn*>(h)
#define SS(h) reinterpret_cast<xlpp::SortState*>(h)
XLPP_API void xlpp_autofilter_set_reference(xlpp_autofilter f, const char* v) { AF(f)->setReference(v); }
XLPP_API void xlpp_autofilter_reference(xlpp_autofilter f, char* out, int outSize) { copyStr(AF(f)->reference(), out, outSize); }
XLPP_API int xlpp_autofilter_enabled(xlpp_autofilter f) { return AF(f)->enabled() ? 1 : 0; }
XLPP_API void xlpp_autofilter_clear(xlpp_autofilter f) { AF(f)->clear(); }
XLPP_API xlpp_filtercol xlpp_autofilter_column(xlpp_autofilter f, uint64_t columnId) {
    return reinterpret_cast<xlpp_filtercol>(&AF(f)->column(static_cast<std::size_t>(columnId)));
}
XLPP_API xlpp_sortstate xlpp_autofilter_sort_state(xlpp_autofilter f) { return reinterpret_cast<xlpp_sortstate>(&AF(f)->sortState()); }

XLPP_API uint64_t xlpp_filtercol_column_id(xlpp_filtercol c) { return FC(c)->columnId(); }
XLPP_API void xlpp_filtercol_add_value(xlpp_filtercol c, const char* v) { FC(c)->addValue(v); }
XLPP_API void xlpp_filtercol_clear_values(xlpp_filtercol c) { FC(c)->clearValues(); }
XLPP_API int xlpp_filtercol_value_count(xlpp_filtercol c) { return static_cast<int>(FC(c)->values().size()); }
XLPP_API void xlpp_filtercol_value_at(xlpp_filtercol c, int index, char* out, int outSize) {
    try { copyStr(FC(c)->values()[static_cast<std::size_t>(index)], out, outSize); } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}
XLPP_API void xlpp_filtercol_set_and_mode(xlpp_filtercol c, int v) { FC(c)->setAndMode(v != 0); }
XLPP_API int xlpp_filtercol_and_mode(xlpp_filtercol c) { return FC(c)->andMode() ? 1 : 0; }
XLPP_API void xlpp_filtercol_set_include_blank(xlpp_filtercol c, int v) { FC(c)->setIncludeBlank(v != 0); }
XLPP_API int xlpp_filtercol_include_blank(xlpp_filtercol c) { return FC(c)->includeBlank() ? 1 : 0; }

XLPP_API void xlpp_sortstate_set_reference(xlpp_sortstate s, const char* v) { SS(s)->setReference(v); }
XLPP_API void xlpp_sortstate_reference(xlpp_sortstate s, char* out, int outSize) { copyStr(SS(s)->reference(), out, outSize); }
XLPP_API void xlpp_sortstate_set_case_sensitive(xlpp_sortstate s, int v) { SS(s)->setCaseSensitive(v != 0); }
XLPP_API int xlpp_sortstate_case_sensitive(xlpp_sortstate s) { return SS(s)->caseSensitive() ? 1 : 0; }
XLPP_API void xlpp_sortstate_add_condition(xlpp_sortstate s, const char* reference, int descending) { SS(s)->addCondition(reference, descending != 0); }
XLPP_API void xlpp_sortstate_clear(xlpp_sortstate s) { SS(s)->clear(); }

// ============================================================
// Page setup / margins / print options / header footer
// ============================================================
#define PS(h) reinterpret_cast<xlpp::PageSetup*>(h)
#define PM(h) reinterpret_cast<xlpp::PageMargins*>(h)
#define PO(h) reinterpret_cast<xlpp::PrintOptions*>(h)
#define HF(h) reinterpret_cast<xlpp::HeaderFooter*>(h)
XLPP_API void xlpp_pagesetup_set_orientation(xlpp_pagesetup p, int v) { PS(p)->setOrientation(static_cast<xlpp::PageOrientation>(v)); }
XLPP_API int xlpp_pagesetup_orientation(xlpp_pagesetup p) { return static_cast<int>(PS(p)->orientation()); }
XLPP_API void xlpp_pagesetup_set_paper_size(xlpp_pagesetup p, int v) { PS(p)->setPaperSize(static_cast<xlpp::PaperSize>(v)); }
XLPP_API int xlpp_pagesetup_paper_size(xlpp_pagesetup p) { return static_cast<int>(PS(p)->paperSize()); }
XLPP_API void xlpp_pagesetup_set_scale(xlpp_pagesetup p, unsigned v) { PS(p)->setScale(v); }
XLPP_API unsigned xlpp_pagesetup_scale(xlpp_pagesetup p) { return PS(p)->scale(); }
XLPP_API void xlpp_pagesetup_set_fit_to_width(xlpp_pagesetup p, unsigned v) { PS(p)->setFitToWidth(v); }
XLPP_API unsigned xlpp_pagesetup_fit_to_width(xlpp_pagesetup p) { return PS(p)->fitToWidth(); }
XLPP_API void xlpp_pagesetup_set_fit_to_height(xlpp_pagesetup p, unsigned v) { PS(p)->setFitToHeight(v); }
XLPP_API unsigned xlpp_pagesetup_fit_to_height(xlpp_pagesetup p) { return PS(p)->fitToHeight(); }
XLPP_API void xlpp_pagesetup_set_fit_to_page(xlpp_pagesetup p, int v) { PS(p)->setFitToPage(v != 0); }
XLPP_API int xlpp_pagesetup_fit_to_page(xlpp_pagesetup p) { return PS(p)->fitToPage() ? 1 : 0; }
XLPP_API void xlpp_pagesetup_set_black_and_white(xlpp_pagesetup p, int v) { PS(p)->setBlackAndWhite(v != 0); }
XLPP_API int xlpp_pagesetup_black_and_white(xlpp_pagesetup p) { return PS(p)->blackAndWhite() ? 1 : 0; }
XLPP_API void xlpp_pagesetup_set_draft(xlpp_pagesetup p, int v) { PS(p)->setDraft(v != 0); }
XLPP_API int xlpp_pagesetup_draft(xlpp_pagesetup p) { return PS(p)->draft() ? 1 : 0; }

XLPP_API void xlpp_pagemargins_set_left(xlpp_pagemargins m, double v) { PM(m)->setLeft(v); }
XLPP_API double xlpp_pagemargins_left(xlpp_pagemargins m) { return PM(m)->left(); }
XLPP_API void xlpp_pagemargins_set_right(xlpp_pagemargins m, double v) { PM(m)->setRight(v); }
XLPP_API double xlpp_pagemargins_right(xlpp_pagemargins m) { return PM(m)->right(); }
XLPP_API void xlpp_pagemargins_set_top(xlpp_pagemargins m, double v) { PM(m)->setTop(v); }
XLPP_API double xlpp_pagemargins_top(xlpp_pagemargins m) { return PM(m)->top(); }
XLPP_API void xlpp_pagemargins_set_bottom(xlpp_pagemargins m, double v) { PM(m)->setBottom(v); }
XLPP_API double xlpp_pagemargins_bottom(xlpp_pagemargins m) { return PM(m)->bottom(); }
XLPP_API void xlpp_pagemargins_set_header(xlpp_pagemargins m, double v) { PM(m)->setHeader(v); }
XLPP_API double xlpp_pagemargins_header(xlpp_pagemargins m) { return PM(m)->header(); }
XLPP_API void xlpp_pagemargins_set_footer(xlpp_pagemargins m, double v) { PM(m)->setFooter(v); }
XLPP_API double xlpp_pagemargins_footer(xlpp_pagemargins m) { return PM(m)->footer(); }

XLPP_API void xlpp_printopts_set_horizontal_centered(xlpp_printopts p, int v) { PO(p)->setHorizontalCentered(v != 0); }
XLPP_API int xlpp_printopts_horizontal_centered(xlpp_printopts p) { return PO(p)->horizontalCentered() ? 1 : 0; }
XLPP_API void xlpp_printopts_set_vertical_centered(xlpp_printopts p, int v) { PO(p)->setVerticalCentered(v != 0); }
XLPP_API int xlpp_printopts_vertical_centered(xlpp_printopts p) { return PO(p)->verticalCentered() ? 1 : 0; }
XLPP_API void xlpp_printopts_set_headings(xlpp_printopts p, int v) { PO(p)->setHeadings(v != 0); }
XLPP_API int xlpp_printopts_headings(xlpp_printopts p) { return PO(p)->headings() ? 1 : 0; }
XLPP_API void xlpp_printopts_set_grid_lines(xlpp_printopts p, int v) { PO(p)->setGridLines(v != 0); }
XLPP_API int xlpp_printopts_grid_lines(xlpp_printopts p) { return PO(p)->gridLines() ? 1 : 0; }

XLPP_API void xlpp_headerfooter_set_odd_header(xlpp_headerfooter h, const char* v) { HF(h)->setOddHeader(v); }
XLPP_API void xlpp_headerfooter_set_odd_footer(xlpp_headerfooter h, const char* v) { HF(h)->setOddFooter(v); }
XLPP_API void xlpp_headerfooter_set_even_header(xlpp_headerfooter h, const char* v) { HF(h)->setEvenHeader(v); }
XLPP_API void xlpp_headerfooter_set_even_footer(xlpp_headerfooter h, const char* v) { HF(h)->setEvenFooter(v); }
XLPP_API void xlpp_headerfooter_set_different_odd_even(xlpp_headerfooter h, int v) { HF(h)->setDifferentOddEven(v != 0); }
XLPP_API int xlpp_headerfooter_different_odd_even(xlpp_headerfooter h) { return HF(h)->differentOddEven() ? 1 : 0; }
XLPP_API void xlpp_headerfooter_set_different_first(xlpp_headerfooter h, int v) { HF(h)->setDifferentFirst(v != 0); }
XLPP_API int xlpp_headerfooter_different_first(xlpp_headerfooter h) { return HF(h)->differentFirst() ? 1 : 0; }

// ============================================================
// Worksheet protection
// ============================================================
#define WSP(h) reinterpret_cast<xlpp::WorksheetProtection*>(h)
XLPP_API void xlpp_wssprotection_set_enabled(xlpp_wssprotection p, int v) { WSP(p)->setEnabled(v != 0); }
XLPP_API int xlpp_wssprotection_enabled(xlpp_wssprotection p) { return WSP(p)->enabled() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_password_hash(xlpp_wssprotection p, const char* v) { WSP(p)->setPasswordHash(v); }
XLPP_API const char* xlpp_wssprotection_password_hash(xlpp_wssprotection p) { return WSP(p)->passwordHash().c_str(); }
XLPP_API void xlpp_wssprotection_set_select_locked(xlpp_wssprotection p, int v) { WSP(p)->setSelectLockedCells(v != 0); }
XLPP_API int xlpp_wssprotection_select_locked(xlpp_wssprotection p) { return WSP(p)->selectLockedCells() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_select_unlocked(xlpp_wssprotection p, int v) { WSP(p)->setSelectUnlockedCells(v != 0); }
XLPP_API int xlpp_wssprotection_select_unlocked(xlpp_wssprotection p) { return WSP(p)->selectUnlockedCells() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_format_cells(xlpp_wssprotection p, int v) { WSP(p)->setFormatCells(v != 0); }
XLPP_API int xlpp_wssprotection_format_cells(xlpp_wssprotection p) { return WSP(p)->formatCells() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_format_columns(xlpp_wssprotection p, int v) { WSP(p)->setFormatColumns(v != 0); }
XLPP_API int xlpp_wssprotection_format_columns(xlpp_wssprotection p) { return WSP(p)->formatColumns() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_format_rows(xlpp_wssprotection p, int v) { WSP(p)->setFormatRows(v != 0); }
XLPP_API int xlpp_wssprotection_format_rows(xlpp_wssprotection p) { return WSP(p)->formatRows() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_insert_rows(xlpp_wssprotection p, int v) { WSP(p)->setInsertRows(v != 0); }
XLPP_API int xlpp_wssprotection_insert_rows(xlpp_wssprotection p) { return WSP(p)->insertRows() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_insert_columns(xlpp_wssprotection p, int v) { WSP(p)->setInsertColumns(v != 0); }
XLPP_API int xlpp_wssprotection_insert_columns(xlpp_wssprotection p) { return WSP(p)->insertColumns() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_delete_rows(xlpp_wssprotection p, int v) { WSP(p)->setDeleteRows(v != 0); }
XLPP_API int xlpp_wssprotection_delete_rows(xlpp_wssprotection p) { return WSP(p)->deleteRows() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_delete_columns(xlpp_wssprotection p, int v) { WSP(p)->setDeleteColumns(v != 0); }
XLPP_API int xlpp_wssprotection_delete_columns(xlpp_wssprotection p) { return WSP(p)->deleteColumns() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_sort(xlpp_wssprotection p, int v) { WSP(p)->setSort(v != 0); }
XLPP_API int xlpp_wssprotection_sort(xlpp_wssprotection p) { return WSP(p)->sort() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_auto_filter(xlpp_wssprotection p, int v) { WSP(p)->setAutoFilter(v != 0); }
XLPP_API int xlpp_wssprotection_auto_filter(xlpp_wssprotection p) { return WSP(p)->autoFilter() ? 1 : 0; }

// ============================================================
// SheetView
// ============================================================
#define SV(h) reinterpret_cast<xlpp::SheetView*>(h)
XLPP_API void xlpp_sheetview_set_workbook_view_id(xlpp_sheetview s, int v) { SV(s)->setWorkbookViewId(v); }
XLPP_API int xlpp_sheetview_workbook_view_id(xlpp_sheetview s) { return SV(s)->workbookViewId(); }
XLPP_API void xlpp_sheetview_set_tab_color(xlpp_sheetview s, const char* v) { SV(s)->setTabColor(v); }
XLPP_API void xlpp_sheetview_clear_tab_color(xlpp_sheetview s) { SV(s)->clearTabColor(); }
XLPP_API void xlpp_sheetview_tab_color(xlpp_sheetview s, char* out, int outSize) {
    const auto& c = SV(s)->tabColor();
    if (c) copyStr(*c, out, outSize);
    else if (out && outSize > 0) out[0] = '\0';
}
XLPP_API void xlpp_sheetview_set_zoom_scale(xlpp_sheetview s, int v) { SV(s)->setZoomScale(v); }
XLPP_API int xlpp_sheetview_zoom_scale(xlpp_sheetview s) { return SV(s)->zoomScale(); }
XLPP_API void xlpp_sheetview_set_zoom_scale_normal(xlpp_sheetview s, int v) { SV(s)->setZoomScaleNormal(v); }
XLPP_API int xlpp_sheetview_zoom_scale_normal(xlpp_sheetview s) { return SV(s)->zoomScaleNormal(); }
XLPP_API void xlpp_sheetview_set_show_grid_lines(xlpp_sheetview s, int v) { SV(s)->setShowGridLines(v != 0); }
XLPP_API int xlpp_sheetview_show_grid_lines(xlpp_sheetview s) { return SV(s)->showGridLines() ? 1 : 0; }
XLPP_API void xlpp_sheetview_set_tab_selected(xlpp_sheetview s, int v) { SV(s)->setTabSelected(v != 0); }
XLPP_API int xlpp_sheetview_tab_selected(xlpp_sheetview s) { return SV(s)->tabSelected() ? 1 : 0; }
XLPP_API void xlpp_sheetview_set_right_to_left(xlpp_sheetview s, int v) { SV(s)->setRightToLeft(v != 0); }
XLPP_API int xlpp_sheetview_right_to_left(xlpp_sheetview s) { return SV(s)->rightToLeft() ? 1 : 0; }
XLPP_API void xlpp_sheetview_set_show_outline_symbols(xlpp_sheetview s, int v) { SV(s)->setShowOutlineSymbols(v != 0); }
XLPP_API int xlpp_sheetview_show_outline_symbols(xlpp_sheetview s) { return SV(s)->showOutlineSymbols() ? 1 : 0; }
XLPP_API void xlpp_sheetview_set_pane(xlpp_sheetview s, const char* v) { SV(s)->setPane(v); }
XLPP_API void xlpp_sheetview_pane(xlpp_sheetview s, char* out, int outSize) { copyStr(SV(s)->pane(), out, outSize); }
XLPP_API void xlpp_sheetview_set_top_left_cell(xlpp_sheetview s, const char* v) { SV(s)->setTopLeftCell(v); }
XLPP_API void xlpp_sheetview_top_left_cell(xlpp_sheetview s, char* out, int outSize) { copyStr(SV(s)->topLeftCell(), out, outSize); }
XLPP_API void xlpp_sheetview_set_x_split(xlpp_sheetview s, int v) { SV(s)->setXSplit(v); }
XLPP_API int xlpp_sheetview_x_split(xlpp_sheetview s) { return SV(s)->xSplit(); }
XLPP_API void xlpp_sheetview_set_y_split(xlpp_sheetview s, int v) { SV(s)->setYSplit(v); }
XLPP_API int xlpp_sheetview_y_split(xlpp_sheetview s) { return SV(s)->ySplit(); }

// ============================================================
// Image
// ============================================================
#define IMG(h) reinterpret_cast<xlpp::Image*>(h)
XLPP_API void xlpp_image_set_anchor(xlpp_image img, const char* v) { IMG(img)->setAnchor(v); }
XLPP_API void xlpp_image_anchor(xlpp_image img, char* out, int outSize) { copyStr(IMG(img)->anchor(), out, outSize); }
XLPP_API void xlpp_image_extension(xlpp_image img, char* out, int outSize) { copyStr(IMG(img)->extension(), out, outSize); }
XLPP_API void xlpp_image_set_width(xlpp_image img, double v) { IMG(img)->setWidthPixels(v); }
XLPP_API double xlpp_image_width(xlpp_image img) { return IMG(img)->widthPixels(); }
XLPP_API void xlpp_image_set_height(xlpp_image img, double v) { IMG(img)->setHeightPixels(v); }
XLPP_API double xlpp_image_height(xlpp_image img) { return IMG(img)->heightPixels(); }
XLPP_API void xlpp_image_set_name(xlpp_image img, const char* v) { IMG(img)->setName(v); }
XLPP_API void xlpp_image_name(xlpp_image img, char* out, int outSize) { copyStr(IMG(img)->name(), out, outSize); }

// ============================================================
// Table
// ============================================================
#define TBL(h) reinterpret_cast<xlpp::Table*>(h)
#define TC(h) reinterpret_cast<xlpp::TableColumn*>(h)
#define TSI(h) reinterpret_cast<xlpp::TableStyleInfo*>(h)
XLPP_API void xlpp_table_set_display_name(xlpp_table t, const char* v) { try { TBL(t)->setDisplayName(v); } catch (...) {} }
XLPP_API void xlpp_table_name(xlpp_table t, char* out, int outSize) { copyStr(TBL(t)->name(), out, outSize); }
XLPP_API void xlpp_table_display_name(xlpp_table t, char* out, int outSize) { copyStr(TBL(t)->displayName(), out, outSize); }
XLPP_API void xlpp_table_set_reference(xlpp_table t, const char* v) { try { TBL(t)->setReference(v); } catch (...) {} }
XLPP_API void xlpp_table_reference(xlpp_table t, char* out, int outSize) { copyStr(TBL(t)->reference(), out, outSize); }
XLPP_API void xlpp_table_set_show_header_row(xlpp_table t, int v) { TBL(t)->setShowHeaderRow(v != 0); }
XLPP_API int xlpp_table_show_header_row(xlpp_table t) { return TBL(t)->showHeaderRow() ? 1 : 0; }
XLPP_API void xlpp_table_set_show_totals_row(xlpp_table t, int v) { TBL(t)->setShowTotalsRow(v != 0); }
XLPP_API int xlpp_table_show_totals_row(xlpp_table t) { return TBL(t)->showTotalsRow() ? 1 : 0; }
XLPP_API int xlpp_table_column_count(xlpp_table t) { return static_cast<int>(TBL(t)->columns().size()); }
XLPP_API xlpp_tablecolumn xlpp_table_column_at(xlpp_table t, int index) {
    try { return reinterpret_cast<xlpp_tablecolumn>(&TBL(t)->columns()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API xlpp_tablecolumn xlpp_table_add_column(xlpp_table t, const char* name) {
    return reinterpret_cast<xlpp_tablecolumn>(&TBL(t)->addColumn(name));
}
XLPP_API xlpp_tablestyle xlpp_table_style_info(xlpp_table t) { return reinterpret_cast<xlpp_tablestyle>(&TBL(t)->styleInfo()); }

XLPP_API void xlpp_tablecolumn_set_name(xlpp_tablecolumn c, const char* v) { try { TC(c)->setName(v); } catch (...) {} }
XLPP_API void xlpp_tablecolumn_name(xlpp_tablecolumn c, char* out, int outSize) { copyStr(TC(c)->name(), out, outSize); }
XLPP_API uint64_t xlpp_tablecolumn_id(xlpp_tablecolumn c) { return TC(c)->id(); }

XLPP_API void xlpp_tablestyle_set_name(xlpp_tablestyle s, const char* v) { TSI(s)->setName(v); }
XLPP_API void xlpp_tablestyle_name(xlpp_tablestyle s, char* out, int outSize) { copyStr(TSI(s)->name(), out, outSize); }
XLPP_API void xlpp_tablestyle_set_show_first(xlpp_tablestyle s, int v) { TSI(s)->setShowFirstColumn(v != 0); }
XLPP_API int xlpp_tablestyle_show_first(xlpp_tablestyle s) { return TSI(s)->showFirstColumn() ? 1 : 0; }
XLPP_API void xlpp_tablestyle_set_show_last(xlpp_tablestyle s, int v) { TSI(s)->setShowLastColumn(v != 0); }
XLPP_API int xlpp_tablestyle_show_last(xlpp_tablestyle s) { return TSI(s)->showLastColumn() ? 1 : 0; }
XLPP_API void xlpp_tablestyle_set_show_row_stripes(xlpp_tablestyle s, int v) { TSI(s)->setShowRowStripes(v != 0); }
XLPP_API int xlpp_tablestyle_show_row_stripes(xlpp_tablestyle s) { return TSI(s)->showRowStripes() ? 1 : 0; }
XLPP_API void xlpp_tablestyle_set_show_column_stripes(xlpp_tablestyle s, int v) { TSI(s)->setShowColumnStripes(v != 0); }
XLPP_API int xlpp_tablestyle_show_column_stripes(xlpp_tablestyle s) { return TSI(s)->showColumnStripes() ? 1 : 0; }

// ============================================================
// Chart
// ============================================================
#define CH(h) reinterpret_cast<xlpp::Chart*>(h)
#define CS(h) reinterpret_cast<xlpp::ChartSeries*>(h)
XLPP_API void xlpp_chart_set_grouping(xlpp_chart c, int v) { CH(c)->setGrouping(static_cast<xlpp::Chart::Grouping>(v)); }
XLPP_API void xlpp_chart_set_scatter_style(xlpp_chart c, const char* v) { if (c) CH(c)->setScatterStyle(v ? v : "lineMarker"); }
XLPP_API void xlpp_chart_scatter_style(xlpp_chart c, char* out, int outSize) { if (c) copyStr(CH(c)->scatterStyle(), out, outSize); }
XLPP_API int xlpp_chart_add_plot(xlpp_chart c, int type, int grouping, int secondary_axes) {
    try {
        if (!c) return -1;
        CH(c)->addPlot(static_cast<xlpp::Chart::Type>(type), static_cast<xlpp::Chart::Grouping>(grouping), secondary_axes != 0);
        return static_cast<int>(CH(c)->plots().size() - 1);
    } catch (const std::exception& e) { setError(e); return -1; }
}
XLPP_API int xlpp_chart_plot_count(xlpp_chart c) { return c ? static_cast<int>(CH(c)->plots().size()) : 0; }
XLPP_API int xlpp_chart_plot_type(xlpp_chart c, int plot_index) {
    try { return static_cast<int>(CH(c)->plots().at(static_cast<std::size_t>(plot_index)).type); } catch (...) { return -1; }
}
XLPP_API int xlpp_chart_plot_uses_secondary_axes(xlpp_chart c, int plot_index) {
    try { return CH(c)->plots().at(static_cast<std::size_t>(plot_index)).usesSecondaryAxes ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API xlpp_chartseries xlpp_chart_add_series_to_plot(xlpp_chart c, int plot_index, const char* title) {
    try {
        if (!c || plot_index < 0) return nullptr;
        return reinterpret_cast<xlpp_chartseries>(&CH(c)->addSeriesToPlot(static_cast<std::size_t>(plot_index), xlpp::ChartSeries(title ? title : "")));
    } catch (const std::exception& e) { setError(e); return nullptr; }
}
XLPP_API int xlpp_chart_grouping(xlpp_chart c) { return static_cast<int>(CH(c)->grouping()); }
XLPP_API void xlpp_chart_set_title(xlpp_chart c, const char* v) { CH(c)->setTitle(v); }
XLPP_API void xlpp_chart_title(xlpp_chart c, char* out, int outSize) { copyStr(CH(c)->title(), out, outSize); }
XLPP_API void xlpp_chart_set_x_axis_title(xlpp_chart c, const char* v) { CH(c)->setXAxisTitle(v); }
XLPP_API void xlpp_chart_set_y_axis_title(xlpp_chart c, const char* v) { CH(c)->setYAxisTitle(v); }
XLPP_API void xlpp_chart_set_style(xlpp_chart c, const char* v) { CH(c)->setStyle(v); }
XLPP_API void xlpp_chart_set_width(xlpp_chart c, int v) { CH(c)->setWidth(v); }
XLPP_API int xlpp_chart_width(xlpp_chart c) { return CH(c)->width(); }
XLPP_API void xlpp_chart_set_height(xlpp_chart c, int v) { CH(c)->setHeight(v); }
XLPP_API int xlpp_chart_height(xlpp_chart c) { return CH(c)->height(); }
XLPP_API void xlpp_chart_set_show_legend(xlpp_chart c, int v) { CH(c)->setShowLegend(v != 0); }
XLPP_API int xlpp_chart_show_legend(xlpp_chart c) { return CH(c)->showLegend() ? 1 : 0; }
XLPP_API void xlpp_chart_set_legend_position(xlpp_chart c, const char* v) { CH(c)->setLegendPosition(v); }
XLPP_API int xlpp_chart_series_count(xlpp_chart c) { return static_cast<int>(CH(c)->series().size()); }
XLPP_API xlpp_chartseries xlpp_chart_series_at(xlpp_chart c, int index) {
    try { return reinterpret_cast<xlpp_chartseries>(&CH(c)->series()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API xlpp_chartseries xlpp_chart_add_series(xlpp_chart c, const char* title) {
    return reinterpret_cast<xlpp_chartseries>(&CH(c)->addSeries(xlpp::ChartSeries(title ? title : "")));
}

XLPP_API void xlpp_chartseries_set_title(xlpp_chartseries s, const char* v) { CS(s)->setTitle(v); }
XLPP_API void xlpp_chartseries_set_values_reference(xlpp_chartseries s, const char* v) { CS(s)->setValuesReference(v); }
XLPP_API void xlpp_chartseries_set_categories_reference(xlpp_chartseries s, const char* v) { CS(s)->setCategoriesReference(v); }
XLPP_API void xlpp_chartseries_set_bubble_size_reference(xlpp_chartseries s, const char* v) { if (s) CS(s)->setBubbleSizeReference(v ? v : ""); }
XLPP_API void xlpp_chartseries_bubble_size_reference(xlpp_chartseries s, char* out, int outSize) { if (s) copyStr(CS(s)->bubbleSizeReference(), out, outSize); }

// ============================================================
// Pivot table
// ============================================================
#define PT(h) reinterpret_cast<xlpp::PivotTable*>(h)
#define PC(h) reinterpret_cast<xlpp::PivotCache*>(h)
XLPP_API void xlpp_pivottable_set_name(xlpp_pivottable p, const char* v) { PT(p)->setName(v); }
XLPP_API void xlpp_pivottable_name(xlpp_pivottable p, char* out, int outSize) { copyStr(PT(p)->name(), out, outSize); }
XLPP_API void xlpp_pivottable_set_location(xlpp_pivottable p, const char* v) { PT(p)->setLocation(v); }
XLPP_API void xlpp_pivottable_location(xlpp_pivottable p, char* out, int outSize) { copyStr(PT(p)->location(), out, outSize); }
XLPP_API xlpp_pivotcache xlpp_pivottable_cache(xlpp_pivottable p) { return reinterpret_cast<xlpp_pivotcache>(&PT(p)->cache()); }
XLPP_API void xlpp_pivottable_add_row_field(xlpp_pivottable p, const char* name) { PT(p)->addRowField(name); }
XLPP_API void xlpp_pivottable_add_column_field(xlpp_pivottable p, const char* name) { PT(p)->addColumnField(name); }
XLPP_API void xlpp_pivottable_add_page_field(xlpp_pivottable p, const char* name) { PT(p)->addPageField(name); }
XLPP_API void xlpp_pivottable_add_data_field(xlpp_pivottable p) { PT(p)->addDataField(); }

XLPP_API void xlpp_pivotcache_set_cache_id(xlpp_pivotcache c, int v) { PC(c)->setCacheId(v); }
XLPP_API int xlpp_pivotcache_cache_id(xlpp_pivotcache c) { return PC(c)->cacheId(); }
XLPP_API void xlpp_pivotcache_set_source_data(xlpp_pivotcache c, const char* v) { PC(c)->setSourceData(v); }
XLPP_API void xlpp_pivotcache_source_data(xlpp_pivotcache c, char* out, int outSize) { copyStr(PC(c)->sourceData(), out, outSize); }

// ============================================================
// Conditional formatting
// ============================================================
#define CF(h) reinterpret_cast<xlpp::ConditionalFormattingCollection*>(h)
#define CFE(h) reinterpret_cast<xlpp::ConditionalFormattingEntry*>(h)
#define CFR(h) reinterpret_cast<xlpp::ConditionalRule*>(h)
XLPP_API int xlpp_cfcollection_entry_count(xlpp_cfcollection c) { return static_cast<int>(CF(c)->entries().size()); }
XLPP_API xlpp_cfentry xlpp_cfcollection_add_entry(xlpp_cfcollection c, const char* reference) {
    try { return reinterpret_cast<xlpp_cfentry>(&CF(c)->add(reference)); } catch (...) { return nullptr; }
}
XLPP_API xlpp_cfentry xlpp_cfcollection_entry_at(xlpp_cfcollection c, int index) {
    try { return reinterpret_cast<xlpp_cfentry>(&CF(c)->entries()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API void xlpp_cfcollection_clear(xlpp_cfcollection c) { CF(c)->clear(); }
XLPP_API int xlpp_cfcollection_empty(xlpp_cfcollection c) { return CF(c)->empty() ? 1 : 0; }

XLPP_API void xlpp_cfentry_reference(xlpp_cfentry e, char* out, int outSize) { copyStr(CFE(e)->reference(), out, outSize); }
XLPP_API void xlpp_cfentry_set_reference(xlpp_cfentry e, const char* v) { try { CFE(e)->setReference(v); } catch (...) {} }
XLPP_API int xlpp_cfentry_rule_count(xlpp_cfentry e) { return static_cast<int>(CFE(e)->rules().size()); }
XLPP_API xlpp_cfrule xlpp_cfentry_add_rule(xlpp_cfentry e, int type) {
    xlpp::ConditionalRule rule;
    switch (type) {
    case 0: rule = xlpp::ConditionalRule::formula(""); break;
    case 1: rule = xlpp::ConditionalRule::cellIs(xlpp::ConditionalOperator::Equal, ""); break;
    case 2: rule = xlpp::ConditionalRule::dataBar(); break;
    case 3: rule = xlpp::ConditionalRule::colorScale(); break;
    case 4: rule = xlpp::ConditionalRule::iconSet(); break;
    default: rule = xlpp::ConditionalRule::formula(""); break;
    }
    return reinterpret_cast<xlpp_cfrule>(&CFE(e)->addRule(std::move(rule)));
}
XLPP_API xlpp_cfrule xlpp_cfentry_rule_at(xlpp_cfentry e, int index) {
    try { return reinterpret_cast<xlpp_cfrule>(&CFE(e)->rules()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}

XLPP_API void xlpp_cfrule_add_formula(xlpp_cfrule r, const char* f) { CFR(r)->addFormula(f); }
XLPP_API int xlpp_cfrule_formula_count(xlpp_cfrule r) { return static_cast<int>(CFR(r)->formulas().size()); }
XLPP_API void xlpp_cfrule_formula_at(xlpp_cfrule r, int index, char* out, int outSize) {
    try { copyStr(CFR(r)->formulas()[static_cast<std::size_t>(index)], out, outSize); } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}
XLPP_API void xlpp_cfrule_set_operator(xlpp_cfrule r, int op) { CFR(r)->setOperator(static_cast<xlpp::ConditionalOperator>(op)); }
XLPP_API int xlpp_cfrule_operator(xlpp_cfrule r) { return static_cast<int>(CFR(r)->op()); }
XLPP_API int xlpp_cfrule_type(xlpp_cfrule r) { return static_cast<int>(CFR(r)->type()); }
XLPP_API void xlpp_cfrule_set_priority(xlpp_cfrule r, uint64_t v) { CFR(r)->setPriority(static_cast<std::size_t>(v)); }
XLPP_API uint64_t xlpp_cfrule_priority(xlpp_cfrule r) { return CFR(r)->priority(); }
XLPP_API void xlpp_cfrule_set_stop_if_true(xlpp_cfrule r, int v) { CFR(r)->setStopIfTrue(v != 0); }
XLPP_API int xlpp_cfrule_stop_if_true(xlpp_cfrule r) { return CFR(r)->stopIfTrue() ? 1 : 0; }
XLPP_API void xlpp_cfrule_set_differential_style(xlpp_cfrule r, xlpp_style s) { CFR(r)->setDifferentialStyle(*STY(s)); }

// ============================================================
// Data validation
// ============================================================
#define DV(h) reinterpret_cast<xlpp::DataValidationCollection*>(h)
#define DVX(h) reinterpret_cast<xlpp::DataValidation*>(h)
XLPP_API int xlpp_dvcollection_count(xlpp_dvcollection c) { return static_cast<int>(DV(c)->items().size()); }
XLPP_API xlpp_datavalidation xlpp_dvcollection_add(xlpp_dvcollection c, int type, const char* reference) {
    try { return reinterpret_cast<xlpp_datavalidation>(&DV(c)->add(static_cast<xlpp::DataValidationType>(type), reference)); } catch (...) { return nullptr; }
}
XLPP_API xlpp_datavalidation xlpp_dvcollection_at(xlpp_dvcollection c, int index) {
    try { return reinterpret_cast<xlpp_datavalidation>(&DV(c)->items()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API void xlpp_dvcollection_clear(xlpp_dvcollection c) { DV(c)->clear(); }

XLPP_API int xlpp_datavalidation_type(xlpp_datavalidation d) { return static_cast<int>(DVX(d)->type()); }
XLPP_API void xlpp_datavalidation_set_type(xlpp_datavalidation d, int v) { DVX(d)->setType(static_cast<xlpp::DataValidationType>(v)); }
XLPP_API int xlpp_datavalidation_operator(xlpp_datavalidation d) { return static_cast<int>(DVX(d)->op()); }
XLPP_API void xlpp_datavalidation_set_operator(xlpp_datavalidation d, int v) { DVX(d)->setOperator(static_cast<xlpp::DataValidationOperator>(v)); }
XLPP_API int xlpp_datavalidation_error_style(xlpp_datavalidation d) { return static_cast<int>(DVX(d)->errorStyle()); }
XLPP_API void xlpp_datavalidation_set_error_style(xlpp_datavalidation d, int v) { DVX(d)->setErrorStyle(static_cast<xlpp::DataValidationErrorStyle>(v)); }
XLPP_API void xlpp_datavalidation_set_formula1(xlpp_datavalidation d, const char* v) { DVX(d)->setFormula1(v); }
XLPP_API void xlpp_datavalidation_formula1(xlpp_datavalidation d, char* out, int outSize) { copyStr(DVX(d)->formula1(), out, outSize); }
XLPP_API void xlpp_datavalidation_set_formula2(xlpp_datavalidation d, const char* v) { DVX(d)->setFormula2(v); }
XLPP_API void xlpp_datavalidation_formula2(xlpp_datavalidation d, char* out, int outSize) { copyStr(DVX(d)->formula2(), out, outSize); }
XLPP_API void xlpp_datavalidation_set_reference(xlpp_datavalidation d, const char* v) { try { DVX(d)->setReference(v); } catch (...) {} }
XLPP_API void xlpp_datavalidation_reference(xlpp_datavalidation d, char* out, int outSize) { copyStr(DVX(d)->reference(), out, outSize); }
XLPP_API void xlpp_datavalidation_set_allow_blank(xlpp_datavalidation d, int v) { DVX(d)->setAllowBlank(v != 0); }
XLPP_API int xlpp_datavalidation_allow_blank(xlpp_datavalidation d) { return DVX(d)->allowBlank() ? 1 : 0; }
XLPP_API void xlpp_datavalidation_set_show_drop_down(xlpp_datavalidation d, int v) { DVX(d)->setShowDropDown(v != 0); }
XLPP_API int xlpp_datavalidation_show_drop_down(xlpp_datavalidation d) { return DVX(d)->showDropDown() ? 1 : 0; }
XLPP_API void xlpp_datavalidation_set_show_input_message(xlpp_datavalidation d, int v) { DVX(d)->setShowInputMessage(v != 0); }
XLPP_API void xlpp_datavalidation_set_show_error_message(xlpp_datavalidation d, int v) { DVX(d)->setShowErrorMessage(v != 0); }
XLPP_API void xlpp_datavalidation_set_prompt_title(xlpp_datavalidation d, const char* v) { DVX(d)->setPromptTitle(v); }
XLPP_API void xlpp_datavalidation_set_prompt(xlpp_datavalidation d, const char* v) { DVX(d)->setPrompt(v); }
XLPP_API void xlpp_datavalidation_set_error_title(xlpp_datavalidation d, const char* v) { DVX(d)->setErrorTitle(v); }
XLPP_API void xlpp_datavalidation_set_error(xlpp_datavalidation d, const char* v) { DVX(d)->setError(v); }

// ============================================================
// Font
// ============================================================
XLPP_API void xlpp_font_set_name(xlpp_font f, const char* v)       { FONT(f)->setName(v); }
XLPP_API void xlpp_font_set_size(xlpp_font f, double v)            { FONT(f)->setSize(v); }
XLPP_API void xlpp_font_set_bold(xlpp_font f, int v)               { FONT(f)->setBold(v != 0); }
XLPP_API void xlpp_font_set_italic(xlpp_font f, int v)             { FONT(f)->setItalic(v != 0); }
XLPP_API void xlpp_font_set_underline(xlpp_font f, int v)          { FONT(f)->setUnderline(v != 0); }
XLPP_API void xlpp_font_set_strike(xlpp_font f, int v)             { FONT(f)->setStrike(v != 0); }
XLPP_API void xlpp_font_set_color(xlpp_font f, const char* argb)   { FONT(f)->color().setArgb(argb); }
XLPP_API const char* xlpp_font_get_name(xlpp_font f)               { return FONT(f)->name().c_str(); }
XLPP_API double xlpp_font_get_size(xlpp_font f)                    { return FONT(f)->size(); }
XLPP_API int xlpp_font_get_bold(xlpp_font f)                       { return FONT(f)->bold() ? 1 : 0; }
XLPP_API int xlpp_font_get_italic(xlpp_font f)                     { return FONT(f)->italic() ? 1 : 0; }
XLPP_API int xlpp_font_get_underline(xlpp_font f)                  { return FONT(f)->underline() ? 1 : 0; }
XLPP_API int xlpp_font_get_strike(xlpp_font f)                     { return FONT(f)->strike() ? 1 : 0; }
XLPP_API void xlpp_font_get_color(xlpp_font f, char* out, int outSize) { copyStr(FONT(f)->color().argb(), out, outSize); }

// ============================================================
// Fill
// ============================================================
XLPP_API void xlpp_fill_set_pattern(xlpp_fill f, const char* v)     { FILL(f)->setPatternType(v); }
XLPP_API void xlpp_fill_set_fg_color(xlpp_fill f, const char* argb) { FILL(f)->foregroundColor().setArgb(argb); }
XLPP_API void xlpp_fill_set_bg_color(xlpp_fill f, const char* argb) { FILL(f)->backgroundColor().setArgb(argb); }
XLPP_API void xlpp_fill_get_pattern(xlpp_fill f, char* out, int outSize) { copyStr(FILL(f)->patternType(), out, outSize); }
XLPP_API void xlpp_fill_get_fg_color(xlpp_fill f, char* out, int outSize) { copyStr(FILL(f)->foregroundColor().argb(), out, outSize); }
XLPP_API void xlpp_fill_get_bg_color(xlpp_fill f, char* out, int outSize) { copyStr(FILL(f)->backgroundColor().argb(), out, outSize); }

// ============================================================
// Border
// ============================================================
XLPP_API xlpp_borderside xlpp_border_left(xlpp_border b)    { return reinterpret_cast<xlpp_borderside>(&BDR(b)->left()); }
XLPP_API xlpp_borderside xlpp_border_right(xlpp_border b)   { return reinterpret_cast<xlpp_borderside>(&BDR(b)->right()); }
XLPP_API xlpp_borderside xlpp_border_top(xlpp_border b)     { return reinterpret_cast<xlpp_borderside>(&BDR(b)->top()); }
XLPP_API xlpp_borderside xlpp_border_bottom(xlpp_border b)  { return reinterpret_cast<xlpp_borderside>(&BDR(b)->bottom()); }
XLPP_API xlpp_borderside xlpp_border_diagonal(xlpp_border b){ return reinterpret_cast<xlpp_borderside>(&BDR(b)->diagonal()); }
XLPP_API void xlpp_borderside_set_style(xlpp_borderside s, const char* v) { BS(s)->setStyle(v); }
XLPP_API void xlpp_borderside_set_color(xlpp_borderside s, const char* argb) { BS(s)->color().setArgb(argb); }
XLPP_API void xlpp_borderside_get_style(xlpp_borderside s, char* out, int outSize) { copyStr(BS(s)->style(), out, outSize); }
XLPP_API void xlpp_borderside_get_color(xlpp_borderside s, char* out, int outSize) { copyStr(BS(s)->color().argb(), out, outSize); }

// ============================================================
// Alignment
// ============================================================
XLPP_API void xlpp_alignment_set_horizontal(xlpp_alignment a, const char* v) { ALN(a)->setHorizontal(v); }
XLPP_API void xlpp_alignment_set_vertical(xlpp_alignment a, const char* v)   { ALN(a)->setVertical(v); }
XLPP_API void xlpp_alignment_set_wrap_text(xlpp_alignment a, int v)           { ALN(a)->setWrapText(v != 0); }
XLPP_API void xlpp_alignment_set_shrink_to_fit(xlpp_alignment a, int v)       { ALN(a)->setShrinkToFit(v != 0); }
XLPP_API void xlpp_alignment_set_text_rotation(xlpp_alignment a, int v)       { ALN(a)->setTextRotation(v); }
XLPP_API void xlpp_alignment_set_indent(xlpp_alignment a, int v)              { ALN(a)->setIndent(v); }
XLPP_API void xlpp_alignment_get_horizontal(xlpp_alignment a, char* out, int outSize) { copyStr(ALN(a)->horizontal(), out, outSize); }
XLPP_API void xlpp_alignment_get_vertical(xlpp_alignment a, char* out, int outSize) { copyStr(ALN(a)->vertical(), out, outSize); }
XLPP_API int xlpp_alignment_get_wrap_text(xlpp_alignment a) { return ALN(a)->wrapText() ? 1 : 0; }
XLPP_API int xlpp_alignment_get_shrink_to_fit(xlpp_alignment a) { return ALN(a)->shrinkToFit() ? 1 : 0; }
XLPP_API int xlpp_alignment_get_text_rotation(xlpp_alignment a) { return ALN(a)->textRotation(); }
XLPP_API int xlpp_alignment_get_indent(xlpp_alignment a) { return ALN(a)->indent(); }

// ============================================================
// Style
// ============================================================
XLPP_API xlpp_font xlpp_style_font(xlpp_style s)            { return reinterpret_cast<xlpp_font>(&STY(s)->font()); }
XLPP_API xlpp_fill xlpp_style_fill(xlpp_style s)            { return reinterpret_cast<xlpp_fill>(&STY(s)->fill()); }
XLPP_API xlpp_border xlpp_style_border(xlpp_style s)        { return reinterpret_cast<xlpp_border>(&STY(s)->border()); }
XLPP_API xlpp_alignment xlpp_style_alignment(xlpp_style s)   { return reinterpret_cast<xlpp_alignment>(&STY(s)->alignment()); }
XLPP_API void xlpp_style_set_number_format(xlpp_style s, const char* v) { STY(s)->setNumberFormat(v); }
XLPP_API void xlpp_style_number_format(xlpp_style s, char* out, int outSize) { copyStr(STY(s)->numberFormat(), out, outSize); }
XLPP_API void xlpp_style_set_num_fmt_id(xlpp_style s, int v) { STY(s)->setNumFmtId(v); }
XLPP_API int xlpp_style_num_fmt_id(xlpp_style s) { return STY(s)->numFmtId(); }
XLPP_API void xlpp_style_set_locked(xlpp_style s, int v) { STY(s)->setLocked(v != 0); }
XLPP_API int xlpp_style_locked(xlpp_style s) { return STY(s)->locked() ? 1 : 0; }
XLPP_API void xlpp_style_set_hidden(xlpp_style s, int v) { STY(s)->setHidden(v != 0); }
XLPP_API int xlpp_style_hidden(xlpp_style s) { return STY(s)->hidden() ? 1 : 0; }
XLPP_API int xlpp_style_is_default(xlpp_style s) { return STY(s)->isDefault() ? 1 : 0; }

// ============================================================
// Streaming writer
// ============================================================

struct xlpp_stream_writer_t {
    std::unique_ptr<xlpp::StreamingWorkbookWriter> writer;
};

XLPP_API xlpp_stream_writer xlpp_stream_create(const char* path) {
    try {
        auto* handle = new xlpp_stream_writer_t;
        handle->writer = std::make_unique<xlpp::StreamingWorkbookWriter>(
            std::filesystem::path(path), xlpp::SharedStringMode::Disabled);
        return handle;
    } catch (...) { return nullptr; }
}

XLPP_API void xlpp_stream_destroy(xlpp_stream_writer w) {
    delete w;
}

XLPP_API uint64_t xlpp_stream_add_sheet(xlpp_stream_writer w, const char* name) {
    try {
        w->writer->addWorksheet(name);
        return static_cast<uint64_t>(w->writer->sheetCount() - 1);
    } catch (...) { return static_cast<uint64_t>(-1); }
}

XLPP_API void xlpp_stream_append_row(xlpp_stream_writer w, uint64_t sheetIndex, const char** values, int count) {
    try {
        auto& sheet = w->writer->worksheet(static_cast<std::size_t>(sheetIndex));
        std::vector<xlpp::CellValue> row;
        row.reserve(static_cast<std::size_t>(count));
        for (int i = 0; i < count; ++i) {
            if (values[i] && values[i][0]) row.push_back(std::string(values[i]));
            else row.push_back(xlpp::CellValue{});
        }
        sheet.append(row);
    } catch (...) {}
}

XLPP_API void xlpp_stream_append_doubles(xlpp_stream_writer w, uint64_t sheetIndex, const double* values, int count) {
    try {
        auto& sheet = w->writer->worksheet(static_cast<std::size_t>(sheetIndex));
        std::vector<xlpp::CellValue> row;
        row.reserve(static_cast<std::size_t>(count));
        for (int i = 0; i < count; ++i) row.push_back(values[i]);
        sheet.append(row);
    } catch (...) {}
}

XLPP_API uint64_t xlpp_stream_row_count(xlpp_stream_writer w, uint64_t sheetIndex) {
    try { return static_cast<uint64_t>(w->writer->worksheet(static_cast<std::size_t>(sheetIndex)).rowCount()); }
    catch (...) { return 0; }
}

XLPP_API uint64_t xlpp_stream_sheet_count(xlpp_stream_writer w) {
    try { return static_cast<uint64_t>(w->writer->sheetCount()); } catch (...) { return 0; }
}

XLPP_API void xlpp_stream_set_date1904(xlpp_stream_writer w, int v) {
    try { w->writer->setDate1904(v != 0); } catch (...) {}
}

XLPP_API void xlpp_stream_set_compression_level(xlpp_stream_writer w, int level) {
    try { w->writer->setCompressionLevel(static_cast<xlpp::CompressionLevel>(level)); } catch (...) {}
}

XLPP_API void xlpp_stream_set_parallel_workers(xlpp_stream_writer w, uint64_t workers) {
    try { w->writer->setParallelWorkers(static_cast<std::size_t>(workers)); } catch (...) {}
}

XLPP_API void xlpp_stream_close(xlpp_stream_writer w) {
    try { if (w && w->writer) w->writer->close(); } catch (...) {}
}

// ============================================================
// Streaming reader
// ============================================================

struct xlpp_stream_reader_t {
    std::unique_ptr<xlpp::StreamingWorkbookReader> reader;
};

XLPP_API xlpp_stream_reader xlpp_stream_reader_open(const char* path) {
    try {
        auto* handle = new xlpp_stream_reader_t;
        handle->reader = std::make_unique<xlpp::StreamingWorkbookReader>(std::filesystem::path(path));
        return handle;
    } catch (...) { return nullptr; }
}

XLPP_API void xlpp_stream_reader_destroy(xlpp_stream_reader r) {
    delete r;
}

XLPP_API int xlpp_stream_reader_sheet_count(xlpp_stream_reader r) {
    try { return static_cast<int>(r->reader->worksheetNames().size()); } catch (...) { return 0; }
}

XLPP_API void xlpp_stream_reader_sheet_name(xlpp_stream_reader r, int index, char* out, int outSize) {
    try {
        const auto& names = r->reader->worksheetNames();
        if (index >= 0 && static_cast<std::size_t>(index) < names.size())
            copyStr(names[static_cast<std::size_t>(index)], out, outSize);
        else if (out && outSize > 0) out[0] = '\0';
    } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}

XLPP_API int xlpp_stream_reader_read_sheet(xlpp_stream_reader r, int index,
                                           xlpp_stream_row_callback callback, void* user) {
    try {
        const auto& names = r->reader->worksheetNames();
        if (index < 0 || static_cast<std::size_t>(index) >= names.size()) return 0;
        const auto& name = names[static_cast<std::size_t>(index)];
        bool keepGoing = true;
        r->reader->forEachRow(name, [&](std::size_t rowNumber, const xlpp::StreamingRow& row) {
            if (!keepGoing) return false;
            std::vector<const char*> addresses, strings, formulas;
            std::vector<double> numbers;
            std::vector<int> valueTypes, styleIndexes;
            addresses.reserve(row.size());
            strings.reserve(row.size());
            formulas.reserve(row.size());
            numbers.reserve(row.size());
            valueTypes.reserve(row.size());
            styleIndexes.reserve(row.size());
            for (const auto& cell : row) {
                addresses.push_back(cell.address.c_str());
                int vt = XLPP_VALUE_EMPTY;
                if (std::holds_alternative<bool>(cell.value)) vt = XLPP_VALUE_BOOL;
                else if (std::holds_alternative<double>(cell.value)) vt = XLPP_VALUE_NUMBER;
                else if (std::holds_alternative<std::string>(cell.value)) vt = XLPP_VALUE_STRING;
                else if (std::holds_alternative<xlpp::CellError>(cell.value)) vt = XLPP_VALUE_ERROR;
                else if (std::holds_alternative<xlpp::DateTime>(cell.value)) vt = XLPP_VALUE_DATE;
                valueTypes.push_back(vt);
                if (auto* v = std::get_if<double>(&cell.value)) numbers.push_back(*v);
                else if (auto* v = std::get_if<xlpp::DateTime>(&cell.value)) numbers.push_back(xlpp::toExcelSerial(*v, false));
                else numbers.push_back(0.0);
                if (auto* v = std::get_if<std::string>(&cell.value)) strings.push_back(v->c_str());
                else strings.push_back("");
                formulas.push_back(cell.formula.c_str());
                styleIndexes.push_back(cell.styleIndex ? static_cast<int>(*cell.styleIndex) : -1);
            }
            keepGoing = callback(user, rowNumber, static_cast<int>(row.size()),
                                 addresses.data(), numbers.data(), valueTypes.data(),
                                 strings.data(), formulas.data(), styleIndexes.data()) != 0;
            return keepGoing;
        });
        return 1;
    } catch (...) { return 0; }
}

XLPP_API void xlpp_free_string(const char* str) {
    (void)str;
}

XLPP_API const char* xlpp_last_error(void) {
    return g_lastError.c_str();
}

XLPP_API void xlpp_clear_error(void) {
    clearError();
}

} // extern "C"
