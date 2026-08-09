# VBA, encryption, and enterprise workbooks

This chapter covers workbook features that require special preservation or
security care.

## VBA projects

### Preserve an existing project

```python
book = xlpp.Workbook.open("template.xlsm")
print(book.has_vba_project)
print(book.has_vba_signature)
print(book.vba_source_editable)

book["Input"]["A1"].value = "Updated"
book.save("result.xlsm")
```

Saving with an `.xlsm` filename does not create VBA by itself. The workbook
must contain a project or one must be attached.

### Read and edit modules

```python
for module in book.vba_modules:
    print(module.name, module.type, module.read_only)

source = book.vba_module_text("Module1")
book.set_vba_module_text(
    "Module1",
    "Public Sub Run()\n    MsgBox \"Hello\"\nEnd Sub\n",
)
```

Create specific module kinds:

```python
book.set_vba_class_module_text("ReportClass", "Option Explicit\n")
book.set_vba_document_module_text("ThisWorkbook", "Option Explicit\n")
book.remove_vba_module("OldModule")
```

### Binary project access

```python
with open("vbaProject.bin", "rb") as stream:
    book.set_vba_project(stream.read())

payload = book.vba_project_bytes
book.save_vba_project("extracted-vbaProject.bin")
```

Project properties are available through `vba_project_properties`.

!!! danger "Signatures"

    Editing VBA content generally invalidates a digital signature. Inspect
    `has_vba_signature` and design an explicit re-signing process outside XL++.

!!! note "UserForms and ActiveX"

    Binary project preservation and module authoring do not imply complete
    semantic authoring for UserForms or ActiveX controls. Those assets are
    inventoried and preserved when possible.

## Encryption

Worksheet/workbook protection is not encryption. Password-to-open encryption
wraps the OOXML package and protects the file content.

### Inspect encryption

```python
info = xlpp.inspect_office_encryption("input.xlsx")
print(info.encrypted)
print(info.supported)
print(info.mode)
print(info.cipher_algorithm, info.hash_algorithm)
print(info.key_bits, info.spin_count)
```

### Open an encrypted workbook

```python
options = xlpp.LoadOptions()
options.password = password_from_secret_store
options.verify_encryption_integrity = True

book = xlpp.Workbook.open("encrypted.xlsx", options)
```

### Save with Agile AES-256/SHA-512

```python
options = xlpp.SaveOptions()
options.encryption_password = password_from_secret_store
options.encryption_mode = xlpp.OfficeEncryptionMode.AGILE_AES256_SHA512
options.encryption_spin_count = 100_000
options.encryption_key_bits = 256

book.save("encrypted.xlsx", options)
```

Standard AES/SHA-1 mode is also exposed for compatibility:

```python
options.encryption_mode = xlpp.OfficeEncryptionMode.STANDARD_AES_SHA1
options.encryption_key_bits = 256
```

!!! tip "Password handling"

    Do not hard-code passwords or place them in command-line arguments. Load
    them from a secret manager, keep them in memory for the shortest practical
    time, and never log `LoadOptions`/`SaveOptions` indiscriminately.

## External data inspection

```python
external = book.inspect_external_data()

print(external.has_external_workbooks)
print(external.has_connections)
print(external.has_query_tables)
print(external.warnings)
```

Inspection reports expose connection/query/external-link package metadata.
XL++ does not execute external database queries.

## Data Model and OLAP

```python
model = book.inspect_data_model()
print(model.present)
print(model.has_olap_pivot_caches)
print(model.model_parts)
print(model.olap_pivot_cache_parts)
```

Use this inventory before structural edits or package transformations on
workbooks backed by Power Pivot or OLAP.

## Enterprise feature inventory

```python
inspection = book.inspect_enterprise_features()

for feature in inspection.features:
    print(
        feature.kind,
        feature.name,
        feature.part_name,
        feature.semantic_editable,
        feature.refresh_on_load if feature.has_refresh_on_load else None,
    )
```

Feature kinds include:

- PivotChart
- Slicer and slicer cache
- Timeline and timeline cache
- OLAP pivot cache
- Data Model
- Power Query
- SmartArt
- ActiveX
- VBA UserForm

### Targeted metadata edits

```python
report = book.set_connection_refresh_on_load("1", True)
print(report.success, report.modified, report.warnings)

book.set_query_table_refresh_on_load("SalesQuery", True)
book.set_olap_pivot_cache_refresh_on_load(
    "xl/pivotCache/pivotCacheDefinition1.xml",
    True,
)
book.set_pivot_chart_source_name(
    "xl/charts/chart1.xml",
    "SalesPivot",
)
```

These APIs intentionally target small metadata fields. They do not turn XL++
into a Power Query, VertiPaq, OLAP, SmartArt, or controls authoring engine.

## Safe enterprise round-trip workflow

1. Keep the original workbook as an immutable input artifact.
2. Inspect external data, Data Model, enterprise features, VBA, and encryption.
3. Make the smallest semantic edit required.
4. Save to a new path with atomic write and validation enabled.
5. Compare package parts and relationships.
6. Open and refresh the result in the target Excel environment.
7. Re-sign VBA where required.

See [Preservation Core](../PRESERVATION_CORE.md) for package-level guarantees
and explicit limitations.
