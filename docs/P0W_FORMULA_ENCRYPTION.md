# P0W — Formula Calculation and OOXML Encryption

> **P0Y supersession note (2026-08-08):** this file preserves the historical P0W acceptance baseline. P0X/v1.2.0 extended Formula from 21/26 to **26/26 scope-defined core capability families** and Encryption from 18/21 to **19/21**; P0Y/v1.3.0 keeps that feature scope and hardens the underlying I/O, ZIP/ZIP64, topology, validation and sanitizer guarantees. See `P0X_90_PERCENT_ENGINE.md` for the feature matrix and `P0Y_CORE_HARDENING.md` for the current reliability baseline.

P0W closes two of the largest functional gaps left after P0V-A: in-process
formula calculation and password-to-open OOXML encryption. The scope in this
milestone is the C++ core. Binding parity for Python/C#/C remains follow-up
work.

## Formula calculation engine

`Workbook::calculateFormulas()` evaluates formula cells and, by default,
updates the cached values stored next to the formula expressions. Formula text
is never rewritten by calculation.

```cpp
xlpp::CalculationOptions options;
options.recursiveDependencies = true;
options.updateCachedValues = true;
options.evaluateVolatileFunctions = true;
options.maxDepth = 512;

auto report = workbook.calculateFormulas(options);
```

The engine supports:

- numeric, string, Boolean and Excel error literals plus two-dimensional array constants;
- arithmetic precedence, unary operators, power, percent, concatenation and comparisons;
- A1 references, rectangular ranges, absolute/mixed markers, local/cross-sheet references and quoted sheet names;
- recursive formula dependencies, workbook/local defined names, maximum-depth protection and circular-reference detection;
- cached-value updates plus diagnostic/no-mutation calculation mode;
- `_xlfn.` / `_xlws.` compatibility prefixes and comma/semicolon argument separators;
- aggregate/statistical functions (`SUM`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `COUNT`, `COUNTA`, `COUNTBLANK`, `MEDIAN`, `STDEV*`, `VAR*`, `SMALL`, `LARGE`, `SUMPRODUCT`, `SUMSQ`);
- logical/error functions (`IF`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`, `AND`, `OR`, `XOR`, `NOT`, `IS*`, `NA`, `TRUE`, `FALSE`);
- common math/trigonometric functions (`ABS`, `SQRT`, `POWER`, `MOD`, rounding, logarithms, `EXP`, `PI`, `SIN/COS/TAN`, inverse trig, `ATAN2`, radians/degrees, ceiling/floor);
- text functions (`LEN`, `LOWER`, `UPPER`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `CONCAT`, `TEXTJOIN`, `FIND`, `SEARCH`, `SUBSTITUTE`, `EXACT`, `VALUE`, `REPT`);
- date/time functions (`DATE`, `YEAR`, `MONTH`, `DAY`, `TODAY`, `NOW`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `DAYS`);
- criteria/wildcard and multi-criteria functions (`COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MINIFS`, `MAXIFS`);
- lookup/reference functions (`INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `CHOOSE`);
- financial core (`PMT`, `FV`, `PV`, `NPV`, `IRR`).

`SaveOptions::calculateFormulasBeforeSave` calculates on a private workbook
copy. The caller remains unchanged. If chart-cache synchronization is also
enabled, formulas are calculated first and chart caches see the calculated
values.

### Formula core capability score

The P0W acceptance matrix is deliberately based on capability families rather
than counting individual Excel function names. Specialized Excel domains
(financial/security/cube/database/engineering functions) are not each counted
as a separate item.

| Capability family | Status |
|---|---|
| Formula tokenizer and scalar/error literals | Complete |
| Operator precedence, unary/power/percent | Complete |
| Comparison and text concatenation | Complete |
| A1 cell and rectangular range references | Complete |
| Cross-sheet and quoted-sheet references | Complete |
| Recursive formula dependencies | Complete |
| Cycle/max-depth protection | Complete |
| Workbook/local defined-name resolution | Complete |
| Cached-result update and diagnostic mode | Complete |
| Calculate-before-save pipeline | Complete |
| Aggregate/statistical core | Complete |
| Logical/error core | Complete |
| Math/trigonometric core | Complete |
| Text core | Complete |
| Date/time core | Complete |
| Criteria/wildcard core | Complete |
| Multi-criteria aggregation | Complete |
| Lookup/reference core | Complete |
| Financial core | Complete |
| Compatibility prefixes and argument separators | Complete |
| Array constants | Complete |
| Dynamic arrays/spill | Missing |
| Structured table references | Missing |
| External-workbook references | Missing |
| Reference-returning `INDIRECT`/`OFFSET` family | Missing |
| Iterative circular calculation | Missing |

**P0W formula-core score: 21/26 = 80.8%.** The implementation intentionally
continues beyond the original arithmetic/lookup target by covering financial
and trigonometric families as well. This score is not a claim that 80% of every
function in the Microsoft Excel catalog is implemented.

## Password-to-open encryption

Saving with `SaveOptions::encryptionPassword` first creates the ordinary OOXML
ZIP package and then wraps it in an OLE/CFB encrypted container. P0W writes
Agile AES-256-CBC/SHA-512 and reads both that profile and Standard AES/SHA-1
(128/192/256-bit key profiles).

```cpp
xlpp::SaveOptions save;
save.encryptionPassword = "correct horse battery staple";
save.encryptionSpinCount = 100000;
workbook.save("secure.xlsx", save);

xlpp::LoadOptions load;
load.password = "correct horse battery staple";
load.verifyEncryptionIntegrity = true;
workbook.load("secure.xlsx", load);
```

The public inspector can distinguish encrypted OOXML without a password:

```cpp
auto info = xlpp::inspectOfficeEncryption("secure.xlsx");
info.mode;
info.keyBits;
info.cipherAlgorithm;
info.hashAlgorithm;
info.spinCount;
```

Implemented security/container behavior includes:

- OLE/CFB reader and writer, mini-stream/miniFAT, FAT and DIFAT;
- DataSpaces / `StrongEncryptionTransform` streams required by Office encrypted packages;
- Agile AES-256-CBC/SHA-512 password verifier and secret-key derivation;
- 4096-byte Agile package segmentation;
- Agile HMAC generation and constant-time integrity verification;
- Standard AES-128/192/256 + SHA-1 password verification/decryption;
- OS CSPRNG use (`BCryptGenRandom`, `getrandom`, or `arc4random_buf` depending on platform);
- constant-time password verifier/HMAC comparison;
- spin-count, package-size, sector-geometry, allocation-table and cyclic-chain hardening;
- path and stream workbook load/save integration;
- password rotation (load with old password, save with a new password);
- encryption removal (load encrypted, save without `encryptionPassword`);
- formula-calculation-before-encrypted-save composition;
- large CFB/DIFAT regression (>9 MiB plaintext package).

### Encryption capability score

The matrix below measures modern password-to-open OOXML lifecycle support. It
does not count legacy `.xls` XOR/RC4 encryption, because XL++ is an OOXML
`.xlsx/.xlsm` library.

| Capability | Status |
|---|---|
| Detect CFB/encrypted OOXML | Complete |
| Inspect `EncryptionInfo` without password | Complete |
| CFB mini-stream/FAT/DIFAT read | Complete |
| CFB mini-stream/FAT/DIFAT write | Complete |
| Required DataSpaces transform metadata | Complete |
| Agile password verification/key derivation | Complete |
| Agile AES-256/SHA-512 decrypt | Complete |
| Agile AES-256/SHA-512 encrypt | Complete |
| Agile 4096-byte package segmentation | Complete |
| Agile HMAC generation/verification | Complete |
| Standard AES-128/SHA-1 decrypt | Complete |
| Standard AES-192/256/SHA-1 decrypt path | Complete |
| Standard password verification | Complete |
| Workbook path + stream integration | Complete |
| Password rotate/remove lifecycle | Complete |
| CSPRNG/constant-time/parser hardening | Complete |
| Independent host interoperability | Complete |
| Large-file CFB/DIFAT path | Complete |
| Standard Encryption writer | Missing |
| Alternate Agile hash/cipher profiles | Missing |
| Certificate/private-key Agile key encryptors | Missing |

**P0W modern-OOXML encryption score: 18/21 = 85.7%.**

Independent interoperability checks used in this development batch:

1. LibreOffice opened an XL++-generated Agile AES-256/SHA-512 workbook with the
   XL++ password and read the expected cells.
2. XL++ opened an independently generated LibreOffice Standard AES-128/SHA-1
   encrypted workbook and read the expected cells.

## Regression state

At P0W closeout the full native unit suite contains **160 suites / 2,838
checks**, all passing. Encryption-specific regression includes wrong/missing
passwords, Unicode password round-trip, stream APIs, HMAC tamper rejection,
password rotation/removal, Standard-encryption fixture load, formula+encryption
save composition, and a >9 MiB DIFAT round-trip.

## Remaining high-value work

Formula: dynamic arrays/spill, structured references, external references,
reference-returning functions, iterative circular calculation, and broader
specialized-function coverage.

Encryption: Standard writer, alternate Agile profiles, certificate key
providers, Microsoft Excel Desktop fixture/automation validation, and binding
parity.
