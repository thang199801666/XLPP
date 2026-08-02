# Embedded zlib source

Run `tools\Fetch-Zlib.ps1` from the repository root. The script downloads the official zlib 1.3.2 source, verifies SHA-256, and copies the source files into this folder.

The source files are intentionally not duplicated in the fixed-files patch ZIP. After running the script, they are built as the `zlibstatic` project inside `Pdf++.sln`.
