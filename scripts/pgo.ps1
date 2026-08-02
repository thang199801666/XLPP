# PGO (Profile-Guided Optimization) build script for XL++
# Usage: Run from the project root directory.
# Step 1: .\scripts\pgo.ps1 instrument  — builds instrumented binaries
# Step 2: .\scripts\pgo.ps1 train      — runs benchmarks to collect profile data
# Step 3: .\scripts\pgo.ps1 optimize   — rebuilds with profile data, creating optimized binary

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("instrument", "train", "optimize", "all")]
    [string]$Action
)

$ErrorActionPreference = "Stop"
$Solution = "XL++.sln"
$MSBuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$OutDir = "x64\Release"
$PgData = "x64\PGO"
$TestExe = "$OutDir\XLPP.UnitTests.exe"
$SampleExe = "$OutDir\XLPP.Sample.exe"

if (-not (Test-Path $MSBuild)) {
    Write-Error "MSBuild not found at $MSBuild"
    exit 1
}

function Invoke-MSBuild {
    param([string[]]$Args)
    $cmd = "& `"$MSBuild`" `"$Solution`" /p:Configuration=Release /p:Platform=x64 /m $($Args -join ' ')"
    Write-Host "> $cmd" -ForegroundColor Cyan
    Invoke-Expression $cmd
    if ($LASTEXITCODE -ne 0) { throw "MSBuild failed" }
}

switch ($Action) {
    "instrument" {
        Write-Host "=== PGO Phase 1: Build Instrumented ===" -ForegroundColor Yellow
        New-Item -ItemType Directory -Force -Path $PgData | Out-Null
        Invoke-MSBuild "/p:WholeProgramOptimization=true" "/p:GeneratePGOData=true" "/p:PGODatabaseDirectory=$(Resolve-Path $PgData)"
        Write-Host "Instrumented build complete. Binaries in $OutDir" -ForegroundColor Green
    }
    "train" {
        Write-Host "=== PGO Phase 2: Train ===" -ForegroundColor Yellow
        Write-Host "Running unit tests to collect profile data..."
        Push-Location $OutDir
        try {
            & ".\XLPP.UnitTests.exe" 2>&1 | Out-Null
            Write-Host "Tests exited with code $LASTEXITCODE"
            Write-Host "Running sample..."
            & ".\XLPP.Sample.exe" 2>&1 | Out-Null
        } finally {
            Pop-Location
        }
        Write-Host "Training complete. Profile data in $PgData" -ForegroundColor Green
    }
    "optimize" {
        Write-Host "=== PGO Phase 3: Build Optimized ===" -ForegroundColor Yellow
        Invoke-MSBuild "/p:WholeProgramOptimization=true" "/p:UsePGOData=true" "/p:PGODatabaseDirectory=$(Resolve-Path $PgData)" "/t:Rebuild"
        Write-Host "PGO-optimized build complete. Binaries in $OutDir" -ForegroundColor Green
    }
    "all" {
        & $PSCommandPath instrument
        & $PSCommandPath train
        & $PSCommandPath optimize
    }
}
