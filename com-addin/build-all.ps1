#Requires -Version 5.0
<#
.SYNOPSIS
    Builds all three SuperQAT COM add-ins (.dotm, .xlam, .ppam).
.DESCRIPTION
    Runs build-word.ps1, build-excel.ps1, and build-powerpoint.ps1 in sequence.
    Output goes to the build/ folder.
.NOTES
    Run from the com-addin folder: .\build-all.ps1
    Requires Word, Excel, and PowerPoint installed on this machine.

    PREREQUISITE (one-time setup):
    In each Office app, go to:
      File > Options > Trust Center > Trust Center Settings
      > Macro Settings > check "Trust access to the VBA project object model"
#>

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "============================================" -ForegroundColor White
Write-Host " SuperQAT v3.0.0 - COM Add-in Builder" -ForegroundColor White
Write-Host "============================================" -ForegroundColor White
Write-Host ""

$failed = @()

# Build Word
Write-Host "--- Building Word add-in (.dotm) ---" -ForegroundColor White
try {
    & "$scriptDir\build-word.ps1"
} catch {
    Write-Host "Word build failed: $($_.Exception.Message)" -ForegroundColor Red
    $failed += "Word"
}
Write-Host ""

# Build Excel
Write-Host "--- Building Excel add-in (.xlam) ---" -ForegroundColor White
try {
    & "$scriptDir\build-excel.ps1"
} catch {
    Write-Host "Excel build failed: $($_.Exception.Message)" -ForegroundColor Red
    $failed += "Excel"
}
Write-Host ""

# Build PowerPoint
Write-Host "--- Building PowerPoint add-in (.ppam) ---" -ForegroundColor White
try {
    & "$scriptDir\build-powerpoint.ps1"
} catch {
    Write-Host "PowerPoint build failed: $($_.Exception.Message)" -ForegroundColor Red
    $failed += "PowerPoint"
}
Write-Host ""

# Summary
Write-Host "============================================" -ForegroundColor White
$buildDir = Join-Path $scriptDir "build"
if ($failed.Count -eq 0) {
    Write-Host " ALL BUILDS SUCCEEDED" -ForegroundColor Green
    Write-Host ""
    Write-Host " Output files:" -ForegroundColor White
    Get-ChildItem $buildDir -File | ForEach-Object {
        $size = [math]::Round($_.Length / 1024)
        Write-Host "   $($_.Name)  ($size KB)" -ForegroundColor Cyan
    }
} else {
    Write-Host " SOME BUILDS FAILED: $($failed -join ', ')" -ForegroundColor Red
    Write-Host ""
    if (Test-Path $buildDir) {
        Write-Host " Successful builds:" -ForegroundColor White
        Get-ChildItem $buildDir -File | ForEach-Object {
            Write-Host "   $($_.Name)" -ForegroundColor Cyan
        }
    }
}
Write-Host "============================================" -ForegroundColor White
