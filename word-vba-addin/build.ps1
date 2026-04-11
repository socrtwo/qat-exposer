# build.ps1 — Assembles SuperQAT.dotm from QATModule.bas
# Usage: powershell -ExecutionPolicy Bypass -File build.ps1
#
# Prerequisites: Microsoft Word must be installed.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$basFile   = Join-Path $scriptDir "QATModule.bas"
$outFile   = Join-Path $scriptDir "SuperQAT.dotm"

# Remove old output if it exists
if (Test-Path $outFile) {
    Remove-Item $outFile -Force
}

if (-not (Test-Path $basFile)) {
    Write-Error "QATModule.bas not found in $scriptDir"
    exit 1
}

Write-Host "Starting Word..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false

try {
    # Suppress all alerts
    $word.DisplayAlerts = 0

    Write-Host "Creating new document..."
    $doc = $word.Documents.Add()

    # Import VBA module
    Write-Host "Importing VBA module..."
    try {
        $doc.VBProject.VBComponents.Import($basFile) | Out-Null
    }
    catch {
        Write-Host ""
        Write-Host "ERROR: Could not import VBA code." -ForegroundColor Red
        Write-Host "Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings"
        Write-Host "Check: 'Trust access to the VBA project object model'"
        throw
    }

    # Save as .dotm (Word Macro-Enabled Template)
    # wdFormatXMLTemplateMacroEnabled = 13
    Write-Host "Saving as $outFile ..."
    $doc.SaveAs([ref]$outFile, [ref]13)
    $doc.Close([ref]0)

    Write-Host ""
    Write-Host "SUCCESS: $outFile created!" -ForegroundColor Green
    Write-Host ""
    Write-Host "To install:"
    Write-Host "  1. Double-click SuperQAT.dotm, or"
    Write-Host "  2. Copy it to: $($env:APPDATA)\Microsoft\Word\STARTUP\"
}
finally {
    $word.Quit([ref]0)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Host "Done."
}
