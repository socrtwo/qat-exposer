#Requires -Version 5.0
<#
.SYNOPSIS
    Builds SuperQAT.dotm for Word from VBA source files.
.DESCRIPTION
    Opens Word via COM automation, creates a new macro-enabled template,
    imports the three VBA files, saves as .dotm, and closes Word.
.NOTES
    Run from the com-addin folder: .\build-word.ps1
    Requires Microsoft Word installed on this machine.
#>

$ErrorActionPreference = "Stop"

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir     = Join-Path $scriptDir "word"
$outDir     = Join-Path $scriptDir "build"
$outFile    = Join-Path $outDir "SuperQAT.dotm"

$basFile    = Join-Path $srcDir "SuperQAT.bas"
$dataFile   = Join-Path $srcDir "SuperQATData.bas"
$frmFile    = Join-Path $srcDir "SuperQATForm.frm"

# Verify source files exist
foreach ($f in @($basFile, $dataFile, $frmFile)) {
    if (-not (Test-Path $f)) {
        Write-Error "Missing file: $f"
        exit 1
    }
}

# Create output directory
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

Write-Host "Starting Word..." -ForegroundColor Cyan
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

try {
    Write-Host "Creating new template..." -ForegroundColor Cyan
    $doc = $word.Documents.Add()

    # Access VBA project (requires "Trust access to the VBA project object model"
    # in Word Options > Trust Center > Trust Center Settings > Macro Settings)
    $vbProj = $doc.VBProject

    Write-Host "Importing SuperQAT.bas..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($basFile) | Out-Null

    Write-Host "Importing SuperQATData.bas..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($dataFile) | Out-Null

    Write-Host "Importing SuperQATForm.frm..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($frmFile) | Out-Null

    # Remove the default empty "ThisDocument" module content
    # (it's fine to leave as-is)

    # Save as macro-enabled template (.dotm = wdFormatXMLTemplateMacroEnabled = 13)
    if (Test-Path $outFile) { Remove-Item $outFile -Force }
    Write-Host "Saving $outFile..." -ForegroundColor Cyan
    $doc.SaveAs2([ref]$outFile, [ref]13)
    $doc.Close(0)  # wdDoNotSaveChanges

    Write-Host ""
    Write-Host "SUCCESS: Built $outFile" -ForegroundColor Green
    Write-Host ""
    Write-Host "To install, copy it to:" -ForegroundColor Yellow
    $startup = Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
    Write-Host "  $startup" -ForegroundColor Yellow
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "If you see 'programmatic access' error, do this:" -ForegroundColor Yellow
    Write-Host "  Word > File > Options > Trust Center > Trust Center Settings" -ForegroundColor Yellow
    Write-Host "  > Macro Settings > check 'Trust access to the VBA project object model'" -ForegroundColor Yellow
}
finally {
    $word.Quit(0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Host "Word closed." -ForegroundColor Cyan
}
