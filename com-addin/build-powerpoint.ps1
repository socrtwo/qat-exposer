#Requires -Version 5.0
<#
.SYNOPSIS
    Builds SuperQAT.ppam for PowerPoint from VBA source files.
.DESCRIPTION
    Opens PowerPoint via COM automation, creates a new presentation,
    imports the three VBA files, saves as .ppam, and closes PowerPoint.
.NOTES
    Run from the com-addin folder: .\build-powerpoint.ps1
    Requires Microsoft PowerPoint installed on this machine.
#>

$ErrorActionPreference = "Stop"

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir     = Join-Path $scriptDir "powerpoint"
$outDir     = Join-Path $scriptDir "build"
$outFile    = Join-Path $outDir "SuperQAT.ppam"

$basFile    = Join-Path $srcDir "SuperQAT.bas"
$dataFile   = Join-Path $srcDir "SuperQATData.bas"
$frmFile    = Join-Path $srcDir "SuperQATForm.frm"

foreach ($f in @($basFile, $dataFile, $frmFile)) {
    if (-not (Test-Path $f)) {
        Write-Error "Missing file: $f"
        exit 1
    }
}

if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

Write-Host "Starting PowerPoint..." -ForegroundColor Cyan
$ppt = New-Object -ComObject PowerPoint.Application

try {
    Write-Host "Creating new presentation..." -ForegroundColor Cyan
    # PowerPoint needs to be visible to add presentations in some versions
    $pres = $ppt.Presentations.Add($true)  # WithWindow = True

    $vbProj = $pres.VBProject

    Write-Host "Importing SuperQAT.bas..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($basFile) | Out-Null

    Write-Host "Importing SuperQATData.bas..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($dataFile) | Out-Null

    Write-Host "Importing SuperQATForm.frm..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($frmFile) | Out-Null

    # Save as PowerPoint Add-In (.ppam = ppSaveAsOpenXMLAddin = 30)
    if (Test-Path $outFile) { Remove-Item $outFile -Force }
    Write-Host "Saving $outFile..." -ForegroundColor Cyan
    $pres.SaveAs($outFile, 30)
    $pres.Close()

    Write-Host ""
    Write-Host "SUCCESS: Built $outFile" -ForegroundColor Green
    Write-Host ""
    Write-Host "To install:" -ForegroundColor Yellow
    $addinsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
    Write-Host "  1. Copy to: $addinsDir" -ForegroundColor Yellow
    Write-Host "  2. PowerPoint > File > Options > Add-Ins > Manage: PowerPoint Add-ins > Go" -ForegroundColor Yellow
    Write-Host "  3. Check 'SuperQAT' and click OK" -ForegroundColor Yellow
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "If you see 'programmatic access' error, do this:" -ForegroundColor Yellow
    Write-Host "  PowerPoint > File > Options > Trust Center > Trust Center Settings" -ForegroundColor Yellow
    Write-Host "  > Macro Settings > check 'Trust access to the VBA project object model'" -ForegroundColor Yellow
}
finally {
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    Write-Host "PowerPoint closed." -ForegroundColor Cyan
}
