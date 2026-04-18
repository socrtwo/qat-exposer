#Requires -Version 5.0
<#
.SYNOPSIS
    Builds SuperQAT.xlam for Excel from VBA source files.
.DESCRIPTION
    Opens Excel via COM automation, creates a new workbook,
    imports the three VBA files, saves as .xlam, and closes Excel.
.NOTES
    Run from the com-addin folder: .\build-excel.ps1
    Requires Microsoft Excel installed on this machine.
#>

$ErrorActionPreference = "Stop"

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir     = Join-Path $scriptDir "excel"
$outDir     = Join-Path $scriptDir "build"
$outFile    = Join-Path $outDir "SuperQAT.xlam"

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

Write-Host "Starting Excel..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Creating new workbook..." -ForegroundColor Cyan
    $wb = $excel.Workbooks.Add()

    $vbProj = $wb.VBProject

    Write-Host "Importing SuperQAT.bas..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($basFile) | Out-Null

    Write-Host "Importing SuperQATData.bas..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($dataFile) | Out-Null

    Write-Host "Importing SuperQATForm.frm..." -ForegroundColor Cyan
    $vbProj.VBComponents.Import($frmFile) | Out-Null

    # Save as Excel Add-In (.xlam = xlAddIn = 55)
    if (Test-Path $outFile) { Remove-Item $outFile -Force }
    Write-Host "Saving $outFile..." -ForegroundColor Cyan
    $wb.SaveAs($outFile, 55)
    $wb.Close($false)

    Write-Host ""
    Write-Host "SUCCESS: Built $outFile" -ForegroundColor Green
    Write-Host ""
    Write-Host "To install:" -ForegroundColor Yellow
    $addinsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
    Write-Host "  1. Copy to: $addinsDir" -ForegroundColor Yellow
    Write-Host "  2. Excel > File > Options > Add-Ins > Manage: Excel Add-ins > Go" -ForegroundColor Yellow
    Write-Host "  3. Check 'SuperQAT' and click OK" -ForegroundColor Yellow
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "If you see 'programmatic access' error, do this:" -ForegroundColor Yellow
    Write-Host "  Excel > File > Options > Trust Center > Trust Center Settings" -ForegroundColor Yellow
    Write-Host "  > Macro Settings > check 'Trust access to the VBA project object model'" -ForegroundColor Yellow
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel closed." -ForegroundColor Cyan
}
