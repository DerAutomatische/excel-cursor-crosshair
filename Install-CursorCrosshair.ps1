# Install-CursorCrosshair.ps1
# Copies CursorCrosshair.xlam to Excel's XLSTART folder so it auto-loads
# in every Excel session.

[CmdletBinding()]
param(
    [string]$AddinName = "CursorCrosshair"
)

$ErrorActionPreference = "Stop"

$srcXlam = Join-Path $PSScriptRoot "$AddinName.xlam"
if (-not (Test-Path $srcXlam)) {
    Write-Host "ERROR: $AddinName.xlam not found at $srcXlam" -ForegroundColor Red
    Write-Host "       Run Build.ps1 first."
    exit 1
}

$xlstart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
if (-not (Test-Path $xlstart)) {
    New-Item -ItemType Directory -Path $xlstart -Force | Out-Null
}

$destXlam = Join-Path $xlstart "$AddinName.xlam"
Copy-Item $srcXlam $destXlam -Force

Write-Host "Installed: $destXlam" -ForegroundColor Green
Write-Host ""
Write-Host "The add-in will load automatically next time Excel starts."
Write-Host "If Excel is currently open, close all Excel windows and reopen."
Write-Host ""
Write-Host "If Excel disables macros on first load, add the XLSTART folder"
Write-Host "as a Trusted Location:"
Write-Host "  Excel > File > Options > Trust Center > Trust Center Settings"
Write-Host "        > Trusted Locations > Add new location"
Write-Host "  Path: $xlstart"
