# Uninstall-CursorCrosshair.ps1
# Removes the add-in from Excel's XLSTART folder.

[CmdletBinding()]
param(
    [string]$AddinName = "CursorCrosshair"
)

$ErrorActionPreference = "Stop"

$destXlam = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART\$AddinName.xlam"

if (Test-Path $destXlam) {
    Remove-Item $destXlam -Force
    Write-Host "Removed: $destXlam" -ForegroundColor Green
    Write-Host "Restart Excel to fully unload the add-in."
} else {
    Write-Host "Not installed (not found at $destXlam)" -ForegroundColor Yellow
}
