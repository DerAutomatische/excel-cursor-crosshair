# Build.ps1
# Compiles src\*.cls + ThisWorkbook bootstrap into CursorCrosshair.xlam
# via Excel COM automation.
#
# Prerequisites:
#   - Excel installed on this machine
#   - "Trust access to the VBA project object model" enabled in
#     Excel > File > Options > Trust Center > Trust Center Settings >
#     Macro Settings.

[CmdletBinding()]
param(
    [string]$OutputDir = $PSScriptRoot,
    [string]$AddinName = "CursorCrosshair"
)

$ErrorActionPreference = "Stop"

$srcDir = Join-Path $PSScriptRoot "src"
$outputPath = Join-Path $OutputDir "$AddinName.xlam"

if (-not (Test-Path $srcDir)) {
    Write-Host "ERROR: src folder not found at $srcDir" -ForegroundColor Red
    exit 1
}

if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force
}

Write-Host "Launching Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null
try {
    $wb = $excel.Workbooks.Add()

    $vbProj = $null
    try { $vbProj = $wb.VBProject } catch { }
    if ($null -eq $vbProj) {
        Write-Host ""
        Write-Host "ERROR: Cannot access Excel's VBA project model." -ForegroundColor Red
        Write-Host ""
        Write-Host "To fix:"
        Write-Host "  1. Open Excel"
        Write-Host "  2. File > Optionen > Trust Center > Einstellungen fuer das Trust Center"
        Write-Host "     (File > Options > Trust Center > Trust Center Settings)"
        Write-Host "  3. Makroeinstellungen (Macro Settings)"
        Write-Host "  4. CHECK 'Zugriff auf das VBA-Projektobjektmodell vertrauen'"
        Write-Host "     ('Trust access to the VBA project object model')"
        Write-Host "  5. OK, close Excel, then re-run Build.ps1"
        Write-Host ""
        throw "VBA project access denied."
    }

    # VBA's .cls / .bas importer requires CRLF line endings to parse the
    # header block (VERSION / BEGIN / Attribute lines). LF-only files cause
    # the headers to leak into the code body and break at compile time.
    # Stage each source file through a temp file with normalized CRLF.
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("CursorCrosshair_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    try {
        $sources = @(Get-ChildItem -Path $srcDir -File | Where-Object { $_.Extension -in '.cls', '.bas' })
        foreach ($src in $sources) {
            Write-Host "  Importing $($src.Name)"
            $text = [System.IO.File]::ReadAllText($src.FullName)
            $text = $text -replace "`r`n", "`n" -replace "`n", "`r`n"
            $tempPath = Join-Path $tempDir $src.Name
            [System.IO.File]::WriteAllText($tempPath, $text, [System.Text.Encoding]::ASCII)
            $vbProj.VBComponents.Import($tempPath) | Out-Null
        }
    }
    finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Workbook-level VBA component is locale-named in some Excel builds; look
    # it up via the workbook's own CodeName instead of a hardcoded string.
    $thisWb = $vbProj.VBComponents.Item($wb.CodeName)
    $bootstrap = @'
Option Explicit

Private mCrosshair As CCrosshairEvents

Private Sub Workbook_Open()
    Set mCrosshair = New CCrosshairEvents
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Set mCrosshair = Nothing
End Sub
'@
    $thisWb.CodeModule.AddFromString($bootstrap)

    $wb.IsAddin = $true
    $wb.Title = "Cursor Crosshair"
    $wb.Subject = "Highlights active row + column across all workbooks"

    # FileFormat 55 = xlOpenXMLAddIn (.xlam)
    $wb.SaveAs($outputPath, 55)
    $wb.Close($false)
    $wb = $null

    Write-Host ""
    Write-Host "Built: $outputPath" -ForegroundColor Green

    # Refresh the shareable deploy bundle if it exists.
    $deployDir = Join-Path $PSScriptRoot "deploy"
    if (Test-Path $deployDir) {
        Copy-Item $outputPath (Join-Path $deployDir "$AddinName.xlam") -Force
        Write-Host "Copied to: $deployDir\$AddinName.xlam" -ForegroundColor Green
    }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
