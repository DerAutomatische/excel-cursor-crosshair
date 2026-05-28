@echo off
setlocal

set "TARGET=%APPDATA%\Microsoft\Excel\XLSTART"
set "FILE=CursorCrosshair.xlam"

if not exist "%TARGET%\%FILE%" (
    echo Cursor Crosshair is not installed.
    echo Looked for: %TARGET%\%FILE%
    pause
    exit /b 0
)

tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
    echo Excel is currently running.
    echo Close all Excel windows, then run Uninstall.bat again.
    pause
    exit /b 1
)

del "%TARGET%\%FILE%"
if errorlevel 1 (
    echo ERROR: Could not delete %TARGET%\%FILE%
    pause
    exit /b 1
)

echo Cursor Crosshair uninstalled.
pause
