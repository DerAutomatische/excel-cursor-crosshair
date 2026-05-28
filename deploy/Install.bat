@echo off
setlocal

set "TARGET=%APPDATA%\Microsoft\Excel\XLSTART"
set "FILE=CursorCrosshair.xlam"

if not exist "%~dp0%FILE%" (
    echo ERROR: %FILE% not found in this folder.
    echo Make sure Install.bat and %FILE% are in the same folder.
    pause
    exit /b 1
)

tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
    echo Excel is currently running.
    echo Close all Excel windows, then run Install.bat again.
    pause
    exit /b 1
)

if not exist "%TARGET%" mkdir "%TARGET%"

REM Strip Mark-of-the-Web in case the file was downloaded from the internet.
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Unblock-File -Path '%~dp0%FILE%' -ErrorAction SilentlyContinue } catch {}" >nul 2>nul

copy /Y "%~dp0%FILE%" "%TARGET%\%FILE%" >nul
if errorlevel 1 (
    echo ERROR: Could not copy %FILE% to:
    echo   %TARGET%
    pause
    exit /b 1
)

echo.
echo Cursor Crosshair installed.
echo.
echo The add-in will load automatically next time you open Excel.
echo.
echo If Excel blocks macros on first load, add this folder as a Trusted Location:
echo   %TARGET%
echo (Excel ^> File ^> Options ^> Trust Center ^> Trust Center Settings
echo  ^> Trusted Locations ^> Add new location)
echo.
pause
