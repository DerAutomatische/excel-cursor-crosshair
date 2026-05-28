Cursor Crosshair for Excel
==========================

Highlights the active row and column in every Excel workbook you open,
in mint green, leaving the selected cell itself clear so the selection
indicator stays visible. The highlight is removed before any save, so
it never gets baked into your files.

Install
-------
  1. Close all Excel windows.
  2. Double-click Install.bat.
  3. Open Excel.

That's it. The add-in loads automatically in every Excel session from
now on.

If Excel blocks macros on first load
------------------------------------
Add the XLSTART folder as a Trusted Location:
  Excel > File > Options > Trust Center > Trust Center Settings
        > Trusted Locations > Add new location
Path:
  %APPDATA%\Microsoft\Excel\XLSTART

Uninstall
---------
Double-click Uninstall.bat. Or manually delete:
  %APPDATA%\Microsoft\Excel\XLSTART\CursorCrosshair.xlam
