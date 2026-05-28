# Cursor Crosshair for Excel

Highlights the active row and column in every Excel workbook you open,
in mint green, leaving the selected cell itself clear so the cursor stays
visible. The highlight is an overlay — your existing cell fills and
conditional formatting are preserved, and the crosshair is stripped
before any save, so it never persists into your files.

## Install

1. Click **Code → Download ZIP** at the top of this page.
2. Extract the zip.
3. Open the `deploy` folder inside.
4. Close all Excel windows.
5. Double-click **`Install.bat`**.
6. Open Excel — the crosshair appears in any workbook you open.

No admin rights, no installer wizard, no rebuilding required.

### If Excel blocks macros on first load

Add the XLSTART folder as a Trusted Location:

> Excel → File → Options → Trust Center → Trust Center Settings → Trusted Locations → Add new location

Path:

```
%APPDATA%\Microsoft\Excel\XLSTART
```

## Uninstall

Double-click `deploy/Uninstall.bat`, or manually delete:

```
%APPDATA%\Microsoft\Excel\XLSTART\CursorCrosshair.xlam
```

## How it works

A small `.xlam` add-in lives in Excel's XLSTART folder and auto-loads
on every Excel startup. It listens to `Application.SheetSelectionChange`
and adds two short-lived Conditional Formatting rules (active-row and
active-column) on each cursor move. The rules are stripped before any
workbook save via `WorkbookBeforeSave`, so they never persist into the
saved file.

The active cell itself is excluded from the highlight via an inline
formula trick — `(COLUMN()<>n) * (LEN(marker)>0)` — so the selection
indicator stays clearly visible against the cell's underlying fill.

The marker string embedded in each formula lets the add-in identify and
remove its own rules reliably without holding `FormatCondition` object
references (which Excel can invalidate when the rules collection is
modified).

## Build from source

Prerequisites:

- Excel installed on this machine.
- "Trust access to the VBA project object model" enabled in
  Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings.

Build and install for development:

```powershell
.\Build.ps1
.\Install-CursorCrosshair.ps1
```

`Build.ps1` compiles `src/CCrosshairEvents.cls` into a `.xlam` via Excel
COM automation, and refreshes both the project-root `.xlam` and the
shareable copy under `deploy/`.

## Repo layout

| Path | Purpose |
| --- | --- |
| `src/CCrosshairEvents.cls` | VBA class with the event handlers |
| `Build.ps1` | Compiles the `.cls` into `CursorCrosshair.xlam` |
| `Install-CursorCrosshair.ps1` / `Uninstall-CursorCrosshair.ps1` | Dev convenience scripts |
| `deploy/` | Shareable bundle — what end users download |

## License

[MIT](LICENSE)
