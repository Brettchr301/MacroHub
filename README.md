# MacroHub

Office productivity toolkit built in PowerShell + WPF. One script, 7 tabs, no external dependencies.

I built this because I got tired of manually repeating the same Excel workflows every quarter. It started as a clipboard manager and macro runner, then kept growing as I added more stuff I needed.

## What's in it

MacroHub runs as a single WPF window with a dark-mode UI (Chrome/Teams style). The tabs:

- **Clipboard** - Multi-slot clipboard manager with paste sequencing. Saves slots between sessions.
- **Macros** - Browse and run PowerShell/VBA macros from the `Macros/` folder. Organized into Data, Format, Export, Quality, Ops, and VBA categories.
- **Scheduler** - Schedule macros to run on a timer or at specific times.
- **Navigator** - Workbook/sheet navigator for jumping between open Excel files. Drag to reorder sheets.
- **Templates** - Manage and insert reusable templates.
- **QSync** - QuarterSync: tracks quarterly file snapshots, compares across periods.
- **QTasks** - Quarterly task checklist with per-quarter JSON/CSV tracking (lives in `quarters/`).

## Requirements

- Windows 10 or 11
- PowerShell 5.1 (ships with Windows, you already have it)
- Office 2016 or later for the COM automation stuff (Excel). The app will still open without Office but the Excel features won't work obviously.

## How to run

Easiest way:

```
powershell -ExecutionPolicy Bypass -File MacroHub.ps1
```

Or right-click `Install.ps1` > "Run with PowerShell" to create a desktop shortcut. The shortcut launches with `-WindowStyle Hidden` so you get the WPF window without a console flash.

No admin rights needed. No modules to install.

## Macro library

Macros live in `Macros/` and are organized into subfolders:

```
Macros/
├── Modular/
│   ├── Data/          # dedup, pivot, normalize dates, etc.
│   ├── Format/        # auto-fit, freeze rows, conditional format
│   ├── Export/        # CSV export, HTML range export
│   ├── Quality/       # highlight blanks, formula errors
│   ├── Ops/           # recalculate workbook
│   └── VBA/           # .bas files for direct VBA injection
├── AddTimestamp.bas
├── ExportSheetToCsv.ps1
├── FormatNumbers.ps1
└── FormatReportHeader.bas
```

PowerShell macros (`.ps1`) run directly. VBA macros (`.bas`) get injected into the active workbook through Excel COM.

## Config files

- `qs_config.json` - QuarterSync configuration (active quarter, sync paths)
- `macro_chains.json` - Saved macro chain sequences for batch runs

## Known issues / TODO

- The Navigator tab sometimes loses its sheet list if Excel closes while MacroHub is open. Just click Refresh.
- Want to add a "Favorites" system for pinning frequently-used macros to the top. The config file exists (`favorites.json`) but the UI isn't wired up yet.
- Scheduler doesn't survive MacroHub restarts. Would be nice to persist scheduled tasks.

## License

Personal use. Not published anywhere.
