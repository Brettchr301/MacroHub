# MacroHub

Office productivity toolkit built in PowerShell + WPF. One script, 11 tabs, no external dependencies.

I built this because I got tired of manually repeating the same Excel/Outlook workflows every quarter. It started as a clipboard manager and macro runner, then kept growing as I added more stuff I needed.

## What's in it

MacroHub runs as a single WPF window with a dark-mode UI (Chrome/Teams style). The tabs:

- **Clipboard** - Multi-slot clipboard manager with paste sequencing. Saves slots between sessions.
- **Macros** - Browse and run PowerShell/VBA macros from the `Macros/` folder. Organized into Data, Format, Export, Quality, Ops, and VBA categories.
- **Scheduler** - Schedule macros to run on a timer or at specific times.
- **Navigator** - Workbook/sheet navigator for jumping between open Excel files. Drag to reorder sheets.
- **Templates** - Manage and insert reusable templates.
- **QSync** - QuarterSync: tracks quarterly file snapshots, compares across periods.
- **QTasks** - Quarterly task checklist with per-quarter JSON/CSV tracking (lives in `quarters/`).
- **Email Helper** - Compose and send through Outlook COM automation. Has templates for follow-ups, meeting requests, weekly status, etc.
- **Email Dashboard** - Read/search your Outlook inbox without switching windows.
- **IDE** - Built-in code editor with syntax validation for PowerShell and VBA. Edit macros in-app.
- **AI IDE** - Connects to AI chat services (Copilot, Claude, ChatGPT) via the Chrome extension bridge. Send code context, get suggestions back.

## Requirements

- Windows 10 or 11
- PowerShell 5.1 (ships with Windows, you already have it)
- Office 2016 or later for the COM automation stuff (Excel, Outlook). The app will still open without Office but the Excel/Outlook features won't work obviously.

## How to run

Easiest way:

```
powershell -ExecutionPolicy Bypass -File MacroHub.ps1
```

Or right-click `Install.ps1` > "Run with PowerShell" to create a desktop shortcut. The shortcut launches with `-WindowStyle Hidden` so you get the WPF window without a console flash.

No admin rights needed. No modules to install.

## Chrome extension setup

The `ChromeExtension/` folder has a Manifest V3 extension that bridges AI chat pages to the AI IDE tab.

1. Open `chrome://extensions/`
2. Enable "Developer mode" (top right toggle)
3. Click "Load unpacked" and point it at the `ChromeExtension/` folder
4. The extension connects to MacroHub's local server on `localhost:9876`
5. Works with Copilot, Claude, and ChatGPT pages

The bridge injects a content script that watches for AI responses and relays them back to the IDE tab. You can also push code context from MacroHub to the AI chat.

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
├── EmailTemplates/    # follow-up, meeting request, weekly status
├── AddTimestamp.bas
├── ExportSheetToCsv.ps1
├── FormatNumbers.ps1
└── FormatReportHeader.bas
```

PowerShell macros (`.ps1`) run directly. VBA macros (`.bas`) get injected into the active workbook through Excel COM.

## Config files

- `email_config.json` - Outlook email settings (default signatures, folders, etc.)
- `qs_config.json` - QuarterSync configuration (active quarter, sync paths)
- `macro_chains.json` - Saved macro chain sequences for batch runs

## Known issues / TODO

- The Navigator tab sometimes loses its sheet list if Excel closes while MacroHub is open. Just click Refresh.
- Email Dashboard search is slow on large mailboxes (>10k items). Need to add date-range filtering.
- AI IDE bridge occasionally drops the websocket connection - just reconnect from the extension popup.
- Want to add a "Favorites" system for pinning frequently-used macros to the top. The config file exists (`favorites.json`) but the UI isn't wired up yet.
- Scheduler doesn't survive MacroHub restarts. Would be nice to persist scheduled tasks.

## License

Personal use. Not published anywhere.
