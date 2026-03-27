# MacroHub v3.2

Office productivity toolkit built in PowerShell + WPF. One script, 8 tabs, no external dependencies.

Built to stop manually repeating the same Excel workflows every quarter. Started as a clipboard manager and macro runner, then kept growing.

## Requirements

- Windows 10 or 11
- PowerShell 5.1 (ships with Windows)
- Office 2016+ for Excel COM automation features (app opens fine without Office — Excel features simply won't work)

## How to run

```powershell
powershell -ExecutionPolicy Bypass -File MacroHub.ps1
```

Or right-click `Install.ps1` → "Run with PowerShell" to create a desktop shortcut. The shortcut uses `-WindowStyle Hidden` so there's no console flash.

No admin rights needed. No modules to install.

---

## Tabs

### Clipboard

Multi-slot clipboard manager with per-slot paste targets.

- Add as many slots as you need. Each slot captures one clipboard item (text, Excel range, etc.).
- Each slot has its own **sheet**, **cell**, and optional **timestamp** settings.
- Click **Record** on a slot to start capturing — the next thing you copy goes into that slot.
- **Record Sequence**: auto-captures successive clipboard copies into slots in order (Slot 1 → 2 → 3…).
- Click **Paste** to write the captured content to Excel at the configured destination.
- **Save Defaults**: persists the current workbook and slot-1 settings so new slots start pre-filled.
- Paste preserves full Excel cell formatting (fonts, fill colors, borders, number formats) when the source workbook is available.

### Macros

Browse and run PowerShell (`.ps1`) and VBA (`.bas`) macro files from the `Macros/` folder.

- **Left panel**: file tree of all discovered macros. Click to select.
- **Right panel**: configure target workbook and sheet, see file description, then click **Run Macro**.
- VBA macros are injected into the workbook's VBA project, run, and removed automatically.
  - Requires Excel's "Trust Access to the VBA Project Object Model" setting enabled.
- PowerShell macros run in-process with optional `-WorkbookName` and `-SheetName` parameters.
- **Toggle Favorite**: marks a macro as a favorite (stored in `favorites.json`).

### Scheduler

Create and manage macros that run on a Windows Task Scheduler schedule.

- Tasks are stored under the `\MacroHub` folder in Windows Task Scheduler.
- Tasks run even when MacroHub is closed (backed by the OS scheduler).
- Configure: task name, script to run, time (HH:mm), and frequency (Daily / Weekly / Monthly).
- On startup, MacroHub checks for missed tasks and offers to run them immediately.
- Delete tasks from the list at the bottom of the tab.

### Navigator

Workbook and worksheet management panel for open Excel files.

- **Left**: list of open workbooks in the current Excel session. Click to see its sheets.
- **Middle**: worksheet list for the selected workbook. Supports multi-select.
  - **Drag** to reorder sheets.
  - **Ctrl+E** to rename the selected sheet.
- **Right**: actions panel —
  - Activate / bring to front / minimize / close workbook
  - Move or copy sheets between workbooks
  - Export a sheet to CSV or XLSX
  - Hide / Unhide sheets
  - Delete sheets
  - Set, remove, or open-with password protection
  - Adjust Excel engine options: calculation mode, events on/off
  - Browse and run VBA modules registered in the workbook

### Templates

Create, save, and reuse text templates that paste directly into Excel.

- **Left panel**: saved template list. Click to load into the editor.
- **Right panel**: name + content editor.
  - Supported placeholders: `{DATE}`, `{SHEET}`, `{USER}` — substituted at paste time.
- **Preview**: shows the rendered output with placeholders filled in.
- **Paste to Excel**: inserts the template content into the currently active cell.
- Templates are saved to `templates.json` in the app root.

### QSync

Quarter-over-quarter folder comparison and to-do generator.

- Point it at **last quarter's root folder** and **this quarter's root folder**.
- Click **Run Sync** to compare the two: new folders are created, new to-do items are added for folders/files that exist in the baseline but not in the current quarter.
- Shows a progress dashboard: folders created, folders skipped, new to-do items, errors.
- **Completion by Folder**: visual progress bars per subfolder.
- **Sync Log**: real-time log of every action taken during the sync.
- Results and to-do lists are saved to `qs_synclog.json` and `qs_compare_results.json`.

### QTasks

Quarterly task checklist and file delivery tracker.

- Compare **Folder A** (baseline) against **Folder B** (target/delivery) to find missing files.
- **Find Missing Files**: scans both folders and adds any files present in A but missing in B as pending tasks.
- Filter the task list by status (All / Pending / Done), file name, or folder.
- Mark tasks done inline. Notes are preserved per item.
- **Export to Excel**: exports the full task list as a formatted workbook.
- Task data stored as JSON in `quarters/` (one file per quarter).

### File Index

DLP-safe metadata-only drive search tool. Scans a folder tree and caches file metadata for instant filtering.

- Select a **Root Folder** and click **Index Now** to scan. Only reads file name, path, size, and modified date — never opens or reads file contents.
- **60-minute cooldown** after each scan to prevent repeated hammering of large network drives.
- **Real-time search**: type in the search box to instantly filter by name, path, or extension.
- Results shown in a sortable, resizable grid with columns: File Name, Path, Size, Modified, Ext.
- **Double-click** or press **Enter** on any result to open the file.
- Index cached as `fileindex_cache.json` — loaded automatically on next app start.
- Status bar shows total file count, match count, and last scan time.

- Compare **Folder A** (baseline) against **Folder B** (target/delivery) to find missing files.
- **Find Missing Files**: scans both folders and adds any files present in A but missing in B as pending tasks.
- Filter the task list by status (All / Pending / Done), file name, or folder.
- Mark tasks done inline. Notes are preserved per item.
- **Export to Excel**: exports the full task list as a formatted workbook.
- Task data stored as JSON in `quarters/` (one file per quarter).

---

## Macro library

Macros live in `Macros/` organized into subfolders:

```
Macros/
├── Modular/
│   ├── Data/          # dedup, pivot normalize dates, etc.
│   ├── Format/        # auto-fit columns, freeze rows, conditional formatting
│   ├── Export/        # CSV export, HTML range export
│   ├── Quality/       # highlight blanks, flag formula errors
│   ├── Ops/           # recalculate workbook, toggle settings
│   └── VBA/           # .bas files for direct VBA injection
├── AddTimestamp.bas
├── ExportSheetToCsv.ps1
├── FormatNumbers.ps1
└── FormatReportHeader.bas
```

PowerShell macros (`.ps1`) run directly. VBA macros (`.bas`) are injected through Excel COM, run, then removed.

---

## Config files

| File | Purpose |
|------|---------|
| `qs_config.json` | QuarterSync: active quarter, compare folder paths |
| `clip_defaults.json` | Clipboard tab default workbook/sheet/cell/timestamp settings |
| `templates.json` | Saved templates |
| `favorites.json` | Favorited macro names |
| `quarters/` | QTasks per-quarter JSON data files |

---

## Known issues / Notes

- **Navigator**: if Excel closes while MacroHub is open, the sheet list goes stale. Click Refresh.
- **Favorites**: the favorites config exists but the Macros tab doesn't yet filter by favorites in the list.
- **Scheduler**: tasks created here persist in Windows Task Scheduler between app restarts.

## License

Personal use.
