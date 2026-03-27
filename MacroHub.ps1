#Requires -Version 5.1
<#
.SYNOPSIS
    MacroHub v3.2 - Office Productivity Super App
.DESCRIPTION
    Merged: MacroHub automation suite + QuarterSync quarterly tracker + File Index.
    9 tabs: Clipboard | Macros | Scheduler | Navigator | Templates | QSync | QTasks | File Index | Formula Auditor

    Chrome/Teams dark-mode UI. Pure PowerShell 5.1 / WPF. No external modules.

    Color palette: black/grey/white text, blue accents, red for destructive only.
    Contrast: all labels ≥ #909090 on dark backgrounds for WCAG AA readability.

    -----------------------------------------------------------------------
    WHAT THIS FILE IS
    -----------------------------------------------------------------------
    MacroHub is a single-window WPF desktop app that centralises common
    Excel/Office automation workflows:
      - Clipboard multi-slot capture and paste to Excel
      - VBA (.bas) and PowerShell (.ps1) macro library runner
      - Windows Task Scheduler integration for recurring macros
      - Excel workbook/sheet navigator (rename, move, hide, export, VBA runner)
      - Reusable text template library
      - QuarterSync: folder structure replication quarter-to-quarter
      - QTasks: folder-comparison deliverables tracker with Excel export

    -----------------------------------------------------------------------
    HOW THE APP IS STRUCTURED
    -----------------------------------------------------------------------
    Single WPF Window containing a TabControl with 9 tabs:
      Tab 0 = Clipboard   -- multi-slot clipboard recording and paste to Excel
      Tab 1 = Macros      -- discover and run .bas/.ps1 macros from the Macros/ folder
      Tab 2 = Scheduler   -- create/remove Windows Task Scheduler entries (\MacroHub)
      Tab 3 = Navigator   -- manage workbooks and sheets in a Navigator Excel session
      Tab 4 = Templates   -- save/load/paste reusable text snippets
      Tab 5 = QSync       -- copy last-quarter folder structure to this quarter
      Tab 6 = QTasks      -- compare two folders and track missing deliverables
      Tab 7 = File Index  -- DLP-safe metadata-only drive search
      Tab 8 = Formula Auditor -- read-only formula inspection and temp highlighting

    -----------------------------------------------------------------------
    CODE ORGANIZATION (top to bottom)
    -----------------------------------------------------------------------
    1.  ASSEMBLIES        -- Add-Type for WPF, Windows.Forms, System.Drawing
    2.  CONFIGURATION     -- $script: path/file globals and version constant
    3.  GLOBAL STATE      -- runtime state variables (slots, timers, COM refs)
    4.  HELPER FUNCTIONS  -- HexBrush, Write-ActivityLog, Normalize-List, UI helpers
    5.  EXCEL COM HELPERS -- Get-ExcelApp, workbook/worksheet accessors
    6.  MACRO DISCOVERY   -- Get-MacroFiles scans the Macros/ folder
    7.  MACRO EXECUTION   -- Invoke-VbaMacro / Invoke-PsScript / Invoke-SelectedMacro
    8.  TASK SCHEDULER    -- Get/Register/Remove-HubTask via Schedule.Service COM
    9.  CLIPBOARD HELPERS -- Get-ClipboardPacket, dimensions, matrix, Paste-ClipboardPacketToExcel
    10. ADD CLIP SLOT UI  -- Add-ClipSlotUI builds the dynamic per-slot card
    11. TEMPLATE MGMT     -- Load/Save-Templates JSON
    12. QUARTERSYNC       -- date-stripping, folder compare, sync engine, Excel export
    13. QUARTER FILE MGMT -- quarter JSON file helpers
    14. FILE INDEX        -- metadata-only folder scan, cache, filter
    15. XAML              -- entire window XAML (inside Start-MacroHub here-string)
    16. G() BINDINGS      -- FindName() calls that populate $ClipWbCombo etc.
    17. REFRESH FUNCTIONS -- Refresh-* functions that populate UI controls from data
    18. EVENT HANDLERS    -- Add_Click / Add_SelectionChanged for every control
    19. KEYBOARD / KEYTIPS -- Alt overlay, Ctrl+PgUp/PgDn, Ctrl+F search
    20. INITIAL LOAD      -- startup data population and lazy-tab loading
    21. SHOW WINDOW       -- $Window.ShowDialog()
    22. LAUNCH            -- Start-MacroHub (entry point at file end)
#>

# ============================================================
#  ASSEMBLIES
# ============================================================
# Load WPF (PresentationFramework/Core/WindowsBase) and Windows.Forms for
# dialogs, plus System.Drawing for the system-tray NotifyIcon.
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
#  CONFIGURATION
# ============================================================
# All file paths are derived from $HubRoot (the folder that contains this script).
# JSON is the canonical storage format; CSV paths are kept only for legacy migration.
$script:HubRoot      = Split-Path $MyInvocation.MyCommand.Path
$script:MacroFolder  = Join-Path $script:HubRoot 'Macros'
$script:LogFile      = Join-Path $script:HubRoot 'MacroHub.log'
$script:TemplateJson = Join-Path $script:HubRoot 'templates.json'
$script:TemplateCSV  = Join-Path $script:HubRoot 'templates.csv'      # legacy import only
$script:QsTodoCSV    = Join-Path $script:HubRoot 'qs_todos.csv'       # legacy import only
$script:QsSyncJson   = Join-Path $script:HubRoot 'qs_synclog.json'
$script:QsSyncCSV    = Join-Path $script:HubRoot 'qs_synclog.csv'     # legacy import only
$script:QsCfgJson    = Join-Path $script:HubRoot 'qs_config.json'
$script:QsCfgCSV     = Join-Path $script:HubRoot 'qs_config.csv'      # legacy import only
$script:QsCompareJson = Join-Path $script:HubRoot 'qs_compare_results.json'
$script:QuartersDir  = Join-Path $script:HubRoot 'quarters'
$script:ActiveQuarterPath = ''
$script:ActiveQuarterCSV = ''  # legacy alias
$script:FavJson      = Join-Path $script:HubRoot 'favorites.json'
$script:FavCSV       = Join-Path $script:HubRoot 'favorites.csv'      # legacy import only
$script:ClipDefaultsJson = Join-Path $script:HubRoot 'clip_defaults.json'
$script:AppVersion   = '3.2.0'

# ============================================================
#  GLOBAL STATE
# ============================================================
# Runtime-mutable variables shared across functions.
# ClipSlots is a hashtable keyed by "Slot_N" holding per-slot state objects.
# ExcelMainApp / ExcelNavApp cache live COM handles; they are validated before use.
$script:ClipSlots      = @{}
$script:ClipSlotIdx    = 0
$script:ClipSequenceEnabled = $false
$script:ClipSequenceLastSig = ''
$script:ClipSequenceNextIndex = 0
$script:ClipSequenceTimer = $null
$script:NavSheetDragStart = $null
$script:NavSheetDragSourceItem = $null
$script:KeyTipsActive  = $false
$script:BusyCount      = 0
$script:Window         = $null
$script:BusyOverlay    = $null
$script:BusyText       = $null
$script:StatusBar      = $null
$script:ExcelMainApp   = $null
$script:ExcelNavApp    = $null

# ============================================================
#  HELPER: Solid color brush from hex string (cached + frozen)
# ============================================================
# Returns a frozen SolidColorBrush for the given hex color string (e.g. '#4C9FE6').
# Brushes are cached by hex string so repeated calls are allocation-free.
$script:BrushCache = @{}
function HexBrush([string]$hex) {
    $b = $script:BrushCache[$hex]
    if ($b) { return $b }
    $b = [System.Windows.Media.SolidColorBrush](
        [System.Windows.Media.ColorConverter]::ConvertFromString($hex))
    $b.Freeze()
    $script:BrushCache[$hex] = $b
    return $b
}

# ============================================================
#  ACTIVITY LOGGER (CSV-safe, ASCII only)
# ============================================================
# Appends a timestamped CSV row to MacroHub.log for every user action.
function Write-ActivityLog {
    param([string]$Action)
    $safe = $Action -replace '"','""'
    $line = '"{0}","{1}"' -f (Get-Date -f 'yyyy-MM-dd HH:mm:ss'), $safe
    try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 } catch {}
}

# Strips $null entries from a loaded JSON/CSV array and guarantees an array return.
function Normalize-List {
    param($InputObject)
    $items = @($InputObject | Where-Object { $null -ne $_ })
    return ,$items
}

# ============================================================
#  UI HELPERS: busy overlay, status bar, dispatcher flush
# ============================================================
# Processes all pending WPF dispatcher messages, keeping the UI responsive
# during long-running synchronous operations.
function Update-UI {
    if ($script:Window) {
        $script:Window.Dispatcher.Invoke(
            [action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }
}

# Shows the modal busy overlay with $Message and increments a nesting counter so
# nested Show-Busy/Hide-Busy pairs don't prematurely hide the overlay.
function Show-Busy {
    param([string]$Message = 'Working...')
    $script:BusyCount++
    if ($script:BusyOverlay) {
        $script:BusyText.Text = $Message
        $script:BusyOverlay.Visibility = 'Visible'
    }
    Update-UI
}

# Decrements the busy counter and hides the overlay when it reaches zero.
function Hide-Busy {
    $script:BusyCount = [Math]::Max(0, $script:BusyCount - 1)
    if ($script:BusyCount -eq 0 -and $script:BusyOverlay) {
        $script:BusyOverlay.Visibility = 'Collapsed'
    }
    Update-UI
}

# Updates the status bar text and foreground color; defaults to muted grey.
function Set-Status {
    param([string]$Text, [string]$Color = '#B0B0B0')
    if ($script:StatusBar) {
        $script:StatusBar.Text = "  $Text"
        $script:StatusBar.Foreground = HexBrush $Color
    }
    Update-UI
}

# ============================================================
#  EXCEL COM HELPERS
# ============================================================
# These functions centralise all COM interop with Excel.
# Two COM sessions are tracked: 'Main' (used by Clipboard/Macros tabs) and
# 'Navigator' (a dedicated or fallback instance for the Navigator tab).
# Validates the cached handle with Test-ExcelComObject before each use.

# Returns $true if $ExcelApp is a live COM object with an accessible Workbooks collection.
function Test-ExcelComObject {
    param($ExcelApp)
    if (-not $ExcelApp) { return $false }
    try {
        $null = $ExcelApp.Workbooks.Count
        return $true
    } catch {
        return $false
    }
}

# Returns the cached or running Excel COM object for the given session ('Main',
# 'Navigator', or 'Any'); creates a new visible instance when -Create is specified.
function Get-ExcelApp {
    param(
        [ValidateSet('Main','Navigator','Any')]
        [string]$Session = 'Main',
        [switch]$Create
    )

    if ($Session -eq 'Any') {
        if (Test-ExcelComObject $script:ExcelMainApp) { return $script:ExcelMainApp }
        if (Test-ExcelComObject $script:ExcelNavApp)  { return $script:ExcelNavApp }
    }

    if ($Session -eq 'Main' -and -not (Test-ExcelComObject $script:ExcelMainApp)) { $script:ExcelMainApp = $null }
    if ($Session -eq 'Navigator' -and -not (Test-ExcelComObject $script:ExcelNavApp)) { $script:ExcelNavApp = $null }

    if ($Session -eq 'Main' -and $script:ExcelMainApp) { return $script:ExcelMainApp }
    if ($Session -eq 'Navigator' -and $script:ExcelNavApp) { return $script:ExcelNavApp }
    # Navigator fallback: if no dedicated navigator session exists yet, use main session if available.
    if ($Session -eq 'Navigator' -and -not $Create -and (Test-ExcelComObject $script:ExcelMainApp)) {
        return $script:ExcelMainApp
    }

    # For Navigator requests that explicitly need a session, force a dedicated Excel instance.
    if ($Create -and $Session -eq 'Navigator') {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $true
        $script:ExcelNavApp = $xl
        return $xl
    }

    try {
        # GetActiveObject retrieves the running Excel COM instance registered in the ROT
        # (Running Object Table); throws if no instance is currently running.
        $active = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        $active = $null
    }

    if ($active) {
        if ($Session -eq 'Navigator') { $script:ExcelNavApp = $active }
        elseif ($Session -eq 'Main')  { $script:ExcelMainApp = $active }
        return $active
    }

    if ($Create) {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $true
        if ($Session -eq 'Navigator') { $script:ExcelNavApp = $xl }
        else { $script:ExcelMainApp = $xl }
        return $xl
    }

    return $null
}

# Returns an array of open workbook names from the specified Excel session.
function Get-OpenWorkbooks {
    param(
        [ValidateSet('Main','Navigator','Any')]
        [string]$Session = 'Main'
    )
    $xl = Get-ExcelApp -Session $Session
    if (-not $xl) { return @() }
    $names = @()
    try {
        for ($i = 1; $i -le $xl.Workbooks.Count; $i++) {
            # Workbooks collection is 1-based in Excel COM
            $names += $xl.Workbooks.Item($i).Name
        }
    } catch {}
    return $names
}

# Returns an array of worksheet names for the named open workbook.
function Get-WorksheetNames {
    param(
        [string]$WorkbookName,
        [ValidateSet('Main','Navigator','Any')]
        [string]$Session = 'Main'
    )
    $xl = Get-ExcelApp -Session $Session
    if (-not $xl) { return @() }
    $sheets = @()
    try {
        $wb = $xl.Workbooks.Item($WorkbookName)
        for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
            $sheets += $wb.Worksheets.Item($i).Name
        }
    } catch {}
    return $sheets
}

function Get-WorkbookSheets {
    param(
        [string]$WorkbookName,
        [ValidateSet('Main','Navigator','Any')]
        [string]$Session = 'Main'
    )
    return Get-WorksheetNames -WorkbookName $WorkbookName -Session $Session
}

# Opens $FilePath in the specified Excel session, creating that session if needed.
function Open-WorkbookInExcel {
    param(
        [string]$FilePath,
        [ValidateSet('Main','Navigator')]
        [string]$Session = 'Main'
    )
    $xl = Get-ExcelApp -Session $Session -Create
    try {
        [void]($xl.Workbooks.Open($FilePath))
        return $true
    } catch {
        return $false
    }
}

# Writes $Text line by line into the worksheet starting at $CellAddress; used by
# the Templates tab paste action (plain text only, no formatting).
function Paste-TextToSheet {
    param(
        [string]$WorkbookName,
        [string]$SheetName,
        [string]$CellAddress,
        [string]$Text,
        [ValidateSet('Main','Navigator')]
        [string]$Session = 'Main'
    )
    $xl = Get-ExcelApp -Session $Session
    if (-not $xl) { throw 'Excel is not open.' }
    $wb = $xl.Workbooks.Item($WorkbookName)
    $ws = $wb.Worksheets.Item($SheetName)
    $cell = $ws.Range($CellAddress)
    # Multi-line: split into rows starting at the target cell
    $lines = $Text -split "`r?`n"
    $startRow = $cell.Row
    $startCol = $cell.Column
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $ws.Cells.Item($startRow + $i, $startCol).Value2 = $lines[$i]
    }
    return $true
}

# ============================================================
#  MACRO DISCOVERY
# ============================================================
# Returns all .ps1 and .bas files found under $script:MacroFolder (recursive).
# Scans the Macros/ subfolder recursively for .bas (VBA) and .ps1 (PowerShell)
# files that can be run by the Macros tab or scheduled by the Scheduler tab.
function Get-MacroFiles {
    if (-not (Test-Path $script:MacroFolder)) { return @() }
    $files = Get-ChildItem -Path $script:MacroFolder -Recurse -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Extension -match '^\.(bas|ps1)$' }
    return @($files)
}

# ============================================================
#  MACRO EXECUTION
# ============================================================
# Invoke-VbaMacro: Imports a .bas module into the target workbook, runs the first
#   Public Sub, then removes the module. Requires Excel's Trust Access to VBA Project.
#
# Invoke-PsScript: Runs a .ps1 script, passing -WorkbookName and -SheetName as
#   optional parameters. Falls back to workbook-only if the script doesn't accept SheetName.
#
# Invoke-SelectedMacro: Dispatcher — routes to Invoke-VbaMacro or Invoke-PsScript
#   based on the file extension.
function Invoke-VbaMacro {
    param(
        [string]$MacroFile,
        [string]$WorkbookName,
        [string]$WorksheetName
    )
    $xl = Get-ExcelApp -Session Main
    if (-not $xl) { throw 'Excel is not open.' }
    $wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
    if (-not $wb) { throw 'No workbook is open.' }
    if ($WorksheetName) {
        try {
            $wsRun = $wb.Worksheets.Item($WorksheetName)
            if ($wsRun.Visible -ne -1) { $wsRun.Visible = -1 }
            $wsRun.Activate()
        } catch {}
    }

    # Find the entry-point Sub name from the .bas file
    $content = Get-Content $MacroFile -Raw -Encoding ASCII
    $entryPoint = $null
    if ($content -match 'Public\s+Sub\s+(\w+)') { $entryPoint = $matches[1] }
    elseif ($content -match 'Sub\s+(\w+)\s*\(') { $entryPoint = $matches[1] }
    if (-not $entryPoint) { throw 'No Sub found in .bas file.' }

    # Import module, run, then remove so the workbook is not left dirty with the macro.
    $vbProj = $wb.VBProject
    $comp = $vbProj.VBComponents.Import($MacroFile)
    try {
        # Run format: "WorkbookName!ModuleName.SubName"
        $xl.Run("$($wb.Name)!$($comp.Name).$entryPoint")
    } finally {
        try { $vbProj.VBComponents.Remove($comp) } catch {}
    }
    Write-ActivityLog "Ran VBA macro: $entryPoint from $(Split-Path $MacroFile -Leaf)"
}

# Executes a .ps1 script, forwarding -WorkbookName / -SheetName parameters when
# the script accepts them; silently retries without -SheetName on parameter error.
function Invoke-PsScript {
    param(
        [string]$ScriptFile,
        [string]$WorkbookName,
        [string]$WorksheetName
    )
    $prevPref = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'
    try {
        $callArgs = @()
        if ($WorkbookName) { $callArgs += '-WorkbookName'; $callArgs += $WorkbookName }
        if ($WorksheetName) { $callArgs += '-SheetName'; $callArgs += $WorksheetName }
        try {
            & $ScriptFile @callArgs
        } catch {
            if ($WorksheetName -and $_.Exception.Message -match "parameter name 'SheetName'") {
                $retryArgs = @()
                if ($WorkbookName) { $retryArgs += '-WorkbookName'; $retryArgs += $WorkbookName }
                & $ScriptFile @retryArgs
            } else {
                throw
            }
        }
    } finally {
        $ErrorActionPreference = $prevPref
    }
    Write-ActivityLog "Ran PS script: $(Split-Path $ScriptFile -Leaf)"
}

# Dispatches $MacroFile to either Invoke-VbaMacro (.bas) or Invoke-PsScript (.ps1).
function Invoke-SelectedMacro {
    param(
        [string]$MacroFile,
        [string]$WorkbookName,
        [string]$WorksheetName
    )
    if (-not (Test-Path $MacroFile)) { throw "Macro file not found: $MacroFile" }
    $ext = [System.IO.Path]::GetExtension($MacroFile).ToLower()
    switch ($ext) {
        '.bas' { Invoke-VbaMacro -MacroFile $MacroFile -WorkbookName $WorkbookName -WorksheetName $WorksheetName }
        '.ps1' { Invoke-PsScript -ScriptFile $MacroFile -WorkbookName $WorkbookName -WorksheetName $WorksheetName }
        default { throw "Unsupported macro type: $ext" }
    }
}

# ============================================================
#  TASK SCHEDULER (Windows Task Scheduler COM)
# ============================================================
# All scheduled tasks are stored under the '\MacroHub' folder in Windows Task Scheduler.
# Schedule.Service is the COM interface for the Task Scheduler engine; tasks are
# defined via ITaskDefinition objects with time triggers and powershell.exe actions.
#
# Get-HubTasks:          Returns a list of task objects (name, state, next/last run times).
# Get-MissedTasks:       Returns tasks that should have run but didn't (PC was off).
# Invoke-MissedTaskCheck: Startup check — prompts to run each missed task.
# Register-HubTask:      Creates a new daily/weekly/monthly trigger task via Schedule.Service COM.
# Remove-HubTask:        Deletes a task from the '\MacroHub' scheduler folder.
function Get-HubTasks {
    $tasks = @()
    try {
        $svc = New-Object -ComObject Schedule.Service
        $svc.Connect()
        $folder = $null
        try { $folder = $svc.GetFolder('\MacroHub') } catch { return @() }
        $collection = $folder.GetTasks(0)
        foreach ($t in $collection) {
            # Extract the action path/args for display
            $actPath = ''
            $actArgs = ''
            try {
                $acts = $t.Definition.Actions
                if ($acts.Count -gt 0) {
                    $actPath = $acts.Item(1).Path
                    $actArgs = $acts.Item(1).Arguments
                }
            } catch {}
            $tasks += [PSCustomObject]@{
                Name          = $t.Name
                State         = switch ($t.State) {
                    1 {'Disabled'} 2 {'Queued'} 3 {'Ready'} 4 {'Running'} default {'Unknown'}
                }
                NextRunTime   = $t.NextRunTime
                LastRunTime   = $t.LastRunTime
                LastRunResult = $t.LastRunResult
                ActionPath    = $actPath
                ActionArgs    = $actArgs
                Enabled       = $t.Enabled
            }
        }
    } catch {}
    return $tasks
}

function Get-MissedTasks {
    <#
      Returns tasks that should have run but didn't because the PC
      was off / disconnected.  Checks:
        - LastRunResult = 0x41303  (task has not run yet) or
        - LastRunTime is more than 25 hours old for daily tasks
        - State = Ready  (not currently running or disabled)
    #>
    $missed = @()
    $tasks = Get-HubTasks
    $now = Get-Date
    foreach ($t in $tasks) {
        if ($t.Enabled -eq $false) { continue }
        # 0x00041303 = Task has never run; 0x00041301 = Task is running;
        # 0x00000001 = Incorrect function (failed); 0x800710E0 = operator declined
        $neverRan   = ($t.LastRunResult -eq 0x41303)
        $staleRun   = $false
        if (-not $neverRan -and $t.NextRunTime -is [datetime]) {
            # If NextRunTime is in the past, the trigger was missed
            try { $staleRun = ($t.NextRunTime -lt $now.AddMinutes(-2)) } catch {}
        }
        if ($neverRan -or $staleRun) {
            $missed += $t
        }
    }
    return $missed
}

function Invoke-MissedTaskCheck {
    <# Startup check: pop a MessageBox for each missed task #>
    $missed = Get-MissedTasks
    foreach ($t in $missed) {
        $reason = if ($t.LastRunResult -eq 0x41303) {
            'has never run (computer may have been off)'
        } else {
            "missed its scheduled run at $($t.NextRunTime.ToString('g')) (computer was off or disconnected)"
        }
        $msg  = "Scheduled task `"$($t.Name)`" $reason.`n`n"
        $msg += "Would you like to run it now?"
        $result = [System.Windows.MessageBox]::Show(
            $msg,
            'MacroHub - Missed Task',
            'YesNo',
            'Warning'
        )
        if ($result -eq 'Yes') {
            try {
                # Determine what the task runs
                $args = $t.ActionArgs
                if ($args -match '-File\s+"?([^"]+)"?') {
                    $scriptPath = $matches[1].Trim()
                    if (Test-Path $scriptPath) {
                        Show-Busy "Running missed task: $($t.Name)..."
                        & powershell.exe -ExecutionPolicy Bypass -File $scriptPath
                        Hide-Busy
                        Set-Status "Ran missed task: $($t.Name)"
                        Write-ActivityLog "Ran missed task on startup: $($t.Name)"
                    } else {
                        [System.Windows.MessageBox]::Show(
                            "Script not found: $scriptPath",
                            'Error', 'OK', 'Error')
                    }
                } else {
                    [System.Windows.MessageBox]::Show(
                        'Could not determine script for this task.',
                        'Error', 'OK', 'Error')
                }
            } catch {
                [System.Windows.MessageBox]::Show(
                    "Error running task: $_",
                    'Error', 'OK', 'Error')
            }
        }
    }
}

# Creates a Windows Task Scheduler entry under \MacroHub that runs $MacroFile
# at $TriggerTime (HH:mm) on the given $Frequency (Daily / Weekly / Monthly).
# For .bas files, generates a PowerShell wrapper script in sched_wrappers/ first.
function Register-HubTask {
    param(
        [string]$TaskName,
        [string]$MacroFile,
        [string]$TriggerTime,         # HH:mm
        [string]$Frequency = 'Daily'  # Daily, Weekly, or Monthly
    )

    # If the file is a .bas VBA macro, generate a wrapper .ps1
    $runFile = $MacroFile
    $ext = [System.IO.Path]::GetExtension($MacroFile).ToLower()
    if ($ext -eq '.bas') {
        $wrapperDir = Join-Path $script:HubRoot 'sched_wrappers'
        if (-not (Test-Path $wrapperDir)) { [void](New-Item -ItemType Directory -Path $wrapperDir -Force) }
        $safeName = ($TaskName -replace '[^\w]','_')
        $wrapperPath = Join-Path $wrapperDir "$safeName.ps1"
        $wrapperCode = @"
# Auto-generated by MacroHub for scheduled task: $TaskName
# Runs VBA macro: $MacroFile
try {
    `$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    `$xl = New-Object -ComObject Excel.Application
    `$xl.Visible = `$true
}
`$content = Get-Content '$MacroFile' -Raw -Encoding ASCII
`$entryPoint = `$null
if (`$content -match 'Public\s+Sub\s+(\w+)') { `$entryPoint = `$matches[1] }
elseif (`$content -match 'Sub\s+(\w+)\s*\(') { `$entryPoint = `$matches[1] }
if (-not `$entryPoint) { Write-Error 'No Sub found in .bas file'; exit 1 }
`$wb = `$xl.ActiveWorkbook
if (-not `$wb) { Write-Error 'No workbook open'; exit 1 }
`$comp = `$wb.VBProject.VBComponents.Import('$MacroFile')
try { `$xl.Run("`$(`$wb.Name)!`$(`$comp.Name).`$entryPoint") }
finally { try { `$wb.VBProject.VBComponents.Remove(`$comp) } catch {} }
"@
        Set-Content -Path $wrapperPath -Value $wrapperCode -Encoding ASCII
        $runFile = $wrapperPath
        Write-ActivityLog "Generated VBA wrapper: $wrapperPath"
    }

    $svc = New-Object -ComObject Schedule.Service
    $svc.Connect()
    try { $svc.GetFolder('\MacroHub') } catch {
        $svc.GetFolder('\').CreateFolder('MacroHub')
    }
    $folder = $svc.GetFolder('\MacroHub')

    $def = $svc.NewTask(0)   # 0 = TASK_VALIDATE_ONLY flag; required by the API
    $def.RegistrationInfo.Description = "MacroHub scheduled: $TaskName -- File: $MacroFile"
    $def.Settings.Enabled = $true
    $def.Settings.StartWhenAvailable = $true   # run missed trigger when PC wakes
    $def.Settings.RunOnlyIfNetworkAvailable = $false
    $def.Settings.StopIfGoingOnBatteries   = $false
    $def.Settings.DisallowStartIfOnBatteries = $false

    $now = Get-Date
    $parts = $TriggerTime -split ':'
    $start = Get-Date -Year $now.Year -Month $now.Month -Day $now.Day `
                       -Hour ([int]$parts[0]) -Minute ([int]$parts[1]) -Second 0
    if ($start -lt $now) { $start = $start.AddDays(1) }

    if ($Frequency -eq 'Monthly') {
        # Monthly trigger (type 4)
        $trigger = $def.Triggers.Create(4)
        $trigger.StartBoundary = $start.ToString('yyyy-MM-ddTHH:mm:ss')
        $trigger.DaysOfMonth   = [Math]::Pow(2, $start.Day - 1)  # bitmask: bit N = day N+1
        $trigger.MonthsOfYear  = 4095  # all 12 months (bitmask: bits 0-11)
    } else {
        # Daily trigger (type 2); DaysInterval=7 simulates weekly frequency
        $trigger = $def.Triggers.Create(2)
        $trigger.StartBoundary = $start.ToString('yyyy-MM-ddTHH:mm:ss')
        if ($Frequency -eq 'Weekly') { $trigger.DaysInterval = 7 }
        else { $trigger.DaysInterval = 1 }
    }

    $action = $def.Actions.Create(0)   # 0 = TASK_ACTION_EXEC
    $action.Path = 'powershell.exe'
    $action.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$runFile`""
    $action.WorkingDirectory = $script:HubRoot

    # RegisterTaskDefinition flags: 6 = CREATE_OR_UPDATE, 3 = TASK_LOGON_INTERACTIVE_TOKEN
    $folder.RegisterTaskDefinition($TaskName, $def, 6, $null, $null, 3)
    Write-ActivityLog "Scheduled task: $TaskName ($Frequency at $TriggerTime) -> $runFile"
}

# Deletes the named task from the \MacroHub Task Scheduler folder.
function Remove-HubTask {
    param([string]$TaskName)
    try {
        $svc = New-Object -ComObject Schedule.Service
        $svc.Connect()
        $folder = $svc.GetFolder('\MacroHub')
        $folder.DeleteTask($TaskName, 0)
        Write-ActivityLog "Removed scheduled task: $TaskName"
    } catch {
        throw "Could not remove task: $_"
    }
}

# ============================================================
#  FAVORITES PERSISTENCE
# ============================================================
# Favorites are macro names that appear at the top of the Macros list with a [*]
# prefix.  They are persisted in favorites.json as a sorted unique string array.

# Loads the favorites list from JSON (or migrates from legacy CSV); returns an
# array of PSCustomObjects with a Name property.
function Load-Favorites {
    if (Test-Path $script:FavJson) {
        try {
            $loaded = Normalize-List (Get-Content $script:FavJson -Raw | ConvertFrom-Json)
            $out = @()
            foreach ($item in $loaded) {
                if ($item -is [string]) {
                    $out += [PSCustomObject]@{ Name = $item }
                } elseif ($item.PSObject.Properties['Name']) {
                    $name = [string]$item.Name
                    if ($name) { $out += [PSCustomObject]@{ Name = $name } }
                } else {
                    $name = [string]$item
                    if ($name) { $out += [PSCustomObject]@{ Name = $name } }
                }
            }
            return ,(Normalize-List $out)
        } catch {
            return ,@()
        }
    }
    if (Test-Path $script:FavCSV) {
        try {
            $legacy = @(Import-Csv $script:FavCSV | ForEach-Object { $_.Name })
            Save-Favorites $legacy
            return ,(@(@($legacy | ForEach-Object { [PSCustomObject]@{ Name = $_ } }) | Where-Object { $_.Name }))
        } catch { return ,@() }
    }
    return ,@()
}

function Save-Favorites([string[]]$Names) {
    if ($Names.Count -eq 0) {
        if (Test-Path $script:FavJson) { Remove-Item $script:FavJson -Force }
        return
    }
    @($Names | Sort-Object -Unique) | ConvertTo-Json | Set-Content $script:FavJson -Encoding UTF8
}

function Toggle-Favorite([string]$MacroName) {
    $favs = Load-Favorites
    $names = @($favs | ForEach-Object { $_.Name })
    if ($names -contains $MacroName) {
        $names = @($names | Where-Object { $_ -ne $MacroName })
    } else {
        $names += $MacroName
    }
    Save-Favorites $names
}

function Get-FavoriteNames {
    $favs = Load-Favorites
    return ,@($favs | ForEach-Object { $_.Name })
}

# ============================================================
#  GLOBAL SEARCH
# ============================================================
# Searches macros, templates, QTasks todos, and scheduled tasks for $query and
# returns a list of result objects with Tab, Item, and Detail properties.
function Search-AllTabs([string]$query) {
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $q = $query.ToLower()
    # Macros
    $macros = Get-MacroFiles
    foreach ($m in $macros) {
        if ($m.Name.ToLower() -like "*$q*") {
            $results.Add([PSCustomObject]@{ Tab='Macros'; Item=$m.Name; Detail='Macro file' })
        }
    }
    # Templates
    $templates = Load-Templates
    foreach ($t in $templates) {
        if ($t.Name.ToLower() -like "*$q*" -or $t.Content.ToLower() -like "*$q*") {
            $results.Add([PSCustomObject]@{ Tab='Templates'; Item=$t.Name; Detail='Template' })
        }
    }
    # Todos
    $todos = Load-QsTodos
    foreach ($t in $todos) {
        if ($t.OriginalName.ToLower() -like "*$q*" -or $t.RelFolder.ToLower() -like "*$q*") {
            $results.Add([PSCustomObject]@{ Tab='QTasks'; Item=$t.OriginalName; Detail="Folder: $($t.RelFolder) | $($t.Status)" })
        }
    }
    # Scheduled tasks
    try {
        $tasks = Get-HubTasks
        foreach ($t in $tasks) {
            if ($t.Name.ToLower() -like "*$q*") {
                $results.Add([PSCustomObject]@{ Tab='Scheduler'; Item=$t.Name; Detail="State: $($t.State)" })
            }
        }
    } catch {}
    return ,$results
}

# ============================================================
#  CLIPBOARD DEFAULT SETTINGS
# ============================================================
# Saves/loads the last-used workbook, sheet, cell, and timestamp settings so
# they are restored on the next MacroHub launch.

# Loads saved clipboard defaults from clip_defaults.json; returns $null if none saved.
function Load-ClipDefaults {
    if (Test-Path $script:ClipDefaultsJson) {
        try { return (Get-Content $script:ClipDefaultsJson -Raw | ConvertFrom-Json) } catch {}
    }
    return $null
}

function Save-ClipDefaults {
    param([string]$Workbook, [string]$Sheet, [string]$Cell,
          [bool]$Timestamp, [string]$DateCell, [string]$TimeCell)
    $obj = @{
        Workbook  = $Workbook
        Sheet     = $Sheet
        Cell      = $Cell
        Timestamp = $Timestamp
        DateCell  = $DateCell
        TimeCell  = $TimeCell
        SavedOn   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
    $obj | ConvertTo-Json | Set-Content $script:ClipDefaultsJson -Encoding UTF8
}

# ============================================================
#  CLIPBOARD PACKET CAPTURE + PASTE ENGINE
# ============================================================
# Get-ClipboardPacket:
#   Captures the current Windows clipboard as a replayable packet.
#   Stores: DataObject (IDataObject), plain Text, format list, SHA-1 Signature, and
#   ExcelSource metadata (workbook/sheet/address) when Excel has an active copy selection.
#
# Paste-ClipboardPacketToExcel:
#   Path 1 — Direct Range.Copy(Destination) via ExcelSource metadata (copies all formatting).
#   Path 2 — Native clipboard paste: restores DataObject → PasteSpecial(xlPasteAll).
#   Path 3 — Plain-text Value2 fallback + PasteSpecial(xlPasteFormats) overlay from source.
#   Post-paste house-style rules are only applied when no Excel formatting was preserved.
# ============================================================
# These functions handle clipboard reading, dimension detection, matrix
# conversion, and the multi-strategy paste pipeline to Excel.

# Reads the current clipboard and returns a packet object containing the raw
# IDataObject, extracted text, format list, SHA-1 signature (for change detection),
# and Excel source metadata (workbook/sheet/address) when a copy range is active.
# Retries up to 5 times with short sleeps to handle transient clipboard lock errors.
$script:ClipSlotIdx = 0

function Get-ClipboardPacket {
    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        try {
            $dataObj = [System.Windows.Clipboard]::GetDataObject()
            if (-not $dataObj) { return $null }
            $txt = ''
            if ($dataObj.GetDataPresent([System.Windows.DataFormats]::UnicodeText)) {
                $txt = [string]$dataObj.GetData([System.Windows.DataFormats]::UnicodeText)
            } elseif ($dataObj.GetDataPresent([System.Windows.DataFormats]::Text)) {
                $txt = [string]$dataObj.GetData([System.Windows.DataFormats]::Text)
            }
            $formats = ''
            try { $formats = @($dataObj.GetFormats()) -join '|' } catch {}
            $sig = Get-TextSha1 ([string]$txt + '|' + [string]$formats)
            $excelSource = $null
            try {
                $xlCandidates = @()
                foreach ($sess in @('Main', 'Navigator', 'Any')) {
                    try {
                        $app = Get-ExcelApp -Session $sess
                        if ($app) { $xlCandidates += $app }
                    } catch {}
                }
                foreach ($xlApp in $xlCandidates) {
                    $cutCopy = 0
                    try { $cutCopy = [int]$xlApp.CutCopyMode } catch { $cutCopy = 0 }
                    # CutCopyMode: 1 = xlCopy (marching ants active), 2 = xlCut; skip if no copy in progress
                    if ($cutCopy -ne 1) { continue }

                    $sel = $null
                    try { if ($xlApp.ActiveWindow) { $sel = $xlApp.ActiveWindow.RangeSelection } } catch {}
                    if (-not $sel) {
                        try { $sel = $xlApp.Selection } catch {}
                    }
                    if (-not $sel -or -not $sel.Worksheet -or -not $sel.Worksheet.Parent) { continue }

                    $wsSel = $sel.Worksheet
                    $wbSel = $wsSel.Parent
                    $rowsSel = 1
                    $colsSel = 1
                    try { $rowsSel = [Math]::Max(1, [int]$sel.Rows.Count) } catch {}
                    try { $colsSel = [Math]::Max(1, [int]$sel.Columns.Count) } catch {}

                    $candidate = [PSCustomObject]@{
                        WorkbookName = [string]$wbSel.Name
                        SheetName    = [string]$wsSel.Name
                        Address      = [string]$sel.Address($true, $true)
                        Rows         = $rowsSel
                        Cols         = $colsSel
                    }
                    if (-not $excelSource) {
                        $excelSource = $candidate
                    } else {
                        $curArea = [int]$excelSource.Rows * [int]$excelSource.Cols
                        $newArea = [int]$candidate.Rows * [int]$candidate.Cols
                        if ($newArea -gt $curArea) { $excelSource = $candidate }
                    }
                }
            } catch {}
            return [PSCustomObject]@{
                DataObject  = $dataObj
                Text        = $txt
                Formats     = $formats
                Signature   = $sig
                ExcelSource = $excelSource
                CapturedOn  = (Get-Date -f 'yyyy-MM-dd HH:mm:ss')
            }
        } catch {
            Start-Sleep -Milliseconds 60
        }
    }
    return $null
}

# Parses tab-delimited clipboard text and returns an object with Rows and Cols
# counts (trailing blank rows are trimmed).
function Get-ClipboardDimensions {
    param([string]$Text)
    if (-not $Text) { return [PSCustomObject]@{ Rows = 1; Cols = 1 } }
    $rows = @($Text -split "`r?`n")
    while ($rows.Count -gt 1 -and [string]::IsNullOrWhiteSpace($rows[-1])) {
        $rows = $rows[0..($rows.Count - 2)]
    }
    if ($rows.Count -eq 0) { $rows = @('') }
    $maxCols = 1
    foreach ($ln in $rows) {
        $cols = @($ln -split "`t", -1).Count
        if ($cols -gt $maxCols) { $maxCols = $cols }
    }
    return [PSCustomObject]@{ Rows = $rows.Count; Cols = $maxCols }
}

# Converts tab-delimited clipboard text into a 2-D object array suitable for
# bulk assignment to Excel's Range.Value2.
# Returns a 2-D array dimensioned [1..Rows, 1..Cols] (1-based lower bounds).
function Convert-ClipboardTextToMatrix {
    param(
        [string]$Text,
        [int]$Rows,
        [int]$Cols
    )
    $safeRows = [Math]::Max(1, [int]$Rows)
    $safeCols = [Math]::Max(1, [int]$Cols)
    # Use 1-based lower bounds: Excel's COM Value2 setter expects a 1-based SAFEARRAY.
    # A 0-based array causes Excel to misread the dimensions and repeat the first cell.
    $matrix = [System.Array]::CreateInstance([object], @($safeRows, $safeCols), @(1, 1))
    if (-not $Text) { return $matrix }

    $lines = @($Text -split "`r?`n")
    while ($lines.Count -gt 1 -and [string]::IsNullOrWhiteSpace($lines[-1])) {
        $lines = $lines[0..($lines.Count - 2)]
    }
    $rowMax = [Math]::Min($safeRows, $lines.Count)
    for ($ri = 0; $ri -lt $rowMax; $ri++) {
        $parts = @($lines[$ri] -split "`t", -1)
        $colMax = [Math]::Min($safeCols, $parts.Count)
        for ($ci = 0; $ci -lt $colMax; $ci++) {
            $matrix[$ri + 1, $ci + 1] = $parts[$ci]
        }
    }
    return $matrix
}

function Test-RangeHasCellData {
    param($RangeObj)
    try {
        $v = $RangeObj.Value2
    } catch {
        return $false
    }
    if ($v -is [array]) {
        foreach ($item in $v) {
            if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace([string]$item)) { return $true }
        }
        return $false
    }
    return ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v))
}

function Test-ExcelFormatIsDateOrPercent {
    param([string]$NumberFormat)
    if ([string]::IsNullOrWhiteSpace($NumberFormat)) { return $false }
    $fmt = $NumberFormat.ToLower()
    if ($fmt.Contains('%')) { return $true }
    if ($fmt -match 'am/pm') { return $true }
    # Strip quoted literals and bracket sections (colors/locales/conditions) before token scan.
    $clean = $fmt -replace '".*?"','' -replace '\[[^\]]+\]',''
    if ($clean -match 'y{2,4}') { return $true }
    if ($clean -match 'd{1,4}') { return $true }
    if ($clean -match 'h{1,2}') { return $true }
    if ($clean -match 's{1,2}') { return $true }
    if ($clean -match 'm{1,4}') { return $true }
    return $false
}

function Test-ExcelTextLooksDateOrPercent {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $t = $Text.Trim()
    if ($t -match '%') { return $true }
    if ($t -match '^\d{1,4}[/-]\d{1,2}[/-]\d{1,4}(\s+\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM)?)?$') { return $true }
    if ($t -match '^[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{2,4}$') { return $true }
    return $false
}

function Try-ParseExcelNumericText {
    param(
        [string]$Text,
        [ref]$OutNumber
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $styles = [System.Globalization.NumberStyles]::Float -bor [System.Globalization.NumberStyles]::AllowThousands
    $n = 0.0
    if ([double]::TryParse($Text.Trim(), $styles, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$n)) {
        $OutNumber.Value = $n
        return $true
    }
    if ([double]::TryParse($Text.Trim(), $styles, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
        $OutNumber.Value = $n
        return $true
    }
    return $false
}

function Apply-PostPasteCellRules {
    param(
        $Worksheet,
        [int]$StartRow,
        [int]$StartCol,
        [int]$Rows,
        [int]$Cols
    )
    if (-not $Worksheet) { return }
    $safeRows = [Math]::Max(1, [int]$Rows)
    $safeCols = [Math]::Max(1, [int]$Cols)

    # Keep header row wrapped.
    try {
        $hdr = $Worksheet.Range(
            $Worksheet.Cells.Item($StartRow, $StartCol),
            $Worksheet.Cells.Item($StartRow, $StartCol + $safeCols - 1)
        )
        $hdr.WrapText = $true
    } catch {}

    if ($safeRows -le 1) { return }

    # All rows below header: unwrap text.
    try {
        $unwrap = $Worksheet.Range(
            $Worksheet.Cells.Item($StartRow + 1, $StartCol),
            $Worksheet.Cells.Item($StartRow + $safeRows - 1, $StartCol + $safeCols - 1)
        )
        $unwrap.WrapText = $false
    } catch {}

    # Rows below header: apply comma format to numeric cells except date/percent formats.
    for ($ri = $StartRow + 1; $ri -le ($StartRow + $safeRows - 1); $ri++) {
        for ($ci = $StartCol; $ci -le ($StartCol + $safeCols - 1); $ci++) {
            try {
                $cell = $Worksheet.Cells.Item($ri, $ci)
                $val = $cell.Value2
                if ($null -eq $val) { continue }
                $fmt = [string]$cell.NumberFormat
                if (Test-ExcelFormatIsDateOrPercent -NumberFormat $fmt) { continue }

                $displayText = ''
                try { $displayText = [string]$cell.Text } catch {}
                if (Test-ExcelTextLooksDateOrPercent -Text $displayText) { continue }

                $isNumeric = $false
                $parsed = 0.0
                if ($val -is [DateTime]) {
                    continue
                } elseif (($val -is [byte]) -or ($val -is [int16]) -or ($val -is [int32]) -or ($val -is [int64]) -or
                          ($val -is [uint16]) -or ($val -is [uint32]) -or ($val -is [uint64]) -or
                          ($val -is [single]) -or ($val -is [double]) -or ($val -is [decimal])) {
                    $parsed = [double]$val
                    $isNumeric = $true
                } else {
                    $txt = ''
                    try { $txt = [string]$val } catch {}
                    if ($txt -eq 'System.__ComObject') { $txt = '' }
                    if ([string]::IsNullOrWhiteSpace($txt)) { $txt = $displayText }
                    if (Test-ExcelTextLooksDateOrPercent -Text $txt) { continue }
                    if (Try-ParseExcelNumericText -Text $txt -OutNumber ([ref]$parsed)) {
                        $isNumeric = $true
                        try { $cell.Value2 = $parsed } catch {}
                    }
                }
                if (-not $isNumeric) { continue }
                $cell.NumberFormat = '#,##0.00'
            } catch {}
        }
    }
}

# Pastes a captured clipboard packet into the target Excel cell using a
# three-tier strategy: (1) direct source range Copy if Excel source metadata is
# available; (2) native PasteSpecial/Paste via the live clipboard; (3) bulk
# Value2 array assignment for plain-text fallback.
# When $AddTimestamp is true, writes date/time into rows above the paste target.
function Paste-ClipboardPacketToExcel {
    param(
        [object]$Packet,
        [string]$WorkbookName,
        [string]$SheetName,
        [string]$CellAddress,
        [bool]$AddTimestamp,
        [string]$DateCell,
        [string]$TimeCell
    )
    if (-not $Packet) { throw 'Clipboard slot is empty — nothing to paste.' }

    # Guard: packet has usable content
    $hasText   = $Packet.Text -and $Packet.Text.Trim().Length -gt 0
    $hasData   = $null -ne $Packet.DataObject
    $hasSrc    = $null -ne $Packet.ExcelSource
    if (-not $hasText -and -not $hasData -and -not $hasSrc) {
        throw 'Clipboard packet contains no text, data, or Excel range — nothing to paste.'
    }

    # Guard: Excel session
    $xl = $null
    try { $xl = Get-ExcelApp -Session Main } catch {}
    if (-not $xl) { throw 'Excel is not open or not responding.' }

    # Guard: workbook
    $wbObj = $null
    try { $wbObj = $xl.Workbooks.Item($WorkbookName) } catch {}
    if (-not $wbObj) { throw "Workbook '$WorkbookName' is not open." }

    # Guard: read-only
    $isReadOnly = $false
    try { $isReadOnly = [bool]$wbObj.ReadOnly } catch {}
    if ($isReadOnly) { throw "Workbook '$WorkbookName' is read-only — cannot paste." }

    # Guard: worksheet
    $ws = $null
    try { $ws = $wbObj.Worksheets.Item($SheetName) } catch {}
    if (-not $ws) { throw "Sheet '$SheetName' not found in '$WorkbookName'." }

    # Guard: target cell
    $target = $null
    try { $target = $ws.Range($CellAddress) } catch {}
    if (-not $target) { throw "Cell address '$CellAddress' is not valid." }

    $r = 1; $c = 1
    try { $r = [int]$target.Row    } catch {}
    try { $c = [int]$target.Column } catch {}
    $dims = Get-ClipboardDimensions -Text $Packet.Text
    $rows = [Math]::Max(1, [int]$dims.Rows)
    $cols = [Math]::Max(1, [int]$dims.Cols)
    $textRows = $rows
    $textCols = $cols

    $dest = $null
    try {
        $dest = $ws.Range(
            $ws.Cells.Item($r, $c),
            $ws.Cells.Item($r + $rows - 1, $c + $cols - 1)
        )
    } catch { throw "Could not build destination range at ${CellAddress}: $($_.Exception.Message)" }
    if (-not $dest) { throw "Destination range is null for address '$CellAddress'." }

    # Warn (but continue) if target contains merged cells — paste may partially succeed.
    $hasMerge = $false
    try { $hasMerge = [bool]$dest.MergeCells } catch {}
    if ($hasMerge) {
        Set-Status "Warning: target range has merged cells — paste results may vary." '#E0A050'
    }

    # Required order: clear values, clear formats, then paste.
    try { $dest.ClearContents() } catch {}
    try { $dest.ClearFormats()  } catch {}

    # Track whether native Excel formatting was preserved (suppresses post-paste style rules).
    $usedClipboardPaste = $false
    $formattingApplied  = $false
    $source = $Packet.ExcelSource

    # ── PATH 1: Direct Excel range copy — copies values + ALL formatting in one shot ──
    # Requires source workbook still open in any Excel session.
    if ($source -and $source.WorkbookName -and $source.SheetName -and $source.Address) {
        try {
            # Try all known Excel sessions so cross-session copies (e.g. Navigator) work.
            $srcWb = $null
            foreach ($srcSess in @('Main', 'Navigator', 'Any')) {
                try {
                    $srcXl = Get-ExcelApp -Session $srcSess
                    if ($srcXl) { $srcWb = $srcXl.Workbooks.Item([string]$source.WorkbookName); break }
                } catch {}
            }
            if (-not $srcWb) { throw 'Source workbook not found in any Excel session.' }
            $srcWs    = $srcWb.Worksheets.Item([string]$source.SheetName)
            $srcRange = $srcWs.Range([string]$source.Address)
            $srcRows  = [Math]::Max(1, [int]$srcRange.Rows.Count)
            $srcCols  = [Math]::Max(1, [int]$srcRange.Columns.Count)
            # Some Excel sessions report Selection as 1×1 even when clipboard holds a larger range.
            # Expand from the source anchor cell when text dimensions indicate a larger copy.
            if ($srcRows -eq 1 -and $srcCols -eq 1 -and ($textCols -gt 1 -or $textRows -gt 1)) {
                $maxRow = [Math]::Min(1048576, [int]$srcRange.Row + $textRows - 1)
                $maxCol = [Math]::Min(16384,   [int]$srcRange.Column + $textCols - 1)
                $srcRange = $srcWs.Range(
                    $srcWs.Cells.Item([int]$srcRange.Row, [int]$srcRange.Column),
                    $srcWs.Cells.Item($maxRow, $maxCol)
                )
                $srcRows = [Math]::Max(1, [int]$srcRange.Rows.Count)
                $srcCols = [Math]::Max(1, [int]$srcRange.Columns.Count)
            }
            $rows = $srcRows
            $cols = $srcCols
            $dest = $ws.Range(
                $ws.Cells.Item($r, $c),
                $ws.Cells.Item($r + $rows - 1, $c + $cols - 1)
            )
            $dest.ClearContents()
            $dest.ClearFormats()
            $srcRange.Copy($dest)           # Range.Copy(Destination) copies values + all formatting
            $usedClipboardPaste = $true
            $formattingApplied  = $true
        } catch {}
    }

    # ── PATH 2: Native clipboard paste (preserves Excel formats when clipboard is live) ──
    $packetSig = [string]$Packet.Signature
    if (-not $packetSig) { $packetSig = Get-TextSha1 ([string]$Packet.Text + '|' + [string]$Packet.Formats) }
    $liveClipboardSame = $false
    if (-not $usedClipboardPaste) {
        try {
            $livePacket = Get-ClipboardPacket
            if ($livePacket -and [string]$livePacket.Signature -eq $packetSig) {
                $liveClipboardSame = $true
            }
        } catch {}
    }

    $maxPasteAttempts = 4
    for ($attempt = 1; $attempt -le $maxPasteAttempts -and -not $usedClipboardPaste; $attempt++) {
        try {
            $shapeBefore = 0
            try { $shapeBefore = [int]$ws.Shapes.Count } catch {}
            if (-not $liveClipboardSame) {
                if ($Packet.DataObject) {
                    [System.Windows.Clipboard]::SetDataObject($Packet.DataObject, $true)
                } elseif ($Packet.Text) {
                    [System.Windows.Clipboard]::SetText($Packet.Text)
                } else {
                    break
                }
            }

            Start-Sleep -Milliseconds (40 * $attempt)
            $wbObj.Activate()
            $ws.Activate()
            $target.Select()
            try {
                $target.PasteSpecial(-4104)  # xlPasteAll — includes values, formats, and formulas
            } catch {
                $ws.Paste($target)
            }

            # Guard against image-only paste (metafile) when the intent was tabular cells.
            if ($Packet.Text -and $Packet.Text.Length -gt 0) {
                $shapeAfter = $shapeBefore
                try { $shapeAfter = [int]$ws.Shapes.Count } catch {}
                $hasCellData = Test-RangeHasCellData -RangeObj $dest
                if ($shapeAfter -gt $shapeBefore -and -not $hasCellData) {
                    for ($si = $shapeAfter; $si -gt $shapeBefore; $si--) {
                        try { $ws.Shapes.Item($si).Delete() } catch {}
                    }
                    throw 'Clipboard paste resolved to an image object. Retrying as cells.'
                }
                if (-not $hasCellData) {
                    throw 'Clipboard paste produced no cell data. Retrying.'
                }
            }

            $usedClipboardPaste = $true
            $formattingApplied  = $true     # Native xlPasteAll includes formatting
        } catch {
            # If live clipboard attempt failed, fall back to packet replay attempts.
            if ($liveClipboardSame) { $liveClipboardSame = $false }
            Start-Sleep -Milliseconds (60 * $attempt)
        }
    }

    # ── PATH 3: Plain-text fallback — values only via 2D array assignment ──
    if (-not $usedClipboardPaste) {
        # Fallback for plain text payloads: bulk-assign the entire matrix in one COM call,
        # which is far faster than writing cells one at a time for large ranges.
        $matrix = Convert-ClipboardTextToMatrix -Text $Packet.Text -Rows $rows -Cols $cols
        $dest.Value2 = $matrix

        # ── PATH 3b: Overlay source formatting after text paste ──
        # If the source workbook is still open, re-copy the source range to clipboard and
        # paste only formats (xlPasteFormats = -4122) on top of the already-pasted values.
        # This preserves fonts, fill colors, borders, and number formats even when the
        # DataObject clipboard path was unavailable.
        if ($source -and $source.WorkbookName -and $source.SheetName -and $source.Address) {
            try {
                $fmtWb = $null
                foreach ($fmtSess in @('Main', 'Navigator', 'Any')) {
                    try {
                        $fmtXl = Get-ExcelApp -Session $fmtSess
                        if ($fmtXl) { $fmtWb = $fmtXl.Workbooks.Item([string]$source.WorkbookName); break }
                    } catch {}
                }
                if ($fmtWb) {
                    $fmtWs    = $fmtWb.Worksheets.Item([string]$source.SheetName)
                    $fmtRange = $fmtWs.Range([string]$source.Address)
                    # Expand 1×1 anchor to match text dimensions.
                    $fmtRows = [Math]::Max(1, [int]$fmtRange.Rows.Count)
                    $fmtCols = [Math]::Max(1, [int]$fmtRange.Columns.Count)
                    if ($fmtRows -eq 1 -and $fmtCols -eq 1 -and ($textCols -gt 1 -or $textRows -gt 1)) {
                        $maxRow = [Math]::Min(1048576, [int]$fmtRange.Row + $textRows - 1)
                        $maxCol = [Math]::Min(16384,   [int]$fmtRange.Column + $textCols - 1)
                        $fmtRange = $fmtWs.Range(
                            $fmtWs.Cells.Item([int]$fmtRange.Row, [int]$fmtRange.Column),
                            $fmtWs.Cells.Item($maxRow, $maxCol)
                        )
                    }
                    $fmtRange.Copy()                    # Puts full Excel range (with formats) on clipboard
                    $dest.PasteSpecial(-4122)           # xlPasteFormats — applies only formatting, not values
                    try { $xl.CutCopyMode = $false } catch {}
                    $formattingApplied = $true
                }
            } catch {}
        }
    }

    # Only apply post-paste house-style rules (unwrap data rows, apply comma number format)
    # when no Excel source formatting was preserved. When formatting comes from Excel, we
    # respect the original cell styles rather than overwriting them.
    if (-not $formattingApplied) {
        Apply-PostPasteCellRules -Worksheet $ws -StartRow $r -StartCol $c -Rows $rows -Cols $cols
    }

    if ($AddTimestamp) {
        $effDateCell = if ([string]::IsNullOrWhiteSpace($DateCell)) { 'A3' } else { $DateCell.Trim() }
        $effTimeCell = if ([string]::IsNullOrWhiteSpace($TimeCell)) { 'A2' } else { $TimeCell.Trim() }
        if ($effDateCell) {
            try {
                $dtRange = $null
                try { $dtRange = $ws.Range($effDateCell) } catch {}
                if ($dtRange) {
                    $dtRange.Value2 = (Get-Date -Format 'MM/dd/yyyy')
                } else {
                    Set-Status "Timestamp: date cell '$effDateCell' is not a valid address — skipped." '#E0A050'
                }
            } catch {
                Set-Status "Timestamp (date) write failed for '$effDateCell': $($_.Exception.Message)" '#E0A050'
            }
        }
        if ($effTimeCell) {
            try {
                $tmRange = $null
                try { $tmRange = $ws.Range($effTimeCell) } catch {}
                if ($tmRange) {
                    $tmRange.Value2 = (Get-Date -Format 'hh:mm tt')
                } else {
                    Set-Status "Timestamp: time cell '$effTimeCell' is not a valid address — skipped." '#E0A050'
                }
            } catch {
                Set-Status "Timestamp (time) write failed for '$effTimeCell': $($_.Exception.Message)" '#E0A050'
            }
        }
    }
}

# ============================================================
#  DYNAMIC CLIPBOARD SLOT UI BUILDER
# ============================================================
# Add-ClipSlotUI: Programmatically builds one clipboard slot card and appends it to
# ClipSlotsPanel. Each card contains a text preview box, per-slot sheet/cell/timestamp
# controls, and Record/Paste/Remove buttons. State is tracked in $script:ClipSlots[$tag].
# Each slot is an independent recording unit with its own DispatcherTimer,
# sheet/cell/timestamp controls, and paste button.
# The slot's state object is stored in $script:ClipSlots keyed by "Slot_N".
function Add-ClipSlotUI {
    param(
        $Panel,
        $WbCombo,
        [string]$DefaultSheet,
        [string]$DefaultCell,
        [bool]$DefaultTimestamp,
        [string]$DefaultDateCell,
        [string]$DefaultTimeCell,
        $CountLabel
    )
    $script:ClipSlotIdx++
    $num = $script:ClipSlotIdx
    $tag = "Slot_$num"
    $script:ClipSlots[$tag] = [PSCustomObject]@{
        Packet         = $null
        Text           = ''
        IsRecording    = $false
        LastSig        = ''
        TextBoxCtrl    = $null
        SheetCombo     = $null
        CellBox        = $null
        TimestampChk   = $null
        DateCellBox    = $null
        TimeCellBox    = $null
    }
    # Capture a direct reference to the hashtable so WPF event closures (which run
    # in a child scope) can still read/write the live slot state after Add_Click fires.
    $clipSlotsRef = $script:ClipSlots

    # Card border
    $border = [System.Windows.Controls.Border]::new()
    $border.Tag = $tag
    $border.Background   = (HexBrush '#252528')
    $border.BorderBrush  = (HexBrush '#3E3E42')
    $border.BorderThickness  = [System.Windows.Thickness]::new(1)
    $border.CornerRadius     = [System.Windows.CornerRadius]::new(6)
    $border.Padding          = [System.Windows.Thickness]::new(14,10,14,10)
    $border.Margin           = [System.Windows.Thickness]::new(0,0,0,8)

    $stack = [System.Windows.Controls.StackPanel]::new()

    # Header row: label + recording indicator + remove button
    $hdr = [System.Windows.Controls.Grid]::new()
    $colA = [System.Windows.Controls.ColumnDefinition]::new()
    $colA.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $colB = [System.Windows.Controls.ColumnDefinition]::new()
    $colB.Width = [System.Windows.GridLength]::Auto
    $colC = [System.Windows.Controls.ColumnDefinition]::new()
    $colC.Width = [System.Windows.GridLength]::Auto
    [void]($hdr.ColumnDefinitions.Add($colA))
    [void]($hdr.ColumnDefinitions.Add($colB))
    [void]($hdr.ColumnDefinitions.Add($colC))

    $lbl = [System.Windows.Controls.TextBlock]::new()
    $lbl.Text = "Slot $num"
    $lbl.Foreground = (HexBrush '#6E6E6E')
    $lbl.FontSize = 10
    $lbl.FontWeight = [System.Windows.FontWeights]::SemiBold
    $lbl.VerticalAlignment = 'Center'
    [System.Windows.Controls.Grid]::SetColumn($lbl, 0)
    [void]($hdr.Children.Add($lbl))

    $recTxt = [System.Windows.Controls.TextBlock]::new()
    $recTxt.Text = 'REC OFF'
    $recTxt.Foreground = (HexBrush '#6E6E6E')
    $recTxt.FontSize = 10
    $recTxt.FontWeight = [System.Windows.FontWeights]::Bold
    $recTxt.VerticalAlignment = 'Center'
    $recTxt.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    $recTxt.Visibility = 'Visible'
    [System.Windows.Controls.Grid]::SetColumn($recTxt, 1)
    [void]($hdr.Children.Add($recTxt))

    $removeBtn = [System.Windows.Controls.Button]::new()
    $removeBtn.Content = 'X'
    $removeBtn.Foreground = (HexBrush '#E05050')
    $removeBtn.Background = (HexBrush '#2E1F1F')
    $removeBtn.BorderThickness = [System.Windows.Thickness]::new(0)
    $removeBtn.Padding = [System.Windows.Thickness]::new(6,2,6,2)
    $removeBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $removeBtn.Tag = $tag
    [System.Windows.Controls.Grid]::SetColumn($removeBtn, 2)
    [void]($hdr.Children.Add($removeBtn))

    [void]($stack.Children.Add($hdr))

    # TextBox
    $tb = [System.Windows.Controls.TextBox]::new()
    $tb.Height = 60
    $tb.AcceptsReturn = $true
    $tb.VerticalScrollBarVisibility = 'Auto'
    $tb.TextWrapping = 'Wrap'
    $tb.Margin = [System.Windows.Thickness]::new(0,6,0,0)
    $tb.Tag = $tag
    [void]($stack.Children.Add($tb))

    # Button + per-slot target row
    $btnRow = [System.Windows.Controls.StackPanel]::new()
    $btnRow.Orientation = 'Horizontal'
    $btnRow.Margin = [System.Windows.Thickness]::new(0,8,0,0)

    $copyBtn = [System.Windows.Controls.Button]::new()
    $copyBtn.Content = 'Record Clipboard'
    $copyBtn.Background = (HexBrush '#3E3E42')
    $copyBtn.Foreground = (HexBrush '#FFFFFF')
    $copyBtn.BorderThickness = [System.Windows.Thickness]::new(0)
    $copyBtn.Padding = [System.Windows.Thickness]::new(10,4,10,4)
    $copyBtn.Margin  = [System.Windows.Thickness]::new(0,0,6,0)
    $copyBtn.Cursor  = [System.Windows.Input.Cursors]::Hand
    [void]($btnRow.Children.Add($copyBtn))

    $copyOutBtn = [System.Windows.Controls.Button]::new()
    $copyOutBtn.Content = 'Copy Slot'
    $copyOutBtn.Background = (HexBrush '#2F3A30')
    $copyOutBtn.Foreground = (HexBrush '#FFFFFF')
    $copyOutBtn.BorderThickness = [System.Windows.Thickness]::new(0)
    $copyOutBtn.Padding = [System.Windows.Thickness]::new(10,4,10,4)
    $copyOutBtn.Margin  = [System.Windows.Thickness]::new(0,0,8,0)
    $copyOutBtn.Cursor  = [System.Windows.Input.Cursors]::Hand
    [void]($btnRow.Children.Add($copyOutBtn))

    $slotSheetLbl = [System.Windows.Controls.TextBlock]::new()
    $slotSheetLbl.Text = 'Sheet'
    $slotSheetLbl.Foreground = (HexBrush '#B0B0B0')
    $slotSheetLbl.FontSize = 11
    $slotSheetLbl.VerticalAlignment = 'Center'
    $slotSheetLbl.Margin = [System.Windows.Thickness]::new(0,0,4,0)
    [void]($btnRow.Children.Add($slotSheetLbl))

    $slotSheetCombo = [System.Windows.Controls.ComboBox]::new()
    $slotSheetCombo.Width = 170
    $slotSheetCombo.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    [void]($btnRow.Children.Add($slotSheetCombo))

    $slotCellLbl = [System.Windows.Controls.TextBlock]::new()
    $slotCellLbl.Text = 'Cell'
    $slotCellLbl.Foreground = (HexBrush '#B0B0B0')
    $slotCellLbl.FontSize = 11
    $slotCellLbl.VerticalAlignment = 'Center'
    $slotCellLbl.Margin = [System.Windows.Thickness]::new(0,0,4,0)
    [void]($btnRow.Children.Add($slotCellLbl))

    $slotCellBox = [System.Windows.Controls.TextBox]::new()
    $slotCellBox.Width = 72
    $slotCellBox.Text = if ([string]::IsNullOrWhiteSpace($DefaultCell)) { 'A1' } else { [string]$DefaultCell }
    $slotCellBox.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    [void]($btnRow.Children.Add($slotCellBox))

    $slotTsChk = [System.Windows.Controls.CheckBox]::new()
    $slotTsChk.Content = 'Timestamp'
    $slotTsChk.Foreground = (HexBrush '#B0B0B0')
    $slotTsChk.IsChecked = [bool]$DefaultTimestamp
    $slotTsChk.VerticalAlignment = 'Center'
    $slotTsChk.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    [void]($btnRow.Children.Add($slotTsChk))

    $slotDateLbl = [System.Windows.Controls.TextBlock]::new()
    $slotDateLbl.Text = 'Date→'
    $slotDateLbl.Foreground = (HexBrush '#FFFFFF')
    $slotDateLbl.FontSize = 11
    $slotDateLbl.VerticalAlignment = 'Center'
    $slotDateLbl.Margin = [System.Windows.Thickness]::new(0,0,4,0)
    $slotDateLbl.ToolTip = 'Cell address where the date timestamp will be written (e.g. A3)'
    [void]($btnRow.Children.Add($slotDateLbl))

    $slotDateCellBox = [System.Windows.Controls.TextBox]::new()
    $slotDateCellBox.Width = 60
    $slotDateCellBox.TextAlignment = 'Center'
    $slotDateCellBox.Text = if ([string]::IsNullOrWhiteSpace($DefaultDateCell)) { 'A3' } else { [string]$DefaultDateCell }
    $slotDateCellBox.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    $slotDateCellBox.ToolTip = 'Cell address where the date timestamp will be written (e.g. A3)'
    [void]($btnRow.Children.Add($slotDateCellBox))

    $slotTimeLbl = [System.Windows.Controls.TextBlock]::new()
    $slotTimeLbl.Text = 'Time→'
    $slotTimeLbl.Foreground = (HexBrush '#FFFFFF')
    $slotTimeLbl.FontSize = 11
    $slotTimeLbl.VerticalAlignment = 'Center'
    $slotTimeLbl.Margin = [System.Windows.Thickness]::new(0,0,4,0)
    $slotTimeLbl.ToolTip = 'Cell address where the time timestamp will be written (e.g. A2)'
    [void]($btnRow.Children.Add($slotTimeLbl))

    $slotTimeCellBox = [System.Windows.Controls.TextBox]::new()
    $slotTimeCellBox.Width = 60
    $slotTimeCellBox.TextAlignment = 'Center'
    $slotTimeCellBox.Text = if ([string]::IsNullOrWhiteSpace($DefaultTimeCell)) { 'A2' } else { [string]$DefaultTimeCell }
    $slotTimeCellBox.Margin = [System.Windows.Thickness]::new(0,0,10,0)
    $slotTimeCellBox.ToolTip = 'Cell address where the time timestamp will be written (e.g. A2)'
    [void]($btnRow.Children.Add($slotTimeCellBox))

    $pasteBtn = [System.Windows.Controls.Button]::new()
    $pasteBtn.Content = 'Paste to Excel'
    $pasteBtn.Background = (HexBrush '#4C9FE6')
    $pasteBtn.Foreground = (HexBrush '#FFFFFF')
    $pasteBtn.BorderThickness = [System.Windows.Thickness]::new(0)
    $pasteBtn.Padding = [System.Windows.Thickness]::new(10,4,10,4)
    $pasteBtn.Cursor  = [System.Windows.Input.Cursors]::Hand
    [void]($btnRow.Children.Add($pasteBtn))

    [void]($stack.Children.Add($btnRow))
    $border.Child = $stack
    [void]($Panel.Children.Add($border))

    $slotRef = $script:ClipSlots[$tag]
    if ($slotRef) {
        $slotRef.TextBoxCtrl   = $tb
        $slotRef.SheetCombo    = $slotSheetCombo
        $slotRef.CellBox       = $slotCellBox
        $slotRef.TimestampChk  = $slotTsChk
        $slotRef.DateCellBox   = $slotDateCellBox
        $slotRef.TimeCellBox   = $slotTimeCellBox
    }

    # Seed slot sheet options from selected workbook.
    try {
        $wbSel = [string]$WbCombo.SelectedItem
        if ($wbSel) {
            $sheets = Get-WorksheetNames -WorkbookName $wbSel -Session Main
            foreach ($s in $sheets) { [void]($slotSheetCombo.Items.Add($s)) }
            $pick = if ($DefaultSheet) { [string]$DefaultSheet } else { '' }
            if ($pick -and $slotSheetCombo.Items.Contains($pick)) {
                $slotSheetCombo.SelectedItem = $pick
            } elseif ($slotSheetCombo.Items.Count -gt 0) {
                $slotSheetCombo.SelectedIndex = 0
            }
        }
    } catch {}

    # Update count label
    $CountLabel.Text = "$($Panel.Children.Count)"

    # DispatcherTimer fires on the UI thread, so clipboard reads and textbox
    # updates are safe without explicit Dispatcher.Invoke marshalling.
    $recTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $recTimer.Interval = [TimeSpan]::FromMilliseconds(350)   # poll clipboard ~3×/second

    $setRecUi = {
        param([bool]$on)
        if ($on) {
            $copyBtn.Content = 'Stop Recording'
            $copyBtn.Background = (HexBrush '#E05050')
            $copyBtn.Foreground = (HexBrush '#FFFFFF')
            $recTxt.Text = 'REC ON'
            $recTxt.Foreground = (HexBrush '#E05050')
        } else {
            $copyBtn.Content = 'Record Clipboard'
            $copyBtn.Background = (HexBrush '#3E3E42')
            $copyBtn.Foreground = (HexBrush '#FFFFFF')
            $recTxt.Text = 'REC OFF'
            $recTxt.Foreground = (HexBrush '#6E6E6E')
        }
    # GetNewClosure() captures the current values of $copyBtn, $recTxt etc. so the
    # script block references this slot's controls even after later slots are added.
    }.GetNewClosure()

    $recTimer.Add_Tick({
        try {
            $slotState = $clipSlotsRef[$tag]
            if (-not $slotState) { return }
            if (-not $slotState.IsRecording) { return }
            $packet = Get-ClipboardPacket
            if (-not $packet) { return }
            $sig = [string]$packet.Signature
            if (-not $sig) { $sig = Get-TextSha1 ([string]$packet.Text + '|' + [string]$packet.Formats) }
            if ($sig -eq [string]$slotState.LastSig) { return }
            $slotState.LastSig = $sig
            $slotState.Packet = $packet
            $slotState.Text = $packet.Text
            $tb.Text = $packet.Text
            Set-Status "Recording slot $num -- clipboard updated"
        } catch {
            try { $recTimer.Stop() } catch {}
            $slotState = $clipSlotsRef[$tag]
            if ($slotState) { $slotState.IsRecording = $false }
            & $setRecUi $false
            Set-Status "Recording error in slot ${num}: $($_.Exception.Message)" '#E05050'
        }
    }.GetNewClosure())

    # Keep editable text synchronized with slot state.
    $tb.Add_TextChanged({
        if ($clipSlotsRef.ContainsKey($tag)) {
            $slot = $clipSlotsRef[$tag]
            if ($slot) {
                $newText = [string]$tb.Text
                $slot.Text = $newText
                if ($slot.IsRecording) { return }
                $packetText = ''
                try { if ($slot.Packet) { $packetText = [string]$slot.Packet.Text } } catch {}
                # If user manually edited the textbox, stale rich clipboard payload is no longer valid.
                if ($slot.Packet -and $packetText -ne $newText) {
                    $slot.Packet = $null
                    $slot.LastSig = Get-TextSha1 ($newText + '|manual')
                }
            }
        }
    }.GetNewClosure())

    # -- Wire Record/Stop toggle (continuous live clipboard capture)
    $copyBtn.Add_Click({
        $step = 'start'
        try {
            if ($script:ClipSequenceEnabled) {
                Stop-ClipSequenceCapture
                Set-Status 'Stopped sequence capture (switched to per-slot recording).'
            }
            $slotState = $clipSlotsRef[$tag]
            if (-not $slotState) { throw 'Slot state was not found.' }
            if ($slotState.IsRecording) {
                $step = 'stop-path'
                $slotState.IsRecording = $false
                $recTimer.Stop()
                & $setRecUi $false
                Set-Status "Stopped recording slot $num"
                return
            }
            $step = 'get-clipboard'
            $packet = Get-ClipboardPacket
            if (-not $packet) { Set-Status 'Clipboard is empty' '#E05050'; return }
            $step = 'slot-state'
            $slotState.Packet = $packet
            $slotState.Text = $packet.Text
            $step = 'textbox-update'
            $tb.Text = $packet.Text
            $step = 'hash'
            $slotState.LastSig = if ($packet.Signature) { [string]$packet.Signature } else { Get-TextSha1 ([string]$packet.Text + '|' + [string]$packet.Formats) }
            $step = 'flag-on'
            $slotState.IsRecording = $true
            $step = 'timer-start'
            $recTimer.Start()
            $step = 'ui-on'
            & $setRecUi $true
            $step = 'status'
            Set-Status "Recording clipboard to slot $num (click Stop Recording to stop)"
        } catch {
            try { $recTimer.Stop() } catch {}
            $slotState = $clipSlotsRef[$tag]
            if ($slotState) { $slotState.IsRecording = $false }
            & $setRecUi $false
            Set-Status "Record click error in slot ${num} [$step]: $($_.Exception.Message)" '#E05050'
            Write-ActivityLog "Record click error slot ${num} step=$step : $($_.Exception.Message) | $($_.ScriptStackTrace)"
        }
    }.GetNewClosure())

    # -- Wire Copy Slot (slot -> clipboard)
    $copyOutBtn.Add_Click({
        $slot = $clipSlotsRef[$tag]
        if (-not $slot) { Set-Status 'Slot is empty' '#E05050'; return }
        try {
            if ($slot.Packet -and $slot.Packet.DataObject) {
                [System.Windows.Clipboard]::SetDataObject($slot.Packet.DataObject, $true)
            } elseif ($slot.Text) {
                [System.Windows.Clipboard]::SetText($slot.Text)
            } else {
                Set-Status 'Slot is empty' '#E05050'
                return
            }
            Set-Status "Copied slot $num back to clipboard"
        } catch {
            Set-Status "Clipboard copy error: $_" '#E05050'
        }
    }.GetNewClosure())

    # -- Wire Paste (per-slot sheet/cell/timestamp controls)
    $pasteBtn.Add_Click({
        $wb = $WbCombo.SelectedItem
        $slot = $clipSlotsRef[$tag]
        if (-not $slot) { Set-Status 'Slot is empty' '#E05050'; return }
        $sh = if ($slot.SheetCombo) { $slot.SheetCombo.SelectedItem } else { $null }
        $cell = if ($slot.CellBox) { [string]$slot.CellBox.Text } else { 'A1' }
        $cell = $cell.Trim()
        if (-not $wb -or -not $sh) { Set-Status 'Select workbook and slot sheet first' '#E05050'; return }
        if (-not $cell) { $cell = 'A1' }

        $packet = $slot.Packet
        if (-not $packet) {
            # Fallback to current live clipboard if slot has not been recorded yet.
            $packet = Get-ClipboardPacket
        }
        if (-not $packet) { Set-Status 'Clipboard is empty' '#E05050'; return }
        if ($tb.Text -and ($tb.Text -ne $slot.Text)) {
            $packet = [PSCustomObject]@{ DataObject = $null; Text = $tb.Text; CapturedOn = (Get-Date -f 'yyyy-MM-dd HH:mm:ss') }
        }

        $dateCellVal = if ($slot.DateCellBox) { [string]$slot.DateCellBox.Text } else { 'A3' }
        $timeCellVal = if ($slot.TimeCellBox) { [string]$slot.TimeCellBox.Text } else { 'A2' }
        $addTs = if ($slot.TimestampChk) { [bool]$slot.TimestampChk.IsChecked } else { $false }

        Show-Busy 'Pasting to Excel...'
        try {
            Paste-ClipboardPacketToExcel -Packet $packet -WorkbookName ([string]$wb) -SheetName ([string]$sh) `
                -CellAddress $cell -AddTimestamp $addTs -DateCell $dateCellVal -TimeCell $timeCellVal
            $slot.Packet = $packet
            $slot.Text = $tb.Text
            Set-Status "Pasted slot $num to $wb > $sh > $cell"
            Write-ActivityLog "Pasted slot $num to $wb > $sh > $cell"
        } catch { Set-Status "Paste error: $_" '#E05050' }
        finally { Hide-Busy }
    }.GetNewClosure())

    # -- Wire Remove
    $removeBtn.Add_Click({
        try { $recTimer.Stop() } catch {}
        if ($script:ClipSequenceEnabled) {
            $rmIdx = [int]$Panel.Children.IndexOf($border)
            if ($rmIdx -ge 0 -and $rmIdx -lt [int]$script:ClipSequenceNextIndex) {
                $script:ClipSequenceNextIndex = [Math]::Max(0, [int]$script:ClipSequenceNextIndex - 1)
            }
        }
        $Panel.Children.Remove($border)
        if ($clipSlotsRef.ContainsKey($tag)) { $clipSlotsRef.Remove($tag) }
        $CountLabel.Text = "$($Panel.Children.Count)"
        Set-Status "Removed slot $num"
    }.GetNewClosure())
}

# ============================================================
#  MACRO NAME HELPERS (favorites prefix)
# ============================================================
# Strips the "[*] " or "    " prefix that Refresh-MacroList adds to list items.
function Get-CleanMacroName([string]$DisplayName) {
    if ($DisplayName -match '^\[.\]\s+(.+)$') { return $Matches[1] }
    return $DisplayName
}

# ============================================================
#  TEMPLATE MANAGEMENT
# ============================================================
# Text templates are persisted in templates.json as objects with Name and
# Content properties.  Content uses literal "\n" for newlines to stay JSON-safe.

# Loads all templates from JSON (migrating from CSV if needed); returns array.
function Load-Templates {
    if (Test-Path $script:TemplateJson) {
        try { return ,(Normalize-List (Get-Content $script:TemplateJson -Raw | ConvertFrom-Json)) } catch { return ,@() }
    }
    if (Test-Path $script:TemplateCSV) {
        try {
            $legacy = @(Import-Csv $script:TemplateCSV)
            Save-Templates $legacy
            return ,(Normalize-List $legacy)
        } catch { return ,@() }
    }
    return ,@()
}

# Serializes the $Templates array to templates.json (writes "[]" when empty).
function Save-Templates {
    param([array]$Templates)
    if ($Templates.Count -eq 0) {
        '[]' | Set-Content $script:TemplateJson -Encoding UTF8
        return
    }
    $Templates | ConvertTo-Json | Set-Content $script:TemplateJson -Encoding UTF8
}

# ============================================================
#  QUARTERSYNC - DATE STRIPPING REGEX
# ============================================================
# Regex patterns used to normalise file names before comparison so that
# "Report_Q1_2024.xlsx" and "Report_Q2_2025.xlsx" are treated as the same file.
$script:DatePatterns = @(
    'Q[1-4][-_\s]?20\d{2}', '20\d{2}[-_\s]?Q[1-4]',
    '20\d{2}[-_\s]\d{2}[-_\s]\d{2}', '\d{2}[-_\s]\d{2}[-_\s]20\d{2}',
    '\d{8}',
    '(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-_\s]?20\d{2}',
    '20\d{2}[-_\s]?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*',
    'FY20\d{2}', 'YE20\d{2}', '20\d{2}'
)

# Removes all date/quarter/year tokens from a filename and returns the normalised
# lowercase base name used as the comparison key.
function Strip-DateTokens([string]$name) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
    foreach ($pat in $script:DatePatterns) {
        $base = [regex]::Replace($base, $pat, '', 'IgnoreCase')
    }
    $base = $base -replace '[-_\s]{2,}','_' -replace '^[-_\s]+|[-_\s]+$',''
    return $base.ToLower().Trim()
}

# ============================================================
#  QUARTERSYNC - JSON HELPERS
# ============================================================
# These functions handle persistence for QTasks (todo lists per quarter/compare
# run), the sync log, and the QSync configuration file.
# All data is stored as JSON; CSV paths are only used for one-time legacy migration.

# Resolves a path to its canonical .json form, accepting .csv or bare name input.
function Resolve-QsTodoPath {
    param([string]$PathLike)
    if (-not $PathLike) { return $null }
    $ext = [System.IO.Path]::GetExtension($PathLike).ToLower()
    if ($ext -eq '.json') { return $PathLike }
    if ($ext -eq '.csv') { return [System.IO.Path]::ChangeExtension($PathLike, '.json') }
    return "$PathLike.json"
}

# Loads the todo item array from the JSON file at $Path (defaults to
# $ActiveQuarterPath); migrates from CSV if the JSON does not yet exist.
function Load-QsTodos {
    param([string]$Path)
    if (-not $Path) { $Path = $script:ActiveQuarterPath }
    if (-not $Path) { return ,@() }

    $jsonPath = Resolve-QsTodoPath $Path
    $csvPath  = [System.IO.Path]::ChangeExtension($jsonPath, '.csv')

    if (Test-Path $jsonPath) {
        try {
            return ,(Normalize-List (Get-Content $jsonPath -Raw | ConvertFrom-Json))
        } catch {
            return ,@()
        }
    }
    if (Test-Path $csvPath) {
        try {
            $legacy = @(Import-Csv $csvPath)
            Save-QsTodos -todos $legacy -Path $jsonPath
            return ,(Normalize-List $legacy)
        } catch { return ,@() }
    }
    return ,@()
}

# Serializes $todos to JSON at $Path (defaults to $ActiveQuarterPath).
function Save-QsTodos([array]$todos, [string]$Path) {
    if (-not $Path) { $Path = $script:ActiveQuarterPath }
    if (-not $Path) { return }
    $jsonPath = Resolve-QsTodoPath $Path
    $cleanTodos = Normalize-List $todos
    if ($cleanTodos.Count -eq 0) {
        '[]' | Set-Content $jsonPath -Encoding UTF8
        return
    }
    $cleanTodos | ConvertTo-Json -Depth 5 | Set-Content $jsonPath -Encoding UTF8
}

function Load-QsSyncLog {
    if (Test-Path $script:QsSyncJson) {
        try { return ,(Normalize-List (Get-Content $script:QsSyncJson -Raw | ConvertFrom-Json)) } catch { return ,@() }
    }
    if (Test-Path $script:QsSyncCSV) {
        try {
            $legacy = @(Import-Csv $script:QsSyncCSV)
            Save-QsSyncLog $legacy
            return ,(Normalize-List $legacy)
        } catch { return ,@() }
    }
    return ,@()
}

function Save-QsSyncLog([array]$rows) {
    if ($rows.Count -eq 0) {
        '[]' | Set-Content $script:QsSyncJson -Encoding UTF8
        return
    }
    $rows | ConvertTo-Json -Depth 5 | Set-Content $script:QsSyncJson -Encoding UTF8
}

function Load-QsConfig {
    if (Test-Path $script:QsCfgJson) {
        try {
            $r = Get-Content $script:QsCfgJson -Raw | ConvertFrom-Json
            return [PSCustomObject]@{
                LastQuarterPath = [string]$r.LastQuarterPath
                ThisQuarterPath = [string]$r.ThisQuarterPath
                CompareSourcePath = [string]$r.CompareSourcePath
                CompareTargetPath = [string]$r.CompareTargetPath
            }
        } catch {}
    }
    if (Test-Path $script:QsCfgCSV) {
        try {
            $legacy = Import-Csv $script:QsCfgCSV
            $cfg = [PSCustomObject]@{
                LastQuarterPath = ($legacy | Where-Object Key -eq 'LastQuarterPath').Value
                ThisQuarterPath = ($legacy | Where-Object Key -eq 'ThisQuarterPath').Value
                CompareSourcePath = ''
                CompareTargetPath = ''
            }
            Save-QsConfig $cfg
            return $cfg
        } catch {}
    }
    return [PSCustomObject]@{
        LastQuarterPath=''; ThisQuarterPath=''
        CompareSourcePath=''; CompareTargetPath=''
    }
}

function Save-QsConfig($cfg) {
    [PSCustomObject]@{
        LastQuarterPath = [string]$cfg.LastQuarterPath
        ThisQuarterPath = [string]$cfg.ThisQuarterPath
        CompareSourcePath = [string]$cfg.CompareSourcePath
        CompareTargetPath = [string]$cfg.CompareTargetPath
        SavedOn         = (Get-Date -f 'yyyy-MM-dd HH:mm:ss')
    } | ConvertTo-Json | Set-Content $script:QsCfgJson -Encoding UTF8
}

# Loads the compare-mode missing-file list from qs_compare_results.json; this is
# the dataset displayed in the QTasks tab after a folder comparison.
function Load-QsCompareTodos {
    if (-not (Test-Path $script:QsCompareJson)) { return ,@() }
    try {
        return ,(Normalize-List (Get-Content $script:QsCompareJson -Raw | ConvertFrom-Json))
    } catch {
        return ,@()
    }
}

# Persists the compare-mode missing-file list to qs_compare_results.json.
function Save-QsCompareTodos([array]$todos) {
    $clean = Normalize-List $todos
    if ($clean.Count -eq 0) {
        '[]' | Set-Content $script:QsCompareJson -Encoding UTF8
        return
    }
    $clean | ConvertTo-Json -Depth 6 | Set-Content $script:QsCompareJson -Encoding UTF8
}

function Resolve-QsFolderPath([string]$PathText) {
    if (-not $PathText) { return '' }
    $p = $PathText.Trim().Trim('"')
    if (-not $p) { return '' }
    try { return [System.IO.Path]::GetFullPath($p) } catch { return $p }
}

function Get-QsCompareFileKey($Root, $FileInfo) {
    $root = [string]$Root
    $relFolder = ''
    if ($FileInfo.DirectoryName.Length -gt $root.TrimEnd('\').Length) {
        $relFolder = $FileInfo.DirectoryName.Substring($root.TrimEnd('\').Length).TrimStart('\')
    }
    $base = [System.IO.Path]::GetFileNameWithoutExtension([string]$FileInfo.Name)
    $normBase = (Strip-DateTokens $base).ToLower()
    $ext = [string]$FileInfo.Extension
    return [PSCustomObject]@{
        Key = "$($normBase)|$($ext.ToLower())|$($relFolder.ToLower())"
        RelFolder = $(if ($relFolder) { $relFolder } else { '(root)' })
    }
}

# Compares two folder trees by normalised (date-stripped) file key, returning a
# result object whose MissingTodos array lists files present in $SourceRoot but
# absent from $TargetRoot; used to populate the QTasks deliverables list.
function Compare-QsFolders {
    param([string]$SourceRoot, [string]$TargetRoot)

    $src = Resolve-QsFolderPath $SourceRoot
    $tgt = Resolve-QsFolderPath $TargetRoot
    if (-not $src -or -not (Test-Path $src)) { throw "Source folder not found: $SourceRoot" }
    if (-not $tgt -or -not (Test-Path $tgt)) { throw "Target folder not found: $TargetRoot" }

    $srcFiles = @(Get-ChildItem -Path $src -Recurse -File -ErrorAction SilentlyContinue)
    $tgtFiles = @(Get-ChildItem -Path $tgt -Recurse -File -ErrorAction SilentlyContinue)

    $targetKeyMap = @{}
    foreach ($f in $tgtFiles) {
        $meta = Get-QsCompareFileKey -Root $tgt -FileInfo $f
        if (-not $targetKeyMap.ContainsKey($meta.Key)) { $targetKeyMap[$meta.Key] = $true }
    }

    $missing = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $srcFiles) {
        $meta = Get-QsCompareFileKey -Root $src -FileInfo $f
        if ($targetKeyMap.ContainsKey($meta.Key)) { continue }
        $addedOn = $f.CreationTime
        if (-not $addedOn -or $addedOn.Year -lt 1980) { $addedOn = $f.LastWriteTime }
        $modifiedOn = $f.LastWriteTime
        $missing.Add([PSCustomObject]@{
            Key = $meta.Key
            OriginalName = [string]$f.Name
            RelFolder = [string]$meta.RelFolder
            LastDoneDate = $addedOn.ToString('yyyy-MM-dd')
            DueDate = $modifiedOn.ToString('yyyy-MM-dd')
            Status = 'Pending'
            AddedOn = (Get-Date -f 'yyyy-MM-dd')
            Note = 'Missing in target'
            SourceRoot = $src
            TargetRoot = $tgt
            SourcePath = [string]$f.FullName
        })
    }

    return [PSCustomObject]@{
        SourceRoot = $src
        TargetRoot = $tgt
        SourceCount = $srcFiles.Count
        TargetCount = $tgtFiles.Count
        MissingTodos = @($missing)
    }
}

# ============================================================
#  QUARTERSYNC - CORE SYNC ENGINE
# ============================================================
# Invoke-QuarterSync mirrors the folder structure from the previous quarter into
# the current quarter folder and creates todo items for files that exist in the
# previous quarter but are missing from the new one.

# Replicates missing subdirectories from $LastRoot into $ThisRoot, then generates
# todo entries for files present in $LastRoot but not yet in $ThisRoot.
# Results and new todos are merged into the active quarter JSON and the sync log.
function Invoke-QuarterSync {
    param([string]$LastRoot, [string]$ThisRoot, [string]$QuarterPath)

    $r = [PSCustomObject]@{
        FoldersCreated = [System.Collections.Generic.List[string]]::new()
        FoldersSkipped = [System.Collections.Generic.List[string]]::new()
        NewTodos       = [System.Collections.Generic.List[object]]::new()
        Errors         = [System.Collections.Generic.List[string]]::new()
    }

    if (-not (Test-Path $LastRoot)) {
        $r.Errors.Add("Last quarter path not found: $LastRoot"); return $r
    }
    if (-not (Test-Path $ThisRoot)) {
        $r.Errors.Add("This quarter path not found: $ThisRoot"); return $r
    }

    # Folder sync
    $lastFolders = Get-ChildItem -Path $LastRoot -Recurse -Directory -ErrorAction SilentlyContinue
    foreach ($f in $lastFolders) {
        $rel    = $f.FullName.Substring($LastRoot.TrimEnd('\').Length).TrimStart('\')
        $target = Join-Path $ThisRoot $rel
        if (Test-Path $target) { $r.FoldersSkipped.Add($rel) }
        else {
            try   { [void](New-Item -ItemType Directory -Path $target -Force); $r.FoldersCreated.Add($rel) }
            catch { $r.Errors.Add("Could not create: $rel") }
        }
    }

    # File comparison
    $lastFiles = Get-ChildItem -Path $LastRoot -Recurse -File -ErrorAction SilentlyContinue
    $thisFiles = Get-ChildItem -Path $ThisRoot -Recurse -File -ErrorAction SilentlyContinue

    $lastMap = @{}
    foreach ($f in $lastFiles) {
        $key = (Strip-DateTokens $f.Name) + $f.Extension.ToLower()
        $rel = $f.DirectoryName.Substring($LastRoot.TrimEnd('\').Length).TrimStart('\')
        if (-not $lastMap.ContainsKey($key)) {
            $lastMap[$key] = [PSCustomObject]@{
                OriginalName  = $f.Name
                RelFolder     = $rel
                LastWriteTime = $f.LastWriteTime
            }
        }
    }

    $thisMap = @{}
    foreach ($f in $thisFiles) {
        $key = (Strip-DateTokens $f.Name) + $f.Extension.ToLower()
        $rel = $f.DirectoryName.Substring($ThisRoot.TrimEnd('\').Length).TrimStart('\')
        $thisMap["$key|$rel"] = $true
    }

    $todoTarget   = if ($QuarterPath) { $QuarterPath } else { $script:ActiveQuarterPath }
    $existing     = Load-QsTodos $todoTarget
    $existingKeys = $existing | ForEach-Object { $_.Key }

    foreach ($key in $lastMap.Keys) {
        $src     = $lastMap[$key]
        $todoKey = "$key|$($src.RelFolder)"
        if (-not $thisMap.ContainsKey($todoKey)) {
            if ($existingKeys -notcontains $todoKey) {
                $due = $src.LastWriteTime.AddMonths(3).ToString('yyyy-MM-dd')
                $r.NewTodos.Add([PSCustomObject]@{
                    Key          = $todoKey
                    OriginalName = $src.OriginalName
                    RelFolder    = if ($src.RelFolder) { $src.RelFolder } else { '(root)' }
                    LastDoneDate = $src.LastWriteTime.ToString('yyyy-MM-dd')
                    DueDate      = $due
                    Status       = 'Pending'
                    AddedOn      = (Get-Date -f 'yyyy-MM-dd')
                    Note         = ''
                })
            }
        }
    }

    $merged = [System.Collections.Generic.List[object]]::new()
    foreach ($e in $existing) { $merged.Add($e) }
    foreach ($t in $r.NewTodos) { $merged.Add($t) }
    Save-QsTodos $merged $todoTarget

    # Save sync log
    $logData = Load-QsSyncLog
    $syncRows = [System.Collections.Generic.List[object]]::new()
    foreach ($sr in $logData) { $syncRows.Add($sr) }
    foreach ($f in $r.FoldersCreated) {
        $syncRows.Add([PSCustomObject]@{
            RunDate = (Get-Date -f 'yyyy-MM-dd'); Type = 'Created'; Path = $f
        })
    }
    foreach ($f in $r.FoldersSkipped) {
        $syncRows.Add([PSCustomObject]@{
            RunDate = (Get-Date -f 'yyyy-MM-dd'); Type = 'Skipped'; Path = $f
        })
    }
    Save-QsSyncLog $syncRows

    return $r
}

# ============================================================
#  QUARTERSYNC - EXCEL EXPORT (COM, 3-sheet workbook)
# ============================================================
# Creates a formatted 3-sheet Excel workbook summarising the compare results:
# Sheet 1 = missing deliverables list, Sheet 2 = folder sync log, Sheet 3 = stats.

# Builds and saves a 3-sheet Excel report for the current compare results to
# $SavePath; uses its own transient Excel COM instance (Visible=false).
function Export-QsToExcel {
    param([string]$SavePath, [string]$SourcePath = '', [string]$TargetPath = '')

    $todos   = Load-QsCompareTodos
    $syncLog = Load-QsSyncLog
    $srcPath = if ($SourcePath) { $SourcePath } else { Resolve-QsFolderPath $SourcePath }
    $tgtPath = if ($TargetPath) { $TargetPath } else { Resolve-QsFolderPath $TargetPath }

    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()

        # -- helper: style a header row --
        $styleHeader = {
            param($ws, $row, $cols, $color)
            for ($c = 1; $c -le $cols; $c++) {
                $ws.Cells($row, $c).Interior.Color = $color
                $ws.Cells($row, $c).Font.Bold  = $true
                $ws.Cells($row, $c).Font.Color = 0xFFFFFF
                $ws.Cells($row, $c).Font.Size  = 11
            }
        }

        $navy = 0x2D3561

        # == SHEET 1: Missing Deliverables ==
        $ws1 = $wb.Worksheets(1)
        $ws1.Name = 'Missing Deliverables'
        $ws1.Cells(1,1) = 'Folder Compare -- Missing In Target'
        $ws1.Cells(1,1).Font.Size  = 14
        $ws1.Cells(1,1).Font.Bold  = $true
        $ws1.Cells(1,1).Font.Color = $navy
        $ws1.Cells(2,1) = "Exported: $(Get-Date -f 'yyyy-MM-dd HH:mm')"
        $ws1.Cells(2,1).Font.Color = 0x888888
        $ws1.Cells(3,1) = "Source A: $srcPath"
        $ws1.Cells(3,2) = "Target B: $tgtPath"
        $ws1.Cells(3,1).Font.Color = 0x666666
        $ws1.Cells(3,2).Font.Color = 0x666666

        $h1 = @('File Name','Folder Path','Added In Source (A)','Updated In Source (A)','Status','Note')
        for ($c = 0; $c -lt $h1.Count; $c++) { $ws1.Cells(5, $c+1) = $h1[$c] }
        & $styleHeader $ws1 5 $h1.Count $navy

        $pendingItems = @($todos | Where-Object { $_.Status -eq 'Pending' })
        $row = 6
        foreach ($t in $pendingItems) {
            $ws1.Cells($row,1) = $t.OriginalName
            $ws1.Cells($row,2) = $t.RelFolder
            $ws1.Cells($row,3) = $t.LastDoneDate
            $ws1.Cells($row,4) = $t.DueDate
            $ws1.Cells($row,5) = $t.Status
            $ws1.Cells($row,6) = $t.Note
            $row++
        }
        $ws1.Columns('A:F').AutoFit()

        # == SHEET 2: Folder Sync Summary ==
        $ws2 = $wb.Worksheets.Add()
        $ws2.Move([System.Reflection.Missing]::Value, $wb.Worksheets($wb.Worksheets.Count))
        $ws2.Name = 'Folder Sync Summary'
        $ws2.Cells(1,1) = 'QuarterSync -- Folder Sync Summary'
        $ws2.Cells(1,1).Font.Size  = 14
        $ws2.Cells(1,1).Font.Bold  = $true
        $ws2.Cells(1,1).Font.Color = $navy

        $h2 = @('Run Date','Type','Folder Path')
        for ($c = 0; $c -lt $h2.Count; $c++) { $ws2.Cells(3,$c+1) = $h2[$c] }
        & $styleHeader $ws2 3 $h2.Count $navy

        $row = 4
        foreach ($s in $syncLog) {
            $ws2.Cells($row,1) = $s.RunDate
            $ws2.Cells($row,2) = $s.Type
            $ws2.Cells($row,3) = $s.Path
            if ($s.Type -eq 'Created') {
                $ws2.Cells($row,2).Font.Color = 0x1A7A4A
                $ws2.Cells($row,2).Font.Bold  = $true
            }
            $row++
        }
        $ws2.Columns('A:C').AutoFit()

        # == SHEET 3: Stats by Folder ==
        $ws3 = $wb.Worksheets.Add()
        $ws3.Move([System.Reflection.Missing]::Value, $wb.Worksheets($wb.Worksheets.Count))
        $ws3.Name = 'Stats by Folder'
        $ws3.Cells(1,1) = 'QuarterSync -- Completion Stats by Subfolder'
        $ws3.Cells(1,1).Font.Size  = 14
        $ws3.Cells(1,1).Font.Bold  = $true
        $ws3.Cells(1,1).Font.Color = $navy

        $h3 = @('Subfolder','Total Items','Pending','Done / Cleared','% Complete')
        for ($c = 0; $c -lt $h3.Count; $c++) { $ws3.Cells(3,$c+1) = $h3[$c] }
        & $styleHeader $ws3 3 $h3.Count $navy

        $folders = $todos | Group-Object RelFolder | Sort-Object Name
        $row = 4
        foreach ($grp in $folders) {
            $total3   = $grp.Count
            $pending3 = ($grp.Group | Where-Object { $_.Status -eq 'Pending' }).Count
            $done3    = $total3 - $pending3
            $pct3     = if ($total3 -gt 0) { [Math]::Round(($done3 / $total3) * 100, 1) } else { 100 }

            $ws3.Cells($row,1) = $(if ($grp.Name) { $grp.Name } else { '(root)' })
            $ws3.Cells($row,2) = $total3
            $ws3.Cells($row,3) = $pending3
            $ws3.Cells($row,4) = $done3
            $ws3.Cells($row,5) = "$pct3%"
            $ws3.Cells($row,5).Interior.Color = $(if ($pct3 -ge 80) { 0xC8F0D8 } elseif ($pct3 -ge 40) { 0xFFF0C0 } else { 0xFFD0D0 })
            $ws3.Cells($row,5).Font.Bold = $true
            $row++
        }
        $ws3.Columns('A:E').AutoFit()

        $wb.SaveAs($SavePath, 51)   # 51 = xlOpenXMLWorkbook
        $wb.Close($false)
        $xl.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        return $true
    } catch {
        try { $xl.Quit() } catch {}
        return "Error: $_"
    }
}

# ============================================================
#  QUARTER FILE MANAGEMENT
# ============================================================
# Helpers for managing the per-quarter JSON todo files stored in quarters/.

# Returns an array of FileInfo objects for all quarter JSON files, newest first;
# migrates any legacy CSV quarter files to JSON on first call.
function Get-QuarterList {
    if (-not (Test-Path $script:QuartersDir)) { return ,@() }
    $d = @(Get-ChildItem $script:QuartersDir -Filter '*.json' | Sort-Object Name -Descending)
    if ($d.Count -eq 0) {
        # Legacy migration: promote existing CSV quarter files to JSON.
        $legacy = @(Get-ChildItem $script:QuartersDir -Filter '*.csv' -ErrorAction SilentlyContinue)
        foreach ($f in $legacy) {
            try {
                $rows = @(Import-Csv $f.FullName)
                $jsonPath = [System.IO.Path]::ChangeExtension($f.FullName, '.json')
                Save-QsTodos -todos $rows -Path $jsonPath
            } catch {}
        }
        $d = @(Get-ChildItem $script:QuartersDir -Filter '*.json' | Sort-Object Name -Descending)
    }
    return ,$d
}

# Scans $FolderPath recursively and returns a todo array with one Pending entry
# per file; used to seed a brand-new quarter file from an existing folder.
function Scan-FolderToTodos {
    param([string]$FolderPath)
    if (-not (Test-Path $FolderPath)) { return @() }
    $files = Get-ChildItem $FolderPath -Recurse -File -ErrorAction SilentlyContinue
    $todos = @()
    foreach ($f in $files) {
        $rel = ''
        if ($f.DirectoryName.Length -gt $FolderPath.TrimEnd('\').Length) {
            $rel = $f.DirectoryName.Substring($FolderPath.TrimEnd('\').Length).TrimStart('\')
        }
        $key = (Strip-DateTokens $f.Name) + $f.Extension.ToLower()
        $todoKey = "$key|$rel"
        $todos += [PSCustomObject]@{
            Key          = $todoKey
            OriginalName = $f.Name
            RelFolder    = if ($rel) { $rel } else { '(root)' }
            LastDoneDate = $f.LastWriteTime.ToString('yyyy-MM-dd')
            DueDate      = ''
            Status       = 'Pending'
            AddedOn      = (Get-Date -f 'yyyy-MM-dd')
            Note         = ''
        }
    }
    return ,$todos
}

function Switch-ActiveQuarter([string]$path) {
    $script:ActiveQuarterPath = (Resolve-QsTodoPath $path)
    $script:ActiveQuarterCSV = $script:ActiveQuarterPath
}

function New-QuarterFile([string]$name) {
    if (-not (Test-Path $script:QuartersDir)) {
        [void](New-Item -ItemType Directory -Path $script:QuartersDir -Force)
    }
    $safe = $name -replace '[^\w\s.-]','_'
    $path = Join-Path $script:QuartersDir "$safe.json"
    if (-not (Test-Path $path)) {
        '[]' | Set-Content $path -Encoding UTF8
    }
    return $path
}

function Get-ComboText($Combo) {
    if (-not $Combo) { return '' }
    if ($Combo.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
        return [string]$Combo.SelectedItem.Content
    }
    if ($Combo.SelectedItem) {
        return [string]$Combo.SelectedItem
    }
    return [string]$Combo.Text
}

# Returns a hex SHA-1 hash of $Text; used to fingerprint clipboard contents for
# change detection without storing the full payload.
function Get-TextSha1([string]$Text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $hash = $sha1.ComputeHash($bytes)
        return ([BitConverter]::ToString($hash) -replace '-','').ToLower()
    } finally {
        try { $sha1.Dispose() } catch {}
    }
}

# ================================================================
#  XAML UI - Chrome/Teams Dark Theme
# ================================================================
# ============================================================
#  FILE INDEX — DLP-safe metadata-only drive search
# ============================================================
# Scans a folder tree, caches file metadata (name/path/size/dates) as JSON, and
# provides real-time in-memory filtering. Never reads file contents.
# A 60-minute cooldown prevents accidental repeated full scans of large trees.

$script:FidxCacheJson   = Join-Path $script:HubRoot 'fileindex_cache.json'
$script:FidxLastScan    = $null          # [datetime] of last successful scan
$script:FidxAllItems    = @()            # full unfiltered index (array of PSCustomObjects)
$script:FidxCooldownSec = 3600           # seconds between allowed rescans (60 min)
$script:FidxCooldownTimer = $null        # DispatcherTimer for cooldown countdown

function Get-FidxCacheAgeSec {
    if (-not $script:FidxLastScan) { return [int]::MaxValue }
    return [int]((Get-Date) - $script:FidxLastScan).TotalSeconds
}

function Format-FidxSize([long]$bytes) {
    if ($bytes -ge 1GB) { return '{0:N1} GB' -f ($bytes / 1GB) }
    if ($bytes -ge 1MB) { return '{0:N1} MB' -f ($bytes / 1MB) }
    if ($bytes -ge 1KB) { return '{0:N1} KB' -f ($bytes / 1KB) }
    return '{0} B' -f $bytes
}

function Invoke-FidxScan {
    param([string]$RootPath)
    if (-not (Test-Path $RootPath -PathType Container)) {
        throw "Folder not found: $RootPath"
    }
    $items = @()
    $files = Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        $items += [PSCustomObject]@{
            Name     = $f.Name
            Path     = $f.FullName
            SizeStr  = (Format-FidxSize $f.Length)
            Modified = $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
            Ext      = $f.Extension.ToLower()
            SizeBytes = [long]$f.Length
        }
    }
    return ,$items
}

function Save-FidxCache {
    param([array]$Items, [string]$Root)
    try {
        $cache = [ordered]@{
            Root      = $Root
            ScannedOn = (Get-Date -f 'yyyy-MM-dd HH:mm:ss')
            Count     = $Items.Count
            Items     = @($Items | Select-Object Name, Path, SizeStr, Modified, Ext, SizeBytes)
        }
        $cache | ConvertTo-Json -Depth 3 -Compress | Set-Content $script:FidxCacheJson -Encoding UTF8
    } catch {}
}

function Load-FidxCache {
    if (-not (Test-Path $script:FidxCacheJson)) { return $null }
    try {
        $raw = Get-Content $script:FidxCacheJson -Raw | ConvertFrom-Json
        return $raw
    } catch { return $null }
}

function Get-FidxFiltered {
    param([string]$Query)
    if ([string]::IsNullOrWhiteSpace($Query)) { return ,$script:FidxAllItems }
    $q = $Query.Trim().ToLower()
    return ,@($script:FidxAllItems | Where-Object {
        $_.Name.ToLower().Contains($q) -or
        $_.Path.ToLower().Contains($q) -or
        $_.Ext.ToLower().Contains($q)
    })
}

function Start-FidxCooldownTimer {
    if ($script:FidxCooldownTimer) {
        try { $script:FidxCooldownTimer.Stop() } catch {}
    }
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromSeconds(10)
    $timer.Add_Tick({
        $age = Get-FidxCacheAgeSec
        $remaining = $script:FidxCooldownSec - $age
        if ($remaining -le 0) {
            $FidxIndexBtn.IsEnabled = $true
            $FidxCooldownTxt.Text = ''
            $script:FidxCooldownTimer.Stop()
        } else {
            $FidxIndexBtn.IsEnabled = $false
            $mins = [Math]::Ceiling($remaining / 60)
            $FidxCooldownTxt.Text = "  Refresh available in ~${mins}m"
        }
    })
    $script:FidxCooldownTimer = $timer
    $timer.Start()
}

# ============================================================

# Start-MacroHub is the single entry point for the entire application.
# It defines the XAML here-string, parses it into a WPF Window, wires up all
# G() element bindings, defines all refresh and event-handler functions, runs
# startup data loading, and then calls $Window.ShowDialog() to block until close.
# Does not return until the user closes the window.
function Start-MacroHub {
    param([string]$HideTabs = '')

$xamlStr = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="MacroHub v3.2" Height="860" Width="1200"
    MinHeight="640" MinWidth="960"
    WindowStartupLocation="CenterScreen"
    Background="#1B1B1F"
    FontFamily="Segoe UI">

  <Window.Resources>
    <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}"        Color="#252528"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.WindowTextBrushKey}"    Color="#FFFFFF"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}"     Color="#4C9FE6"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}"       Color="#252528"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.ControlTextBrushKey}"   Color="#FFFFFF"/>

    <SolidColorBrush x:Key="BgPage"    Color="#1B1B1F"/>
    <SolidColorBrush x:Key="BgCard"    Color="#252528"/>
    <SolidColorBrush x:Key="BgInput"   Color="#1B1B1F"/>
    <SolidColorBrush x:Key="BgHover"   Color="#2D2D30"/>
    <SolidColorBrush x:Key="Border"    Color="#3E3E42"/>
    <SolidColorBrush x:Key="Accent"    Color="#4C9FE6"/>
    <SolidColorBrush x:Key="TextHi"    Color="#FFFFFF"/>
    <SolidColorBrush x:Key="TextMid"   Color="#B0B0B0"/>
    <SolidColorBrush x:Key="TextLo"    Color="#6E6E6E"/>
    <SolidColorBrush x:Key="Red"       Color="#E05050"/>

    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="6"/>
      <Setter Property="MinWidth" Value="6"/>
    </Style>

    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Background"      Value="#252528"/>
      <Setter Property="Foreground"      Value="#FFFFFF"/>
      <Setter Property="BorderBrush"     Value="#3E3E42"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="14,7"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                    RecognizesAccessKey="True"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background"  Value="#2D2D30"/>
                <Setter Property="BorderBrush" Value="#4C9FE6"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#333338"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnAccent" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background"  Value="#1A2A3E"/>
      <Setter Property="BorderBrush" Value="#4C9FE6"/>
      <Setter Property="Foreground"  Value="#4C9FE6"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
    </Style>

    <Style x:Key="BtnRed" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background"  Value="#2E1F1F"/>
      <Setter Property="BorderBrush" Value="#E05050"/>
      <Setter Property="Foreground"  Value="#E05050"/>
    </Style>

    <Style x:Key="BtnGreen" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background"  Value="#1A2E1A"/>
      <Setter Property="BorderBrush" Value="#50A050"/>
      <Setter Property="Foreground"  Value="#50A050"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background"      Value="#1B1B1F"/>
      <Setter Property="Foreground"      Value="#FFFFFF"/>
      <Setter Property="BorderBrush"     Value="#3E3E42"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="8,6"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="CaretBrush"      Value="#4C9FE6"/>
      <Setter Property="SelectionBrush"  Value="#4C9FE6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsFocused" Value="True">
                <Setter Property="BorderBrush" Value="#4C9FE6"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ComboBox">
      <Setter Property="Background"      Value="#252528"/>
      <Setter Property="Foreground"      Value="#FFFFFF"/>
      <Setter Property="BorderBrush"     Value="#3E3E42"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="8,6"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton x:Name="PART_ToggleButton" Focusable="False"
                            ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="tBd" Background="#252528" BorderBrush="#3E3E42"
                            BorderThickness="1" CornerRadius="4">
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="24"/>
                        </Grid.ColumnDefinitions>
                        <Path Grid.Column="1" Data="M0,0 L4,4 8,0" Stroke="#B0B0B0"
                              StrokeThickness="1.5" HorizontalAlignment="Center"
                              VerticalAlignment="Center"/>
                      </Grid>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="tBd" Property="BorderBrush" Value="#4C9FE6"/>
                        <Setter TargetName="tBd" Property="Background"  Value="#2D2D30"/>
                      </Trigger>
                      <Trigger Property="IsChecked" Value="True">
                        <Setter TargetName="tBd" Property="BorderBrush" Value="#4C9FE6"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter x:Name="PART_ContentSite"
                                IsHitTestVisible="False"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                Margin="10,6,28,6" VerticalAlignment="Center"
                                HorizontalAlignment="Left"/>
              <TextBox x:Name="PART_EditableTextBox" Visibility="Hidden"
                       IsReadOnly="{TemplateBinding IsReadOnly}"
                       Background="Transparent" Foreground="#FFFFFF"
                       CaretBrush="#FFFFFF" Margin="10,6,28,6"/>
              <Popup x:Name="PART_Popup" Placement="Bottom"
                     IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Slide">
                <Border x:Name="DropDownBorder" Background="#252528"
                        BorderBrush="#3E3E42" BorderThickness="1"
                        CornerRadius="0,0,4,4" Margin="0,1,0,0"
                        MinWidth="{TemplateBinding ActualWidth}"
                        MaxHeight="{TemplateBinding MaxDropDownHeight}">
                  <ScrollViewer SnapsToDevicePixels="True">
                    <StackPanel IsItemsHost="True"
                                KeyboardNavigation.DirectionalNavigation="Contained"/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Background"  Value="#252528"/>
      <Setter Property="Foreground"  Value="#FFFFFF"/>
      <Setter Property="Padding"     Value="8,6"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="#4C9FE6"/>
          <Setter Property="Foreground" Value="#FFFFFF"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#37373D"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="ListBoxItem">
      <Setter Property="Background"  Value="Transparent"/>
      <Setter Property="Foreground"  Value="#FFFFFF"/>
      <Setter Property="Padding"     Value="6,4"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#4C9FE6"/>
          <Setter Property="Foreground" Value="#FFFFFF"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#37373D"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="TabItem">
      <Setter Property="Foreground"      Value="#A0A0A0"/>
      <Setter Property="Background"      Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"         Value="18,12"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="TabBd" Background="Transparent"
                    BorderThickness="0,0,0,2" BorderBrush="Transparent"
                    Padding="{TemplateBinding Padding}" Margin="0,0,2,0">
              <ContentPresenter ContentSource="Header" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="TabBd" Property="BorderBrush" Value="#4C9FE6"/>
                <Setter Property="Foreground" Value="#FFFFFF"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="TabBd" Property="Background" Value="#252528"/>
                <Setter Property="Foreground" Value="#B0B0B0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Background"      Value="#1B1B1F"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"         Value="0"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="50"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="26"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#252528" BorderBrush="#3E3E42" BorderThickness="0,0,0,1">
      <Grid Margin="20,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="MacroHub" Foreground="#FFFFFF" FontSize="16" FontWeight="SemiBold"
                     VerticalAlignment="Center"/>
          <TextBlock Text="  v3.1" Foreground="#6E6E6E" FontSize="11"
                     VerticalAlignment="Center" Margin="4,2,0,0"/>
        </StackPanel>

        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="HelpBtn" Content="?" Style="{StaticResource Btn}"
                  Width="26" Height="22" Margin="0,0,6,0" Padding="0"
                  ToolTip="Tab workflow guide"/>
          <Border Background="#1A2A3E" CornerRadius="12" Padding="10,3" Margin="4,0">
            <TextBlock x:Name="ExcelStatus" Text="Excel: --" Foreground="#4C9FE6" FontSize="11"/>
          </Border>
          <Border Background="#1A2A3E" CornerRadius="12" Padding="10,3" Margin="4,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock x:Name="QsStatTotal" Text="0" Foreground="#4C9FE6" FontWeight="Bold" FontSize="11"/>
              <TextBlock Text=" items" Foreground="#6E6E6E" FontSize="11" Margin="2,0,0,0"/>
            </StackPanel>
          </Border>
          <Border Background="#2E1F1F" CornerRadius="12" Padding="10,3" Margin="4,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock x:Name="QsStatPending" Text="0" Foreground="#E05050" FontWeight="Bold" FontSize="11"/>
              <TextBlock Text=" pending" Foreground="#6E6E6E" FontSize="11" Margin="2,0,0,0"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </Grid>
    </Border>

    <TabControl Grid.Row="1" x:Name="MainTabs">

      <TabItem Header="Clipboard">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="14"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- Target workbook selector + global controls -->
          <Border Grid.Row="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="16,12">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="260"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Target Workbook" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Center"
                         ToolTip="The Excel workbook all slots will paste into"/>
              <ComboBox  Grid.Column="2" x:Name="ClipWbCombo"
                         ToolTip="Select the open workbook to paste clipboard data into"/>
              <Button    Grid.Column="4" x:Name="ClipRefreshBtn" Content="_Refresh Workbooks"
                         Style="{StaticResource Btn}" Padding="12,6" Margin="0,0,0,0"
                         ToolTip="Reload the list of open Excel workbooks"/>
              <Button    Grid.Column="5" x:Name="ClipLockBtn" Content="_Save Defaults"
                         Style="{StaticResource BtnAccent}" Padding="10,6" Margin="10,0,0,0"
                         ToolTip="Save the current workbook and slot-1 sheet/cell/timestamp settings as defaults for new slots"/>
              <Button    Grid.Column="6" x:Name="ClipClearDefaultsBtn" Content="_Clear Defaults"
                         Style="{StaticResource Btn}" Padding="10,6" Margin="6,0,0,0"
                         Foreground="#E05050" ToolTip="Remove saved default settings"/>
              <TextBlock Grid.Column="7" x:Name="ClipDefaultsIndicator" Text="✓ Defaults loaded"
                         Foreground="#4C9FE6" FontSize="11" VerticalAlignment="Center"
                         Margin="10,0,0,0" Visibility="Collapsed"/>
              <!-- Hidden backing controls for slot defaults — keep these, they are used by code -->
              <StackPanel Grid.Column="8" Orientation="Horizontal" Visibility="Collapsed">
                <ComboBox x:Name="ClipSheetCombo"/>
                <TextBox x:Name="ClipCellBox" Text="A1"/>
                <CheckBox x:Name="ClipTimestampChk"/>
                <TextBox x:Name="ClipDateCell" Text="A3"/>
                <TextBox x:Name="ClipTimeCell" Text="A2"/>
              </StackPanel>
            </Grid>
          </Border>

          <!-- Slot controls toolbar -->
          <Border Grid.Row="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>

              <TextBlock Grid.Column="0"
                         Text="Each slot captures one clipboard item. Set the destination sheet and cell per slot, then click Paste to write to Excel."
                         Foreground="#909090" FontSize="11" VerticalAlignment="Center"
                         TextWrapping="Wrap"/>

              <StackPanel Grid.Column="1" Orientation="Horizontal">
                <TextBlock Text="Slots:" Foreground="#909090" FontSize="11"
                           VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBlock x:Name="ClipSlotCount" Text="0" Foreground="#4C9FE6" FontWeight="Bold"
                           FontSize="13" VerticalAlignment="Center" Margin="0,0,14,0"/>
                <Button x:Name="ClipAddSlotBtn" Content="+ _Add Slot" Style="{StaticResource BtnAccent}"
                        Padding="12,5" ToolTip="Add a new clipboard capture slot"/>
                <Button x:Name="ClipRecordSeqBtn" Content="_Record Sequence"
                        Style="{StaticResource Btn}" Padding="12,5" Margin="8,0,0,0"
                        ToolTip="Auto-capture each successive clipboard copy into the next slot in order (Slot 1, 2, 3…)"/>
                <TextBlock x:Name="ClipRecordSeqState" Text="SEQ OFF" Foreground="#909090"
                           FontWeight="Bold" FontSize="11" VerticalAlignment="Center" Margin="10,0,0,0"/>
              </StackPanel>
            </Grid>
          </Border>

          <!-- Clipboard slots -->
          <ScrollViewer Grid.Row="4" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="ClipSlotsPanel" Orientation="Vertical"/>
          </ScrollViewer>
        </Grid>
      </TabItem>

      <TabItem Header="Macros">
        <Grid Margin="24,18">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="280"/>
            <ColumnDefinition Width="16"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <!-- Macro file list (left panel) -->
          <Border Grid.Column="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <TextBlock Text="MACRO FILES" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"
                           ToolTip="PowerShell (.ps1) and VBA (.bas) files from the Macros folder"/>
              </Border>
              <ListBox Grid.Row="1" x:Name="MacroList" Background="Transparent"
                       Foreground="#FFFFFF" BorderThickness="0" Padding="6,4"
                       FontSize="12"
                       ToolTip="Click to select a macro, then configure and run it on the right"/>
              <Border Grid.Row="2" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="10,8">
                <StackPanel>
                  <Button x:Name="MacroRefreshBtn" Content="_Refresh List" Style="{StaticResource Btn}"
                          Padding="10,6" HorizontalAlignment="Stretch" Margin="0,0,0,6"
                          ToolTip="Reload macro files from the Macros folder"/>
                  <Button x:Name="MacroFavBtn" Content="Toggle _Favorite" Style="{StaticResource BtnGreen}"
                          Padding="10,6" HorizontalAlignment="Stretch"
                          ToolTip="Mark or unmark the selected macro as a favorite"/>
                </StackPanel>
              </Border>
            </Grid>
          </Border>

          <!-- Run macro panel (right panel) -->
          <Border Grid.Column="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="20,18">
            <StackPanel>
              <TextBlock Text="RUN MACRO" Foreground="#A0A0A0" FontSize="11"
                         FontWeight="SemiBold" Margin="0,0,0,16"/>

              <TextBlock Text="Target Workbook" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"
                         ToolTip="The open Excel workbook the macro will run against"/>
              <ComboBox x:Name="MacroWbCombo" Margin="0,0,0,14"/>

              <TextBlock Text="Target Sheet (optional)" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"
                         ToolTip="Specific worksheet to activate before running the macro"/>
              <ComboBox x:Name="MacroSheetCombo" Margin="0,0,0,14"/>

              <TextBlock Text="Selected Macro" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
              <TextBox x:Name="MacroSelectedTxt" IsReadOnly="True" Foreground="#909090"
                       Margin="0,0,0,14"
                       ToolTip="The macro file selected in the list on the left"/>

              <TextBlock Text="Description" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
              <TextBox x:Name="MacroDescTxt" IsReadOnly="True" Height="64"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       Foreground="#909090" Margin="0,0,0,20"/>

              <Button x:Name="MacroRunBtn" Content="_Run Macro" Style="{StaticResource BtnAccent}"
                      Padding="20,10" FontSize="13" HorizontalAlignment="Left"
                      ToolTip="Execute the selected macro against the target workbook/sheet"/>

              <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Margin="0,20,0,0" Padding="0,12,0,0">
                <StackPanel>
                  <TextBlock Text="OUTPUT" Foreground="#A0A0A0" FontSize="11"
                             FontWeight="SemiBold" Margin="0,0,0,8"/>
                  <TextBox x:Name="MacroOutputTxt" Height="120" IsReadOnly="True"
                           TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                           FontFamily="Consolas" FontSize="11" Foreground="#B0B0B0"
                           Background="Transparent" BorderThickness="0"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="Scheduler">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="14"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- Scripts folder info bar -->
          <Border Grid.Row="0" Background="#1A2A3E" CornerRadius="6"
                  BorderBrush="#4C9FE6" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Scripts Folder" Foreground="#4C9FE6" FontSize="12"
                         FontWeight="SemiBold" VerticalAlignment="Center"/>
              <TextBlock Grid.Column="2" x:Name="SchedFolderPath" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Center" TextTrimming="CharacterEllipsis"
                         ToolTip="Drop .ps1 or .bas macro files into this folder — they will appear in the script dropdown below"/>
              <Button Grid.Column="4" x:Name="SchedOpenFolderBtn" Content="_Open Folder"
                      Style="{StaticResource Btn}" Padding="12,5"
                      ToolTip="Open the Macros folder in Windows Explorer"/>
            </Grid>
          </Border>

          <!-- Create new scheduled task form -->
          <Border Grid.Row="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="20,16">
            <StackPanel>
              <Grid Margin="0,0,0,12">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="SCHEDULE A NEW TASK" Foreground="#A0A0A0"
                           FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1"
                           Text="Backed by Windows Task Scheduler — tasks run even when MacroHub is closed"
                           Foreground="#4C9FE6" FontSize="11" FontStyle="Italic"
                           VerticalAlignment="Center" HorizontalAlignment="Right"/>
              </Grid>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="120"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="130"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="6"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Grid.Column="0" Text="Task Name" Foreground="#B0B0B0" FontSize="12"/>
                <TextBlock Grid.Row="0" Grid.Column="2" Text="Script or Macro File" Foreground="#B0B0B0" FontSize="12"/>
                <TextBlock Grid.Row="0" Grid.Column="4" Text="Run Time (HH:mm)" Foreground="#B0B0B0" FontSize="12"/>
                <TextBlock Grid.Row="0" Grid.Column="6" Text="Frequency" Foreground="#B0B0B0" FontSize="12"/>

                <TextBox   Grid.Row="2" Grid.Column="0" x:Name="SchedNameBox"
                           ToolTip="Unique name for this scheduled task (shown in Windows Task Scheduler)"/>
                <ComboBox  Grid.Row="2" Grid.Column="2" x:Name="SchedMacroCombo"
                           ToolTip="Select a .ps1 or .bas script to run on schedule"/>
                <TextBox   Grid.Row="2" Grid.Column="4" x:Name="SchedTimeBox" Text="09:00"
                           ToolTip="24-hour time when the task runs (e.g. 09:00, 17:30)"/>
                <ComboBox  Grid.Row="2" Grid.Column="6" x:Name="SchedFreqCombo">
                  <ComboBoxItem Content="Daily" IsSelected="True"/>
                  <ComboBoxItem Content="Weekly"/>
                  <ComboBoxItem Content="Monthly"/>
                </ComboBox>
              </Grid>
              <StackPanel Orientation="Horizontal" Margin="0,14,0,0">
                <Button x:Name="SchedCreateBtn" Content="_Create Task" Style="{StaticResource BtnAccent}"
                        Padding="16,8" Margin="0,0,10,0"
                        ToolTip="Register this task in Windows Task Scheduler"/>
                <Button x:Name="SchedRefreshBtn" Content="_Refresh List" Style="{StaticResource Btn}"
                        Padding="14,8" ToolTip="Reload scheduled tasks from Windows Task Scheduler"/>
              </StackPanel>
            </StackPanel>
          </Border>

          <!-- Existing tasks list -->
          <Border Grid.Row="4" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10"
                      Background="#222225">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="90"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="TASK NAME" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="1" Text="STATE" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="2" Text="NEXT RUN" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="3" Text="LAST RUN" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="4" Text="ACTION" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Grid>
              </Border>
              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="SchedTaskPanel" Margin="0,4"/>
              </ScrollViewer>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="Navigator">
        <Grid Margin="24,18">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="220"/>
            <ColumnDefinition Width="12"/>
            <ColumnDefinition Width="240"/>
            <ColumnDefinition Width="12"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <!-- Open workbooks list -->
          <Border Grid.Column="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <TextBlock Text="OPEN WORKBOOKS" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"
                           ToolTip="All Excel workbooks currently open. Click to see its sheets."/>
              </Border>
              <ListBox Grid.Row="1" x:Name="NavWbList" Background="Transparent"
                       Foreground="#FFFFFF" BorderThickness="0" Padding="6,4" FontSize="12"/>
              <Border Grid.Row="2" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="10,8">
                <Button x:Name="NavRefreshBtn" Content="_Refresh Workbooks" Style="{StaticResource Btn}"
                        Padding="10,6" HorizontalAlignment="Stretch"
                        ToolTip="Reload the list of open workbooks from Excel"/>
              </Border>
            </Grid>
          </Border>

          <!-- Sheets list for selected workbook -->
          <Border Grid.Column="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <StackPanel>
                  <TextBlock Text="WORKSHEETS" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                  <TextBlock Text="Multi-select   •   Drag to reorder   •   Ctrl+E to rename"
                             Foreground="#6E6E6E" FontSize="10" Margin="0,3,0,0"/>
                </StackPanel>
              </Border>
              <ListBox Grid.Row="1" x:Name="NavSheetList" Background="Transparent"
                       Foreground="#FFFFFF" BorderThickness="0" Padding="6,4" FontSize="12"
                       SelectionMode="Extended" AllowDrop="True"
                       ToolTip="Select one or more sheets. Drag to reorder. Press Ctrl+E to rename."/>
              <Border Grid.Row="2" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="10,8">
                <TextBlock x:Name="NavSheetCount" Text="0 sheets" Foreground="#909090"
                           FontSize="11" HorizontalAlignment="Center"/>
              </Border>
            </Grid>
          </Border>

          <!-- Navigator right-panel actions -->
          <Border Grid.Column="4" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="20,16">
              <StackPanel>

                <TextBlock Text="ACTIVATE / OPEN" Foreground="#A0A0A0" FontSize="11"
                           FontWeight="SemiBold" Margin="0,0,0,10"/>
                <Button x:Name="NavActivateBtn" Content="_Activate Selected Sheet"
                        Style="{StaticResource BtnAccent}" Padding="14,8" Margin="0,0,0,6"
                        HorizontalAlignment="Stretch"
                        ToolTip="Switch Excel focus to the selected sheet"/>
                <Button x:Name="NavOpenBtn" Content="_Open Workbook File..."
                        Style="{StaticResource Btn}" Padding="14,8" Margin="0,0,0,8"
                        HorizontalAlignment="Stretch"
                        ToolTip="Browse for and open an Excel file"/>
                <Grid Margin="0,0,0,16">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Button Grid.Column="0" x:Name="NavBringFrontBtn" Content="_Bring to Front"
                          Style="{StaticResource Btn}" Padding="8,6" HorizontalAlignment="Stretch"
                          ToolTip="Bring the Excel window to the foreground"/>
                  <Button Grid.Column="2" x:Name="NavMinimizeBtn" Content="_Minimize"
                          Style="{StaticResource Btn}" Padding="8,6" HorizontalAlignment="Stretch"
                          ToolTip="Minimize the Excel window"/>
                  <Button Grid.Column="4" x:Name="NavCloseWbBtn" Content="_Close Workbook"
                          Style="{StaticResource BtnRed}" Padding="8,6" HorizontalAlignment="Stretch"
                          ToolTip="Close the selected workbook (prompts to save if unsaved)"/>
                </Grid>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,10">
                  <TextBlock Text="MOVE / COPY SHEET" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <TextBlock Text="Destination Workbook" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,5"/>
                <ComboBox x:Name="NavDestWbCombo" Margin="0,0,0,10"
                          ToolTip="The workbook to move or copy the sheet into"/>
                <TextBlock Text="Insert Before Sheet" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,5"
                           ToolTip="The sheet will be inserted before this one (leave blank to move to end)"/>
                <ComboBox x:Name="NavDestSheetCombo" Margin="0,0,0,8"/>
                <CheckBox x:Name="NavCopyChk" Content="Copy (keep original in source workbook)"
                          Foreground="#FFFFFF" IsChecked="True" Margin="0,0,0,10" FontSize="11"/>
                <Button x:Name="NavMoveCopyBtn" Content="_Move / Copy Sheet"
                        Style="{StaticResource BtnAccent}" Padding="14,8" Margin="0,0,0,6"
                        HorizontalAlignment="Stretch"
                        ToolTip="Move or copy the selected sheet(s) to the destination workbook"/>
                <Button x:Name="NavExportSheetBtn" Content="E_xport Sheet to CSV/XLSX..."
                        Style="{StaticResource Btn}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"
                        ToolTip="Export the selected sheet as a standalone file"/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,10">
                  <TextBlock Text="VISIBILITY" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <Grid Margin="0,0,0,16">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Button Grid.Column="0" x:Name="NavHideBtn" Content="_Hide Sheet"
                          Style="{StaticResource Btn}" Padding="14,8" HorizontalAlignment="Stretch"
                          ToolTip="Hide the selected sheets"/>
                  <Button Grid.Column="2" x:Name="NavUnhideBtn" Content="_Unhide Sheet"
                          Style="{StaticResource BtnGreen}" Padding="14,8" HorizontalAlignment="Stretch"
                          ToolTip="Unhide hidden sheets in the selected workbook"/>
                </Grid>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,10">
                  <TextBlock Text="DELETE" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <Button x:Name="NavDeleteBtn" Content="_Delete Selected Sheets"
                        Style="{StaticResource BtnRed}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"
                        ToolTip="Permanently delete the selected sheets (cannot be undone)"/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,10">
                  <TextBlock Text="PASSWORD PROTECTION" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <TextBlock Text="Password" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,5"/>
                <TextBox x:Name="NavPwdBox" FontSize="12" Foreground="#FFFFFF" Background="#1B1B1F"
                         BorderBrush="#3E3E42" BorderThickness="1" Padding="8,6" Margin="0,0,0,8"
                         ToolTip="Enter the password to set, remove, or use when opening a protected workbook"/>
                <Button x:Name="NavSetPwdBtn" Content="Set _Password on Workbook"
                        Style="{StaticResource BtnAccent}" Padding="14,8" Margin="0,0,0,6"
                        HorizontalAlignment="Stretch"
                        ToolTip="Protect the active workbook with the password above"/>
                <Button x:Name="NavRemPwdBtn" Content="_Remove Password"
                        Style="{StaticResource Btn}" Padding="14,8" Margin="0,0,0,6"
                        HorizontalAlignment="Stretch"
                        ToolTip="Remove the password from the active workbook"/>
                <Button x:Name="NavOpenPwdBtn" Content="Open _With Password..."
                        Style="{StaticResource Btn}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"
                        ToolTip="Open a password-protected workbook file"/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,10">
                  <TextBlock Text="EXCEL ENGINE OPTIONS" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <TextBlock Text="Calculation Mode" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,5"/>
                <ComboBox x:Name="NavCalcModeCombo" Margin="0,0,0,8">
                  <ComboBoxItem Content="Automatic" IsSelected="True"/>
                  <ComboBoxItem Content="Manual"/>
                  <ComboBoxItem Content="Semiautomatic"/>
                </ComboBox>
                <CheckBox x:Name="NavEventsChk" Content="Enable worksheet events"
                          Foreground="#FFFFFF" IsChecked="True" Margin="0,0,0,10" FontSize="11"
                          ToolTip="Toggle Excel's event system (disable to suppress change/calculate triggers)"/>
                <Button x:Name="NavApplyExcelOptsBtn" Content="_Apply Excel Options"
                        Style="{StaticResource BtnAccent}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"
                        ToolTip="Apply the calculation mode and events setting to the active Excel session"/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,10">
                  <TextBlock Text="VBA MODULES / MACROS" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <ListBox x:Name="NavVbaList" Height="120" Background="#1B1B1F" Foreground="#C0C0C0"
                         BorderBrush="#3E3E42" BorderThickness="1" Margin="0,0,0,16" FontSize="11"
                         ToolTip="Right-click a macro entry to run it in the selected navigator workbook."/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,10">
                  <TextBlock Text="WORKBOOK INFO" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <TextBox x:Name="NavInfoTxt" IsReadOnly="True" Height="140"
                         TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                         FontFamily="Consolas" FontSize="11" Foreground="#B0B0B0"
                         Background="Transparent" BorderThickness="0"/>

              </StackPanel>
            </ScrollViewer>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="Templates">
        <Grid Margin="24,18">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="280"/>
            <ColumnDefinition Width="16"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <!-- Saved templates list (left panel) -->
          <Border Grid.Column="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <TextBlock Text="SAVED TEMPLATES" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"
                           ToolTip="Click a template to load it into the editor on the right"/>
              </Border>
              <ListBox Grid.Row="1" x:Name="TplList" Background="Transparent"
                       Foreground="#FFFFFF" BorderThickness="0" Padding="6,4" FontSize="12"/>
              <Border Grid.Row="2" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="10,8">
                <Button x:Name="TplDeleteBtn" Content="_Delete Selected Template"
                        Style="{StaticResource BtnRed}" Padding="10,6" HorizontalAlignment="Stretch"
                        ToolTip="Permanently delete the selected template"/>
              </Border>
            </Grid>
          </Border>

          <!-- Template editor (right panel) -->
          <Border Grid.Column="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="20,18">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="10"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="6"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="6"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="14"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="10"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <TextBlock Grid.Row="0" Text="TEMPLATE EDITOR" Foreground="#A0A0A0"
                         FontSize="11" FontWeight="SemiBold"/>

              <TextBlock Grid.Row="2" Text="Template Name" Foreground="#B0B0B0" FontSize="12"/>
              <TextBox Grid.Row="4" x:Name="TplNameBox"
                       ToolTip="A short identifying name for this template"/>

              <TextBlock Grid.Row="6" Text="Template Content" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Top" Margin="0,0,0,6"/>
              <TextBox Grid.Row="6" x:Name="TplContentBox" AcceptsReturn="True"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Margin="0,22,0,0"
                       ToolTip="Template body — use {DATE}, {SHEET}, or {USER} as dynamic placeholders"
                       FontFamily="Consolas" FontSize="11"/>

              <StackPanel Grid.Row="8" Orientation="Horizontal">
                <Button x:Name="TplSaveBtn" Content="_Save Template" Style="{StaticResource BtnAccent}"
                        Padding="14,8" Margin="0,0,10,0"
                        ToolTip="Save this template (creates new or updates existing by name)"/>
                <Button x:Name="TplPasteBtn" Content="_Paste to Excel" Style="{StaticResource BtnGreen}"
                        Padding="14,8" Margin="0,0,10,0"
                        ToolTip="Paste this template's content into the active Excel cell (substitutes placeholders)"/>
                <Button x:Name="TplPreviewBtn" Content="Pre_view" Style="{StaticResource Btn}"
                        Padding="14,8"
                        ToolTip="Show a preview of the template with placeholders substituted"/>
              </StackPanel>

              <Border Grid.Row="10" x:Name="TplPreviewCard" Visibility="Collapsed"
                      Background="#1E2D1E" CornerRadius="6" BorderBrush="#50A050"
                      BorderThickness="1" Padding="14,12">
                <StackPanel>
                  <TextBlock Text="PREVIEW — placeholder values substituted" Foreground="#50A050"
                             FontSize="11" FontWeight="SemiBold" Margin="0,0,0,8"/>
                  <TextBox x:Name="TplPreviewTxt" IsReadOnly="True" FontFamily="Consolas" FontSize="11"
                           AcceptsReturn="True" TextWrapping="Wrap" MaxHeight="160"
                           VerticalScrollBarVisibility="Auto" Background="#1B1B1F" Foreground="#FFFFFF"
                           BorderThickness="0"/>
                </StackPanel>
              </Border>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="QSync">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel Margin="24,18" MaxWidth="900">

            <!-- Quarter folder comparison setup -->
            <Border Background="#252528" CornerRadius="6" BorderBrush="#3E3E42"
                    BorderThickness="1" Padding="20,16" Margin="0,0,0,12">
              <StackPanel>
                <TextBlock Text="QUARTER FOLDER SYNC" Foreground="#A0A0A0" FontSize="11"
                           FontWeight="SemiBold" Margin="0,0,0,6"/>
                <TextBlock Text="Compare last quarter's folder structure against this quarter to build a new to-do checklist. Run once at the start of each quarter."
                           Foreground="#909090" FontSize="11" FontStyle="Italic" Margin="0,0,0,14"/>

                <TextBlock Text="Quarter Name" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
                <TextBox x:Name="QsSyncQuarterName" ToolTip="e.g. Q1 2025" Margin="0,0,0,12"/>

                <TextBlock Text="Last Quarter Root" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
                <Grid Margin="0,0,0,12">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="90"/>
                  </Grid.ColumnDefinitions>
                  <TextBox Grid.Column="0" x:Name="QsLastRoot" ToolTip="e.g. Z:\Reports\Q3_2024"/>
                  <Button  Grid.Column="2" x:Name="QsBrowseLast" Content="Bro_wse" Style="{StaticResource Btn}"/>
                </Grid>

                <TextBlock Text="This Quarter Root" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
                <Grid Margin="0,0,0,14">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="90"/>
                  </Grid.ColumnDefinitions>
                  <TextBox Grid.Column="0" x:Name="QsThisRoot" ToolTip="e.g. Z:\Reports\Q4_2024"/>
                  <Button  Grid.Column="2" x:Name="QsBrowseThis" Content="_Browse" Style="{StaticResource Btn}"/>
                </Grid>

                <Button x:Name="QsRunSyncBtn" Content="_Run Sync" Style="{StaticResource BtnAccent}"
                        Padding="20,10" FontSize="13" HorizontalAlignment="Left"/>
              </StackPanel>
            </Border>

            <Border x:Name="QsProgressCard" Visibility="Collapsed" Margin="0,0,0,12"
                    Background="#252528" CornerRadius="6" BorderBrush="#3E3E42"
                    BorderThickness="1" Padding="20,12">
              <StackPanel>
                <TextBlock x:Name="QsProgressTxt" Text="Running..." Foreground="#4C9FE6" FontSize="12"/>
                <ProgressBar IsIndeterminate="True" Height="3" Margin="0,8,0,0"
                             Background="#3E3E42" Foreground="#4C9FE6"/>
              </StackPanel>
            </Border>

            <Border x:Name="QsResultCard" Visibility="Collapsed" Margin="0,0,0,12"
                    Background="#252528" CornerRadius="6" BorderBrush="#3E3E42"
                    BorderThickness="1" Padding="20,16">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                  <TextBlock x:Name="QsRCreated" Text="0" Foreground="#4C9FE6" FontSize="28" FontWeight="Bold"/>
                  <TextBlock Text="Folders Created" Foreground="#6E6E6E" FontSize="11"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <TextBlock x:Name="QsRSkipped" Text="0" Foreground="#6E6E6E" FontSize="28" FontWeight="Bold"/>
                  <TextBlock Text="Already Existed" Foreground="#6E6E6E" FontSize="11"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock x:Name="QsRNewTodos" Text="0" Foreground="#FFFFFF" FontSize="28" FontWeight="Bold"/>
                  <TextBlock Text="New To-Do Items" Foreground="#6E6E6E" FontSize="11"/>
                </StackPanel>
                <StackPanel Grid.Column="3">
                  <TextBlock x:Name="QsRErrors" Text="0" Foreground="#E05050" FontSize="28" FontWeight="Bold"/>
                  <TextBlock Text="Errors" Foreground="#6E6E6E" FontSize="11"/>
                </StackPanel>
              </Grid>
            </Border>

            <Grid Margin="0,0,0,12">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>

              <Border Grid.Column="0" Background="#252528" CornerRadius="6"
                      BorderBrush="#3E3E42" BorderThickness="1" Padding="16,12">
                <StackPanel>
                  <TextBlock x:Name="QsBigPct" Text="0%" Foreground="#4C9FE6" FontSize="32" FontWeight="Bold"/>
                  <TextBlock Text="Overall Complete" Foreground="#6E6E6E" FontSize="11"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="2" Background="#252528" CornerRadius="6"
                      BorderBrush="#3E3E42" BorderThickness="1" Padding="16,12">
                <StackPanel>
                  <TextBlock x:Name="QsBigTotal" Text="0" Foreground="#FFFFFF" FontSize="32" FontWeight="Bold"/>
                  <TextBlock Text="Total Tracked" Foreground="#6E6E6E" FontSize="11"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="4" Background="#252528" CornerRadius="6"
                      BorderBrush="#3E3E42" BorderThickness="1" Padding="16,12">
                <StackPanel>
                  <TextBlock x:Name="QsBigPending" Text="0" Foreground="#B0B0B0" FontSize="32" FontWeight="Bold"/>
                  <TextBlock Text="Still Pending" Foreground="#6E6E6E" FontSize="11"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="6" Background="#252528" CornerRadius="6"
                      BorderBrush="#3E3E42" BorderThickness="1" Padding="16,12">
                <StackPanel>
                  <TextBlock x:Name="QsBigDone" Text="0" Foreground="#50A050" FontSize="32" FontWeight="Bold"/>
                  <TextBlock Text="Completed" Foreground="#6E6E6E" FontSize="11"/>
                </StackPanel>
              </Border>
            </Grid>

            <Border Background="#252528" CornerRadius="6" BorderBrush="#3E3E42"
                    BorderThickness="1" Margin="0,0,0,12">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                  <TextBlock Text="COMPLETION BY FOLDER" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="200">
                  <StackPanel x:Name="QsFolderBars" Margin="14,8,14,10"/>
                </ScrollViewer>
              </Grid>
            </Border>

            <Border Background="#252528" CornerRadius="6" BorderBrush="#3E3E42" BorderThickness="1">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                  <TextBlock Text="SYNC LOG" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <TextBox Grid.Row="1" x:Name="QsSyncLogTxt" Height="140" IsReadOnly="True"
                         TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                         Background="Transparent" BorderThickness="0"
                         Foreground="#6E6E6E" FontSize="11" FontFamily="Consolas"
                         Margin="14,8" Padding="0"/>
              </Grid>
            </Border>

          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <TabItem Header="QTasks">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="14"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- Folder pair selector (compare Source A vs Folder B) -->
          <Border Grid.Row="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="16,12">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="24"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="200"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>

              <!-- Source A folder -->
              <TextBlock Grid.Column="0" Text="Folder A (baseline)" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Center" ToolTip="The reference folder — usually last quarter"/>
              <ComboBox  Grid.Column="2" x:Name="QsQuarterCombo" VerticalAlignment="Center" IsEditable="True"
                         ToolTip="Type or browse to the baseline (source) folder path"/>
              <Button    Grid.Column="4" x:Name="QsNewQuarterBtn" Content="_Browse A"
                         Style="{StaticResource Btn}" Padding="10,5"
                         ToolTip="Browse for Folder A (baseline)"/>

              <!-- VS separator -->
              <TextBlock Grid.Column="5" Text="vs" Foreground="#4C9FE6" FontSize="13" FontWeight="Bold"
                         VerticalAlignment="Center" HorizontalAlignment="Center"/>

              <!-- Folder B -->
              <TextBlock Grid.Column="6" Text="Folder B (target)" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Center" ToolTip="The delivery or current-quarter folder"/>
              <ComboBox  Grid.Column="8" x:Name="QsCompareCombo" VerticalAlignment="Center" IsEditable="True"
                         ToolTip="Type or browse to the target (delivery) folder path"/>
              <Button    Grid.Column="10" x:Name="QsScanFolderBtn" Content="_Browse B"
                         Style="{StaticResource Btn}" Padding="10,5"
                         ToolTip="Browse for Folder B (target)"/>

              <!-- Compare action + tag -->
              <Button    Grid.Column="12" x:Name="QsAddTaskBtn" Content="_Find Missing Files"
                         Style="{StaticResource BtnAccent}" Padding="12,5"
                         ToolTip="Compare Folder A vs Folder B and add any missing files as to-do items"/>
              <TextBox   Grid.Column="14" x:Name="QsScanPathBox"
                         ToolTip="Optional tag or note saved with compared items (e.g. 'Q1 delivery')"/>

              <!-- Summary + swap -->
              <TextBlock Grid.Column="16" x:Name="QsQuarterSummary" Foreground="#909090"
                         FontSize="11" VerticalAlignment="Center" Margin="0,0,8,0"/>
              <TextBlock Grid.Column="17" x:Name="QsQuarterPct" Foreground="#4C9FE6"
                         FontSize="13" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0"/>
              <Button    Grid.Column="18" x:Name="QsSwapFoldersBtn" Content="⇄ _Swap"
                         Style="{StaticResource Btn}" Padding="10,5"
                         ToolTip="Swap Folder A and Folder B"/>
            </Grid>
          </Border>

          <!-- Add task inline form (hidden until triggered) -->
          <Border Grid.Row="1" x:Name="QsAddTaskCard" Visibility="Collapsed"
                  Background="#2A2A2E" CornerRadius="6" BorderBrush="#4C9FE6"
                  BorderThickness="1" Padding="14,10" Margin="0,8,0,0">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="130"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox Grid.Column="0" x:Name="QsNewTaskName" ToolTip="File or task name"/>
              <TextBox Grid.Column="2" x:Name="QsNewTaskFolder" ToolTip="Subfolder (optional)"/>
              <TextBox Grid.Column="4" x:Name="QsNewTaskDue" ToolTip="Due date (yyyy-MM-dd)"/>
              <Button Grid.Column="6" x:Name="QsSaveTaskBtn" Content="_Save Task"
                      Style="{StaticResource BtnAccent}" Padding="12,5"/>
              <Button Grid.Column="8" x:Name="QsCancelTaskBtn" Content="_Cancel"
                      Style="{StaticResource Btn}" Padding="12,5"/>
            </Grid>
          </Border>

          <!-- Filter bar -->
          <Border Grid.Row="3" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="120"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>

              <TextBlock Grid.Column="0" Text="Status" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Center"/>
              <ComboBox  Grid.Column="2" x:Name="QsFltStatus" VerticalAlignment="Center">
                <ComboBoxItem Content="All" IsSelected="True"/>
                <ComboBoxItem Content="Pending"/>
                <ComboBoxItem Content="Done"/>
              </ComboBox>

              <TextBlock Grid.Column="4" Text="File name" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Center"/>
              <TextBox   Grid.Column="6" x:Name="QsFltName" ToolTip="Filter results by file name (partial match)"/>

              <TextBlock Grid.Column="8" Text="Folder" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Center"/>
              <TextBox   Grid.Column="10" x:Name="QsFltFolder" ToolTip="Filter results by folder name (partial match)"/>

              <Button Grid.Column="12" x:Name="QsExportBtn" Content="_Export to Excel"
                      Style="{StaticResource BtnAccent}" Padding="12,6"
                      ToolTip="Export the current task list to an Excel workbook"/>
              <Button Grid.Column="14" x:Name="QsRefreshBtn" Content="_Refresh"
                      Style="{StaticResource Btn}" Padding="12,6"
                      ToolTip="Reload the task list applying current filters"/>
            </Grid>
          </Border>

          <!-- Task list grid -->
          <Border Grid.Row="5" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <!-- Column headers -->
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10"
                      Background="#222225">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="180"/>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="210"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="FILE NAME" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="1" Text="FOLDER" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="2" Text="NEW IN A" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"
                             ToolTip="File was added since the last quarter baseline"/>
                  <TextBlock Grid.Column="3" Text="CHANGED IN A" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"
                             ToolTip="File was modified since the last quarter baseline"/>
                  <TextBlock Grid.Column="4" Text="ACTIONS" Foreground="#A0A0A0" FontSize="11" FontWeight="SemiBold"/>
                </Grid>
              </Border>

              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="QsTodoPanel" Margin="0,4"/>
              </ScrollViewer>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <!-- ═══════════════════════════════════════════════════════════ -->
      <!-- FILE INDEX TAB — DLP-safe metadata-only drive search       -->
      <!-- ═══════════════════════════════════════════════════════════ -->
      <TabItem Header="File Index">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Root folder selector + index controls -->
          <Border Grid.Row="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="16,12">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Root Folder" Foreground="#B0B0B0" FontSize="12"
                         VerticalAlignment="Center"
                         ToolTip="The folder to scan for file metadata. Subfolders are included."/>
              <TextBox   Grid.Column="2" x:Name="FidxRootBox"
                         ToolTip="Type or paste a folder path, then click Index Now to scan it"/>
              <Button    Grid.Column="4" x:Name="FidxBrowseBtn" Content="_Browse..."
                         Style="{StaticResource Btn}" Padding="12,5"
                         ToolTip="Open a folder picker dialog"/>
              <Button    Grid.Column="6" x:Name="FidxIndexBtn" Content="_Index Now"
                         Style="{StaticResource BtnAccent}" Padding="14,5"
                         ToolTip="Scan the root folder and cache file metadata (60-minute cooldown after each scan)"/>
              <TextBlock Grid.Column="8" x:Name="FidxCooldownTxt" Foreground="#909090"
                         FontSize="11" VerticalAlignment="Center"
                         ToolTip="Time remaining before next scan is allowed"/>
            </Grid>
          </Border>

          <!-- Search bar -->
          <Border Grid.Row="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox Grid.Column="0" x:Name="FidxSearchBox" FontSize="13"
                       ToolTip="Filter by file name, folder path, or extension — results update as you type"/>
              <TextBlock Grid.Column="2" x:Name="FidxMatchCount" Foreground="#4C9FE6"
                         FontSize="12" FontWeight="Bold" VerticalAlignment="Center"
                         ToolTip="Number of files matching the current search"/>
            </Grid>
          </Border>

          <!-- Results DataGrid -->
          <Border Grid.Row="4" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <DataGrid x:Name="FidxGrid"
                      AutoGenerateColumns="False"
                      IsReadOnly="True"
                      SelectionMode="Single"
                      GridLinesVisibility="Horizontal"
                      Background="Transparent"
                      RowBackground="#252528"
                      AlternatingRowBackground="#202023"
                      Foreground="#FFFFFF"
                      BorderThickness="0"
                      ColumnHeaderHeight="32"
                      RowHeight="26"
                      FontSize="12"
                      CanUserReorderColumns="True"
                      CanUserResizeColumns="True"
                      CanUserSortColumns="True"
                      HeadersVisibility="Column"
                      ToolTip="Double-click or press Enter to open a file. Click a column header to sort.">
              <DataGrid.Resources>
                <Style TargetType="DataGridColumnHeader">
                  <Setter Property="Background"   Value="#222225"/>
                  <Setter Property="Foreground"   Value="#A0A0A0"/>
                  <Setter Property="FontSize"     Value="11"/>
                  <Setter Property="FontWeight"   Value="SemiBold"/>
                  <Setter Property="Padding"      Value="10,0"/>
                  <Setter Property="BorderBrush"  Value="#3E3E42"/>
                  <Setter Property="BorderThickness" Value="0,0,1,1"/>
                </Style>
                <Style TargetType="DataGridRow">
                  <Setter Property="BorderThickness" Value="0"/>
                  <Style.Triggers>
                    <Trigger Property="IsSelected" Value="True">
                      <Setter Property="Background" Value="#1A2A3E"/>
                    </Trigger>
                    <Trigger Property="IsMouseOver" Value="True">
                      <Setter Property="Background" Value="#2D2D30"/>
                    </Trigger>
                  </Style.Triggers>
                </Style>
                <Style TargetType="DataGridCell">
                  <Setter Property="BorderThickness" Value="0"/>
                  <Setter Property="Padding"         Value="8,0"/>
                  <Style.Triggers>
                    <Trigger Property="IsSelected" Value="True">
                      <Setter Property="Background"  Value="Transparent"/>
                      <Setter Property="Foreground"  Value="#FFFFFF"/>
                    </Trigger>
                  </Style.Triggers>
                </Style>
              </DataGrid.Resources>
              <DataGrid.Columns>
                <DataGridTextColumn Header="File Name" Binding="{Binding Name}"   Width="220"/>
                <DataGridTextColumn Header="Path"      Binding="{Binding Path}"   Width="*"/>
                <DataGridTextColumn Header="Size"      Binding="{Binding SizeStr}" Width="80"/>
                <DataGridTextColumn Header="Modified"  Binding="{Binding Modified}" Width="140"/>
                <DataGridTextColumn Header="Ext"       Binding="{Binding Ext}"    Width="60"/>
              </DataGrid.Columns>
            </DataGrid>
          </Border>

          <!-- Status bar -->
          <Border Grid.Row="5" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,8" Margin="0,8,0,0">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" x:Name="FidxStatusTxt"
                         Foreground="#909090" FontSize="11" VerticalAlignment="Center"
                         Text="No index loaded. Select a root folder and click Index Now."/>
              <TextBlock Grid.Column="2" Text="Total files:" Foreground="#6E6E6E"
                         FontSize="11" VerticalAlignment="Center"/>
              <TextBlock Grid.Column="3" x:Name="FidxTotalCount" Foreground="#B0B0B0"
                         FontSize="11" FontWeight="Bold" VerticalAlignment="Center" Text="0"/>
              <TextBlock Grid.Column="5" x:Name="FidxLastScanTxt" Foreground="#6E6E6E"
                         FontSize="11" VerticalAlignment="Center"/>
            </Grid>
          </Border>

        </Grid>
      </TabItem>

      <!-- ════════════════════════════════════════════════════════════ -->
      <!-- FORMULA AUDITOR — read-only inspection + temp highlighting  -->
      <!-- ════════════════════════════════════════════════════════════ -->
      <TabItem Header="Formula Auditor">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Top toolbar: cell address, formula, action buttons -->
          <Border Grid.Row="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="100"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Active Cell:" Foreground="#B0B0B0"
                         FontSize="12" VerticalAlignment="Center"/>
              <TextBlock Grid.Column="2" x:Name="FaActiveCellTxt" Text="—"
                         Foreground="#4C9FE6" FontSize="13" FontWeight="Bold"
                         VerticalAlignment="Center"/>
              <TextBlock Grid.Column="4" Text="Formula / Value:" Foreground="#B0B0B0"
                         FontSize="12" VerticalAlignment="Center"/>
              <TextBlock Grid.Column="6" x:Name="FaFormulaTxt" Text="—"
                         Foreground="#E0E0E0" FontSize="12" VerticalAlignment="Center"
                         TextTrimming="CharacterEllipsis"/>
              <Button Grid.Column="8"  x:Name="FaInspectBtn"  Content="_Inspect"
                      Style="{StaticResource BtnAccent}" Padding="12,5"
                      ToolTip="Inspect the currently selected cell in Excel"/>
              <Button Grid.Column="10" x:Name="FaBackBtn"     Content="◀ Back"
                      Style="{StaticResource Btn}" Padding="10,5" IsEnabled="False"
                      ToolTip="Navigate back to the previous cell"/>
              <Button Grid.Column="12" x:Name="FaAuditRangeBtn" Content="_Audit Range"
                      Style="{StaticResource Btn}" Padding="10,5"
                      ToolTip="Outline formula cells in the selected range with a temporary blue glow. Click Clear to remove highlights."/>
              <Button Grid.Column="14" x:Name="FaClearHighlightsBtn" Content="✕ Clear Highlights"
                      Style="{StaticResource BtnRed}" Padding="10,5" IsEnabled="False"
                      ToolTip="Remove ALL temporary borders and background highlights Formula Auditor added to the spreadsheet"/>
              <CheckBox Grid.Column="16" x:Name="FaAutoBox" Content="Auto"
                        Foreground="#B0B0B0" FontSize="11" IsChecked="True"
                        VerticalAlignment="Center"
                        ToolTip="Automatically re-inspect when you move to a different cell (polls every 500 ms)"/>
            </Grid>
          </Border>

          <!-- Main area: left = inspector + visualizer; right = dependency tree -->
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="270"/>
            </Grid.ColumnDefinitions>

            <!-- Left column: Cell Inspector + SUMIF Visualizer -->
            <Grid Grid.Column="0">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="10"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <!-- Cell Inspector card -->
              <Border Grid.Row="0" Background="#252528" CornerRadius="6"
                      BorderBrush="#3E3E42" BorderThickness="1" Padding="14,12">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <!-- Type / Value / Sheet row -->
                  <StackPanel Grid.Row="0" Orientation="Horizontal">
                    <TextBlock Text="Type: "  Foreground="#909090" FontSize="11"/>
                    <TextBlock x:Name="FaCellTypeTxt"  Text="—" Foreground="#B0B0B0"
                               FontSize="11" FontWeight="SemiBold"/>
                    <TextBlock Text="     Value: " Foreground="#909090" FontSize="11"/>
                    <TextBlock x:Name="FaCellValueTxt" Text="—" Foreground="#B0B0B0" FontSize="11"/>
                    <TextBlock Text="     Sheet: " Foreground="#909090" FontSize="11"/>
                    <TextBlock x:Name="FaCellSheetTxt" Text="—" Foreground="#B0B0B0" FontSize="11"/>
                  </StackPanel>
                  <!-- Precedents label -->
                  <TextBlock Grid.Row="2"
                             Text="Precedents — click → to navigate:"
                             Foreground="#909090" FontSize="11"/>
                  <!-- Precedents container -->
                  <Border Grid.Row="4" Background="#1B1B1F" CornerRadius="4"
                          BorderBrush="#3E3E42" BorderThickness="1"
                          MinHeight="34" MaxHeight="90">
                    <Grid>
                      <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="FaPrecedentsPanel" Margin="6,3"/>
                      </ScrollViewer>
                      <TextBlock x:Name="FaNoPrecedentsTxt"
                                 Text="No precedents — cell is hardcoded or empty."
                                 Foreground="#6E6E6E" FontSize="11"
                                 Margin="10,8" Visibility="Collapsed"/>
                    </Grid>
                  </Border>
                  <!-- Dependents count -->
                  <StackPanel Grid.Row="6" Orientation="Horizontal">
                    <TextBlock Text="Dependents: " Foreground="#909090" FontSize="11"/>
                    <TextBlock x:Name="FaDependentCountTxt" Text="—"
                               Foreground="#B0B0B0" FontSize="11"/>
                  </StackPanel>
                </Grid>
              </Border>

              <!-- SUMIF / COUNTIF / AVERAGEIF Visualizer card -->
              <Border Grid.Row="2" Background="#252528" CornerRadius="6"
                      BorderBrush="#3E3E42" BorderThickness="1">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                  </Grid.RowDefinitions>
                  <!-- Header -->
                  <Border Grid.Row="0" Background="#222225" CornerRadius="6,6,0,0"
                          BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,8">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock Grid.Column="0" x:Name="FaSumifTitleTxt"
                                 Text="SUMIF Visualizer — inspect a SUMIF / SUMIFS / COUNTIF / AVERAGEIF cell"
                                 Foreground="#909090" FontSize="11" FontWeight="SemiBold"/>
                      <TextBlock Grid.Column="1" x:Name="FaSumifTotalTxt"
                                 Text="" Foreground="#4C9FE6" FontSize="12" FontWeight="Bold"/>
                    </Grid>
                  </Border>
                  <!-- Body: hint OR DataGrid -->
                  <Grid Grid.Row="1">
                    <TextBlock x:Name="FaSumifEmptyTxt" Visibility="Visible"
                               Text="Inspect a SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, or AVERAGEIFS cell to see which rows contribute to the result.  Matching rows are highlighted here and (temporarily) in the spreadsheet — no edits are ever made."
                               Foreground="#6E6E6E" FontSize="11" TextWrapping="Wrap"
                               Margin="14,10" VerticalAlignment="Top"/>
                    <DataGrid x:Name="FaSumifGrid"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Single" GridLinesVisibility="Horizontal"
                              Background="Transparent"
                              RowBackground="#252528" AlternatingRowBackground="#202023"
                              Foreground="#E0E0E0" BorderThickness="0"
                              ColumnHeaderHeight="28" RowHeight="24"
                              FontSize="11" HeadersVisibility="Column"
                              Visibility="Collapsed">
                      <DataGrid.Resources>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="Background"      Value="#222225"/>
                          <Setter Property="Foreground"      Value="#A0A0A0"/>
                          <Setter Property="FontSize"        Value="11"/>
                          <Setter Property="FontWeight"      Value="SemiBold"/>
                          <Setter Property="Padding"         Value="8,0"/>
                          <Setter Property="BorderBrush"     Value="#3E3E42"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                        <Style TargetType="DataGridRow">
                          <Setter Property="BorderThickness" Value="0"/>
                          <Style.Triggers>
                            <DataTrigger Binding="{Binding IsMatch}" Value="True">
                              <Setter Property="Background" Value="#1A3A1A"/>
                            </DataTrigger>
                            <Trigger Property="IsSelected" Value="True">
                              <Setter Property="Background" Value="#1A2A3E"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                              <Setter Property="Background" Value="#2D2D30"/>
                            </Trigger>
                          </Style.Triggers>
                        </Style>
                        <Style TargetType="DataGridCell">
                          <Setter Property="BorderThickness" Value="0"/>
                          <Setter Property="Padding"         Value="6,0"/>
                          <Style.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                              <Setter Property="Background" Value="Transparent"/>
                              <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                          </Style.Triggers>
                        </Style>
                      </DataGrid.Resources>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Row"  Binding="{Binding RowNum}"      Width="50"/>
                        <DataGridTextColumn Header="✓"    Binding="{Binding MatchSymbol}" Width="30"/>
                        <DataGridTextColumn Header="Criteria Value"
                                            Binding="{Binding CriteriaVal}" Width="*"/>
                        <DataGridTextColumn Header="Sum / Count Value"
                                            Binding="{Binding SumVal}"      Width="110"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </Grid>
                </Grid>
              </Border>
            </Grid>

            <!-- Right column: Dependency Tree -->
            <Border Grid.Column="2" Background="#252528" CornerRadius="6"
                    BorderBrush="#3E3E42" BorderThickness="1">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Border Grid.Row="0" Background="#222225" CornerRadius="6,6,0,0"
                        BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,8">
                  <TextBlock Text="Dependency Tree"
                             Foreground="#B0B0B0" FontSize="12" FontWeight="SemiBold"/>
                </Border>
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="8,6">
                  <StackPanel>
                    <TextBlock Text="↑ PRECEDENTS (reads from):"
                               Foreground="#909090" FontSize="10" FontWeight="SemiBold"
                               Margin="0,2,0,4"/>
                    <TreeView x:Name="FaPrecedentsTree"
                              Background="Transparent" BorderThickness="0"
                              Foreground="#B0B0B0" FontSize="11">
                      <TreeView.ItemContainerStyle>
                        <Style TargetType="TreeViewItem">
                          <Setter Property="Background"  Value="Transparent"/>
                          <Setter Property="Foreground"  Value="#B0B0B0"/>
                          <Setter Property="Padding"     Value="2,2"/>
                          <Setter Property="IsExpanded"  Value="True"/>
                          <Style.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                              <Setter Property="Background" Value="#1A2A3E"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                              <Setter Property="Background" Value="#2D2D30"/>
                            </Trigger>
                          </Style.Triggers>
                        </Style>
                      </TreeView.ItemContainerStyle>
                    </TreeView>
                    <TextBlock Text="↓ DEPENDENTS (used by):"
                               Foreground="#909090" FontSize="10" FontWeight="SemiBold"
                               Margin="0,14,0,4"/>
                    <TreeView x:Name="FaDependentsTree"
                              Background="Transparent" BorderThickness="0"
                              Foreground="#B0B0B0" FontSize="11">
                      <TreeView.ItemContainerStyle>
                        <Style TargetType="TreeViewItem">
                          <Setter Property="Background"  Value="Transparent"/>
                          <Setter Property="Foreground"  Value="#B0B0B0"/>
                          <Setter Property="Padding"     Value="2,2"/>
                          <Setter Property="IsExpanded"  Value="True"/>
                          <Style.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                              <Setter Property="Background" Value="#1A2A3E"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                              <Setter Property="Background" Value="#2D2D30"/>
                            </Trigger>
                          </Style.Triggers>
                        </Style>
                      </TreeView.ItemContainerStyle>
                    </TreeView>
                  </StackPanel>
                </ScrollViewer>
              </Grid>
            </Border>
          </Grid>

          <!-- Range Audit footer strip -->
          <Border Grid.Row="4" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="6"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <StackPanel Grid.Row="0" Orientation="Horizontal">
                <TextBlock Text="Range Audit:" Foreground="#909090" FontSize="11"
                           FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock x:Name="FaRangeAddrTxt" Text="—" Foreground="#4C9FE6"
                           FontSize="11" VerticalAlignment="Center" Margin="0,0,16,0"/>
                <TextBlock Text="Formulas:" Foreground="#909090" FontSize="11"
                           VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBlock x:Name="FaRangeFormulaTxt" Text="—" Foreground="#B0B0B0"
                           FontSize="11" VerticalAlignment="Center" Margin="0,0,16,0"/>
                <TextBlock Text="Hardcoded:" Foreground="#909090" FontSize="11"
                           VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBlock x:Name="FaRangeHardcodedTxt" Text="—" Foreground="#B0B0B0"
                           FontSize="11" VerticalAlignment="Center" Margin="0,0,16,0"/>
                <TextBlock x:Name="FaRangeExternalTxt" Text="" Foreground="#E07000"
                           FontSize="11" FontWeight="Bold" VerticalAlignment="Center"/>
              </StackPanel>
              <TextBlock Grid.Row="2" x:Name="FaRangeDetailTxt"
                         Text="Select a range in Excel and click Audit Range to analyze formula consistency. Formula cells will be given a temporary blue glow in the spreadsheet — click Clear Highlights to remove."
                         Foreground="#6E6E6E" FontSize="11" TextWrapping="Wrap"/>
            </Grid>
          </Border>

        </Grid>
      </TabItem>

    </TabControl>

    <Border Grid.Row="2" Background="#252528" BorderBrush="#3E3E42" BorderThickness="0,1,0,0">
      <Grid Margin="12,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" x:Name="StatusBarTxt" Text="  Ready" Foreground="#B0B0B0"
                   FontSize="12" VerticalAlignment="Center"/>
        <TextBlock Grid.Column="1" x:Name="StatusClockTxt" Text="" Foreground="#909090"
                   FontSize="12" VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <Border Grid.RowSpan="3" x:Name="SearchOverlayBd" Visibility="Collapsed"
            Background="#E5000000">
      <Border Background="#252528" CornerRadius="8" BorderBrush="#4C9FE6" BorderThickness="2"
              Padding="20,16" HorizontalAlignment="Center" VerticalAlignment="Top"
              Margin="0,60,0,0" MinWidth="500" MaxWidth="700" MaxHeight="500">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Grid.Column="0" x:Name="SearchBox" FontSize="14"
                     ToolTip="Search across all tabs..."/>
            <Button Grid.Column="1" x:Name="SearchCloseBtn" Content="X" Style="{StaticResource BtnRed}"
                    Padding="8,4" Margin="8,0,0,0"/>
          </Grid>
          <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" MaxHeight="350">
            <StackPanel x:Name="SearchResultsPanel"/>
          </ScrollViewer>
          <TextBlock Grid.Row="3" Text="Ctrl+F to open  |  Esc to close" Foreground="#6E6E6E" FontSize="10"
                     HorizontalAlignment="Center" Margin="0,8,0,0"/>
        </Grid>
      </Border>
    </Border>

    <Border Grid.RowSpan="3" x:Name="BusyOverlayBd" Visibility="Collapsed"
            Background="#CC000000">
      <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
        <ProgressBar IsIndeterminate="True" Width="220" Height="4"
                     Background="#3E3E42" Foreground="#4C9FE6" Margin="0,0,0,10"/>
        <TextBlock x:Name="BusyTxt" Text="Working..." Foreground="#B0B0B0"
                   FontSize="13" HorizontalAlignment="Center"/>
      </StackPanel>
    </Border>

    <Canvas Grid.RowSpan="3" x:Name="KeyTipCanvas" IsHitTestVisible="False"
            Visibility="Collapsed"/>

  </Grid>
</Window>
"@

# ================================================================
#  BUILD WINDOW + BIND ELEMENTS
# ================================================================
# Parse the XAML string into a live WPF Window object.
# XamlReader.Load is the standard way to instantiate XAML from a string in PS 5.1.
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlStr))
$Window = [System.Windows.Markup.XamlReader]::Load($reader)
$script:Window = $Window

# Shortcut: FindName looks up a named XAML element; e.g. G 'ClipWbCombo' returns
# the ComboBox whose x:Name="ClipWbCombo" was declared in the XAML above.
function G([string]$n) { $Window.FindName($n) }

# -- Global UI refs --
$script:BusyOverlay = G 'BusyOverlayBd'
$script:BusyText    = G 'BusyTxt'
$script:StatusBar   = G 'StatusBarTxt'
$script:StatusClock  = G 'StatusClockTxt'
$ExcelStatus        = G 'ExcelStatus'
$QsStatTotal        = G 'QsStatTotal'
$QsStatPending      = G 'QsStatPending'
$MainTabs           = G 'MainTabs'
$KeyTipCanvas       = G 'KeyTipCanvas'
$HelpBtn            = G 'HelpBtn'

# -- Clipboard tab --
$ClipWbCombo      = G 'ClipWbCombo'
$ClipSheetCombo   = G 'ClipSheetCombo'
$ClipCellBox      = G 'ClipCellBox'
$ClipRefreshBtn   = G 'ClipRefreshBtn'
$ClipTimestampChk = G 'ClipTimestampChk'
$ClipDateCell     = G 'ClipDateCell'
$ClipTimeCell     = G 'ClipTimeCell'
$ClipSlotsPanel   = G 'ClipSlotsPanel'
$ClipSlotCount    = G 'ClipSlotCount'
$ClipAddSlotBtn   = G 'ClipAddSlotBtn'
$ClipRecordSeqBtn = G 'ClipRecordSeqBtn'
$ClipRecordSeqState = G 'ClipRecordSeqState'
$ClipLockBtn      = G 'ClipLockBtn'
$ClipClearDefaultsBtn = G 'ClipClearDefaultsBtn'
$ClipDefaultsIndicator = G 'ClipDefaultsIndicator'

# -- Macros tab --
$MacroList         = G 'MacroList'
$MacroWbCombo      = G 'MacroWbCombo'
$MacroSheetCombo   = G 'MacroSheetCombo'
$MacroSelectedTxt  = G 'MacroSelectedTxt'
$MacroDescTxt      = G 'MacroDescTxt'
$MacroRunBtn       = G 'MacroRunBtn'
$MacroOutputTxt    = G 'MacroOutputTxt'
$MacroRefreshBtn   = G 'MacroRefreshBtn'
$MacroFavBtn       = G 'MacroFavBtn'

# -- Scheduler tab --
$SchedNameBox       = G 'SchedNameBox'
$SchedMacroCombo    = G 'SchedMacroCombo'
$SchedTimeBox       = G 'SchedTimeBox'
$SchedFreqCombo     = G 'SchedFreqCombo'
$SchedCreateBtn     = G 'SchedCreateBtn'
$SchedRefreshBtn    = G 'SchedRefreshBtn'
$SchedTaskPanel     = G 'SchedTaskPanel'
$SchedFolderPath    = G 'SchedFolderPath'
$SchedOpenFolderBtn = G 'SchedOpenFolderBtn'

# Set the folder path display
$SchedFolderPath.Text = $script:MacroFolder
$SchedFolderPath.ToolTip = $script:MacroFolder

# -- Navigator tab --
$NavWbList         = G 'NavWbList'
$NavSheetList      = G 'NavSheetList'
$NavSheetCount     = G 'NavSheetCount'
$NavRefreshBtn     = G 'NavRefreshBtn'
$NavActivateBtn    = G 'NavActivateBtn'
$NavOpenBtn        = G 'NavOpenBtn'
$NavBringFrontBtn  = G 'NavBringFrontBtn'
$NavMinimizeBtn    = G 'NavMinimizeBtn'
$NavCloseWbBtn     = G 'NavCloseWbBtn'
$NavInfoTxt        = G 'NavInfoTxt'
$NavDestWbCombo    = G 'NavDestWbCombo'
$NavDestSheetCombo = G 'NavDestSheetCombo'
$NavCopyChk        = G 'NavCopyChk'
$NavMoveCopyBtn    = G 'NavMoveCopyBtn'
$NavExportSheetBtn = G 'NavExportSheetBtn'
$NavHideBtn        = G 'NavHideBtn'
$NavUnhideBtn      = G 'NavUnhideBtn'
$NavDeleteBtn      = G 'NavDeleteBtn'
$NavPwdBox         = G 'NavPwdBox'
$NavSetPwdBtn      = G 'NavSetPwdBtn'
$NavRemPwdBtn      = G 'NavRemPwdBtn'
$NavOpenPwdBtn     = G 'NavOpenPwdBtn'
$NavCalcModeCombo  = G 'NavCalcModeCombo'
$NavEventsChk      = G 'NavEventsChk'
$NavApplyExcelOptsBtn = G 'NavApplyExcelOptsBtn'
$NavVbaList        = G 'NavVbaList'

# -- Templates tab --
$TplList           = G 'TplList'
$TplDeleteBtn      = G 'TplDeleteBtn'
$TplNameBox        = G 'TplNameBox'
$TplContentBox     = G 'TplContentBox'
$TplSaveBtn        = G 'TplSaveBtn'
$TplPasteBtn       = G 'TplPasteBtn'
$TplPreviewBtn     = G 'TplPreviewBtn'
$TplPreviewCard    = G 'TplPreviewCard'
$TplPreviewTxt     = G 'TplPreviewTxt'

# -- Search overlay --
$SearchOverlayBd   = G 'SearchOverlayBd'
$SearchBox         = G 'SearchBox'
$SearchCloseBtn    = G 'SearchCloseBtn'
$SearchResultsPanel = G 'SearchResultsPanel'

# -- QSync tab --
$QsLastRoot        = G 'QsLastRoot'
$QsThisRoot        = G 'QsThisRoot'
$QsBrowseLast      = G 'QsBrowseLast'
$QsBrowseThis      = G 'QsBrowseThis'
$QsRunSyncBtn      = G 'QsRunSyncBtn'
$QsProgressCard    = G 'QsProgressCard'
$QsProgressTxt     = G 'QsProgressTxt'
$QsResultCard      = G 'QsResultCard'
$QsRCreated        = G 'QsRCreated'
$QsRSkipped        = G 'QsRSkipped'
$QsRNewTodos       = G 'QsRNewTodos'
$QsRErrors         = G 'QsRErrors'
$QsBigPct          = G 'QsBigPct'
$QsBigTotal        = G 'QsBigTotal'
$QsBigPending      = G 'QsBigPending'
$QsBigDone         = G 'QsBigDone'
$QsFolderBars      = G 'QsFolderBars'
$QsSyncLogTxt      = G 'QsSyncLogTxt'
$QsSyncQuarterName = G 'QsSyncQuarterName'

# -- QTasks tab --
$QsQuarterCombo    = G 'QsQuarterCombo'
$QsCompareCombo    = G 'QsCompareCombo'
$QsNewQuarterBtn   = G 'QsNewQuarterBtn'
$QsScanFolderBtn   = G 'QsScanFolderBtn'
$QsScanPathBox     = G 'QsScanPathBox'
$QsAddTaskBtn      = G 'QsAddTaskBtn'
$QsSwapFoldersBtn  = G 'QsSwapFoldersBtn'
$QsQuarterSummary  = G 'QsQuarterSummary'
$QsQuarterPct      = G 'QsQuarterPct'
$QsAddTaskCard     = G 'QsAddTaskCard'
$QsNewTaskName     = G 'QsNewTaskName'
$QsNewTaskFolder   = G 'QsNewTaskFolder'
$QsNewTaskDue      = G 'QsNewTaskDue'
$QsSaveTaskBtn     = G 'QsSaveTaskBtn'
$QsCancelTaskBtn   = G 'QsCancelTaskBtn'
$QsFltStatus       = G 'QsFltStatus'
$QsFltName         = G 'QsFltName'
$QsFltFolder       = G 'QsFltFolder'
$QsExportBtn       = G 'QsExportBtn'
$QsRefreshBtn      = G 'QsRefreshBtn'
$QsTodoPanel       = G 'QsTodoPanel'

# -- File Index tab --
$FidxRootBox       = G 'FidxRootBox'
$FidxBrowseBtn     = G 'FidxBrowseBtn'
$FidxIndexBtn      = G 'FidxIndexBtn'
$FidxCooldownTxt   = G 'FidxCooldownTxt'
$FidxSearchBox     = G 'FidxSearchBox'
$FidxMatchCount    = G 'FidxMatchCount'
$FidxGrid          = G 'FidxGrid'
$FidxStatusTxt     = G 'FidxStatusTxt'
$FidxTotalCount    = G 'FidxTotalCount'
$FidxLastScanTxt   = G 'FidxLastScanTxt'

# -- Formula Auditor tab --
$FaActiveCellTxt        = G 'FaActiveCellTxt'
$FaFormulaTxt           = G 'FaFormulaTxt'
$FaInspectBtn           = G 'FaInspectBtn'
$FaBackBtn              = G 'FaBackBtn'
$FaAuditRangeBtn        = G 'FaAuditRangeBtn'
$FaClearHighlightsBtn   = G 'FaClearHighlightsBtn'
$FaAutoBox              = G 'FaAutoBox'
$FaCellTypeTxt          = G 'FaCellTypeTxt'
$FaCellValueTxt         = G 'FaCellValueTxt'
$FaCellSheetTxt         = G 'FaCellSheetTxt'
$FaPrecedentsPanel      = G 'FaPrecedentsPanel'
$FaNoPrecedentsTxt      = G 'FaNoPrecedentsTxt'
$FaDependentCountTxt    = G 'FaDependentCountTxt'
$FaSumifTitleTxt        = G 'FaSumifTitleTxt'
$FaSumifTotalTxt        = G 'FaSumifTotalTxt'
$FaSumifEmptyTxt        = G 'FaSumifEmptyTxt'
$FaSumifGrid            = G 'FaSumifGrid'
$FaPrecedentsTree       = G 'FaPrecedentsTree'
$FaDependentsTree       = G 'FaDependentsTree'
$FaRangeAddrTxt         = G 'FaRangeAddrTxt'
$FaRangeFormulaTxt      = G 'FaRangeFormulaTxt'
$FaRangeHardcodedTxt    = G 'FaRangeHardcodedTxt'
$FaRangeExternalTxt     = G 'FaRangeExternalTxt'
$FaRangeDetailTxt       = G 'FaRangeDetailTxt'

# ================================================================
#  SYNC LOG WRITER (appends to the QSync log textbox)
# ================================================================
function QsLog([string]$msg) {
    $line = "[$(Get-Date -f 'HH:mm:ss')]  $msg"
    $QsSyncLogTxt.AppendText("$line`r`n")
    $QsSyncLogTxt.ScrollToEnd()
    Update-UI
}

# (Refresh-QuarterDropdown was removed — QsQuarterCombo is now an editable path combo
#  populated via Add-QsFolderHistoryItem and direct text entry, not JSON file names.)

# ================================================================
#  REFRESH FUNCTIONS: MacroHub tabs
# ================================================================
# Each Refresh-* function re-queries its data source and repopulates the
# corresponding UI control(s), preserving the previously selected item when
# possible.  They are called both on tab activation and by explicit refresh buttons.

# Re-queries open workbooks from Excel Main session and updates the Clipboard and
# Macros workbook combo boxes, then cascades to sheet list refreshes.
function Refresh-WorkbookDropdowns {
    $wbs = Get-OpenWorkbooks -Session Main
    foreach ($combo in @($ClipWbCombo, $MacroWbCombo)) {
        $prev = $combo.SelectedItem
        $combo.Items.Clear()
        foreach ($w in $wbs) { [void]($combo.Items.Add($w)) }
        if ($prev -and $combo.Items.Contains($prev)) { $combo.SelectedItem = $prev }
        elseif ($combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 }
    }
    Refresh-ClipSheets
    Refresh-MacroSheets
    # Update Excel status pill
    $xl = Get-ExcelApp -Session Main
    if ($xl) {
        $ExcelStatus.Text = "Excel Main: $($wbs.Count) open"
        $ExcelStatus.Foreground = HexBrush '#4C9FE6'
    } else {
        $ExcelStatus.Text = 'Excel Main: --'
        $ExcelStatus.Foreground = HexBrush '#6E6E6E'
    }
}

# Populates the hidden default sheet combo and every per-slot sheet combo from
# the currently selected workbook in the Clipboard tab.
function Refresh-ClipSheets {
    if (-not $ClipWbCombo) { return }
    $wb = $ClipWbCombo.SelectedItem

    # Hidden defaults sheet combo (backing values for new slots/default persistence).
    if ($ClipSheetCombo) {
        $prevSheet = $ClipSheetCombo.SelectedItem
        $ClipSheetCombo.Items.Clear()
        if ($wb) {
            $sheets = Get-WorksheetNames -WorkbookName $wb -Session Main
            foreach ($s in $sheets) { [void]($ClipSheetCombo.Items.Add($s)) }
            if ($prevSheet -and $ClipSheetCombo.Items.Contains($prevSheet)) {
                $ClipSheetCombo.SelectedItem = $prevSheet
            } elseif ($ClipSheetCombo.Items.Count -gt 0) {
                $ClipSheetCombo.SelectedIndex = 0
            }
        }
    }

    # Per-slot sheet selectors.
    foreach ($slotKey in @($script:ClipSlots.Keys)) {
        $slot = $script:ClipSlots[$slotKey]
        if (-not $slot -or -not $slot.SheetCombo) { continue }
        $combo = $slot.SheetCombo
        $prev = $combo.SelectedItem
        $combo.Items.Clear()
        if ($wb) {
            $sheets = Get-WorksheetNames -WorkbookName $wb -Session Main
            foreach ($s in $sheets) { [void]($combo.Items.Add($s)) }
            if ($prev -and $combo.Items.Contains($prev)) {
                $combo.SelectedItem = $prev
            } elseif ($ClipSheetCombo -and $ClipSheetCombo.SelectedItem -and $combo.Items.Contains($ClipSheetCombo.SelectedItem)) {
                $combo.SelectedItem = $ClipSheetCombo.SelectedItem
            } elseif ($combo.Items.Count -gt 0) {
                $combo.SelectedIndex = 0
            }
        }
    }
}

function Get-PrimaryClipSlotState {
    if (-not $ClipSlotsPanel) { return $null }
    foreach ($child in @($ClipSlotsPanel.Children)) {
        $tag = [string]$child.Tag
        if (-not [string]::IsNullOrWhiteSpace($tag) -and $script:ClipSlots.ContainsKey($tag)) {
            return $script:ClipSlots[$tag]
        }
    }
    return $null
}

function Get-ClipSlotStateByIndex([int]$Index) {
    if (-not $ClipSlotsPanel) { return $null }
    if ($Index -lt 0 -or $Index -ge $ClipSlotsPanel.Children.Count) { return $null }
    $child = $ClipSlotsPanel.Children[$Index]
    $tag = [string]$child.Tag
    if (-not $tag) { return $null }
    if (-not $script:ClipSlots.ContainsKey($tag)) { return $null }
    return $script:ClipSlots[$tag]
}

function Ensure-ClipSlotStateByIndex([int]$Index) {
    if ($Index -lt 0) { return $null }
    while ($ClipSlotsPanel.Children.Count -le $Index) {
        Add-ClipSlotUI -Panel $ClipSlotsPanel -WbCombo $ClipWbCombo `
            -DefaultSheet ([string]$ClipSheetCombo.SelectedItem) `
            -DefaultCell $ClipCellBox.Text `
            -DefaultTimestamp ([bool]$ClipTimestampChk.IsChecked) `
            -DefaultDateCell $ClipDateCell.Text `
            -DefaultTimeCell $ClipTimeCell.Text `
            -CountLabel $ClipSlotCount
        Refresh-ClipSheets
    }
    return (Get-ClipSlotStateByIndex $Index)
}

function Set-ClipSequenceUiState([bool]$On) {
    if (-not $ClipRecordSeqBtn) { return }
    if ($On) {
        $ClipRecordSeqBtn.Content = 'Stop _Sequence'
        $ClipRecordSeqBtn.Background = HexBrush '#E05050'
        $ClipRecordSeqBtn.Foreground = HexBrush '#FFFFFF'
        if ($ClipRecordSeqState) {
            $ClipRecordSeqState.Text = 'SEQ ON'
            $ClipRecordSeqState.Foreground = HexBrush '#E05050'
        }
    } else {
        $ClipRecordSeqBtn.Content = 'Record _Sequence'
        $ClipRecordSeqBtn.Background = HexBrush '#3E3E42'
        $ClipRecordSeqBtn.Foreground = HexBrush '#FFFFFF'
        if ($ClipRecordSeqState) {
            $ClipRecordSeqState.Text = 'SEQ OFF'
            $ClipRecordSeqState.Foreground = HexBrush '#6E6E6E'
        }
    }
}

function Stop-ClipSequenceCapture {
    $script:ClipSequenceEnabled = $false
    try { if ($script:ClipSequenceTimer) { $script:ClipSequenceTimer.Stop() } } catch {}
    Set-ClipSequenceUiState $false
}

function Refresh-MacroSheets {
    $MacroSheetCombo.Items.Clear()
    $wb = $MacroWbCombo.SelectedItem
    if (-not $wb) { return }
    $sheets = Get-WorksheetNames -WorkbookName $wb -Session Main
    foreach ($s in $sheets) { [void]($MacroSheetCombo.Items.Add($s)) }
    if ($MacroSheetCombo.Items.Count -gt 0) { $MacroSheetCombo.SelectedIndex = 0 }
}

# Scans the Macros/ folder and repopulates the MacroList, placing favorited macros
# at the top with a [*] prefix; also repopulates the Scheduler macro combo.
function Refresh-MacroList {
    $MacroList.Items.Clear()
    $SchedMacroCombo.Items.Clear()
    $macros = Get-MacroFiles
    $favNames = Get-FavoriteNames
    # Sort: favorites first (with star prefix), then rest alphabetically
    $favList  = @()
    $normList = @()
    foreach ($m in ($macros | Sort-Object Name)) {
        if ($favNames -contains $m.Name) {
            $favList += $m
        } else {
            $normList += $m
        }
    }
    foreach ($m in $favList) {
        [void]($MacroList.Items.Add("[*] $($m.Name)"))
    }
    foreach ($m in $normList) {
        [void]($MacroList.Items.Add("    $($m.Name)"))
    }
    foreach ($m in $macros) {
        [void]($SchedMacroCombo.Items.Add($m.FullName))
    }
    if ($SchedMacroCombo.Items.Count -gt 0) { $SchedMacroCombo.SelectedIndex = 0 }
    Set-Status "Found $($macros.Count) macro(s) -- $($favList.Count) favorite(s)"
}

function Refresh-TaskList {
    $SchedTaskPanel.Children.Clear()
    $tasks = Get-HubTasks
    foreach ($t in $tasks) {
        $taskName = $t.Name
        $row = [System.Windows.Controls.Border]::new()
        $row.Padding         = [System.Windows.Thickness]::new(14,8,14,8)
        $row.BorderBrush     = HexBrush '#3E3E42'
        $row.BorderThickness = [System.Windows.Thickness]::new(0,0,0,1)

        $g = [System.Windows.Controls.Grid]::new()
        foreach ($w in @('*','120','150','150','80')) {
            $cd = [System.Windows.Controls.ColumnDefinition]::new()
            if ($w -eq '*') { $cd.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star) }
            else { $cd.Width = [System.Windows.GridLength]::new([int]$w) }
            [void]($g.ColumnDefinitions.Add($cd))
        }

        $ci = 0
        $nextRun = try { $t.NextRunTime.ToString('g') } catch { 'N/A' }
        $lastRun = try { $t.LastRunTime.ToString('g') } catch { 'Never' }
        foreach ($val in @($t.Name, $t.State, $nextRun, $lastRun)) {
            $tb = [System.Windows.Controls.TextBlock]::new()
            $tb.Text = $val; $tb.FontSize = 12; $tb.VerticalAlignment = 'Center'
            $tb.Foreground = $(if ($ci -eq 0) { HexBrush '#FFFFFF' } else { HexBrush '#B0B0B0' })
            [System.Windows.Controls.Grid]::SetColumn($tb, $ci)
            [void]($g.Children.Add($tb))
            $ci++
        }

        $delBtn = [System.Windows.Controls.Button]::new()
        $delBtn.Content = 'Del'
        $delBtn.Style   = $Window.FindResource('BtnRed')
        $delBtn.Padding = [System.Windows.Thickness]::new(8,3,8,3)
        $delBtn.Add_Click({
            try {
                Remove-HubTask -TaskName $taskName
                Refresh-TaskList
                Set-Status "Removed task: $taskName"
            } catch { Set-Status "Error: $_" '#E05050' }
        }.GetNewClosure())
        [System.Windows.Controls.Grid]::SetColumn($delBtn, 4)
        [void]($g.Children.Add($delBtn))

        $row.Child = $g
        [void]($SchedTaskPanel.Children.Add($row))
    }
}

# Populates the Navigator workbook list from the Navigator Excel session and
# cascades to the destination workbook combo, Excel options, and VBA list.
function Refresh-NavWorkbooks {
    $NavWbList.Items.Clear()
    $wbs = Get-OpenWorkbooks -Session Navigator
    foreach ($w in $wbs) { [void]($NavWbList.Items.Add($w)) }
    # Also refresh destination workbook combo (reuse list)
    Refresh-NavDestWbCombo $wbs
    Refresh-NavExcelOptions
    Refresh-NavVbaList
}

# Enumerates all worksheets in the selected workbook and shows each in the
# Navigator sheet list with a colored dot indicating visible/hidden/very-hidden state.
function Refresh-NavSheets {
    $NavSheetList.Items.Clear()
    $wb = $NavWbList.SelectedItem
    if ($wb) {
        $xl = Get-ExcelApp -Session Navigator
        try {
            $wbObj = $xl.Workbooks.Item($wb)
            $visCount = 0; $hidCount = 0; $vhidCount = 0
            for ($i = 1; $i -le $wbObj.Worksheets.Count; $i++) {
                $ws = $wbObj.Worksheets.Item($i)
                $state = [int]$ws.Visible
                # Excel sheet Visible constants: -1=xlSheetVisible, 0=xlSheetHidden, 2=xlSheetVeryHidden
                $dotColor = switch ($state) {
                    -1 { '#50A050' }   # visible  = green dot
                    0  { '#FFFFFF' }   # hidden   = white dot
                    2  { '#000000' }   # very hidden = black dot (cannot be unhidden via UI in Excel)
                    default { '#6E6E6E' }
                }
                $suffix = switch ($state) {
                    -1 { '' }
                    0  { ' [Hidden]' }
                    2  { ' [VeryHidden]' }
                    default { '' }
                }
                if ($state -eq -1) { $visCount++ }
                elseif ($state -eq 2) { $vhidCount++ }
                else { $hidCount++ }

                $item = [System.Windows.Controls.ListBoxItem]::new()
                $item.Tag = [PSCustomObject]@{ Name = $ws.Name; Visible = $state }

                $sp = [System.Windows.Controls.StackPanel]::new()
                $sp.Orientation = 'Horizontal'
                $dot = [System.Windows.Controls.Border]::new()
                $dot.Width = 10
                $dot.Height = 10
                $dot.Margin = [System.Windows.Thickness]::new(0,0,6,0)
                $dot.CornerRadius = [System.Windows.CornerRadius]::new(5)
                $dot.Background = HexBrush $dotColor
                if ($state -eq 2) {
                    $dot.BorderBrush = HexBrush '#3E3E42'
                    $dot.BorderThickness = [System.Windows.Thickness]::new(1)
                }
                $txt = [System.Windows.Controls.TextBlock]::new()
                $txt.Text = "$($ws.Name)$suffix"
                $txt.Foreground = HexBrush '#FFFFFF'
                $txt.FontSize = 12
                [void]($sp.Children.Add($dot))
                [void]($sp.Children.Add($txt))
                $item.Content = $sp
                [void]($NavSheetList.Items.Add($item))
            }
            $NavSheetCount.Text = "$($visCount + $hidCount + $vhidCount) sheets ($hidCount hidden, $vhidCount very hidden)"
        } catch {
            $sheets = Get-WorksheetNames -WorkbookName $wb -Session Navigator
            foreach ($s in $sheets) { [void]($NavSheetList.Items.Add($s)) }
            $NavSheetCount.Text = "$($sheets.Count) sheets"
        }
    } else {
        $NavSheetCount.Text = '0 sheets'
    }
}

function Refresh-NavDestWbCombo {
    param($wbs)
    $NavDestWbCombo.Items.Clear()
    if (-not $wbs) { $wbs = Get-OpenWorkbooks -Session Navigator }
    foreach ($w in $wbs) { [void]($NavDestWbCombo.Items.Add($w)) }
    $NavDestSheetCombo.Items.Clear()
}

function Refresh-NavDestSheets {
    $NavDestSheetCombo.Items.Clear()
    $destWb = $NavDestWbCombo.SelectedItem
    if ($destWb) {
        [void]($NavDestSheetCombo.Items.Add('(Move to End)'))
        $sheets = Get-WorksheetNames -WorkbookName $destWb -Session Navigator
        foreach ($s in $sheets) { [void]($NavDestSheetCombo.Items.Add($s)) }
        if ($NavDestSheetCombo.Items.Count -gt 0) { $NavDestSheetCombo.SelectedIndex = 0 }
    }
}

function Refresh-NavVbaList {
    $NavVbaList.Items.Clear()
    $wb = $NavWbList.SelectedItem
    if (-not $wb) { return }
    try {
        $xl = Get-ExcelApp -Session Navigator
        $wbObj = $xl.Workbooks.Item($wb)
        $proj = $wbObj.VBProject
        $macroCount = 0
        for ($i = 1; $i -le $proj.VBComponents.Count; $i++) {
            $comp = $proj.VBComponents.Item($i)
            $modItem = [System.Windows.Controls.ListBoxItem]::new()
            $modItem.Content = "[Module] $($comp.Name)"
            $modItem.IsEnabled = $false
            [void]($NavVbaList.Items.Add($modItem))
            try {
                $cm = $comp.CodeModule
                $lineCount = [int]$cm.CountOfLines
                if ($lineCount -gt 0) {
                    $code = [string]$cm.Lines(1, $lineCount)
                    foreach ($m in [regex]::Matches($code, '(?im)^\s*(Public\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)')) {
                        $macroName = [string]$m.Groups[3].Value
                        $entry = "$($comp.Name).$macroName"
                        $macroItem = [System.Windows.Controls.ListBoxItem]::new()
                        $macroItem.Content = "  - $macroName"
                        $macroItem.Tag = [PSCustomObject]@{
                            Kind      = 'Macro'
                            Workbook  = [string]$wb
                            Module    = [string]$comp.Name
                            MacroName = $macroName
                            Entry     = $entry
                        }
                        [void]($NavVbaList.Items.Add($macroItem))
                        $macroCount++
                    }
                }
            } catch {}
        }
        if ($NavVbaList.Items.Count -eq 0) {
            [void]($NavVbaList.Items.Add('(No VBA modules found)'))
        } elseif ($macroCount -eq 0) {
            [void]($NavVbaList.Items.Add('(No runnable Sub/Function entries found)'))
        }
    } catch {
        [void]($NavVbaList.Items.Add('(Unable to read VBA project. Enable "Trust access to the VBA project object model".)'))
    }
}

function Refresh-NavExcelOptions {
    try {
        $xl = Get-ExcelApp -Session Navigator
        if (-not $xl) { return }
        $calcMode = [int]$xl.Calculation
        # Excel Calculation constants: -4105=xlAutomatic, -4135=xlManual, -4134=xlSemiautomatic
        $label = switch ($calcMode) {
            -4105 { 'Automatic' }
            -4135 { 'Manual' }
            -4134 { 'Semiautomatic' }
            default { 'Automatic' }
        }
        foreach ($item in $NavCalcModeCombo.Items) {
            if ($item.Content -eq $label) { $NavCalcModeCombo.SelectedItem = $item; break }
        }
        $NavEventsChk.IsChecked = [bool]$xl.EnableEvents
    } catch {}
}

function Refresh-TemplateList {
    $TplList.Items.Clear()
    $templates = Load-Templates
    foreach ($t in $templates) {
        [void]($TplList.Items.Add($t.Name))
    }
}


# ================================================================
#  REFRESH FUNCTIONS: QuarterSync tabs
# ================================================================
# Refresh-QsProgress updates the big stat cards and per-folder progress bars
# on the QSync tab using the current compare results from qs_compare_results.json.
# Refresh-QsTodoList rebuilds the scrollable QTasks list with optional filters.

# Recomputes completion percentages from qs_compare_results.json and updates the
# stat cards, percentage label colors, and folder progress bars on the QSync tab.
function Refresh-QsProgress {
    $todos   = Load-QsCompareTodos
    $total   = $todos.Count
    $pending = ($todos | Where-Object { $_.Status -eq 'Pending' }).Count
    $done    = $total - $pending
    $pct     = if ($total -gt 0) { [Math]::Round(($done / $total) * 100, 1) } else { 0 }

    $QsBigTotal.Text   = $total
    $QsBigPending.Text = $pending
    $QsBigDone.Text    = $done
    $QsBigPct.Text     = "$pct%"

    # Update QTasks quarter summary label
    try { $QsQuarterSummary.Text = "$total items  |  $pending pending  |  $done done" } catch {}
    try { $QsQuarterPct.Text = "$pct% complete" } catch {}

    # Color the percentage based on progress
    if ($pct -ge 80)     { $QsBigPct.Foreground = HexBrush '#50A050' }
    elseif ($pct -ge 40) { $QsBigPct.Foreground = HexBrush '#B0B0B0' }
    else                 { $QsBigPct.Foreground = HexBrush '#E05050' }

    # Header stat pills
    $QsStatTotal.Text   = $total
    $QsStatPending.Text = $pending

    # Folder progress bars
    $QsFolderBars.Children.Clear()
    $groups = $todos | Group-Object RelFolder | Sort-Object Name
    foreach ($grp in $groups) {
        $gt   = $grp.Count
        $gp   = ($grp.Group | Where-Object { $_.Status -eq 'Pending' }).Count
        $gd   = $gt - $gp
        $gpct = if ($gt -gt 0) { [Math]::Round(($gd / $gt) * 100) } else { 100 }
        $fname = if ($grp.Name -and $grp.Name -ne '') { $grp.Name } else { '(root)' }
        $barColor = if ($gpct -ge 80) { '#50A050' } elseif ($gpct -ge 40) { '#B0B0B0' } else { '#E05050' }

        $barRow = [System.Windows.Controls.Grid]::new()
        $barRow.Margin = [System.Windows.Thickness]::new(0,4,0,4)
        $c1 = [System.Windows.Controls.ColumnDefinition]::new()
        $c1.Width = [System.Windows.GridLength]::new(200)
        $c2 = [System.Windows.Controls.ColumnDefinition]::new()
        $c2.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
        $c3 = [System.Windows.Controls.ColumnDefinition]::new()
        $c3.Width = [System.Windows.GridLength]::new(50)
        [void]($barRow.ColumnDefinitions.Add($c1))
        [void]($barRow.ColumnDefinitions.Add($c2))
        [void]($barRow.ColumnDefinitions.Add($c3))

        $lbl = [System.Windows.Controls.TextBlock]::new()
        $lbl.Text = $fname
        $lbl.Foreground = HexBrush '#B0B0B0'
        $lbl.FontSize = 11
        $lbl.VerticalAlignment = 'Center'
        $lbl.TextTrimming = 'CharacterEllipsis'
        [System.Windows.Controls.Grid]::SetColumn($lbl, 0)
        [void]($barRow.Children.Add($lbl))

        $track = [System.Windows.Controls.Border]::new()
        $track.Background = HexBrush '#3E3E42'
        $track.CornerRadius = [System.Windows.CornerRadius]::new(3)
        $track.Height = 6
        $track.VerticalAlignment = 'Center'
        $track.Margin = [System.Windows.Thickness]::new(10,0,10,0)
        $fill = [System.Windows.Controls.Border]::new()
        $fill.Background = HexBrush $barColor
        $fill.CornerRadius = [System.Windows.CornerRadius]::new(3)
        $fill.HorizontalAlignment = 'Left'
        $pctVal = $gpct
        # Add_Loaded fires once after layout pass so ActualWidth is valid; setting
        # fill width in a constructor would give 0 because layout has not run yet.
        $track.Add_Loaded({
            $fill.Width = ($track.ActualWidth * $pctVal / 100)
        }.GetNewClosure())
        $track.Child = $fill
        [System.Windows.Controls.Grid]::SetColumn($track, 1)
        [void]($barRow.Children.Add($track))

        $pctTxt = [System.Windows.Controls.TextBlock]::new()
        $pctTxt.Text = "$gpct%"
        $pctTxt.Foreground = HexBrush $barColor
        $pctTxt.FontSize = 11
        $pctTxt.FontWeight = 'Bold'
        $pctTxt.VerticalAlignment = 'Center'
        $pctTxt.HorizontalAlignment = 'Right'
        [System.Windows.Controls.Grid]::SetColumn($pctTxt, 2)
        [void]($barRow.Children.Add($pctTxt))

        [void]($QsFolderBars.Children.Add($barRow))
    }
}

function New-QsTodoRow($todo, [string]$compareLabel) {
    $isPending = $todo.Status -eq 'Pending'
    $todoKey   = $todo.Key
    $isCompare = [bool]$compareLabel

    $row = [System.Windows.Controls.Border]::new()
    $row.Margin          = [System.Windows.Thickness]::new(4,2,4,2)
    $row.Padding         = [System.Windows.Thickness]::new(14,8,14,8)
    $row.CornerRadius    = [System.Windows.CornerRadius]::new(4)
    $row.BorderThickness = [System.Windows.Thickness]::new(1)
    if ($isCompare) {
        $row.Background  = HexBrush '#1B1F2A'
        $row.BorderBrush = HexBrush '#2A3E5E'
        $row.Opacity     = 0.7
    } elseif ($isPending) {
        $row.Background  = HexBrush '#1E1E22'
        $row.BorderBrush = HexBrush '#3E3E42'
    } else {
        $row.Background  = HexBrush '#1A221A'
        $row.BorderBrush = HexBrush '#2A3E2A'
        $row.Opacity     = 0.55
    }

    $g = [System.Windows.Controls.Grid]::new()
    foreach ($w in @('*','160','100','100','200')) {
        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        if ($w -eq '*') { $cd.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star) }
        else { $cd.Width = [System.Windows.GridLength]::new([int]$w) }
        [void]($g.ColumnDefinitions.Add($cd))
    }

    # Col 0: badge + filename
    $np = [System.Windows.Controls.StackPanel]::new()
    $np.Orientation = 'Horizontal'

    $badge = [System.Windows.Controls.Border]::new()
    $badge.CornerRadius = [System.Windows.CornerRadius]::new(3)
    $badge.Padding      = [System.Windows.Thickness]::new(5,2,5,2)
    $badge.Margin       = [System.Windows.Thickness]::new(0,0,8,0)
    $badge.VerticalAlignment = 'Center'
    $btxt = [System.Windows.Controls.TextBlock]::new()
    $btxt.FontSize = 9; $btxt.FontWeight = 'Bold'
    if ($isCompare) {
        $badge.Background = HexBrush '#1A2040'
        $btxt.Foreground  = HexBrush '#4C9FE6'
        $btxt.Text        = $compareLabel
    } elseif ($isPending) {
        $badge.Background = HexBrush '#2A2A1A'
        $btxt.Foreground  = HexBrush '#B0B0B0'
        $btxt.Text        = 'TODO'
    } else {
        $badge.Background = HexBrush '#1A2E1A'
        $btxt.Foreground  = HexBrush '#50A050'
        $btxt.Text        = 'DONE'
    }
    $badge.Child = $btxt
    [void]($np.Children.Add($badge))

    $ntxt = [System.Windows.Controls.TextBlock]::new()
    $ntxt.Text = $todo.OriginalName
    $ntxt.FontSize = 12
    $ntxt.VerticalAlignment = 'Center'
    $ntxt.TextTrimming = 'CharacterEllipsis'
    $ntxt.Foreground = HexBrush $(if ($isPending) { '#FFFFFF' } else { '#6E6E6E' })
    [void]($np.Children.Add($ntxt))
    [System.Windows.Controls.Grid]::SetColumn($np, 0)
    [void]($g.Children.Add($np))

    # Col 1: folder
    $ft = [System.Windows.Controls.TextBlock]::new()
    $ft.Text = $(if ($todo.RelFolder) { $todo.RelFolder } else { '(root)' })
    $ft.FontSize = 11; $ft.Foreground = HexBrush '#6E6E6E'
    $ft.VerticalAlignment = 'Center'; $ft.TextTrimming = 'CharacterEllipsis'
    $ft.Margin = [System.Windows.Thickness]::new(8,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($ft, 1)
    [void]($g.Children.Add($ft))

    # Col 2: last done
    $lt = [System.Windows.Controls.TextBlock]::new()
    $lt.Text = $todo.LastDoneDate; $lt.FontSize = 11; $lt.Foreground = HexBrush '#6E6E6E'
    $lt.VerticalAlignment = 'Center'
    $lt.Margin = [System.Windows.Thickness]::new(8,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($lt, 2)
    [void]($g.Children.Add($lt))

    # Col 3: updated in source A
    $db = [System.Windows.Controls.TextBlock]::new()
    $db.Text = $todo.DueDate
    $db.FontSize = 11
    $db.Foreground = HexBrush '#6E6E6E'
    $db.VerticalAlignment = 'Center'
    $db.Margin = [System.Windows.Thickness]::new(8,0,8,0)
    [System.Windows.Controls.Grid]::SetColumn($db, 3)
    [void]($g.Children.Add($db))

    # Col 4: action buttons
    $ap = [System.Windows.Controls.StackPanel]::new()
    $ap.Orientation = 'Horizontal'
    $ap.VerticalAlignment = 'Center'
    $ap.Margin = [System.Windows.Thickness]::new(8,0,0,0)

    if ($isPending) {
        $bDone = [System.Windows.Controls.Button]::new()
        $bDone.Content = 'Done'
        $bDone.Style   = $Window.FindResource('BtnGreen')
        $bDone.Padding = [System.Windows.Thickness]::new(10,4,10,4)
        $bDone.Margin  = [System.Windows.Thickness]::new(0,0,6,0)
        $bDone.Add_Click({
            $todos = Load-QsCompareTodos
            $t = $todos | Where-Object { $_.Key -eq $todoKey }
            if ($t) { $t.Status = 'Done'; Save-QsCompareTodos $todos }
            Refresh-QsTodoList; Refresh-QsProgress
        }.GetNewClosure())
        [void]($ap.Children.Add($bDone))
    }

    $bClr = [System.Windows.Controls.Button]::new()
    $bClr.Content = $(if ($isPending) { 'Clear' } else { 'Restore' })
    $bClr.Style   = $(if ($isPending) { $Window.FindResource('BtnRed') } else { $Window.FindResource('Btn') })
    $bClr.Padding = [System.Windows.Thickness]::new(10,4,10,4)
    $bClr.Add_Click({
        $todos = Load-QsCompareTodos
        $t = $todos | Where-Object { $_.Key -eq $todoKey }
        if ($t) {
            $t.Status = if ($t.Status -eq 'Pending') { 'Cleared' } else { 'Pending' }
            Save-QsCompareTodos $todos
        }
        Refresh-QsTodoList; Refresh-QsProgress
    }.GetNewClosure())
    [void]($ap.Children.Add($bClr))

    [System.Windows.Controls.Grid]::SetColumn($ap, 4)
    [void]($g.Children.Add($ap))
    $row.Child = $g
    return $row
}

# Clears and rebuilds the QTasks scrollable panel, applying the current status /
# name / folder filters and sorting pending items before completed ones.
function Refresh-QsTodoList {
    $QsTodoPanel.Children.Clear()
    $todos = Load-QsCompareTodos

    $sf = ($QsFltStatus.SelectedItem.Content).ToString()
    $nf = $QsFltName.Text.Trim().ToLower()
    $ff = $QsFltFolder.Text.Trim().ToLower()

    # Filter helper
    $filterBlock = {
        param($item, $statusFilter, $nameFilter, $folderFilter)
        $ms = switch ($statusFilter) {
            'Pending' { $item.Status -eq 'Pending' }
            'Done'    { $item.Status -ne 'Pending' }
            default   { $true }
        }
        $mn = (-not $nameFilter) -or ($item.OriginalName.ToLower() -like "*$nameFilter*")
        $mf = (-not $folderFilter) -or ($item.RelFolder.ToLower()    -like "*$folderFilter*")
        return ($ms -and $mn -and $mf)
    }

    $filtered = @($todos | Where-Object { & $filterBlock $_ $sf $nf $ff })

    if ($filtered.Count -eq 0) {
        $e = [System.Windows.Controls.TextBlock]::new()
        $src = Resolve-QsFolderPath (Get-ComboText $QsQuarterCombo)
        $tgt = Resolve-QsFolderPath (Get-ComboText $QsCompareCombo)
        if ($src -and $tgt) {
            $e.Text = 'No missing files for this comparison (or none match the filters).'
        } else {
            $e.Text = 'Pick Source A and Target B folders, then click Compare Missing.'
        }
        $e.Foreground = HexBrush '#6E6E6E'
        $e.FontSize = 13
        $e.HorizontalAlignment = 'Center'
        $e.Margin = [System.Windows.Thickness]::new(0,30,0,0)
        [void]($QsTodoPanel.Children.Add($e))
        if ($QsQuarterSummary) { $QsQuarterSummary.Text = '' }
        if ($QsQuarterPct) { $QsQuarterPct.Text = '' }
        return
    }

    $sorted = @($filtered | Where-Object { $_.Status -eq 'Pending' } | Sort-Object LastDoneDate) +
              @($filtered | Where-Object { $_.Status -ne 'Pending' } | Sort-Object LastDoneDate)

    foreach ($t in $sorted) {
        [void]($QsTodoPanel.Children.Add((New-QsTodoRow $t)))
    }

    $srcPath = Resolve-QsFolderPath (Get-ComboText $QsQuarterCombo)
    $tgtPath = Resolve-QsFolderPath (Get-ComboText $QsCompareCombo)
    $srcName = if ($srcPath) { Split-Path $srcPath -Leaf } else { '' }
    $tgtName = if ($tgtPath) { Split-Path $tgtPath -Leaf } else { '' }
    if (-not $srcName) { $srcName = 'A' }
    if (-not $tgtName) { $tgtName = 'B' }
    if ($QsQuarterSummary) { $QsQuarterSummary.Text = "$($filtered.Count) missing in $tgtName from $srcName" }
}

function Add-QsFolderHistoryItem($Combo, [string]$PathText) {
    if (-not $Combo) { return }
    $p = Resolve-QsFolderPath $PathText
    if (-not $p) { return }
    $items = @()
    foreach ($it in @($Combo.Items)) { $items += [string]$it }
    if ($items -contains $p) { $Combo.Items.Remove($p) }
    $Combo.Items.Insert(0, $p)
    $Combo.Text = $p
    $Combo.SelectedItem = $p
}

function Save-QsCompareConfigFromUi {
    $cfg = Load-QsConfig
    $cfg.CompareSourcePath = Resolve-QsFolderPath (Get-ComboText $QsQuarterCombo)
    $cfg.CompareTargetPath = Resolve-QsFolderPath (Get-ComboText $QsCompareCombo)
    Save-QsConfig $cfg
}

function Invoke-QsCompareFromUi {
    $src = Resolve-QsFolderPath (Get-ComboText $QsQuarterCombo)
    $tgt = Resolve-QsFolderPath (Get-ComboText $QsCompareCombo)
    if (-not $src -or -not $tgt) {
        [System.Windows.MessageBox]::Show('Pick both Source A and Target B folders first.','MacroHub','OK','Warning') | Out-Null
        return
    }
    if (-not (Test-Path $src)) { Set-Status "Source folder not found: $src" '#E05050'; return }
    if (-not (Test-Path $tgt)) { Set-Status "Target folder not found: $tgt" '#E05050'; return }

    Show-Busy 'Comparing folders...'
    try {
        $res = Compare-QsFolders -SourceRoot $src -TargetRoot $tgt
        $tag = $QsScanPathBox.Text.Trim()
        $rows = @($res.MissingTodos)
        if ($tag) {
            foreach ($r in $rows) {
                if ($r.Note) { $r.Note = "$($r.Note) | $tag" } else { $r.Note = $tag }
            }
        }
        Save-QsCompareTodos $rows
        Add-QsFolderHistoryItem -Combo $QsQuarterCombo -PathText $src
        Add-QsFolderHistoryItem -Combo $QsCompareCombo -PathText $tgt
        Save-QsCompareConfigFromUi
        Refresh-QsTodoList
        Refresh-QsProgress
        Set-Status "Compare complete: $($rows.Count) file(s) in A missing from B."
    } catch {
        Set-Status "Compare error: $($_.Exception.Message)" '#E05050'
    } finally {
        Hide-Busy
    }
}

Set-ClipSequenceUiState $false

# ================================================================
#  EVENT HANDLERS
# ================================================================
# All Add_Click / Add_SelectionChanged handlers are wired here, after the
# Refresh-* functions are defined so closures can reference them freely.

# -- Clipboard tab --
# Handlers for workbook selection, slot add/remove, record/stop, sequence
# capture, paste, and defaults lock/clear on the Clipboard tab.
$ClipWbCombo.Add_SelectionChanged({ Refresh-ClipSheets })

$ClipRefreshBtn.Add_Click({
    Show-Busy 'Refreshing Excel...'
    try {
        Refresh-WorkbookDropdowns
        Refresh-ClipSheets
        Set-Status 'Refreshed workbook list'
    } catch { Set-Status "Error: $_" '#E05050' }
    finally { Hide-Busy }
})

# Dynamic Add Slot button
$ClipAddSlotBtn.Add_Click({
    Add-ClipSlotUI -Panel $ClipSlotsPanel -WbCombo $ClipWbCombo `
        -DefaultSheet ([string]$ClipSheetCombo.SelectedItem) `
        -DefaultCell $ClipCellBox.Text `
        -DefaultTimestamp ([bool]$ClipTimestampChk.IsChecked) `
        -DefaultDateCell $ClipDateCell.Text `
        -DefaultTimeCell $ClipTimeCell.Text `
        -CountLabel $ClipSlotCount
    Set-Status "Added slot $($script:ClipSlotIdx)"
})

# Sequence capture: each new clipboard copy is recorded into the next slot.
if (-not $script:ClipSequenceTimer) {
    $script:ClipSequenceTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:ClipSequenceTimer.Interval = [TimeSpan]::FromMilliseconds(280)
    $script:ClipSequenceTimer.Add_Tick({
        if (-not $script:ClipSequenceEnabled) { return }
        try {
            $packet = Get-ClipboardPacket
            if (-not $packet) { return }
            $sig = [string]$packet.Signature
            if (-not $sig) { $sig = Get-TextSha1 ([string]$packet.Text + '|' + [string]$packet.Formats) }
            if (-not $sig -or $sig -eq [string]$script:ClipSequenceLastSig) { return }

            $slotIdx = [Math]::Max(0, [int]$script:ClipSequenceNextIndex)
            $slot = Ensure-ClipSlotStateByIndex $slotIdx
            if (-not $slot) { return }

            $slot.Packet = $packet
            $slot.Text = [string]$packet.Text
            $slot.LastSig = $sig
            if ($slot.TextBoxCtrl) { $slot.TextBoxCtrl.Text = $slot.Text }

            $script:ClipSequenceLastSig = $sig
            $script:ClipSequenceNextIndex = $slotIdx + 1
            Set-Status "Sequence captured clipboard into slot $($slotIdx + 1)"
        } catch {
            Stop-ClipSequenceCapture
            Set-Status "Sequence capture error: $($_.Exception.Message)" '#E05050'
        }
    })
}

$ClipRecordSeqBtn.Add_Click({
    if ($script:ClipSequenceEnabled) {
        Stop-ClipSequenceCapture
        Set-Status 'Stopped sequence capture'
        return
    }
    foreach ($k in @($script:ClipSlots.Keys)) {
        $s = $script:ClipSlots[$k]
        if ($s -and $s.IsRecording) {
            Set-Status 'Stop per-slot recording before starting sequence capture.' '#E05050'
            return
        }
    }
    $next = $ClipSlotsPanel.Children.Count
    for ($idx = 0; $idx -lt $ClipSlotsPanel.Children.Count; $idx++) {
        $st = Get-ClipSlotStateByIndex $idx
        if (-not $st -or [string]::IsNullOrWhiteSpace([string]$st.Text)) {
            $next = $idx
            break
        }
    }
    $script:ClipSequenceEnabled = $true
    $script:ClipSequenceLastSig = ''
    $script:ClipSequenceNextIndex = $next
    Set-ClipSequenceUiState $true
    $script:ClipSequenceTimer.Start()
    Set-Status "Sequence capture ON (next target slot: $($next + 1))"
})

# Lock-in / Clear defaults
$ClipLockBtn.Add_Click({
    $wb = $ClipWbCombo.SelectedItem
    $slot = Get-PrimaryClipSlotState
    $sh = if ($slot -and $slot.SheetCombo) { $slot.SheetCombo.SelectedItem } else { $ClipSheetCombo.SelectedItem }
    $cell = if ($slot -and $slot.CellBox) { [string]$slot.CellBox.Text } else { [string]$ClipCellBox.Text }
    $cell = $cell.Trim()
    if (-not $cell) { $cell = 'A1' }
    $ts   = if ($slot -and $slot.TimestampChk) { [bool]$slot.TimestampChk.IsChecked } else { [bool]$ClipTimestampChk.IsChecked }
    $dCell = if ($slot -and $slot.DateCellBox) { [string]$slot.DateCellBox.Text } else { [string]$ClipDateCell.Text }
    $tCell = if ($slot -and $slot.TimeCellBox) { [string]$slot.TimeCellBox.Text } else { [string]$ClipTimeCell.Text }

    # Keep hidden defaults controls in sync (used when creating future slots).
    if ($sh -and $ClipSheetCombo -and $ClipSheetCombo.Items.Contains($sh)) { $ClipSheetCombo.SelectedItem = $sh }
    $ClipCellBox.Text = $cell
    $ClipTimestampChk.IsChecked = $ts
    $ClipDateCell.Text = $dCell
    $ClipTimeCell.Text = $tCell

    Save-ClipDefaults -Workbook ([string]$wb) -Sheet ([string]$sh) -Cell $cell `
        -Timestamp $ts -DateCell $dCell -TimeCell $tCell
    $ClipDefaultsIndicator.Text = "(defaults: $wb > $sh > $cell)"
    $ClipDefaultsIndicator.Visibility = 'Visible'
    Set-Status "Defaults locked (workbook + slot settings): $wb > $sh > $cell | Timestamp=$ts | Date→$dCell Time→$tCell"
})

$ClipClearDefaultsBtn.Add_Click({
    if (Test-Path $script:ClipDefaultsJson) { Remove-Item $script:ClipDefaultsJson -Force }
    $ClipDefaultsIndicator.Visibility = 'Collapsed'
    Set-Status 'Defaults cleared -- will use last-selected settings on next launch'
})

$HelpBtn.Add_Click({
    $guide = @"
MacroHub Quick Guide

0 Clipboard: Record per-slot, or use Record Sequence to capture multiple clipboard copies into slot1/slot2/slot3 before pasting.
1 Macros: Pick macro + workbook + sheet and run.
2 Scheduler: Schedule .ps1/.bas tasks in Windows Task Scheduler.
3 Navigator: Manage workbooks/sheets in a separate Excel Navigator session (Ctrl+E renames selected sheet).
4 Templates: Save reusable text snippets and paste into Excel.
5 QSync: Build quarter checklist from folder comparisons.
6 QTasks: Work a running task list, track completion %, export to Excel.

Typical flow:
Refresh workbook/email data -> filter/select target -> run action -> verify status.
"@
    [System.Windows.MessageBox]::Show($guide, 'MacroHub Guide', 'OK', 'Information') | Out-Null
})

# -- Macros tab --
# Handlers for macro list selection, run, favorite toggle, and workbook/sheet refresh.
$MacroRefreshBtn.Add_Click({
    Refresh-WorkbookDropdowns
    Refresh-MacroList
})

$MacroWbCombo.Add_SelectionChanged({
    Refresh-MacroSheets
})

$MacroList.Add_SelectionChanged({
    $sel = $MacroList.SelectedItem
    if ($sel) {
        $cleanName = Get-CleanMacroName $sel
        $MacroSelectedTxt.Text = $cleanName
        $file = Get-MacroFiles | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
        if ($file) {
            $ext = $file.Extension.ToLower()
            $desc = if ($ext -eq '.bas') { 'VBA macro - will be imported and run in Excel' }
                    elseif ($ext -eq '.ps1') { 'PowerShell script - runs via command line' }
                    else { 'Unknown type' }
            $isFav = if ($sel -match '^\[\*\]') { ' [FAVORITE]' } else { '' }
            $MacroDescTxt.Text = "$desc$isFav`r`nPath: $($file.FullName)"
        }
    }
})

$MacroRunBtn.Add_Click({
    $sel = $MacroList.SelectedItem
    if (-not $sel) { Set-Status 'Select a macro first' '#E05050'; return }
    $cleanName = Get-CleanMacroName $sel
    $file = Get-MacroFiles | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
    if (-not $file) { Set-Status 'Macro file not found' '#E05050'; return }
    $wb = $MacroWbCombo.SelectedItem
    $sh = $MacroSheetCombo.SelectedItem

    Show-Busy "Running $cleanName..."
    $MacroOutputTxt.Text = "Running $cleanName...`r`n"
    Update-UI
    try {
        $output = Invoke-SelectedMacro -MacroFile $file.FullName -WorkbookName $wb -WorksheetName $sh 2>&1 | Out-String
        $MacroOutputTxt.Text += $output
        $MacroOutputTxt.Text += "`r`nDone."
        Set-Status "Macro completed: $cleanName"
    } catch {
        $MacroOutputTxt.Text += "`r`nERROR: $_"
        Set-Status "Macro failed: $_" '#E05050'
    } finally { Hide-Busy }
})

# Macro Favorite toggle
$MacroFavBtn.Add_Click({
    $sel = $MacroList.SelectedItem
    if (-not $sel) { Set-Status 'Select a macro first' '#E05050'; return }
    $cleanName = Get-CleanMacroName $sel
    Toggle-Favorite $cleanName
    Refresh-MacroList
    Set-Status "Toggled favorite: $cleanName"
})

# -- Scheduler tab --
# Handlers for opening the Macros folder, refreshing the task list, and creating
# new scheduled tasks via Register-HubTask.
$SchedOpenFolderBtn.Add_Click({
    # Create the folder if it doesn't exist, then open it
    if (-not (Test-Path $script:MacroFolder)) {
        [void](New-Item -ItemType Directory -Path $script:MacroFolder -Force)
    }
    Start-Process explorer.exe $script:MacroFolder
    Set-Status "Opened: $($script:MacroFolder)"
})

$SchedRefreshBtn.Add_Click({
    Refresh-MacroList   # re-scans the folder for new scripts
    Refresh-TaskList
})

$SchedCreateBtn.Add_Click({
    $name = $SchedNameBox.Text.Trim()
    $file = $SchedMacroCombo.SelectedItem
    $time = $SchedTimeBox.Text.Trim()
    $freq = ($SchedFreqCombo.SelectedItem).Content.ToString()

    if (-not $name) { Set-Status 'Enter a task name' '#E05050'; return }
    if (-not $file) { Set-Status 'Select a macro file' '#E05050'; return }
    if ($time -notmatch '^\d{2}:\d{2}$') { Set-Status 'Time must be HH:mm' '#E05050'; return }

    Show-Busy 'Creating task...'
    try {
        Register-HubTask -TaskName $name -MacroFile $file -TriggerTime $time -Frequency $freq
        Set-Status "Created task: $name ($freq at $time)"
        $SchedNameBox.Text = ''
        Refresh-TaskList
    } catch { Set-Status "Error: $_" '#E05050' }
    finally { Hide-Busy }
})

# -- Navigator tab --
# Handlers for workbook/sheet list refresh, drag-drop sheet reorder, sheet
# activate/open/close/hide/unhide/delete/rename, move-copy, export, password
# management, Excel window controls, and VBA macro execution.
$NavRefreshBtn.Add_Click({
    Refresh-NavWorkbooks
    Refresh-WorkbookDropdowns
})

$NavWbList.Add_SelectionChanged({
    Refresh-NavSheets
    Refresh-NavVbaList
    Refresh-NavExcelOptions
})

$NavDestWbCombo.Add_SelectionChanged({ Refresh-NavDestSheets })

# Navigator sheet reordering via drag/drop.
$NavSheetList.Add_PreviewMouseLeftButtonDown({
    $script:NavSheetDragStart = $_.GetPosition($NavSheetList)
    $script:NavSheetDragSourceItem = Get-NavSheetListItemFromElement $_.OriginalSource
})

$NavSheetList.Add_PreviewMouseMove({
    if ($_.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }
    if (-not $script:NavSheetDragStart -or -not $script:NavSheetDragSourceItem) { return }
    $pt = $_.GetPosition($NavSheetList)
    $dx = [Math]::Abs($pt.X - $script:NavSheetDragStart.X)
    $dy = [Math]::Abs($pt.Y - $script:NavSheetDragStart.Y)
    if ($dx -lt [SystemParameters]::MinimumHorizontalDragDistance -and $dy -lt [SystemParameters]::MinimumVerticalDragDistance) { return }

    $srcItem = $script:NavSheetDragSourceItem
    $srcName = Get-CleanSheetName $srcItem
    $srcWb = [string]$NavWbList.SelectedItem
    if (-not $srcName -or -not $srcWb) { return }

    $data = [System.Windows.DataObject]::new()
    [void]$data.SetData('MacroHub.NavSheetName', $srcName)
    [void]$data.SetData('MacroHub.NavSheetIndex', [int]$NavSheetList.Items.IndexOf($srcItem))
    [void]$data.SetData('MacroHub.NavWorkbook', $srcWb)
    [void][System.Windows.DragDrop]::DoDragDrop($srcItem, $data, [System.Windows.DragDropEffects]::Move)

    $script:NavSheetDragStart = $null
    $script:NavSheetDragSourceItem = $null
})

$NavSheetList.Add_DragOver({
    if ($_.Data.GetDataPresent('MacroHub.NavSheetName')) {
        $_.Effects = [System.Windows.DragDropEffects]::Move
        $_.Handled = $true
    }
})

$NavSheetList.Add_Drop({
    try {
        if (-not $_.Data.GetDataPresent('MacroHub.NavSheetName')) { return }
        $wbCurrent = [string]$NavWbList.SelectedItem
        $wbFromData = [string]$_.Data.GetData('MacroHub.NavWorkbook')
        if (-not $wbCurrent) { return }
        if ($wbFromData -and $wbFromData -ne $wbCurrent) {
            Set-Status 'Sheet reorder only supports the currently selected workbook.' '#E05050'
            return
        }

        $srcName = [string]$_.Data.GetData('MacroHub.NavSheetName')
        $srcIndex = -1
        try { $srcIndex = [int]$_.Data.GetData('MacroHub.NavSheetIndex') } catch {}
        if (-not $srcName) { return }

        $targetItem = Get-NavSheetListItemFromElement $_.OriginalSource
        $moveToEnd = $false
        $targetName = ''
        $placeAfter = $false
        if ($targetItem) {
            $targetName = Get-CleanSheetName $targetItem
            if (-not $targetName -or $targetName -eq $srcName) { return }
            $targetIndex = [int]$NavSheetList.Items.IndexOf($targetItem)
            if ($srcIndex -ge 0 -and $targetIndex -ge 0 -and $srcIndex -lt $targetIndex) {
                $placeAfter = $true
            }
        } else {
            $moveToEnd = $true
        }

        Show-Busy 'Reordering sheets...'
        try {
            $ok = Move-NavSheetWithinWorkbook -WorkbookName $wbCurrent -SourceSheetName $srcName `
                -TargetSheetName $targetName -PlaceAfter $placeAfter -MoveToEnd $moveToEnd
            if ($ok) {
                Refresh-NavSheets
                foreach ($it in @($NavSheetList.Items)) {
                    if ((Get-CleanSheetName $it) -eq $srcName) {
                        $NavSheetList.SelectedItem = $it
                        try { $NavSheetList.ScrollIntoView($it) } catch {}
                        break
                    }
                }
                if ($moveToEnd) { Set-Status "Moved sheet to end: $srcName" }
                elseif ($placeAfter) { Set-Status "Moved sheet after ${targetName}: $srcName" }
                else { Set-Status "Moved sheet before ${targetName}: $srcName" }
            }
        } catch {
            Set-Status "Sheet reorder error: $($_.Exception.Message)" '#E05050'
        } finally { Hide-Busy }
        $_.Handled = $true
    } finally {
        $script:NavSheetDragStart = $null
        $script:NavSheetDragSourceItem = $null
    }
})

# Navigator VBA list: right-click or double-click macro -> run it.
$navVbaCtx = [System.Windows.Controls.ContextMenu]::new()
$navVbaRunMi = [System.Windows.Controls.MenuItem]::new()
$navVbaRunMi.Header = '_Run Selected Macro'
$navVbaRunMi.Add_Click({ Invoke-NavSelectedMacro }.GetNewClosure())
[void]($navVbaCtx.Items.Add($navVbaRunMi))
$NavVbaList.ContextMenu = $navVbaCtx

$NavVbaList.Add_ContextMenuOpening({
    $ctx = Get-NavSelectedMacroContext
    $navVbaRunMi.IsEnabled = [bool]$ctx
}.GetNewClosure())

$NavVbaList.Add_PreviewMouseRightButtonDown({
    $dep = $_.OriginalSource
    while ($dep -and -not ($dep -is [System.Windows.Controls.ListBoxItem])) {
        try { $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep) } catch { $dep = $null }
    }
    if ($dep -is [System.Windows.Controls.ListBoxItem]) {
        $dep.IsSelected = $true
        $NavVbaList.Focus() | Out-Null
    }
}.GetNewClosure())

$NavVbaList.Add_MouseDoubleClick({
    if (Get-NavSelectedMacroContext) {
        Invoke-NavSelectedMacro
    }
})

# Helper: support both legacy text rows and ListBoxItem sheet rows.
function Get-CleanSheetName([object]$raw) {
    if ($raw -is [System.Windows.Controls.ListBoxItem] -and $raw.Tag -and $raw.Tag.Name) {
        return [string]$raw.Tag.Name
    }
    return ([string]$raw -replace '\s*\[(Hidden|VeryHidden)\]$','')
}

function Get-NavSheetListItemFromElement([object]$depObj) {
    $dep = $depObj
    while ($dep -and -not ($dep -is [System.Windows.Controls.ListBoxItem])) {
        try { $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep) } catch { $dep = $null }
    }
    if ($dep -is [System.Windows.Controls.ListBoxItem]) { return $dep }
    return $null
}

function Move-NavSheetWithinWorkbook {
    param(
        [string]$WorkbookName,
        [string]$SourceSheetName,
        [string]$TargetSheetName,
        [bool]$PlaceAfter = $false,
        [bool]$MoveToEnd = $false
    )
    if (-not $WorkbookName -or -not $SourceSheetName) { return $false }
    $xl = Get-ExcelApp -Session Navigator
    if (-not $xl) { throw 'Navigator Excel session is not open.' }
    $wbObj = $xl.Workbooks.Item($WorkbookName)
    $srcWs = $wbObj.Worksheets.Item($SourceSheetName)

    if ($MoveToEnd) {
        $last = $wbObj.Worksheets.Item($wbObj.Worksheets.Count)
        if ([string]$last.Name -eq [string]$srcWs.Name) { return $true }
        $srcWs.Move([System.Type]::Missing, $last)
        return $true
    }

    if (-not $TargetSheetName) { return $false }
    if ([string]$TargetSheetName -eq [string]$SourceSheetName) { return $false }

    $targetWs = $wbObj.Worksheets.Item($TargetSheetName)
    if ($PlaceAfter) {
        $srcWs.Move([System.Type]::Missing, $targetWs)
    } else {
        $srcWs.Move($targetWs, [System.Type]::Missing)
    }
    return $true
}

function Invoke-NavSheetRename {
    param(
        [string]$NewName
    )
    $wbName = [string]$NavWbList.SelectedItem
    $selItem = $NavSheetList.SelectedItem
    if (-not $wbName -or -not $selItem) {
        Set-Status 'Select a workbook and one sheet to rename.' '#E05050'
        return $false
    }

    $oldName = Get-CleanSheetName $selItem
    if (-not $oldName) {
        Set-Status 'Could not resolve selected sheet name.' '#E05050'
        return $false
    }

    $newRaw = $NewName
    if (-not $PSBoundParameters.ContainsKey('NewName')) {
        try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue } catch {}
        $newRaw = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Rename selected sheet in '$wbName':",
            'Rename Sheet',
            $oldName
        )
    }

    if ($null -eq $newRaw) { return $false }
    $newName = [string]$newRaw
    $newName = $newName.Trim()
    if (-not $newName) { return $false }
    if ($newName -eq $oldName) {
        Set-Status "Sheet name unchanged: $oldName"
        return $true
    }
    if ($newName.Length -gt 31) {
        Set-Status 'Sheet name must be 31 characters or fewer.' '#E05050'
        return $false
    }
    if ($newName -match '[:\\/\?\*\[\]]') {
        Set-Status 'Sheet name contains invalid characters: : \ / ? * [ ]' '#E05050'
        return $false
    }
    if ($newName.StartsWith("'") -or $newName.EndsWith("'")) {
        Set-Status "Sheet name cannot start or end with apostrophe (')." '#E05050'
        return $false
    }

    try {
        $xl = Get-ExcelApp -Session Navigator
        if (-not $xl) { throw 'Navigator Excel session is not open.' }
        $wbObj = $xl.Workbooks.Item($wbName)
        $ws = $wbObj.Worksheets.Item($oldName)

        for ($i = 1; $i -le $wbObj.Worksheets.Count; $i++) {
            $other = $wbObj.Worksheets.Item($i)
            if ([int]$other.Index -eq [int]$ws.Index) { continue }
            if ([string]$other.Name -ieq $newName) {
                Set-Status "A sheet named '$newName' already exists in $wbName." '#E05050'
                return $false
            }
        }

        $ws.Name = $newName
        Refresh-NavSheets
        foreach ($it in @($NavSheetList.Items)) {
            if ((Get-CleanSheetName $it) -eq $newName) {
                $NavSheetList.SelectedItem = $it
                try { $NavSheetList.ScrollIntoView($it) } catch {}
                break
            }
        }
        Set-Status "Renamed sheet: $oldName -> $newName"
        Write-ActivityLog "Renamed sheet in ${wbName}: $oldName -> $newName"
        return $true
    } catch {
        Set-Status "Rename sheet error: $($_.Exception.Message)" '#E05050'
        return $false
    }
}

function Get-NavSelectedMacroContext {
    $sel = $NavVbaList.SelectedItem
    if (-not ($sel -is [System.Windows.Controls.ListBoxItem])) { return $null }
    $tag = $sel.Tag
    if (-not $tag) { return $null }
    if ([string]$tag.Kind -ne 'Macro') { return $null }
    if (-not $tag.Entry) { return $null }
    return $tag
}

function Invoke-NavSelectedMacro {
    $ctx = Get-NavSelectedMacroContext
    if (-not $ctx) {
        Set-Status 'Select a specific macro entry from Navigator VBA list.' '#E05050'
        return
    }
    $wbName = if ($ctx.Workbook) { [string]$ctx.Workbook } else { [string]$NavWbList.SelectedItem }
    if (-not $wbName) {
        Set-Status 'Select a workbook first.' '#E05050'
        return
    }
    try {
        $xl = Get-ExcelApp -Session Navigator
        if (-not $xl) { throw 'Navigator Excel session is not open.' }
        $wbObj = $xl.Workbooks.Item($wbName)
        $xl.Visible = $true
        $xl.WindowState = -4143 # xlNormal
        $wbObj.Activate()
        $entry = [string]$ctx.Entry
        $xl.Run("$($wbObj.Name)!$entry")
        Set-Status "Ran Navigator macro: $entry"
        Write-ActivityLog "Navigator ran macro: $($wbObj.Name)!$entry"
    } catch {
        Set-Status "Navigator macro run failed: $($_.Exception.Message)" '#E05050'
    }
}

$NavActivateBtn.Add_Click({
    $wb = $NavWbList.SelectedItem
    $sh = $NavSheetList.SelectedItem
    if (-not $wb -or -not $sh) { Set-Status 'Select a workbook and sheet' '#E05050'; return }
    $shName = Get-CleanSheetName $sh
    try {
        $xl = Get-ExcelApp -Session Navigator
        $ws = $xl.Workbooks.Item($wb).Worksheets.Item($shName)
        if ($ws.Visible -ne -1) { $ws.Visible = -1 }
        $xl.Visible = $true
        $xl.WindowState = -4143 # xlNormal (restore from minimized/maximized)
        $ws.Activate()
        Set-Status "Activated: $wb > $shName"
        Write-ActivityLog "Activated sheet: $wb > $shName"
    } catch { Set-Status "Error: $_" '#E05050' }
})

$NavOpenBtn.Add_Click({
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Title  = 'Open Workbook'
    $dlg.Filter = 'Excel Files|*.xlsx;*.xlsm;*.xls;*.xlsb|All Files|*.*'
    if ($dlg.ShowDialog() -eq 'OK') {
        Show-Busy 'Opening workbook...'
        try {
            Open-WorkbookInExcel -FilePath $dlg.FileName -Session Navigator
            Start-Sleep -Milliseconds 500
            Refresh-NavWorkbooks
            Refresh-WorkbookDropdowns
            Set-Status "Opened: $(Split-Path $dlg.FileName -Leaf)"
        } catch { Set-Status "Error: $_" '#E05050' }
        finally { Hide-Busy }
    }
})

$NavBringFrontBtn.Add_Click({
    try {
        $xl = Get-ExcelApp -Session Navigator
        if (-not $xl) { Set-Status 'Navigator Excel session is not open' '#E05050'; return }
        $xl.Visible = $true
        $xl.WindowState = -4143 # xlNormal
        if ($xl.ActiveWorkbook) { $xl.ActiveWorkbook.Activate() }
        Set-Status 'Navigator Excel brought forward'
    } catch { Set-Status "Error: $_" '#E05050' }
})

$NavMinimizeBtn.Add_Click({
    try {
        $xl = Get-ExcelApp -Session Navigator
        if (-not $xl) { Set-Status 'Navigator Excel session is not open' '#E05050'; return }
        $xl.WindowState = -4140 # xlMinimized
        Set-Status 'Navigator Excel minimized'
    } catch { Set-Status "Error: $_" '#E05050' }
})

$NavCloseWbBtn.Add_Click({
    $wb = $NavWbList.SelectedItem
    if (-not $wb) { Set-Status 'Select a workbook first' '#E05050'; return }
    $ans = [System.Windows.MessageBox]::Show(
        "Close workbook '$wb'?`nYes = save then close`nNo = close without saving",
        'Close Workbook', 'YesNoCancel', 'Question')
    if ($ans -eq 'Cancel') { return }
    $save = ($ans -eq 'Yes')
    try {
        $xl = Get-ExcelApp -Session Navigator
        $xl.Workbooks.Item($wb).Close($save)
        Refresh-NavWorkbooks
        Refresh-NavSheets
        Set-Status "Closed workbook: $wb"
        Write-ActivityLog "Closed workbook: $wb (save=$save)"
    } catch { Set-Status "Error: $_" '#E05050' }
})

# Move / Copy selected sheet to destination workbook
$NavMoveCopyBtn.Add_Click({
    $srcWbName = $NavWbList.SelectedItem
    $selItem   = $NavSheetList.SelectedItem
    $destWbName = $NavDestWbCombo.SelectedItem
    $destSheet  = $NavDestSheetCombo.SelectedItem
    if (-not $srcWbName -or -not $selItem) { Set-Status 'Select a source workbook and sheet' '#E05050'; return }
    if (-not $destWbName) { Set-Status 'Select a destination workbook' '#E05050'; return }
    $shName = Get-CleanSheetName $selItem
    $asCopy = $NavCopyChk.IsChecked
    $verb   = if ($asCopy) { 'Copied' } else { 'Moved' }
    Show-Busy "$verb $shName..."
    try {
        $xl     = Get-ExcelApp -Session Navigator
        $srcWs  = $xl.Workbooks.Item($srcWbName).Worksheets.Item($shName)
        $destWb = $xl.Workbooks.Item($destWbName)
        if ($destSheet -and $destSheet -ne '(Move to End)') {
            $before = $destWb.Worksheets.Item($destSheet)
            if ($asCopy) { $srcWs.Copy($before) } else { $srcWs.Move($before) }
        } else {
            # Move/Copy to end: use After parameter with the last sheet
            $lastSheet = $destWb.Worksheets.Item($destWb.Worksheets.Count)
            if ($asCopy) { $srcWs.Copy([System.Type]::Missing, $lastSheet) }
            else         { $srcWs.Move([System.Type]::Missing, $lastSheet) }
        }
        Refresh-NavSheets
        Refresh-NavDestSheets
        Set-Status "$verb '$shName' -> $destWbName"
        Write-ActivityLog "$verb sheet '$shName' from $srcWbName to $destWbName"
    } catch { Set-Status "Error: $_" '#E05050' }
    finally { Hide-Busy }
})

$NavExportSheetBtn.Add_Click({
    $wb = $NavWbList.SelectedItem
    $sh = $NavSheetList.SelectedItem
    if (-not $wb -or -not $sh) { Set-Status 'Select workbook/sheet first' '#E05050'; return }
    $shName = Get-CleanSheetName $sh
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Title = 'Export Selected Sheet'
    $dlg.Filter = 'Excel Workbook|*.xlsx|CSV (active sheet)|*.csv'
    $dlg.FileName = "$shName-export"
    if ($dlg.ShowDialog() -ne 'OK') { return }
    Show-Busy 'Exporting sheet...'
    try {
        $xl = Get-ExcelApp -Session Navigator
        $ws = $xl.Workbooks.Item($wb).Worksheets.Item($shName)
        $ws.Copy()
        $newWb = $xl.ActiveWorkbook
        $ext = [System.IO.Path]::GetExtension($dlg.FileName).ToLower()
        $format = if ($ext -eq '.csv') { 6 } else { 51 }  # 6=xlCSV, 51=xlOpenXMLWorkbook
        $newWb.SaveAs($dlg.FileName, $format)
        $newWb.Close($false)
        Set-Status "Exported sheet: $shName"
        Write-ActivityLog "Exported sheet: $wb > $shName -> $($dlg.FileName)"
    } catch { Set-Status "Export error: $_" '#E05050' }
    finally { Hide-Busy }
})

# Hide selected sheets
$NavHideBtn.Add_Click({
    $wb = $NavWbList.SelectedItem
    if (-not $wb) { Set-Status 'Select a workbook first' '#E05050'; return }
    $selected = @($NavSheetList.SelectedItems)
    if ($selected.Count -eq 0) { Set-Status 'Select sheets to hide' '#E05050'; return }
    try {
        $xl    = Get-ExcelApp -Session Navigator
        $wbObj = $xl.Workbooks.Item($wb)
        # Check we're not hiding ALL visible sheets
        $visCount = 0
        for ($i = 1; $i -le $wbObj.Worksheets.Count; $i++) {
            if ($wbObj.Worksheets.Item($i).Visible -eq -1) { $visCount++ }
        }
        $selectedVisible = 0
        foreach ($s in $selected) {
            $nm = Get-CleanSheetName $s
            if ($wbObj.Worksheets.Item($nm).Visible -eq -1) { $selectedVisible++ }
        }
        if ($selectedVisible -ge $visCount) {
            Set-Status 'Cannot hide all sheets -- at least one must remain visible' '#E05050'
            return
        }
        $count = 0
        foreach ($shRaw in $selected) {
            $shName = Get-CleanSheetName $shRaw
            $ws = $wbObj.Worksheets.Item($shName)
            if ($ws.Visible -eq -1) {
                $ws.Visible = 0  # xlSheetHidden
                $count++
            }
        }
        Refresh-NavSheets
        Set-Status "Hidden $count sheet(s)"
        Write-ActivityLog "Hidden $count sheet(s) in $wb"
    } catch { Set-Status "Error: $_" '#E05050' }
})

# Unhide selected sheets
$NavUnhideBtn.Add_Click({
    $wb = $NavWbList.SelectedItem
    if (-not $wb) { Set-Status 'Select a workbook first' '#E05050'; return }
    $selected = @($NavSheetList.SelectedItems)
    if ($selected.Count -eq 0) { Set-Status 'Select sheets to unhide' '#E05050'; return }
    try {
        $xl    = Get-ExcelApp -Session Navigator
        $wbObj = $xl.Workbooks.Item($wb)
        $count = 0
        foreach ($shRaw in $selected) {
            $shName = Get-CleanSheetName $shRaw
            $ws = $wbObj.Worksheets.Item($shName)
            if ($ws.Visible -ne -1) {
                $ws.Visible = -1  # xlSheetVisible
                $count++
            }
        }
        Refresh-NavSheets
        Set-Status "Unhidden $count sheet(s)"
        Write-ActivityLog "Unhidden $count sheet(s) in $wb"
    } catch { Set-Status "Error: $_" '#E05050' }
})

# Delete selected sheets
$NavDeleteBtn.Add_Click({
    $wb = $NavWbList.SelectedItem
    if (-not $wb) { Set-Status 'Select a workbook first' '#E05050'; return }
    $selected = @($NavSheetList.SelectedItems)
    if ($selected.Count -eq 0) { Set-Status 'Select sheets to delete' '#E05050'; return }
    $names = ($selected | ForEach-Object { Get-CleanSheetName $_ }) -join ', '
    $confirm = [System.Windows.MessageBox]::Show(
        "Permanently delete $($selected.Count) sheet(s)?`n$names`n`nThis cannot be undone.",
        'Confirm Delete', 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    try {
        $xl    = Get-ExcelApp -Session Navigator
        $wbObj = $xl.Workbooks.Item($wb)
        $xl.DisplayAlerts = $false
        $count = 0
        foreach ($shRaw in $selected) {
            $shName = Get-CleanSheetName $shRaw
            $wbObj.Worksheets.Item($shName).Delete()
            $count++
        }
        $xl.DisplayAlerts = $true
        Refresh-NavSheets
        Set-Status "Deleted $count sheet(s)"
        Write-ActivityLog "Deleted $count sheet(s) from ${wb}: $names"
    } catch {
        try { $xl.DisplayAlerts = $true } catch {}
        Set-Status "Error: $_" '#E05050'
    }
})

# Set password protection on selected workbook
$NavSetPwdBtn.Add_Click({
    $wb  = $NavWbList.SelectedItem
    $pwd = $NavPwdBox.Text
    if (-not $wb) { Set-Status 'Select a workbook first' '#E05050'; return }
    if (-not $pwd) { Set-Status 'Enter a password first' '#E05050'; return }
    try {
        $xl    = Get-ExcelApp -Session Navigator
        $wbObj = $xl.Workbooks.Item($wb)
        $wbObj.Password = $pwd
        $NavPwdBox.Text = ''
        Set-Status "Password set on $wb -- save the workbook to persist"
        Write-ActivityLog "Password protection set on $wb"
    } catch { Set-Status "Error: $_" '#E05050' }
})

# Remove password from selected workbook (requires current password)
$NavRemPwdBtn.Add_Click({
    $wb  = $NavWbList.SelectedItem
    $pwd = $NavPwdBox.Text
    if (-not $wb) { Set-Status 'Select a workbook first' '#E05050'; return }
    if (-not $pwd) { Set-Status 'Enter the current password' '#E05050'; return }
    try {
        $xl    = Get-ExcelApp -Session Navigator
        $wbObj = $xl.Workbooks.Item($wb)
        $wbObj.Unprotect($pwd)
        $wbObj.Password = ''
        $NavPwdBox.Text = ''
        Set-Status "Password removed from $wb -- save to persist"
        Write-ActivityLog "Password protection removed from $wb"
    } catch { Set-Status "Error: $_" '#E05050' }
})

# Open a file with password (read-only option via dialog)
$NavOpenPwdBtn.Add_Click({
    $pwd = $NavPwdBox.Text
    if (-not $pwd) { Set-Status 'Enter the password in the box above first' '#E05050'; return }
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Title  = 'Open Password-Protected Workbook'
    $dlg.Filter = 'Excel Files|*.xlsx;*.xlsm;*.xls;*.xlsb|All Files|*.*'
    if ($dlg.ShowDialog() -eq 'OK') {
        $readOnly = [System.Windows.MessageBox]::Show(
            "Open as read-only?`n`nYes = read-only`nNo = full edit access",
            'Open Mode', 'YesNo', 'Question')
        $ro = ($readOnly -eq 'Yes')
        Show-Busy 'Opening protected workbook...'
        try {
            $xl = Get-ExcelApp -Session Navigator -Create
            # Workbooks.Open(Filename, UpdateLinks, ReadOnly, Format, Password)
            [void]($xl.Workbooks.Open($dlg.FileName, 0, $ro, [System.Type]::Missing, $pwd))
            Start-Sleep -Milliseconds 500
            Refresh-NavWorkbooks
            Refresh-WorkbookDropdowns
            $mode = if ($ro) { 'read-only' } else { 'edit' }
            Set-Status "Opened $(Split-Path $dlg.FileName -Leaf) ($mode)"
            Write-ActivityLog "Opened protected file: $(Split-Path $dlg.FileName -Leaf) ($mode)"
        } catch { Set-Status "Error: $_" '#E05050' }
        finally { Hide-Busy }
        $NavPwdBox.Text = ''
    }
})

$NavApplyExcelOptsBtn.Add_Click({
    try {
        $xl = Get-ExcelApp -Session Navigator
        if (-not $xl) { Set-Status 'Navigator Excel session is not open' '#E05050'; return }
        $mode = if ($NavCalcModeCombo.SelectedItem) { $NavCalcModeCombo.SelectedItem.Content.ToString() } else { 'Automatic' }
        # Map label back to Excel xlCalculation constants
        $xl.Calculation = switch ($mode) {
            'Manual'        { -4135 }
            'Semiautomatic' { -4134 }
            default         { -4105 }   # xlAutomatic
        }
        $xl.EnableEvents = [bool]$NavEventsChk.IsChecked
        Set-Status "Excel options updated: Calc=$mode, Events=$($xl.EnableEvents)"
        Write-ActivityLog "Navigator Excel options updated: Calc=$mode, Events=$($xl.EnableEvents)"
    } catch { Set-Status "Excel options error: $_" '#E05050' }
})

# -- Templates tab --
# Handlers for template list selection, save (upsert), delete, paste to Excel,
# and preview toggle with {DATE}/{SHEET}/{USER} placeholder substitution.
$TplList.Add_SelectionChanged({
    $sel = $TplList.SelectedItem
    if ($sel) {
        $templates = Load-Templates
        $t = $templates | Where-Object { $_.Name -eq $sel } | Select-Object -First 1
        if ($t) {
            $TplNameBox.Text    = $t.Name
            $TplContentBox.Text = $t.Content -replace '\\n',"`r`n"
        }
    }
})

$TplSaveBtn.Add_Click({
    $name    = $TplNameBox.Text.Trim()
    $content = $TplContentBox.Text
    if (-not $name) { Set-Status 'Enter a template name' '#E05050'; return }

    $templates = Load-Templates
    $existing = $templates | Where-Object { $_.Name -eq $name }
    if ($existing) { $existing.Content = $content -replace "`r?`n",'\\n' }
    else {
        $templates += [PSCustomObject]@{
            Name    = $name
            Content = $content -replace "`r?`n",'\\n'
        }
    }
    Save-Templates $templates
    Refresh-TemplateList
    Set-Status "Saved template: $name"
    Write-ActivityLog "Saved template: $name"
})

$TplDeleteBtn.Add_Click({
    $sel = $TplList.SelectedItem
    if (-not $sel) { Set-Status 'Select a template to delete' '#E05050'; return }
    $confirm = [System.Windows.MessageBox]::Show(
        "Delete template '$sel'?", 'Confirm Delete', 'YesNo', 'Warning')
    if ($confirm -eq 'Yes') {
        $allTemplates = Load-Templates
        $templates = @($allTemplates | Where-Object { $_.Name -ne $sel })
        Save-Templates @($templates)
        Refresh-TemplateList
        $TplNameBox.Text = ''; $TplContentBox.Text = ''
        Set-Status "Deleted template: $sel"
    }
})

$TplPasteBtn.Add_Click({
    $content = $TplContentBox.Text
    if (-not $content) { Set-Status 'No template content to paste' '#E05050'; return }
    $wb = $ClipWbCombo.SelectedItem
    $slot = Get-PrimaryClipSlotState
    $sh = if ($slot -and $slot.SheetCombo) { $slot.SheetCombo.SelectedItem } else { $ClipSheetCombo.SelectedItem }
    if (-not $wb -or -not $sh) { Set-Status 'Select workbook/sheet on Clipboard tab first' '#E05050'; return }

    # Resolve placeholders
    $content = $content -replace '\{DATE\}', (Get-Date -f 'yyyy-MM-dd')
    $content = $content -replace '\{SHEET\}', $sh
    $content = $content -replace '\{USER\}', $env:USERNAME

    Show-Busy 'Pasting template...'
    try {
        Paste-TextToSheet -WorkbookName $wb -SheetName $sh -CellAddress 'A1' -Text $content
        Set-Status "Pasted template to $wb > $sh"
        Write-ActivityLog "Pasted template to $wb > $sh"
    } catch { Set-Status "Error: $_" '#E05050' }
    finally { Hide-Busy }
})

# Template Preview button
$TplPreviewBtn.Add_Click({
    $content = $TplContentBox.Text
    if (-not $content) {
        $TplPreviewCard.Visibility = 'Collapsed'
        Set-Status 'No content to preview' '#E05050'
        return
    }
    # Resolve placeholders
    $preview = $content -replace '\{DATE\}', (Get-Date -f 'yyyy-MM-dd')
    $slot = Get-PrimaryClipSlotState
    $sh = if ($slot -and $slot.SheetCombo) { $slot.SheetCombo.SelectedItem } else { $ClipSheetCombo.SelectedItem }
    if (-not $sh) { $sh = '(no sheet selected)' }
    $preview = $preview -replace '\{SHEET\}', $sh
    $preview = $preview -replace '\{USER\}', $env:USERNAME
    $TplPreviewTxt.Text = $preview
    if ($TplPreviewCard.Visibility -eq 'Visible') {
        $TplPreviewCard.Visibility = 'Collapsed'
    } else {
        $TplPreviewCard.Visibility = 'Visible'
    }
})

# -- QSync tab --
# Handlers for folder browse buttons and the Run Sync button that calls
# Invoke-QuarterSync and displays the result cards.
$QsBrowseLast.Add_Click({
    $d = [System.Windows.Forms.FolderBrowserDialog]::new()
    $d.Description = 'Select Last Quarter Root Folder'
    if ($d.ShowDialog() -eq 'OK') { $QsLastRoot.Text = $d.SelectedPath }
})

$QsBrowseThis.Add_Click({
    $d = [System.Windows.Forms.FolderBrowserDialog]::new()
    $d.Description = 'Select This Quarter Root Folder'
    if ($d.ShowDialog() -eq 'OK') { $QsThisRoot.Text = $d.SelectedPath }
})

$QsRunSyncBtn.Add_Click({
    $lr = $QsLastRoot.Text.Trim()
    $tr = $QsThisRoot.Text.Trim()
    if (-not $lr -or -not $tr) {
        [System.Windows.MessageBox]::Show('Set both folder paths first.','MacroHub','OK','Warning')
        return
    }
    $cfg = Load-QsConfig
    $cfg.LastQuarterPath = $lr
    $cfg.ThisQuarterPath = $tr
    Save-QsConfig $cfg

    $QsProgressCard.Visibility = 'Visible'
    $QsResultCard.Visibility   = 'Collapsed'
    $QsSyncLogTxt.Clear()
    $QsRunSyncBtn.IsEnabled = $false
    Update-UI

    QsLog "Starting sync..."
    QsLog "Last: $lr"
    QsLog "This: $tr"

    $qName = $QsSyncQuarterName.Text.Trim()
    if (-not $qName) { $qName = "Q_$(Get-Date -f 'yyyy_Qq')" }
    $qPath = New-QuarterFile $qName
    Switch-ActiveQuarter $qPath
    QsLog "Quarter: $qName"
    $res = Invoke-QuarterSync -LastRoot $lr -ThisRoot $tr -QuarterPath $qPath

    foreach ($f in $res.FoldersCreated) { QsLog "  [+] $f" }
    foreach ($f in $res.FoldersSkipped) { QsLog "  [.] $f (exists)" }
    foreach ($t in $res.NewTodos)       { QsLog "  [!] Missing: $($t.OriginalName) in $($t.RelFolder)" }
    foreach ($e in $res.Errors)         { QsLog "  [ERR] $e" }
    QsLog "Done -- $($res.FoldersCreated.Count) created, $($res.NewTodos.Count) new items."

    $QsRCreated.Text  = $res.FoldersCreated.Count
    $QsRSkipped.Text  = $res.FoldersSkipped.Count
    $QsRNewTodos.Text = $res.NewTodos.Count
    $QsRErrors.Text   = $res.Errors.Count

    $QsProgressCard.Visibility = 'Collapsed'
    $QsResultCard.Visibility   = 'Visible'
    $QsRunSyncBtn.IsEnabled = $true

    Refresh-QsTodoList
    Refresh-QsProgress
    Set-Status "Sync complete -- $($res.NewTodos.Count) new items found"
    Write-ActivityLog "QuarterSync: $($res.FoldersCreated.Count) folders, $($res.NewTodos.Count) todos"
})

# -- QTasks tab --
# Handlers for Source A / Target B folder selection, Compare Missing button,
# swap folders, filter controls, status toggle on individual todo rows, and
# the Export to Excel button.
$QsQuarterCombo.Add_SelectionChanged({
    if (Get-ComboText $QsQuarterCombo) {
        Save-QsCompareConfigFromUi
        Refresh-QsTodoList
        Refresh-QsProgress
    }
})
$QsQuarterCombo.Add_LostFocus({ Save-QsCompareConfigFromUi })

$QsCompareCombo.Add_SelectionChanged({
    Save-QsCompareConfigFromUi
    Refresh-QsTodoList
    Refresh-QsProgress
})
$QsCompareCombo.Add_LostFocus({ Save-QsCompareConfigFromUi })

$QsNewQuarterBtn.Add_Click({
    $d = [System.Windows.Forms.FolderBrowserDialog]::new()
    $d.Description = 'Select Source A folder (baseline)'
    if ($d.ShowDialog() -eq 'OK') {
        Add-QsFolderHistoryItem -Combo $QsQuarterCombo -PathText $d.SelectedPath
        Save-QsCompareConfigFromUi
        Refresh-QsTodoList
        Refresh-QsProgress
    }
})

$QsScanFolderBtn.Add_Click({
    $d = [System.Windows.Forms.FolderBrowserDialog]::new()
    $d.Description = 'Select Target B folder (delivery)'
    if ($d.ShowDialog() -eq 'OK') {
        Add-QsFolderHistoryItem -Combo $QsCompareCombo -PathText $d.SelectedPath
        Save-QsCompareConfigFromUi
        Refresh-QsTodoList
        Refresh-QsProgress
    }
})

$QsAddTaskBtn.Add_Click({
    Invoke-QsCompareFromUi
})

$QsSwapFoldersBtn.Add_Click({
    $a = Resolve-QsFolderPath (Get-ComboText $QsQuarterCombo)
    $b = Resolve-QsFolderPath (Get-ComboText $QsCompareCombo)
    if (-not $a -and -not $b) { return }
    if ($b) { Add-QsFolderHistoryItem -Combo $QsQuarterCombo -PathText $b } else { $QsQuarterCombo.Text = '' }
    if ($a) { Add-QsFolderHistoryItem -Combo $QsCompareCombo -PathText $a } else { $QsCompareCombo.Text = '' }
    Save-QsCompareConfigFromUi
    Refresh-QsTodoList
    Refresh-QsProgress
    Set-Status 'Swapped Source A and Target B'
})

$QsSaveTaskBtn.Add_Click({
    # QTasks now centers on folder-to-folder compare output.
    $QsAddTaskCard.Visibility = 'Collapsed'
})

$QsCancelTaskBtn.Add_Click({
    $QsNewTaskName.Text = ''
    $QsNewTaskFolder.Text = ''
    $QsNewTaskDue.Text = ''
    $QsAddTaskCard.Visibility = 'Collapsed'
})

$QsRefreshBtn.Add_Click({ Refresh-QsTodoList; Refresh-QsProgress })
$QsFltStatus.Add_SelectionChanged({ Refresh-QsTodoList })
$QsFltName.Add_TextChanged({ Refresh-QsTodoList })
$QsFltFolder.Add_TextChanged({ Refresh-QsTodoList })

$QsExportBtn.Add_Click({
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Title    = 'Export Missing Deliverables Report'
    $dlg.Filter   = 'Excel Workbook|*.xlsx'
    $dlg.FileName = "FolderCompare_Missing_$(Get-Date -f 'yyyy-MM-dd')"
    if ($dlg.ShowDialog() -eq 'OK') {
        Show-Busy 'Exporting to Excel...'
        try {
            $result = Export-QsToExcel -SavePath $dlg.FileName `
                -SourcePath (Resolve-QsFolderPath (Get-ComboText $QsQuarterCombo)) `
                -TargetPath (Resolve-QsFolderPath (Get-ComboText $QsCompareCombo))
            if ($result -eq $true) {
                [System.Windows.MessageBox]::Show(
                    "Report saved:`n$($dlg.FileName)", 'Export Complete', 'OK', 'Information')
                Set-Status 'Excel export complete'
            } else {
                [System.Windows.MessageBox]::Show("Export failed:`n$result", 'Error', 'OK', 'Error')
            }
        } catch { Set-Status "Export error: $_" '#E05050' }
        finally { Hide-Busy }
    }
})

# ================================================================
#  ALT KEYTIP OVERLAY (Excel-style letter badges)
# ================================================================
# Pressing Alt alone shows coloured letter badges over every interactive control
# on the active tab.  Pressing the indicated letter then activates that control.
# Buttons derive their letter from the underscore prefix in their Content string
# (e.g. '_Run' -> R); other controls use the $ExtraKeyTips registry below.
$script:KeyTipsActive = $false

# Registry: extra keytips for non-button controls (name -> letter)
# These are in ADDITION to buttons which auto-derive from _X content
$script:ExtraKeyTips = @{
    # Clipboard tab
    ClipWbCombo       = 'W'   # Workbook
    ClipSheetCombo    = 'S'   # Sheet  (only on Clipboard tab)
    ClipCellBox       = 'E'   # cEll
    ClipTimestampChk  = 'T'   # Timestamp
    ClipLockBtn       = 'L'   # Lock defaults
    ClipClearDefaultsBtn = 'F' # clear deFaults
    # Macros tab
    MacroWbCombo      = 'W'   # Workbook
    MacroSheetCombo   = 'H'   # sHeet
    MacroList         = 'K'   # picK macro
    # Scheduler tab
    SchedNameBox      = 'N'   # Name
    SchedMacroCombo   = 'M'   # Macro
    SchedTimeBox      = 'I'   # tIme
    SchedFreqCombo    = 'F'   # Frequency
    # Navigator tab
    NavWbList         = 'W'   # Workbooks list
    NavDestWbCombo    = 'B'   # dest wB
    NavDestSheetCombo = 'S'   # dest Sheet
    NavCopyChk        = 'Y'   # copY
    NavPwdBox         = 'P'   # Password box
    NavCalcModeCombo  = 'C'   # Calculation
    NavEventsChk      = 'E'   # Events
    # QTasks tab
    QsQuarterCombo    = 'Q'   # Quarter
    QsCompareCombo    = 'P'   # comPare
    QsFltStatus       = 'Z'   # status (Z to avoid conflicts)
    QsFltName         = 'N'   # Name filter
    QsFltFolder       = 'F'   # Folder filter
    QsNewTaskName     = 'K'   # tasK name
    QsNewTaskDue      = 'D'   # Due date
    QsScanPathBox     = 'X'   # path boX
}

# Walks the WPF visual tree, collects visible interactive controls, positions a
# coloured letter badge over each one, and stores the letter->control mapping in
# $script:KeyTipMap for dispatch by the PreviewKeyDown handler.
function Show-KeyTips {
    $KeyTipCanvas.Children.Clear()
    $KeyTipCanvas.Visibility = 'Visible'
    $script:KeyTipsActive = $true
    $script:KeyTipMap = @{}

    # Collect ALL visible interactive controls from the visual tree
    $controls = New-Object System.Collections.Generic.List[System.Windows.FrameworkElement]
    $stack = New-Object System.Collections.Generic.Stack[System.Windows.DependencyObject]
    $stack.Push($Window)
    while ($stack.Count -gt 0) {
        $parent = $stack.Pop()
        $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent)
        for ($i = 0; $i -lt $count; $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)
            if ($child -is [System.Windows.FrameworkElement] -and $child.Visibility -eq 'Visible') {
                $isInteractive = $child -is [System.Windows.Controls.Button] -or
                                 $child -is [System.Windows.Controls.ComboBox] -or
                                 $child -is [System.Windows.Controls.CheckBox] -or
                                 ($child -is [System.Windows.Controls.TextBox] -and $child.Name -and -not $child.IsReadOnly) -or
                                 ($child -is [System.Windows.Controls.ListBox] -and $child.Name)
                if ($isInteractive) { $controls.Add($child) }
            }
            $stack.Push($child)
        }
    }

    foreach ($ctrl in $controls) {
        $letter = $null

        # Buttons: derive from _X in Content
        if ($ctrl -is [System.Windows.Controls.Button]) {
            $text = if ($ctrl.Content -is [string]) { $ctrl.Content } else { $null }
            if ($text) {
                $m = [regex]::Match($text, '_([A-Za-z])')
                if ($m.Success) { $letter = $m.Groups[1].Value.ToUpper() }
            }
        }
        # Non-buttons: look up in ExtraKeyTips registry by Name
        elseif ($ctrl.Name -and $script:ExtraKeyTips.ContainsKey($ctrl.Name)) {
            $letter = $script:ExtraKeyTips[$ctrl.Name]
        }

        if (-not $letter) { continue }

        # Only show if control is actually visible on screen
        try {
            $pt = $ctrl.TranslatePoint([System.Windows.Point]::new(0, 0), $Window)
        } catch { continue }
        if ($pt.X -lt 0 -or $pt.Y -lt 0) { continue }
        if ($pt.X -gt $Window.ActualWidth -or $pt.Y -gt $Window.ActualHeight) { continue }

        # Store the control (last one wins for duplicate letters - only visible tab matters)
        $script:KeyTipMap[$letter] = $ctrl

        # Create the keytip badge
        $badge = New-Object System.Windows.Controls.Border
        # Color-code by control type
        $badgeColor = if ($ctrl -is [System.Windows.Controls.Button]) {
            [System.Windows.Media.Color]::FromRgb(0x4C, 0x9F, 0xE6)  # blue for buttons
        } elseif ($ctrl -is [System.Windows.Controls.ComboBox]) {
            [System.Windows.Media.Color]::FromRgb(0x50, 0xC8, 0x78)  # green for combos
        } elseif ($ctrl -is [System.Windows.Controls.CheckBox]) {
            [System.Windows.Media.Color]::FromRgb(0xFF, 0xD7, 0x00)  # gold for checkboxes
        } else {
            [System.Windows.Media.Color]::FromRgb(0xFF, 0xA5, 0x00)  # orange for text/list
        }
        $badge.Background   = [System.Windows.Media.SolidColorBrush]::new($badgeColor)
        $badge.BorderBrush  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xFF, 0xFF, 0xFF))
        $badge.BorderThickness = [System.Windows.Thickness]::new(1)
        $badge.CornerRadius    = [System.Windows.CornerRadius]::new(3)
        $badge.Padding         = [System.Windows.Thickness]::new(4, 1, 4, 1)

        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text       = $letter
        $lbl.Foreground = $(if ($ctrl -is [System.Windows.Controls.CheckBox]) {
            [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x00, 0x00, 0x00))
        } else { [System.Windows.Media.Brushes]::White })
        $lbl.FontSize   = 11
        $lbl.FontWeight = [System.Windows.FontWeights]::Bold
        $lbl.HorizontalAlignment = 'Center'
        $badge.Child = $lbl

        # Position badge at bottom-center of the control
        $cw = $ctrl.ActualWidth
        [System.Windows.Controls.Canvas]::SetLeft($badge, $pt.X + ($cw / 2) - 10)
        [System.Windows.Controls.Canvas]::SetTop($badge,  $pt.Y + $ctrl.ActualHeight - 4)
        [System.Windows.Controls.Panel]::SetZIndex($badge, 999)

        [void]($KeyTipCanvas.Children.Add($badge))
    }
}

function Hide-KeyTips {
    $KeyTipCanvas.Children.Clear()
    $KeyTipCanvas.Visibility = 'Collapsed'
    $script:KeyTipsActive = $false
    $script:KeyTipMap = @{}
}

function Get-TabIndexFromDigitKey([object]$keyObj) {
    $k = [string]$keyObj
    switch ($k) {
        'D1' { 0 } 'D2' { 1 } 'D3' { 2 } 'D4' { 3 } 'D5' { 4 }
        'D6' { 5 } 'D7' { 6 } 'D8' { 7 } 'D9' { 8 } 'D0' { 9 }
        'NumPad1' { 0 } 'NumPad2' { 1 } 'NumPad3' { 2 } 'NumPad4' { 3 } 'NumPad5' { 4 }
        'NumPad6' { 5 } 'NumPad7' { 6 } 'NumPad8' { 7 } 'NumPad9' { 8 } 'NumPad0' { 9 }
        default { -1 }
    }
}

function Select-MainTabIndex([int]$idx) {
    if ($idx -lt 0 -or $idx -ge $MainTabs.Items.Count) { return $false }
    $tab = $MainTabs.Items[$idx]
    if ($tab -is [System.Windows.Controls.TabItem] -and $tab.Visibility -eq 'Collapsed') { return $false }
    $MainTabs.SelectedIndex = $idx
    return $true
}

# Returns a list of visible TabItem objects in order
function Get-VisibleTabs {
    $result = @()
    for ($i = 0; $i -lt $MainTabs.Items.Count; $i++) {
        $t = $MainTabs.Items[$i]
        if (-not ($t -is [System.Windows.Controls.TabItem]) -or $t.Visibility -ne 'Collapsed') {
            $result += [PSCustomObject]@{ Tab = $t; AbsIndex = $i }
        }
    }
    return $result
}

# Select the Nth visible tab (1-based position among visible tabs)
function Select-VisibleTabByPosition([int]$pos) {
    $visible = Get-VisibleTabs
    if ($pos -lt 1 -or $pos -gt $visible.Count) { return $false }
    $MainTabs.SelectedIndex = $visible[$pos - 1].AbsIndex
    return $true
}

function Select-MainTabRelative([int]$delta) {
    $visible = Get-VisibleTabs
    if ($visible.Count -le 0) { return $false }
    # Find current position among visible tabs
    $curAbs = $MainTabs.SelectedIndex
    $curVis = -1
    for ($i = 0; $i -lt $visible.Count; $i++) {
        if ($visible[$i].AbsIndex -eq $curAbs) { $curVis = $i; break }
    }
    if ($curVis -lt 0) { $curVis = 0 }
    $nextVis = ($curVis + $delta + $visible.Count) % $visible.Count
    $MainTabs.SelectedIndex = $visible[$nextVis].AbsIndex
    return $true
}

# -- Keyboard shortcuts: Alt+0-9 / Ctrl+0-9 / Ctrl+PgUp/PgDn tab switch, Ctrl+F search --
# PreviewKeyDown fires before focused controls see the key, allowing tab-switch
# shortcuts to work even when a TextBox has focus.
$Window.Add_PreviewKeyDown({
    $mods = [System.Windows.Input.Keyboard]::Modifiers
    $ctrlDown = (($mods -band [System.Windows.Input.ModifierKeys]::Control) -ne 0)
    $shiftDown = (($mods -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0)

    # Global tab/search shortcuts on PreviewKeyDown so focused controls can't swallow them.
    if ($ctrlDown) {
        if ($_.Key -eq 'Next' -or $_.Key -eq 'PageDown') {
            if (Select-MainTabRelative 1) {
                if ($script:KeyTipsActive) { Hide-KeyTips }
                $_.Handled = $true
                return
            }
        }
        if ($_.Key -eq 'Prior' -or $_.Key -eq 'PageUp') {
            if (Select-MainTabRelative -1) {
                if ($script:KeyTipsActive) { Hide-KeyTips }
                $_.Handled = $true
                return
            }
        }
        if ($_.Key -eq 'Tab') {
            $delta = if ($shiftDown) { -1 } else { 1 }
            if (Select-MainTabRelative $delta) {
                if ($script:KeyTipsActive) { Hide-KeyTips }
                $_.Handled = $true
                return
            }
        }
        if ($_.Key -eq 'F') {
            $SearchOverlayBd.Visibility = 'Visible'
            $SearchBox.Focus() | Out-Null
            $_.Handled = $true
            return
        }
        if ($_.Key -eq 'E') {
            # Navigator quick rename: select sheet then Ctrl+E.
            if ($MainTabs.SelectedIndex -eq 3 -and $NavSheetList -and $NavSheetList.SelectedItem) {
                [void](Invoke-NavSheetRename)
                if ($script:KeyTipsActive) { Hide-KeyTips }
                $_.Handled = $true
                return
            }
        }
        $ctrlNum = Get-TabIndexFromDigitKey $_.Key
        if ($ctrlNum -ge 0 -and (Select-MainTabIndex $ctrlNum)) {
            if ($script:KeyTipsActive) { Hide-KeyTips }
            $_.Handled = $true
            return
        }
    }

    # Alt+1..9 tab switching by visible-tab position (SystemKey path catches Alt+digit reliably)
    if ($_.Key -eq 'System') {
        $altNum = Get-TabIndexFromDigitKey $_.SystemKey
        if ($altNum -ge 0) {
            # altNum is 0-based absolute from digit keys (D1→0, D2→1, …); treat as 1-based visible position
            $visPos = $altNum + 1
            if ($visPos -ge 1 -and (Select-VisibleTabByPosition $visPos)) {
                if ($script:KeyTipsActive) { Hide-KeyTips }
                $_.Handled = $true
                return
            }
        }
    }

    # Alt key alone (System key) shows keytips
    if ($_.Key -eq 'System' -and ($_.SystemKey -eq 'LeftAlt' -or $_.SystemKey -eq 'RightAlt')) {
        if (-not $script:KeyTipsActive) {
            Show-KeyTips
        }
        return
    }

    # If keytips are active and a letter is pressed, invoke that control
    if ($script:KeyTipsActive) {
        $letter = $null
        $k = if ($_.Key -eq 'System') { $_.SystemKey } else { $_.Key }
        $ks = $k.ToString()
        if ($ks.Length -eq 1 -and $ks -match '[A-Z]') {
            $letter = $ks
        }
        if ($letter -and $script:KeyTipMap.ContainsKey($letter)) {
            $target = $script:KeyTipMap[$letter]
            Hide-KeyTips

            # Activate based on control type
            if ($target -is [System.Windows.Controls.Button]) {
                $peer = [System.Windows.Automation.Peers.ButtonAutomationPeer]::new($target)
                $invokeProv = $peer.GetPattern([System.Windows.Automation.Peers.PatternInterface]::Invoke)
                if ($invokeProv) { $invokeProv.Invoke() }
            }
            elseif ($target -is [System.Windows.Controls.ComboBox]) {
                $target.Focus()
                $target.IsDropDownOpen = $true
            }
            elseif ($target -is [System.Windows.Controls.CheckBox]) {
                $target.IsChecked = -not $target.IsChecked
            }
            elseif ($target -is [System.Windows.Controls.TextBox]) {
                $target.Focus()
                $target.SelectAll()
            }
            elseif ($target -is [System.Windows.Controls.ListBox]) {
                $target.Focus()
            }

            $_.Handled = $true
            return
        }
        # Any other key dismisses keytips
        Hide-KeyTips
    }
})

$Window.Add_PreviewKeyUp({
    # Alt released without pressing a letter -- toggle keytips
    if ($_.Key -eq 'System' -and ($_.SystemKey -eq 'LeftAlt' -or $_.SystemKey -eq 'RightAlt')) {
        # Keep keytips visible -- they were shown on KeyDown
        # They'll dismiss on next non-Alt key press or Escape
        return
    }
})

$Window.Add_KeyDown({
    if ($_.Handled) { return }
    $mods = $_.KeyboardDevice.Modifiers

    # Alt+1..9 (and Alt+0 when available) -> tab switch
    if (($mods -band [System.Windows.Input.ModifierKeys]::Alt) -ne 0) {
        $altNum = Get-TabIndexFromDigitKey $_.Key
        if ($altNum -ge 0 -and (Select-MainTabIndex $altNum)) {
            if ($script:KeyTipsActive) { Hide-KeyTips }
            $_.Handled = $true
            return
        }
    }

    if (($mods -band [System.Windows.Input.ModifierKeys]::Control) -ne 0) {
        # Ctrl+PageDown / Ctrl+PageUp -> next/previous tab
        if ($_.Key -eq 'Next' -or $_.Key -eq 'PageDown') {
            if (Select-MainTabRelative 1) {
                $_.Handled = $true
                return
            }
        }
        if ($_.Key -eq 'Prior' -or $_.Key -eq 'PageUp') {
            if (Select-MainTabRelative -1) {
                $_.Handled = $true
                return
            }
        }

        # Ctrl+F -> open search
        if ($_.Key -eq 'F') {
            $SearchOverlayBd.Visibility = 'Visible'
            $SearchBox.Focus()
            $_.Handled = $true
            return
        }
        if ($_.Key -eq 'Tab') {
            $delta = if ((($mods -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0)) { -1 } else { 1 }
            if (Select-MainTabRelative $delta) {
                $_.Handled = $true
                return
            }
        }
        $num = Get-TabIndexFromDigitKey $_.Key
        if ($num -ge 0 -and (Select-MainTabIndex $num)) {
            $_.Handled = $true
            return
        }
    }
    # Esc -> close search or keytips
    if ($_.Key -eq 'Escape') {
        if ($script:KeyTipsActive) { Hide-KeyTips; $_.Handled = $true; return }
        if ($SearchOverlayBd.Visibility -eq 'Visible') {
            $SearchOverlayBd.Visibility = 'Collapsed'
            $_.Handled = $true
        }
    }
})

# -- Search handlers --
$SearchCloseBtn.Add_Click({
    $SearchOverlayBd.Visibility = 'Collapsed'
})

# -- Search debounce (200ms delay) --
$script:SearchTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:SearchTimer.Interval = [TimeSpan]::FromMilliseconds(200)
$script:SearchTimer.Add_Tick({
    $script:SearchTimer.Stop()
    $q = $SearchBox.Text.Trim()
    $SearchResultsPanel.Children.Clear()
    if ($q.Length -lt 2) { return }
    $results = Search-AllTabs $q
    $lastTab = ''
    foreach ($r in $results) {
        # Tab section header
        if ($r.Tab -ne $lastTab) {
            $lastTab = $r.Tab
            $hdr = [System.Windows.Controls.TextBlock]::new()
            $hdr.Text = $r.Tab.ToUpper()
            $hdr.Foreground = (HexBrush '#4C9FE6')
            $hdr.FontSize = 10
            $hdr.FontWeight = [System.Windows.FontWeights]::SemiBold
            $hdr.Margin = [System.Windows.Thickness]::new(0,8,0,4)
            [void]($SearchResultsPanel.Children.Add($hdr))
        }
        # Result row
        $row = [System.Windows.Controls.Border]::new()
        $row.Padding = [System.Windows.Thickness]::new(8,4,8,4)
        $row.BorderBrush = (HexBrush '#3E3E42')
        $row.BorderThickness = [System.Windows.Thickness]::new(0,0,0,1)
        $row.Cursor = [System.Windows.Input.Cursors]::Hand
        $row.Tag = $r.Tab

        $g = [System.Windows.Controls.Grid]::new()
        $cA = [System.Windows.Controls.ColumnDefinition]::new()
        $cA.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $cB = [System.Windows.Controls.ColumnDefinition]::new()
        $cB.Width = [System.Windows.GridLength]::Auto
        [void]($g.ColumnDefinitions.Add($cA))
        [void]($g.ColumnDefinitions.Add($cB))

        $itemTxt = [System.Windows.Controls.TextBlock]::new()
        $itemTxt.Text = $r.Item
        $itemTxt.Foreground = (HexBrush '#FFFFFF')
        $itemTxt.FontSize = 12
        [System.Windows.Controls.Grid]::SetColumn($itemTxt, 0)
        [void]($g.Children.Add($itemTxt))

        $detailTxt = [System.Windows.Controls.TextBlock]::new()
        $detailTxt.Text = $r.Detail
        $detailTxt.Foreground = (HexBrush '#6E6E6E')
        $detailTxt.FontSize = 11
        [System.Windows.Controls.Grid]::SetColumn($detailTxt, 1)
        [void]($g.Children.Add($detailTxt))

        $row.Child = $g

        # Click to navigate to tab
        $tabName = $r.Tab
        $row.Add_MouseLeftButtonDown({
            $tabIdx = switch ($tabName) {
                'Macros'    { 1 }
                'Templates' { 4 }
                'Scheduler' { 2 }
                'QTasks'    { 6 }
                default     { -1 }
            }
            if ($tabIdx -ge 0) { $MainTabs.SelectedIndex = $tabIdx }
            $SearchOverlayBd.Visibility = 'Collapsed'
        }.GetNewClosure())

        [void]($SearchResultsPanel.Children.Add($row))
    }
    if ($results.Count -eq 0) {
        $noRes = [System.Windows.Controls.TextBlock]::new()
        $noRes.Text = 'No results found'
        $noRes.Foreground = (HexBrush '#6E6E6E')
        $noRes.FontStyle = [System.Windows.FontStyles]::Italic
        $noRes.Margin = [System.Windows.Thickness]::new(0,12,0,0)
        $noRes.HorizontalAlignment = 'Center'
        [void]($SearchResultsPanel.Children.Add($noRes))
    }
})

$SearchBox.Add_TextChanged({
    $script:SearchTimer.Stop()
    $script:SearchTimer.Start()
})

# ================================================================
#  STATUS CLOCK TIMER (updates every 30 seconds)
# ================================================================
$clockTimer = [System.Windows.Threading.DispatcherTimer]::new()
$clockTimer.Interval = [TimeSpan]::FromSeconds(30)
$clockTimer.Add_Tick({
    if ($script:StatusClock) { $script:StatusClock.Text = (Get-Date -f 'HH:mm') }
})
$clockTimer.Start()
if ($script:StatusClock) { $script:StatusClock.Text = (Get-Date -f 'HH:mm') }

# ================================================================
#  SCHEDULER NOTIFICATION (balloon toast every 60s check)
# ================================================================
$script:NotifyIcon = [System.Windows.Forms.NotifyIcon]::new()
$script:NotifyIcon.Icon = [System.Drawing.SystemIcons]::Application
$script:NotifyIcon.Text = 'MacroHub'
$script:NotifyIcon.Visible = $false
$script:LastNotifiedTasks = @{}

$notifyTimer = [System.Windows.Threading.DispatcherTimer]::new()
$notifyTimer.Interval = [TimeSpan]::FromSeconds(60)
$notifyTimer.Add_Tick({
    try {
        $tasks = Get-HubTasks
        foreach ($t in $tasks) {
            if ($t.State -eq 'Running') {  # Running state
                $key = $t.Name
                $lastNotify = $script:LastNotifiedTasks[$key]
                $now = Get-Date
                if (-not $lastNotify -or ($now - $lastNotify).TotalMinutes -gt 5) {
                    $script:NotifyIcon.Visible = $true
                    $script:NotifyIcon.ShowBalloonTip(
                        5000,
                        'MacroHub Scheduler',
                        "Task fired: $($t.Name)",
                        [System.Windows.Forms.ToolTipIcon]::Info
                    )
                    $script:LastNotifiedTasks[$key] = $now
                    Write-ActivityLog "Notification: task fired $($t.Name)"
                }
            }
        }
    } catch {}
})
$notifyTimer.Start()

# Clean up NotifyIcon when window closes
$Window.Add_Closed({
    try { Stop-ClipSequenceCapture } catch {}
    $script:NotifyIcon.Visible = $false
    $script:NotifyIcon.Dispose()
})

# ================================================================
#  INITIAL LOAD  (defers heavy tab work until first visit)
# ================================================================
# Runs at window creation time (before ShowDialog).  Heavy per-tab data is loaded
# lazily via the Add_SelectionChanged handler so startup is fast even when Excel
# is not open.  Tabs 0, 1, 4, and 5 are considered always-loaded immediately.
$script:TabsLoaded = @{}   # track which tabs have been lazy-loaded

# ================================================================
#  APPLY -HideTabs: collapse tabs listed in the parameter
# ================================================================
if ($HideTabs -and $HideTabs.Trim()) {
    # Canonical map: header text → TabItem
    $tabHeaderMap = @{}
    for ($i = 0; $i -lt $MainTabs.Items.Count; $i++) {
        $t = $MainTabs.Items[$i]
        if ($t -is [System.Windows.Controls.TabItem]) {
            $tabHeaderMap[[string]$t.Header] = $t
        }
    }
    foreach ($name in ($HideTabs -split ',')) {
        $name = $name.Trim()
        if (-not $name) { continue }
        # Case-insensitive lookup
        $matched = $tabHeaderMap.Keys | Where-Object { $_ -ieq $name } | Select-Object -First 1
        if ($matched) {
            $tabHeaderMap[$matched].Visibility = 'Collapsed'
        }
    }
    # Ensure the currently selected tab is visible; if not, jump to first visible tab
    $curTab = $MainTabs.Items[$MainTabs.SelectedIndex]
    if ($curTab -is [System.Windows.Controls.TabItem] -and $curTab.Visibility -eq 'Collapsed') {
        for ($i = 0; $i -lt $MainTabs.Items.Count; $i++) {
            $t = $MainTabs.Items[$i]
            if (-not ($t -is [System.Windows.Controls.TabItem]) -or $t.Visibility -ne 'Collapsed') {
                $MainTabs.SelectedIndex = $i; break
            }
        }
    }
}

# Load workbook lists first so defaults can be applied safely.
Refresh-WorkbookDropdowns

# Load saved clipboard defaults (if any) into hidden backing controls.
$clipDef = Load-ClipDefaults
if ($clipDef) {
    try {
        if ($clipDef.Workbook) {
            foreach ($it in $ClipWbCombo.Items) {
                if ([string]$it -eq $clipDef.Workbook) { $ClipWbCombo.SelectedItem = $it; break }
            }
        }
        Refresh-ClipSheets
        if ($clipDef.Sheet -and $ClipSheetCombo) {
            foreach ($it in $ClipSheetCombo.Items) {
                if ([string]$it -eq $clipDef.Sheet) { $ClipSheetCombo.SelectedItem = $it; break }
            }
        }
        if ($clipDef.Cell) { $ClipCellBox.Text = $clipDef.Cell }
        if ($null -ne $clipDef.Timestamp) { $ClipTimestampChk.IsChecked = [bool]$clipDef.Timestamp }
        # Load timestamp cell addresses — fall back gracefully from old offset-based format.
        if ($clipDef.DateCell) { $ClipDateCell.Text = $clipDef.DateCell }
        elseif ($clipDef.DateOffset) { $ClipDateCell.Text = 'A3' }   # migrate: old offset → new default
        if ($clipDef.TimeCell) { $ClipTimeCell.Text = $clipDef.TimeCell }
        elseif ($clipDef.TimeOffset) { $ClipTimeCell.Text = 'A2' }   # migrate: old offset → new default
        $wb2 = if ($clipDef.Workbook) { $clipDef.Workbook } else { '?' }
        $sh2 = if ($clipDef.Sheet) { $clipDef.Sheet } else { '?' }
        $cl2 = if ($clipDef.Cell) { $clipDef.Cell } else { 'A1' }
        $ClipDefaultsIndicator.Text = "(defaults: $wb2 > $sh2 > $cl2)"
        $ClipDefaultsIndicator.Visibility = 'Visible'
    } catch {}
}

# Start with 1 clipboard slot (Tab 1 -- always visible), seeded from defaults.
Add-ClipSlotUI -Panel $ClipSlotsPanel -WbCombo $ClipWbCombo `
    -DefaultSheet ([string]$ClipSheetCombo.SelectedItem) `
    -DefaultCell $ClipCellBox.Text `
    -DefaultTimestamp ([bool]$ClipTimestampChk.IsChecked) `
    -DefaultDateCell $ClipDateCell.Text `
    -DefaultTimeCell $ClipTimeCell.Text `
    -CountLabel $ClipSlotCount
Refresh-ClipSheets
Refresh-MacroList
Refresh-TemplateList

# Load QSync saved config (lightweight)
$qsCfg = Load-QsConfig
if ($qsCfg.LastQuarterPath) { $QsLastRoot.Text = $qsCfg.LastQuarterPath }
if ($qsCfg.ThisQuarterPath) { $QsThisRoot.Text = $qsCfg.ThisQuarterPath }

# Create quarters directory
if (-not (Test-Path $script:QuartersDir)) {
    [void](New-Item -ItemType Directory -Path $script:QuartersDir -Force)
}

# Migrate legacy qs_todos.csv if quarters/ is empty
if ((Test-Path $script:QsTodoCSV) -and (Get-QuarterList).Count -eq 0) {
    try {
        $legacyRows = @(Import-Csv $script:QsTodoCSV)
        $legDest = Join-Path $script:QuartersDir 'Legacy_Import.json'
        Save-QsTodos -todos $legacyRows -Path $legDest
    } catch {}
}

# Hydrate QTasks folder compare selectors from config / last compare results.
if ($qsCfg.CompareSourcePath) { Add-QsFolderHistoryItem -Combo $QsQuarterCombo -PathText $qsCfg.CompareSourcePath }
if ($qsCfg.CompareTargetPath) { Add-QsFolderHistoryItem -Combo $QsCompareCombo -PathText $qsCfg.CompareTargetPath }
try {
    $lastCompareRows = Load-QsCompareTodos
    if ($lastCompareRows.Count -gt 0) {
        $lastSource = Resolve-QsFolderPath ([string]$lastCompareRows[0].SourceRoot)
        $lastTarget = Resolve-QsFolderPath ([string]$lastCompareRows[0].TargetRoot)
        if ($lastSource) { Add-QsFolderHistoryItem -Combo $QsQuarterCombo -PathText $lastSource }
        if ($lastTarget) { Add-QsFolderHistoryItem -Combo $QsCompareCombo -PathText $lastTarget }
    }
} catch {}

# Lazy-load tab content on first visit: tabs 2, 3, 6 hit external COM/JSON on
# first activation so they are deferred; all other tabs were loaded at startup.
$MainTabs.Add_SelectionChanged({
    $idx = $MainTabs.SelectedIndex

    if ($script:TabsLoaded[$idx]) { return }
    $script:TabsLoaded[$idx] = $true
    switch ($idx) {
        2 { Refresh-TaskList }                                         # Scheduler (COM + WMI)
        3 { Refresh-NavWorkbooks; Refresh-NavSheets }                  # Navigator (Excel COM)
        6 { Refresh-QsTodoList; Refresh-QsProgress }                   # QTasks (JSON + UI build)
        7 {                                                            # File Index (load from cache)
            $cachedIdx = Load-FidxCache
            if ($cachedIdx -and $cachedIdx.Items) {
                $script:FidxAllItems = @($cachedIdx.Items)
                $script:FidxLastScan = try { [datetime]::Parse($cachedIdx.ScannedOn) } catch { $null }
                $FidxRootBox.Text    = [string]$cachedIdx.Root
                $FidxTotalCount.Text = $script:FidxAllItems.Count
                $FidxLastScanTxt.Text = "Last scanned: $($cachedIdx.ScannedOn)"
                $FidxStatusTxt.Text  = "Index loaded from cache ($($script:FidxAllItems.Count) files)."
                $FidxGrid.ItemsSource = $script:FidxAllItems
                $FidxMatchCount.Text = "$($script:FidxAllItems.Count) files"
                if ((Get-FidxCacheAgeSec) -lt $script:FidxCooldownSec) { Start-FidxCooldownTimer }
            }
        }
    }
})

# ================================================================
#  FILE INDEX EVENT HANDLERS
# ================================================================

# Helper: rebuild DataGrid from filtered results
function Refresh-FidxGrid {
    $q = if ($FidxSearchBox) { [string]$FidxSearchBox.Text } else { '' }
    $filtered = Get-FidxFiltered -Query $q
    $FidxGrid.ItemsSource = $filtered
    $FidxMatchCount.Text  = "$($filtered.Count) files"
}

# Browse button — folder picker
$FidxBrowseBtn.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select root folder to index'
    if ($FidxRootBox.Text -and (Test-Path $FidxRootBox.Text -PathType Container)) {
        $dlg.SelectedPath = $FidxRootBox.Text
    }
    if ($dlg.ShowDialog() -eq 'OK') {
        $FidxRootBox.Text = $dlg.SelectedPath
    }
})

# Index Now button — scan and cache
$FidxIndexBtn.Add_Click({
    $root = $FidxRootBox.Text.Trim()
    if (-not $root) {
        [System.Windows.MessageBox]::Show('Enter a root folder path first.', 'File Index', 'OK', 'Information')
        return
    }
    if (-not (Test-Path $root -PathType Container)) {
        [System.Windows.MessageBox]::Show("Folder not found:`n$root", 'File Index', 'OK', 'Warning')
        return
    }
    $FidxIndexBtn.IsEnabled  = $false
    $FidxStatusTxt.Text      = 'Scanning… this may take a moment for large folders.'
    $FidxStatusTxt.Foreground = HexBrush '#4C9FE6'
    Update-UI

    try {
        Show-Busy 'Building file index…'
        $items = Invoke-FidxScan -RootPath $root
        $script:FidxAllItems = $items
        $script:FidxLastScan = Get-Date
        Save-FidxCache -Items $items -Root $root

        $FidxTotalCount.Text  = $items.Count
        $FidxLastScanTxt.Text = "Last scanned: $(Get-Date -f 'yyyy-MM-dd HH:mm')"
        $FidxStatusTxt.Text   = "Index complete — $($items.Count) files in $(Split-Path $root -Leaf)."
        $FidxStatusTxt.Foreground = HexBrush '#50A050'
        Refresh-FidxGrid
        Start-FidxCooldownTimer
        Set-Status "File index built: $($items.Count) files"
    } catch {
        $FidxStatusTxt.Text = "Error: $_"
        $FidxStatusTxt.Foreground = HexBrush '#E05050'
        $FidxIndexBtn.IsEnabled = $true
        Set-Status "File index error: $_" '#E05050'
    } finally {
        Hide-Busy
    }
})

# Search box — real-time filter as user types
$FidxSearchBox.Add_TextChanged({
    if ($script:FidxAllItems.Count -gt 0) { Refresh-FidxGrid }
})

# DataGrid double-click — open selected file
$FidxGrid.Add_MouseDoubleClick({
    $item = $FidxGrid.SelectedItem
    if ($item -and $item.Path -and (Test-Path $item.Path)) {
        try { Start-Process $item.Path } catch {
            Set-Status "Cannot open: $_" '#E05050'
        }
    }
})

# DataGrid Enter key — open selected file
$FidxGrid.Add_KeyDown({
    if ($_.Key -eq 'Return') {
        $item = $FidxGrid.SelectedItem
        if ($item -and $item.Path -and (Test-Path $item.Path)) {
            try { Start-Process $item.Path } catch {
                Set-Status "Cannot open: $_" '#E05050'
            }
        }
    }
})

# ================================================================
#  FORMULA AUDITOR — state, helpers, logic, event wiring
# ================================================================

# ---- State ----
$script:FaNavHistory            = [System.Collections.ArrayList]::new()
$script:FaLastCellSig           = ''
# Highlight undo stacks — each entry: @{ Cell = <COM Range>; OrigBorders = <hashtable>; OrigInterior = <hashtable> }
$script:FaHighlightBorderCells  = [System.Collections.ArrayList]::new()
$script:FaHighlightInteriorCells = [System.Collections.ArrayList]::new()

# Excel color helpers (R + G*256 + B*65536)
$script:FaBlueBorderColor  = [long](76  + 159*256 + 230*65536)  # #4C9FE6
$script:FaGreenFillColor   = [long](144 + 238*256 + 144*65536)  # light green  (#90EE90)
$script:FaBlueFillColor    = [long](173 + 216*256 + 230*65536)  # light blue   (#ADD8E6)

# ---- Border save/restore ----
function Save-FaCellBorderState {
    param($Cell)
    $s = @{}
    foreach ($edge in @(7, 8, 9, 10)) {   # Left=7 Top=8 Bottom=9 Right=10
        try {
            $b = $Cell.Borders($edge)
            $s[$edge] = @{ LineStyle = $b.LineStyle; Weight = $b.Weight;
                           Color = $b.Color; ColorIndex = $b.ColorIndex }
        } catch { $s[$edge] = $null }
    }
    return $s
}

function Restore-FaCellBorderState {
    param($Cell, $State)
    foreach ($edge in @(7, 8, 9, 10)) {
        $s = $State[$edge]
        if ($null -eq $s) { continue }
        try {
            $b = $Cell.Borders($edge)
            if ($s.LineStyle -eq -4142) {      # xlLineStyleNone
                $b.LineStyle = -4142
            } else {
                $b.LineStyle = $s.LineStyle
                $b.Weight    = $s.Weight
                if ($s.ColorIndex -eq -4105) { # xlColorIndexAutomatic
                    $b.ColorIndex = -4105
                } else {
                    try { $b.Color = $s.Color } catch { $b.ColorIndex = -4105 }
                }
            }
        } catch {}
    }
}

# ---- Interior save/restore ----
function Save-FaCellInteriorState {
    param($Cell)
    try {
        $ci = $Cell.Interior.ColorIndex    # -4142 = xlNone
        @{ ColorIndex = $ci
           Color      = if ($ci -ne -4142) { try { $Cell.Interior.Color } catch { $null } } else { $null }
           Pattern    = try { $Cell.Interior.Pattern } catch { 1 } }
    } catch { $null }
}

function Restore-FaCellInteriorState {
    param($Cell, $State)
    if ($null -eq $State) { return }
    try {
        if ($State.ColorIndex -eq -4142) { # xlColorIndexNone — no fill
            $Cell.Interior.ColorIndex = -4142
        } elseif ($null -ne $State.Color) {
            $Cell.Interior.Color = $State.Color
        } else {
            $Cell.Interior.ColorIndex = $State.ColorIndex
        }
    } catch {}
}

# ---- Clear ALL temporary Excel highlights ----
function Clear-FaHighlights {
    $anyBorders  = $script:FaHighlightBorderCells.Count -gt 0
    $anyInterior = $script:FaHighlightInteriorCells.Count -gt 0
    if (-not $anyBorders -and -not $anyInterior) { return }

    # Restore borders
    foreach ($rec in $script:FaHighlightBorderCells) {
        try { Restore-FaCellBorderState -Cell $rec.Cell -State $rec.OrigBorders } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rec.Cell) } catch {}
    }
    $script:FaHighlightBorderCells.Clear()

    # Restore interiors
    foreach ($rec in $script:FaHighlightInteriorCells) {
        try { Restore-FaCellInteriorState -Cell $rec.Cell -State $rec.OrigInterior } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rec.Cell) } catch {}
    }
    $script:FaHighlightInteriorCells.Clear()

    $FaClearHighlightsBtn.IsEnabled = $false
    Set-Status 'Formula Auditor: highlights cleared'
}

# ---- Apply a thick blue outline to one cell (stores undo) ----
function Apply-FaBorderHighlight {
    param($Cell)
    $orig = Save-FaCellBorderState -Cell $Cell
    try {
        foreach ($edge in @(7, 8, 9, 10)) {
            $b = $Cell.Borders($edge)
            $b.LineStyle = 1       # xlContinuous
            $b.Weight    = 4       # xlThick
            $b.Color     = $script:FaBlueBorderColor
        }
        [void]($script:FaHighlightBorderCells.Add(@{ Cell = $Cell; OrigBorders = $orig }))
        $FaClearHighlightsBtn.IsEnabled = $true
    } catch {}
}

# ---- Apply a background fill to one cell (stores undo) ----
function Apply-FaInteriorHighlight {
    param($Cell, [long]$Color)
    $orig = Save-FaCellInteriorState -Cell $Cell
    try {
        $Cell.Interior.Color = $Color
        [void]($script:FaHighlightInteriorCells.Add(@{ Cell = $Cell; OrigInterior = $orig }))
        $FaClearHighlightsBtn.IsEnabled = $true
    } catch {}
}

# ---- Parse formula arguments (comma-split respecting nested parens) ----
function Parse-FaFormulaArgs {
    param([string]$Formula)
    # Strip = and everything up to the first '('
    $idx = $Formula.IndexOf('(')
    if ($idx -lt 0) { return @() }
    $inner = $Formula.Substring($idx + 1)
    if ($inner.EndsWith(')')) { $inner = $inner.Substring(0, $inner.Length - 1) }
    $result = @(); $depth = 0; $cur = ''
    for ($i = 0; $i -lt $inner.Length; $i++) {
        $c = $inner[$i]
        if     ($c -eq '(') { $depth++; $cur += $c }
        elseif ($c -eq ')') { $depth--; $cur += $c }
        elseif ($c -eq ',' -and $depth -eq 0) { $result += $cur.Trim(); $cur = '' }
        else   { $cur += $c }
    }
    if ($cur.Trim()) { $result += $cur.Trim() }
    return $result
}

# ---- Navigate to a cell in Excel ----
function Invoke-FaNavigateTo {
    param([string]$Address, [switch]$PushCurrent)
    $xl = Get-ExcelApp -Session Any
    if (-not $xl) { return }
    # Optionally save current cell to the back-stack
    if ($PushCurrent) {
        try {
            $entry = @{
                WbName  = [string]$xl.ActiveWorkbook.Name
                WsName  = [string]$xl.ActiveSheet.Name
                Address = [string]$xl.ActiveCell.Address($false, $false)
            }
            [void]($script:FaNavHistory.Add($entry))
            $FaBackBtn.IsEnabled = $true
        } catch {}
    }
    try {
        if ($Address -match "^'?(.+?)'?!(.+)$") {
            $sn = $Matches[1]; $ca = $Matches[2]
            $ws = try { $xl.ActiveWorkbook.Worksheets.Item($sn) } catch { $null }
            if (-not $ws) {
                Set-Status "Cannot navigate: sheet '$sn' not found in active workbook" '#E07000'
                return
            }
            $ws.Activate()
            [void]($ws.Range($ca).Select())
        } else {
            [void]($xl.ActiveSheet.Range($Address).Select())
        }
    } catch {
        Set-Status "Cannot navigate to '$Address': $_" '#E05050'; return
    }
    Invoke-FaInspect
}

# ---- Go back one step in navigation history ----
function Invoke-FaBack {
    if ($script:FaNavHistory.Count -eq 0) { $FaBackBtn.IsEnabled = $false; return }
    $xl = Get-ExcelApp -Session Any
    if (-not $xl) { return }
    $entry = $script:FaNavHistory[$script:FaNavHistory.Count - 1]
    $script:FaNavHistory.RemoveAt($script:FaNavHistory.Count - 1)
    $FaBackBtn.IsEnabled = ($script:FaNavHistory.Count -gt 0)
    try {
        $ws = try { $xl.ActiveWorkbook.Worksheets.Item($entry.WsName) } catch { $null }
        if ($ws) { $ws.Activate(); [void]($ws.Range($entry.Address).Select()) }
    } catch { Set-Status "Back navigation failed: $_" '#E05050'; return }
    Invoke-FaInspect
}

# ---- Build the precedents buttons panel ----
function Build-FaPrecedentsPanel {
    param($Cell, [string]$Formula)
    $FaPrecedentsPanel.Children.Clear()
    $refs = @()
    if ($Formula -and $Formula.StartsWith('=')) {
        try {
            $prec = $Cell.Precedents
            if ($prec) {
                for ($ai = 1; $ai -le $prec.Areas.Count; $ai++) {
                    $area = $prec.Areas.Item($ai)
                    $refs += try { $area.Address($false, $false, 1, $true) } catch { $area.Address($false, $false) }
                }
            }
        } catch {}
    }
    if ($refs.Count -gt 0) {
        $FaNoPrecedentsTxt.Visibility = 'Collapsed'
        foreach ($ref in $refs) {
            $btn = New-Object System.Windows.Controls.Button
            $btn.Content  = "$ref  →"
            $btn.Style    = $Window.FindResource('Btn')
            $btn.Margin   = [System.Windows.Thickness]::new(0, 2, 0, 2)
            $btn.Padding  = [System.Windows.Thickness]::new(8, 3, 8, 3)
            $btn.FontSize = 11
            $btn.Tag      = $ref
            $btn.ToolTip  = "Click to navigate to $ref"
            $btn.Add_Click({
                $clickedRef = [string]($args[0].Tag)
                Invoke-FaNavigateTo -Address $clickedRef -PushCurrent
            })
            [void]($FaPrecedentsPanel.Children.Add($btn))
        }
    } else {
        $FaNoPrecedentsTxt.Visibility = 'Visible'
    }
    return $refs
}

# ---- Build the precedents TreeView ----
function Build-FaPrecedentsTree {
    param($Cell, [string]$Formula)
    $FaPrecedentsTree.Items.Clear()
    if (-not ($Formula -and $Formula.StartsWith('='))) {
        $ni = New-Object System.Windows.Controls.TreeViewItem
        $ni.Header    = '(no formula — hardcoded or empty)'
        $ni.Foreground = HexBrush '#6E6E6E'
        [void]($FaPrecedentsTree.Items.Add($ni)); return
    }
    $areas = @()
    try {
        $prec = $Cell.Precedents
        if ($prec) {
            for ($ai = 1; $ai -le $prec.Areas.Count; $ai++) { $areas += $prec.Areas.Item($ai) }
        }
    } catch {}
    if ($areas.Count -eq 0) {
        $ni = New-Object System.Windows.Controls.TreeViewItem
        $ni.Header    = '(no cell precedents — e.g. =TODAY(), =NOW())'
        $ni.Foreground = HexBrush '#6E6E6E'
        [void]($FaPrecedentsTree.Items.Add($ni)); return
    }
    foreach ($area in $areas) {
        $aAddr = try { $area.Address($false, $false, 1, $true) } catch { $area.Address($false, $false) }
        $item  = New-Object System.Windows.Controls.TreeViewItem
        $item.Header     = $aAddr
        $item.IsExpanded = $true
        $item.Tag        = $aAddr
        $item.Foreground = HexBrush '#4C9FE6'
        $item.ToolTip    = "Click to navigate to $aAddr"
        # Expand up to 10 child cells
        try {
            $cnt = $area.Cells.Count
            $show = [Math]::Min($cnt, 10)
            for ($ci = 1; $ci -le $show; $ci++) {
                $cc    = $area.Cells.Item($ci)
                $cAddr = $cc.Address($false, $false)
                $cFml  = try { [string]$cc.Formula } catch { '' }
                $cVal  = try { [string]$cc.Value2 } catch { '?' }
                $label = if ($cFml -and $cFml.StartsWith('=')) { "$cAddr = $cFml" } else { "$cAddr = $cVal" }
                $child = New-Object System.Windows.Controls.TreeViewItem
                $child.Header    = $label
                $child.Tag       = $cAddr
                $child.Foreground = HexBrush '#B0B0B0'
                $child.ToolTip   = "Click to navigate to $cAddr"
                [void]($item.Items.Add($child))
            }
            if ($cnt -gt 10) {
                $more = New-Object System.Windows.Controls.TreeViewItem
                $more.Header    = "… $($cnt - 10) more cells"
                $more.Foreground = HexBrush '#6E6E6E'
                [void]($item.Items.Add($more))
            }
        } catch {}
        [void]($FaPrecedentsTree.Items.Add($item))
    }
}

# ---- Build the dependents TreeView ----
function Build-FaDependentsTree {
    param($Cell)
    $FaDependentsTree.Items.Clear()
    $areas = @()
    try {
        $deps = $Cell.Dependents
        if ($deps) {
            for ($ai = 1; $ai -le $deps.Areas.Count; $ai++) { $areas += $deps.Areas.Item($ai) }
        }
    } catch {}
    if ($areas.Count -eq 0) {
        $ni = New-Object System.Windows.Controls.TreeViewItem
        $ni.Header    = '(nothing depends on this cell)'
        $ni.Foreground = HexBrush '#6E6E6E'
        [void]($FaDependentsTree.Items.Add($ni)); return
    }
    foreach ($area in $areas) {
        $aAddr = try { $area.Address($false, $false, 1, $true) } catch { $area.Address($false, $false) }
        $item  = New-Object System.Windows.Controls.TreeViewItem
        $item.Header     = $aAddr
        $item.IsExpanded = $false
        $item.Tag        = $aAddr
        $item.Foreground = HexBrush '#50A050'
        $item.ToolTip    = "Click to navigate to $aAddr"
        [void]($FaDependentsTree.Items.Add($item))
    }
}

# ---- SUMIF / COUNTIF / AVERAGEIF visualizer ----
function Invoke-FaSumifVisualize {
    param([string]$Formula, $WorkSheet)
    # Detect function type
    $fn = $null
    if    ($Formula -match '(?i)^=SUMIFS\(')     { $fn = 'SUMIFS' }
    elseif($Formula -match '(?i)^=SUMIF\(')      { $fn = 'SUMIF' }
    elseif($Formula -match '(?i)^=COUNTIFS\(')   { $fn = 'COUNTIFS' }
    elseif($Formula -match '(?i)^=COUNTIF\(')    { $fn = 'COUNTIF' }
    elseif($Formula -match '(?i)^=AVERAGEIFS\(') { $fn = 'AVERAGEIFS' }
    elseif($Formula -match '(?i)^=AVERAGEIF\(')  { $fn = 'AVERAGEIF' }
    if (-not $fn) {
        $FaSumifEmptyTxt.Visibility = 'Visible'; $FaSumifGrid.Visibility = 'Collapsed'
        $FaSumifTitleTxt.Text = 'SUMIF Visualizer — inspect a SUMIF / SUMIFS / COUNTIF / AVERAGEIF cell'
        $FaSumifTotalTxt.Text = ''; return
    }
    $fargs = Parse-FaFormulaArgs -Formula $Formula
    try {
        $cRangeName = $null; $sRangeName = $null; $criteria = $null
        switch ($fn) {
            'SUMIF'      { $cRangeName = $fargs[0]; $criteria = $fargs[1]
                           $sRangeName = if ($fargs.Count -ge 3) { $fargs[2] } else { $fargs[0] } }
            'SUMIFS'     { $sRangeName = $fargs[0]; $cRangeName = $fargs[1]; $criteria = $fargs[2] }
            'COUNTIF'    { $cRangeName = $fargs[0]; $criteria = $fargs[1]; $sRangeName = $null }
            'COUNTIFS'   { $cRangeName = $fargs[0]; $criteria = $fargs[1]; $sRangeName = $null }
            'AVERAGEIF'  { $cRangeName = $fargs[0]; $criteria = $fargs[1]
                           $sRangeName = if ($fargs.Count -ge 3) { $fargs[2] } else { $fargs[0] } }
            'AVERAGEIFS' { $sRangeName = $fargs[0]; $cRangeName = $fargs[1]; $criteria = $fargs[2] }
        }
        $cRange = $WorkSheet.Range($cRangeName)
        $sRange = if ($sRangeName) { $WorkSheet.Range($sRangeName) } else { $null }
        # Clean criteria
        $crit = $criteria -replace '^"(.*)"$','$1'
        $op = 'eq'; $compareVal = $crit
        if ($crit -match '^(>=|<=|<>|>|<)(.+)$') { $op = $Matches[1]; $compareVal = $Matches[2] }
        $isNum = $false; $numVal = 0.0
        [void]([double]::TryParse($compareVal, [ref]$numVal) -and ($isNum = $true))
        # Build rows (cap at 500 for performance)
        $rowCount = $cRange.Rows.Count
        $maxRows  = [Math]::Min($rowCount, 500)
        $rows = @(); $total = 0.0; $matchCount = 0
        for ($r = 1; $r -le $maxRows; $r++) {
            $cCell   = $cRange.Cells.Item($r, 1)
            $cVal    = try { $cCell.Value2 } catch { $null }
            $cValStr = if ($null -eq $cVal) { '' } else { [string]$cVal }
            $sVal    = $null; $sValStr = ''
            if ($sRange) {
                $sc = $sRange.Cells.Item($r, 1)
                $sVal = try { $sc.Value2 } catch { $null }
                $sValStr = if ($null -eq $sVal) { '' } else { [string]$sVal }
            }
            $isMatch = $false
            switch ($op) {
                'eq' {
                    if ($isNum -and $null -ne $cVal) {
                        try { $isMatch = ([double]$cVal -eq $numVal) } catch {}
                    } elseif ($crit -match '\*|\?') {
                        $pat = '^' + ([regex]::Escape($crit) -replace '\\\*','.*' -replace '\\\?','.') + '$'
                        $isMatch = ($cValStr -imatch $pat)
                    } else {
                        $isMatch = ($cValStr -ieq $compareVal)
                    }
                }
                '>=' { if ($null -ne $cVal) { try { $isMatch = ([double]$cVal -ge [double]$compareVal) } catch {} } }
                '<=' { if ($null -ne $cVal) { try { $isMatch = ([double]$cVal -le [double]$compareVal) } catch {} } }
                '<>' { $isMatch = ($cValStr -ine $compareVal) }
                '>'  { if ($null -ne $cVal) { try { $isMatch = ([double]$cVal -gt [double]$compareVal) } catch {} } }
                '<'  { if ($null -ne $cVal) { try { $isMatch = ([double]$cVal -lt [double]$compareVal) } catch {} } }
            }
            if ($isMatch) {
                $matchCount++
                if ($sRange -and $null -ne $sVal) { try { $total += [double]$sVal } catch {} }
                # Highlight matching rows in Excel (criteria=blue, sum=green)
                try {
                    $cCellRef = $cRange.Cells.Item($r, 1)
                    Apply-FaInteriorHighlight -Cell $cCellRef -Color $script:FaBlueFillColor
                } catch {}
                if ($sRange) {
                    try {
                        $sCellRef = $sRange.Cells.Item($r, 1)
                        Apply-FaInteriorHighlight -Cell $sCellRef -Color $script:FaGreenFillColor
                    } catch {}
                }
            }
            $absRow = try { $cRange.Cells.Item($r, 1).Row } catch { $r }
            $rows += [PSCustomObject]@{
                RowNum      = $absRow
                IsMatch     = $isMatch
                MatchSymbol = if ($isMatch) { '✓' } else { '' }
                CriteriaVal = $cValStr
                SumVal      = if ($isMatch) { $sValStr } else { '' }
            }
        }
        $suffix = if ($fn -in @('SUMIF','SUMIFS')) { "SUM = $total" }
                  elseif ($fn -in @('COUNTIF','COUNTIFS')) { "COUNT = $matchCount" }
                  else {
                      $avg = if ($matchCount -gt 0) { [Math]::Round($total/$matchCount,4) } else { 0 }
                      "AVG = $avg"
                  }
        $truncNote = if ($maxRows -lt $rowCount) { " (first $maxRows of $rowCount rows)" } else { " ($rowCount rows)" }
        $FaSumifTitleTxt.Text = "$fn($cRangeName, $criteria, …)$truncNote"
        $FaSumifTotalTxt.Text = "$suffix  •  $matchCount match(es)"
        $FaSumifGrid.ItemsSource = $rows
        $FaSumifEmptyTxt.Visibility = 'Collapsed'
        $FaSumifGrid.Visibility = 'Visible'
    } catch {
        $FaSumifTitleTxt.Text = "$fn — parse error: $_"
        $FaSumifTotalTxt.Text = ''
        $FaSumifEmptyTxt.Visibility = 'Visible'
        $FaSumifGrid.Visibility = 'Collapsed'
    }
}

# ---- Range Audit ----
function Invoke-FaAuditRange {
    $xl = Get-ExcelApp -Session Any
    if (-not $xl) { Set-Status 'Excel not connected' '#E05050'; return }
    try {
        $sel  = $xl.Selection
        $ws   = $xl.ActiveSheet
        $cnt  = $sel.Cells.Count
        if ($cnt -gt 5000) {
            $FaRangeDetailTxt.Text = "Selection too large ($cnt cells). Please select up to 5,000 cells."
            $FaRangeDetailTxt.Foreground = HexBrush '#E07000'; return
        }
        $formulaCells  = @(); $hardcodedCells = @()
        $crossSheetRef = @(); $externalRef    = @()
        $sigMap = @{}
        for ($i = 1; $i -le $cnt; $i++) {
            $c    = $sel.Cells.Item($i)
            $addr = $c.Address($false, $false)
            $fml  = try { [string]$c.Formula } catch { '' }
            $val  = try { $c.Value2 } catch { $null }
            if ($fml -and $fml.StartsWith('=')) {
                $formulaCells += $addr
                $sig = $fml -replace '\$?[A-Za-z]+\$?\d+', 'REF' -replace '\d+', 'N'
                $sigMap[$sig] = ($sigMap[$sig] -as [int]) + 1
                if ($fml -match "(?i)[A-Za-z0-9_' ]+!") { $crossSheetRef += $addr }
                if ($fml -match '\[.+\]')                { $externalRef   += $addr }
                # Temporarily highlight formula cells with thick blue outline
                $cellRef = $sel.Cells.Item($i)
                Apply-FaBorderHighlight -Cell $cellRef
            } elseif ($null -ne $val -and [string]$val -ne '') {
                $hardcodedCells += $addr
            }
        }
        $total = $formulaCells.Count + $hardcodedCells.Count
        $fPct  = if ($total -gt 0) { "$($formulaCells.Count) ($([Math]::Round($formulaCells.Count*100/$total,0))%)" } else { '0' }
        $hPct  = if ($total -gt 0) { "$($hardcodedCells.Count) ($([Math]::Round($hardcodedCells.Count*100/$total,0))%)" } else { '0' }
        $FaRangeAddrTxt.Text      = $sel.Address($false, $false)
        $FaRangeFormulaTxt.Text   = $fPct
        $FaRangeHardcodedTxt.Text = $hPct
        $FaRangeExternalTxt.Text  = if ($externalRef.Count -gt 0) { "⚠ $($externalRef.Count) external ref(s)" } else { '' }
        $details = @()
        if ($crossSheetRef.Count -gt 0) { $details += "Cross-sheet refs: $($crossSheetRef -join ', ')" }
        if ($externalRef.Count   -gt 0) { $details += "External workbook refs: $($externalRef -join ', ')" }
        if ($sigMap.Count -gt 1) {
            $sorted = $sigMap.GetEnumerator() | Sort-Object Value -Descending
            $majority = ($sorted | Select-Object -First 1).Key
            foreach ($m in ($sorted | Select-Object -Skip 1)) {
                $details += "Inconsistent formula ($($m.Value) cell(s) differ): $($m.Key)"
            }
        }
        if ($details.Count -eq 0) {
            $details += "No issues detected.  $($formulaCells.Count) formula cell(s), $($hardcodedCells.Count) hardcoded, $($cnt - $formulaCells.Count - $hardcodedCells.Count) empty."
        }
        $FaRangeDetailTxt.Text       = ($details -join "`n")
        $FaRangeDetailTxt.Foreground = HexBrush '#B0B0B0'
        Set-Status "Range audit complete: $($sel.Address($false,$false))  ($cnt cells, $($formulaCells.Count) formulas highlighted in blue)"
    } catch {
        $FaRangeDetailTxt.Text       = "Range audit error: $_"
        $FaRangeDetailTxt.Foreground = HexBrush '#E05050'
        Set-Status "Range audit error: $_" '#E05050'
    }
}

# ---- Main inspect function ----
function Invoke-FaInspect {
    $xl = Get-ExcelApp -Session Any
    if (-not $xl) {
        $FaActiveCellTxt.Text = '(Excel not connected)'
        $FaFormulaTxt.Text    = '—'
        return
    }
    $cell = $null; $ws = $null; $wb = $null
    try { $wb = $xl.ActiveWorkbook; $ws = $xl.ActiveSheet; $cell = $xl.ActiveCell } catch { return }
    if (-not $cell) { return }
    $addr   = try { [string]$cell.Address($false, $false) } catch { '?' }
    $wsName = try { [string]$ws.Name } catch { '?' }
    $wbName = try { [string]$wb.Name } catch { '?' }
    # Top bar
    $FaActiveCellTxt.Text = "$wsName!$addr"
    $formula = try { [string]$cell.Formula } catch { '' }
    $value   = try { $cell.Value2 } catch { $null }
    $FaFormulaTxt.Text = if ($formula -and $formula.StartsWith('=')) { $formula }
                         elseif ($null -eq $value -or [string]$value -eq '') { '(empty)' }
                         else { [string]$value }
    # Inspector card
    $cellType = if ($formula -and $formula.StartsWith('=')) { 'Formula' }
                elseif ($null -eq $value -or [string]$value -eq '') { 'Empty' }
                else { 'Hardcoded' }
    $FaCellTypeTxt.Text  = $cellType
    $FaCellValueTxt.Text = if ($null -eq $value) { '(null)' } else { [string]$value }
    $FaCellSheetTxt.Text = $wsName
    # Precedents panel + tree
    Build-FaPrecedentsPanel -Cell $cell -Formula $formula
    Build-FaPrecedentsTree  -Cell $cell -Formula $formula
    # Dependents
    $depCount = 0
    try { $deps = $cell.Dependents; if ($deps) { $depCount = $deps.Cells.Count } } catch {}
    $FaDependentCountTxt.Text = if ($depCount -gt 0) { "$depCount cell(s)" } else { 'none' }
    Build-FaDependentsTree -Cell $cell
    # SUMIF visualizer
    if ($formula -match '(?i)^=(SUM|COUNT|AVERAGE)IFS?\(') {
        Invoke-FaSumifVisualize -Formula $formula -WorkSheet $ws
    } else {
        $FaSumifEmptyTxt.Visibility = 'Visible'; $FaSumifGrid.Visibility = 'Collapsed'
        $FaSumifTitleTxt.Text = 'SUMIF Visualizer — inspect a SUMIF / SUMIFS / COUNTIF / AVERAGEIF cell'
        $FaSumifTotalTxt.Text = ''
    }
}

# ================================================================
#  FORMULA AUDITOR — 500 ms polling timer
# ================================================================
$script:FaTimer          = New-Object System.Windows.Threading.DispatcherTimer
$script:FaTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:FaTimer.Add_Tick({
    if ($MainTabs.SelectedIndex -ne 8) { return }
    if (-not $FaAutoBox.IsChecked)     { return }
    $xl = Get-ExcelApp -Session Any
    if (-not $xl) { $FaActiveCellTxt.Text = '(Excel not connected)'; return }
    try {
        $cell = $xl.ActiveCell
        if (-not $cell) { return }
        $sig = "$([string]$xl.ActiveWorkbook.Name)!$([string]$xl.ActiveSheet.Name)!$([string]$cell.Address($false,$false))"
        if ($sig -eq $script:FaLastCellSig) { return }
        $script:FaLastCellSig = $sig
        Invoke-FaInspect
    } catch {}
})
$script:FaTimer.Start()

# ================================================================
#  FORMULA AUDITOR — event handlers
# ================================================================
$FaInspectBtn.Add_Click({ Invoke-FaInspect })
$FaBackBtn.Add_Click({ Invoke-FaBack })
$FaAuditRangeBtn.Add_Click({ Invoke-FaAuditRange })
$FaClearHighlightsBtn.Add_Click({ Clear-FaHighlights })

# Tree node click — navigate to selected cell
$FaPrecedentsTree.Add_SelectedItemChanged({
    $node = $FaPrecedentsTree.SelectedItem
    if ($node -and $node.Tag) { Invoke-FaNavigateTo -Address ([string]$node.Tag) -PushCurrent }
})
$FaDependentsTree.Add_SelectedItemChanged({
    $node = $FaDependentsTree.SelectedItem
    if ($node -and $node.Tag) { Invoke-FaNavigateTo -Address ([string]$node.Tag) -PushCurrent }
})

# Clear highlights automatically when leaving the Formula Auditor tab
$MainTabs.Add_SelectionChanged({
    if ($MainTabs.SelectedIndex -ne 8) { Clear-FaHighlights }
})

# Mark Tab 0 (Clipboard) + Tab 1 (Macros) + Tab 4 (Templates) + Tab 8 (Formula Auditor — timer-driven) as already loaded
$script:TabsLoaded[0] = $true
$script:TabsLoaded[1] = $true
$script:TabsLoaded[4] = $true
$script:TabsLoaded[5]  = $true
$script:TabsLoaded[8] = $true

Set-Status 'MacroHub v3.2 ready — Alt+1–9 tabs  |  Ctrl+Tab cycle  |  Ctrl+F search'

# Check for missed scheduled tasks AFTER window is visible
$Window.Add_ContentRendered({
    Invoke-MissedTaskCheck
})

# -- Dismiss keytips when window loses focus --
$Window.Add_Deactivated({ if ($script:KeyTipsActive) { Hide-KeyTips } })

# -- Formula Auditor: stop timer and clear Excel highlights when window closes --
$Window.Add_Closing({
    if ($script:FaTimer) { $script:FaTimer.Stop() }
    try { Clear-FaHighlights } catch {}
})

# ================================================================
#  SHOW WINDOW
# ================================================================
# ShowDialog() is a blocking call: the script waits here until the window closes.
$Window.ShowDialog()

}  # END Start-MacroHub

# ================================================================
#  LAUNCH
# ================================================================
# Script entry point: simply calls Start-MacroHub which builds the UI and blocks.
Start-MacroHub

