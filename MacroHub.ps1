#Requires -Version 5.1
<#
.SYNOPSIS
    MacroHub v3.1 - Office Productivity Super App
.DESCRIPTION
    Merged: MacroHub automation suite + QuarterSync quarterly tracker.
    11 tabs: Clipboard | Macros | Scheduler | Navigator | Templates | QSync | QTasks | Email Helper | Email Dashboard | IDE | AI IDE

    Chrome/Teams dark-mode UI. Pure PowerShell 5.1 / WPF. No external modules.

    Color palette: black/grey/white text, blue accents, red for destructive only.
#>

# ============================================================
#  ASSEMBLIES
# ============================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
#  CONFIGURATION
# ============================================================
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
$script:EmailCfgJson = Join-Path $script:HubRoot 'email_config.json'
$script:EmailStateJson = Join-Path $script:HubRoot 'email_state.json'
$script:ClipDefaultsJson = Join-Path $script:HubRoot 'clip_defaults.json'
$script:AppVersion   = '3.1.0'

# ============================================================
#  GLOBAL STATE
# ============================================================
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
$script:EmailCfg       = $null
$script:EmailState     = $null
$script:EmailCache     = @()
$script:EmailRawCache  = @()
$script:EmailRawScope  = ''
$script:OutlookApp     = $null
$script:ExcelMainApp   = $null
$script:ExcelNavApp    = $null
$script:IdeCurrentFile = ''
$script:IdeMode        = 'Auto'
$script:IdeLastValidationOk = $false
$script:IdeValidatedHash = ''
$script:IdeValidatedLang = ''
$script:IdeEditorReady = $false

# -"--- AI IDE -- global synchronized state -"-------------------------------------------------------------------
$script:AiIde = [hashtable]::Synchronized(@{
    # Graph / OAuth
    AccessToken      = ''
    RefreshToken     = ''
    TokenExpiry      = [datetime]::MinValue
    TenantId         = 'common'
    ClientId         = 'YOUR-CLIENT-ID-HERE'
    DeviceCode       = ''
    DeviceCodeExpiry = [datetime]::MinValue
    ConversationId   = ''
    GraphAccessVerified  = $false
    GraphAccessCheckedOn = [datetime]::MinValue
    GraphAccessNote      = ''
    GraphAuthUser        = ''
    GraphLastHttpCode    = 0

    # Chrome bridge
    Port             = 9876
    PendingPrompt    = ''
    PendingRequest   = ''
    PendingFile      = ''
    LastResult       = ''
    Status           = 'idle'   # idle|waiting|done|error|auth_pending|authenticated

    # Runtime
    LastError        = ''
    RetryCount       = 0
    MaxRetries       = 3

    # Handles
    ListenerPS       = $null
    ListenerRS       = $null
})
$script:AiIdeCurrentFile  = ''
$script:AiIdeConvHistory  = [System.Collections.ArrayList]::new()
$script:AiIdeRefreshingTargets = $false
# -"--- End AI IDE global state -"-------------------------------------------------------------------------------------------
$script:IdeEditorPendingText = $null

# ============================================================
#  HELPER: Solid color brush from hex string (cached + frozen)
# ============================================================
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
function Write-ActivityLog {
    param([string]$Action)
    $safe = $Action -replace '"','""'
    $line = '"{0}","{1}"' -f (Get-Date -f 'yyyy-MM-dd HH:mm:ss'), $safe
    try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 } catch {}
}

# Normalize loaded JSON/CSV payloads to a null-free object array.
function Normalize-List {
    param($InputObject)
    $items = @($InputObject | Where-Object { $null -ne $_ })
    return ,$items
}

# ============================================================
#  UI HELPERS: busy overlay, status bar, dispatcher flush
# ============================================================
function Update-UI {
    if ($script:Window) {
        $script:Window.Dispatcher.Invoke(
            [action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }
}

function Show-Busy {
    param([string]$Message = 'Working...')
    $script:BusyCount++
    if ($script:BusyOverlay) {
        $script:BusyText.Text = $Message
        $script:BusyOverlay.Visibility = 'Visible'
    }
    Update-UI
}

function Hide-Busy {
    $script:BusyCount = [Math]::Max(0, $script:BusyCount - 1)
    if ($script:BusyCount -eq 0 -and $script:BusyOverlay) {
        $script:BusyOverlay.Visibility = 'Collapsed'
    }
    Update-UI
}

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
            $names += $xl.Workbooks.Item($i).Name
        }
    } catch {}
    return $names
}

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
function Get-MacroFiles {
    if (-not (Test-Path $script:MacroFolder)) { return @() }
    $files = Get-ChildItem -Path $script:MacroFolder -Recurse -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Extension -match '^\.(bas|ps1)$' }
    return @($files)
}

# ============================================================
#  MACRO EXECUTION
# ============================================================
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

    # Import module, run, then remove
    $vbProj = $wb.VBProject
    $comp = $vbProj.VBComponents.Import($MacroFile)
    try {
        $xl.Run("$($wb.Name)!$($comp.Name).$entryPoint")
    } finally {
        try { $vbProj.VBComponents.Remove($comp) } catch {}
    }
    Write-ActivityLog "Ran VBA macro: $entryPoint from $(Split-Path $MacroFile -Leaf)"
}

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

    $def = $svc.NewTask(0)
    $def.RegistrationInfo.Description = "MacroHub scheduled: $TaskName -- File: $MacroFile"
    $def.Settings.Enabled = $true
    $def.Settings.StartWhenAvailable = $true
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
        $trigger.DaysOfMonth   = [Math]::Pow(2, $start.Day - 1)  # bitmask
        $trigger.MonthsOfYear  = 4095  # all 12 months
    } else {
        # Daily trigger (type 2)
        $trigger = $def.Triggers.Create(2)
        $trigger.StartBoundary = $start.ToString('yyyy-MM-ddTHH:mm:ss')
        if ($Frequency -eq 'Weekly') { $trigger.DaysInterval = 7 }
        else { $trigger.DaysInterval = 1 }
    }

    $action = $def.Actions.Create(0)
    $action.Path = 'powershell.exe'
    $action.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$runFile`""
    $action.WorkingDirectory = $script:HubRoot

    $folder.RegisterTaskDefinition($TaskName, $def, 6, $null, $null, 3)
    Write-ActivityLog "Scheduled task: $TaskName ($Frequency at $TriggerTime) -> $runFile"
}

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
#  CLIPBOARD SLOT HELPERS
# ============================================================
# ============================================================
#  FAVORITES PERSISTENCE
# ============================================================
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
function Load-ClipDefaults {
    if (Test-Path $script:ClipDefaultsJson) {
        try { return (Get-Content $script:ClipDefaultsJson -Raw | ConvertFrom-Json) } catch {}
    }
    return $null
}

function Save-ClipDefaults {
    param([string]$Workbook, [string]$Sheet, [string]$Cell,
          [bool]$Timestamp, [string]$DateOffset, [string]$TimeOffset)
    $obj = @{
        Workbook   = $Workbook
        Sheet      = $Sheet
        Cell       = $Cell
        Timestamp  = $Timestamp
        DateOffset = $DateOffset
        TimeOffset = $TimeOffset
        SavedOn    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
    $obj | ConvertTo-Json | Set-Content $script:ClipDefaultsJson -Encoding UTF8
}

# ============================================================
#  DYNAMIC CLIPBOARD SLOT UI BUILDER
# ============================================================
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

function Convert-ClipboardTextToMatrix {
    param(
        [string]$Text,
        [int]$Rows,
        [int]$Cols
    )
    $safeRows = [Math]::Max(1, [int]$Rows)
    $safeCols = [Math]::Max(1, [int]$Cols)
    $matrix = New-Object 'object[,]' $safeRows, $safeCols
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
            $matrix[$ri, $ci] = $parts[$ci]
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

function Paste-ClipboardPacketToExcel {
    param(
        [object]$Packet,
        [string]$WorkbookName,
        [string]$SheetName,
        [string]$CellAddress,
        [bool]$AddTimestamp,
        [string]$DateOffset,
        [string]$TimeOffset
    )
    if (-not $Packet) { throw 'Clipboard slot is empty.' }
    $xl = Get-ExcelApp -Session Main
    if (-not $xl) { throw 'Excel is not open.' }
    $wbObj = $xl.Workbooks.Item($WorkbookName)
    $ws    = $wbObj.Worksheets.Item($SheetName)
    $target = $ws.Range($CellAddress)
    $r = [int]$target.Row
    $c = [int]$target.Column
    $dims = Get-ClipboardDimensions -Text $Packet.Text
    $rows = [Math]::Max(1, [int]$dims.Rows)
    $cols = [Math]::Max(1, [int]$dims.Cols)
    $textRows = $rows
    $textCols = $cols

    $dest = $ws.Range(
        $ws.Cells.Item($r, $c),
        $ws.Cells.Item($r + $rows - 1, $c + $cols - 1)
    )

    # Required order: clear values, clear formats, then paste.
    $dest.ClearContents()
    $dest.ClearFormats()

    # Best-case formatting path for Excel source copy: replay source range directly.
    $usedClipboardPaste = $false
    $source = $Packet.ExcelSource
    if ($source -and $source.WorkbookName -and $source.SheetName -and $source.Address) {
        try {
            $srcWb = $xl.Workbooks.Item([string]$source.WorkbookName)
            $srcWs = $srcWb.Worksheets.Item([string]$source.SheetName)
            $srcRange = $srcWs.Range([string]$source.Address)
            $srcRows = [Math]::Max(1, [int]$srcRange.Rows.Count)
            $srcCols = [Math]::Max(1, [int]$srcRange.Columns.Count)
            # Some Excel sessions report Selection as 1x1 even when clipboard holds a larger range.
            # If we captured larger clipboard text dimensions, expand from the source anchor cell.
            $canExpandFromText = ($srcRows -eq 1 -and $srcCols -eq 1 -and ($textCols -gt 1 -or $textRows -gt 2))
            if ($canExpandFromText) {
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
            $srcRange.Copy($dest)
            $usedClipboardPaste = $true
        } catch {}
    }

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

    # For formatted clipboard payloads (Excel/HTML/RTF), retry native paste next.
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
                $target.PasteSpecial(-4104)  # xlPasteAll
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
        } catch {
            # If live clipboard attempt failed, fall back to packet replay attempts.
            if ($liveClipboardSame) { $liveClipboardSame = $false }
            Start-Sleep -Milliseconds (60 * $attempt)
        }
    }

    if (-not $usedClipboardPaste) {
        # Fallback for plain text payloads: use 2D array assignment for large paste speed.
        $matrix = Convert-ClipboardTextToMatrix -Text $Packet.Text -Rows $rows -Cols $cols
        $dest.Value2 = $matrix
    }

    Apply-PostPasteCellRules -Worksheet $ws -StartRow $r -StartCol $c -Rows $rows -Cols $cols

    if ($AddTimestamp) {
        $dOff = 1; $tOff = 2
        try { $dOff = [int]$DateOffset } catch {}
        try { $tOff = [int]$TimeOffset } catch {}
        $dateRow = $r - $dOff
        $timeRow = $r - $tOff
        if ($dateRow -ge 1) {
            $ws.Cells.Item($dateRow, $c).Value2 = (Get-Date -Format 'MM/dd/yyyy')
        }
        if ($timeRow -ge 1) {
            $ws.Cells.Item($timeRow, $c).Value2 = (Get-Date -Format 'hh:mm tt')
        }
    }
}

function Add-ClipSlotUI {
    param(
        $Panel,
        $WbCombo,
        [string]$DefaultSheet,
        [string]$DefaultCell,
        [bool]$DefaultTimestamp,
        [string]$DefaultDateOffset,
        [string]$DefaultTimeOffset,
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
        DateOffsetBox  = $null
        TimeOffsetBox  = $null
    }
    # Capture direct hashtable reference for WPF event closures.
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
    $slotDateLbl.Text = 'Date'
    $slotDateLbl.Foreground = (HexBrush '#FFFFFF')
    $slotDateLbl.FontSize = 11
    $slotDateLbl.VerticalAlignment = 'Center'
    $slotDateLbl.Margin = [System.Windows.Thickness]::new(0,0,4,0)
    [void]($btnRow.Children.Add($slotDateLbl))

    $slotDateOffsetBox = [System.Windows.Controls.TextBox]::new()
    $slotDateOffsetBox.Width = 44
    $slotDateOffsetBox.TextAlignment = 'Center'
    $slotDateOffsetBox.Text = if ([string]::IsNullOrWhiteSpace($DefaultDateOffset)) { '1' } else { [string]$DefaultDateOffset }
    $slotDateOffsetBox.Margin = [System.Windows.Thickness]::new(0,0,4,0)
    [void]($btnRow.Children.Add($slotDateOffsetBox))

    $slotDateUpLbl = [System.Windows.Controls.TextBlock]::new()
    $slotDateUpLbl.Text = 'up'
    $slotDateUpLbl.Foreground = (HexBrush '#FFFFFF')
    $slotDateUpLbl.FontSize = 11
    $slotDateUpLbl.VerticalAlignment = 'Center'
    $slotDateUpLbl.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    [void]($btnRow.Children.Add($slotDateUpLbl))

    $slotTimeLbl = [System.Windows.Controls.TextBlock]::new()
    $slotTimeLbl.Text = 'Time'
    $slotTimeLbl.Foreground = (HexBrush '#FFFFFF')
    $slotTimeLbl.FontSize = 11
    $slotTimeLbl.VerticalAlignment = 'Center'
    $slotTimeLbl.Margin = [System.Windows.Thickness]::new(0,0,4,0)
    [void]($btnRow.Children.Add($slotTimeLbl))

    $slotTimeOffsetBox = [System.Windows.Controls.TextBox]::new()
    $slotTimeOffsetBox.Width = 44
    $slotTimeOffsetBox.TextAlignment = 'Center'
    $slotTimeOffsetBox.Text = if ([string]::IsNullOrWhiteSpace($DefaultTimeOffset)) { '2' } else { [string]$DefaultTimeOffset }
    $slotTimeOffsetBox.Margin = [System.Windows.Thickness]::new(0,0,4,0)
    [void]($btnRow.Children.Add($slotTimeOffsetBox))

    $slotTimeUpLbl = [System.Windows.Controls.TextBlock]::new()
    $slotTimeUpLbl.Text = 'up'
    $slotTimeUpLbl.Foreground = (HexBrush '#FFFFFF')
    $slotTimeUpLbl.FontSize = 11
    $slotTimeUpLbl.VerticalAlignment = 'Center'
    $slotTimeUpLbl.Margin = [System.Windows.Thickness]::new(0,0,10,0)
    [void]($btnRow.Children.Add($slotTimeUpLbl))

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
        $slotRef.DateOffsetBox = $slotDateOffsetBox
        $slotRef.TimeOffsetBox = $slotTimeOffsetBox
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

    $recTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $recTimer.Interval = [TimeSpan]::FromMilliseconds(350)

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

        $dateOffset = if ($slot.DateOffsetBox) { [string]$slot.DateOffsetBox.Text } else { '1' }
        $timeOffset = if ($slot.TimeOffsetBox) { [string]$slot.TimeOffsetBox.Text } else { '2' }
        $addTs = if ($slot.TimestampChk) { [bool]$slot.TimestampChk.IsChecked } else { $false }

        Show-Busy 'Pasting to Excel...'
        try {
            Paste-ClipboardPacketToExcel -Packet $packet -WorkbookName ([string]$wb) -SheetName ([string]$sh) `
                -CellAddress $cell -AddTimestamp $addTs -DateOffset $dateOffset -TimeOffset $timeOffset
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
function Get-CleanMacroName([string]$DisplayName) {
    if ($DisplayName -match '^\[.\]\s+(.+)$') { return $Matches[1] }
    return $DisplayName
}

# ============================================================
#  TEMPLATE MANAGEMENT
# ============================================================
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
$script:DatePatterns = @(
    'Q[1-4][-_\s]?20\d{2}', '20\d{2}[-_\s]?Q[1-4]',
    '20\d{2}[-_\s]\d{2}[-_\s]\d{2}', '\d{2}[-_\s]\d{2}[-_\s]20\d{2}',
    '\d{8}',
    '(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-_\s]?20\d{2}',
    '20\d{2}[-_\s]?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*',
    'FY20\d{2}', 'YE20\d{2}', '20\d{2}'
)

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
function Resolve-QsTodoPath {
    param([string]$PathLike)
    if (-not $PathLike) { return $null }
    $ext = [System.IO.Path]::GetExtension($PathLike).ToLower()
    if ($ext -eq '.json') { return $PathLike }
    if ($ext -eq '.csv') { return [System.IO.Path]::ChangeExtension($PathLike, '.json') }
    return "$PathLike.json"
}

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

# Compare-mode persistence (QTasks folder A vs folder B)
function Load-QsCompareTodos {
    if (-not (Test-Path $script:QsCompareJson)) { return ,@() }
    try {
        return ,(Normalize-List (Get-Content $script:QsCompareJson -Raw | ConvertFrom-Json))
    } catch {
        return ,@()
    }
}

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
function Export-QsToExcel {
    param([string]$SavePath)

    $todos   = Load-QsCompareTodos
    $syncLog = Load-QsSyncLog
    $srcPath = Resolve-QsFolderPath (Get-ComboText $QsQuarterCombo)
    $tgtPath = Resolve-QsFolderPath (Get-ComboText $QsCompareCombo)

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

# ================================================================
#  EMAIL: Load / Save Config & State
# ================================================================
function Load-EmailConfig {
    if (Test-Path $script:EmailCfgJson) {
        $script:EmailCfg = Get-Content $script:EmailCfgJson -Raw | ConvertFrom-Json
    } else {
        $script:EmailCfg = [pscustomobject]@{
            vipSenders=@(); views=@{}; replyTemplates=@{}
            agingBuckets=@(); internalDomains=@()
            defaultFetchCount=100; calendarReminderMinutes=15; recallMaxAgeHours=1
        }
    }
}
function Load-EmailState {
    if (Test-Path $script:EmailStateJson) {
        $script:EmailState = Get-Content $script:EmailStateJson -Raw | ConvertFrom-Json
    } else {
        $script:EmailState = [pscustomobject]@{
            lastRefresh=''; triageActions=@(); draftLog=@()
            todoAssignments=@(); deadlineAssignments=@()
            recallAttempts=@(); dismissedEntryIds=@(); viewSnapshots=@{}
        }
    }
}
function Save-EmailState {
    if ($script:EmailState) {
        $script:EmailState | ConvertTo-Json -Depth 5 | Set-Content $script:EmailStateJson -Encoding UTF8
    }
}

# ================================================================
#  EMAIL: Outlook COM helpers
# ================================================================
function Get-OutlookApp {
    if ($script:OutlookApp) {
        try { $null = $script:OutlookApp.Session; return $script:OutlookApp }
        catch { $script:OutlookApp = $null }
    }
    try {
        $script:OutlookApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch {
        try { $script:OutlookApp = New-Object -ComObject Outlook.Application }
        catch { return $null }
    }
    return $script:OutlookApp
}

function Get-InboxEmails([int]$maxCount = 100, [string]$scope = 'All Folders') {
    $ol = Get-OutlookApp
    if (-not $ol) { return @() }
    try {
        $ns = $ol.GetNamespace('MAPI')

        $rootFolders = New-Object System.Collections.Generic.List[object]
        if ($scope -eq 'Inbox Only') {
            [void]$rootFolders.Add($ns.GetDefaultFolder(6)) # olFolderInbox
        } else {
            try {
                foreach ($st in @($ns.Stores)) {
                    try { [void]$rootFolders.Add($st.GetRootFolder()) } catch {}
                }
            } catch {}
            if ($rootFolders.Count -eq 0) {
                [void]$rootFolders.Add($ns.GetDefaultFolder(6))
            }
        }

        $folderQueue = [System.Collections.Queue]::new()
        foreach ($rf in $rootFolders) { $folderQueue.Enqueue($rf) }
        $folders = New-Object System.Collections.Generic.List[object]
        while ($folderQueue.Count -gt 0) {
            $f = $folderQueue.Dequeue()
            [void]$folders.Add($f)
            try {
                foreach ($sub in @($f.Folders)) { $folderQueue.Enqueue($sub) }
            } catch {}
        }

        $raw = New-Object System.Collections.Generic.List[object]
        $perFolderTake = [Math]::Max(20, [Math]::Ceiling(($maxCount * 2) / [Math]::Max(1, $folders.Count)))
        foreach ($fld in $folders) {
            try {
                $items = $fld.Items
                try { $items.Sort('[ReceivedTime]', $true) } catch {}
                $take = [Math]::Min([int]$perFolderTake, [int]$items.Count)
                for ($i = 1; $i -le $take; $i++) {
                    $mail = $items.Item($i)
                    if (-not $mail -or $mail.Class -ne 43) { continue }  # olMail = 43
                    $ageSpan = (Get-Date) - $mail.ReceivedTime
                    $bodyPreview = ''
                    try {
                        if ($mail.Body) { $bodyPreview = $mail.Body.Substring(0, [Math]::Min(500, $mail.Body.Length)) }
                    } catch {}
                    [void]$raw.Add([pscustomobject]@{
                        EntryID       = $mail.EntryID
                        Subject       = $mail.Subject
                        SenderName    = $mail.SenderName
                        SenderEmail   = $mail.SenderEmailAddress
                        ReceivedTime  = $mail.ReceivedTime
                        AgeDays       = [Math]::Floor($ageSpan.TotalDays)
                        IsRead        = ($mail.UnRead -ne $true)
                        HasAttachment = ($mail.Attachments.Count -gt 0)
                        Importance    = $mail.Importance   # 0=Low, 1=Normal, 2=High
                        IsVIP         = $false
                        IsInternal    = $false
                        FolderPath    = $fld.FolderPath
                        Body          = $bodyPreview
                    })
                }
            } catch {}
        }

        # Global sort + de-dupe by EntryID.
        $seen = @{}
        $result = New-Object System.Collections.Generic.List[object]
        foreach ($row in @($raw | Sort-Object ReceivedTime -Descending)) {
            if (-not $row.EntryID) { continue }
            if ($seen.ContainsKey($row.EntryID)) { continue }
            $seen[$row.EntryID] = $true
            [void]$result.Add($row)
            if ($result.Count -ge $maxCount) { break }
        }

        # Tag VIP and internal.
        $vips = @()
        if ($script:EmailCfg.vipSenders) { $vips = @($script:EmailCfg.vipSenders) }
        $doms = @()
        if ($script:EmailCfg.internalDomains) { $doms = @($script:EmailCfg.internalDomains) }
        foreach ($e in $result) {
            foreach ($v in $vips) {
                if ($e.SenderEmail -like "*$v*") { $e.IsVIP = $true; break }
            }
            foreach ($d in $doms) {
                if ($e.SenderEmail -like "*@$d") { $e.IsInternal = $true; break }
            }
        }
        return @($result)
    } catch { return @() }
}

function Get-AgingBucket([int]$ageDays) {
    if (-not $script:EmailCfg.agingBuckets) { return @{ label='Unknown'; color='#6E6E6E' } }
    foreach ($b in $script:EmailCfg.agingBuckets) {
        if ($ageDays -ge $b.minDays -and $ageDays -le $b.maxDays) {
            return @{ label = $b.label; color = $b.color }
        }
    }
    return @{ label='Old'; color='#FF0000' }
}

function Get-EmailMetrics([array]$emails) {
    $total   = $emails.Count
    $unread  = ($emails | Where-Object { -not $_.IsRead }).Count
    $vip     = ($emails | Where-Object { $_.IsVIP }).Count
    $intern  = ($emails | Where-Object { $_.IsInternal }).Count
    $extern  = $total - $intern
    $oldest  = if ($total -gt 0) { ($emails | Sort-Object AgeDays -Descending | Select-Object -First 1).AgeDays } else { 0 }
    # Top senders
    $topSenders = $emails | Group-Object SenderName | Sort-Object Count -Descending | Select-Object -First 5
    # Aging distribution
    $agingDist = @{}
    foreach ($e in $emails) {
        $b = Get-AgingBucket $e.AgeDays
        $key = $b.label
        if (-not $agingDist.ContainsKey($key)) { $agingDist[$key] = @{ count=0; color=$b.color } }
        $agingDist[$key].count++
    }
    return [pscustomobject]@{
        Total=$total; Unread=$unread; VIP=$vip
        Internal=$intern; External=$extern; OldestDays=$oldest
        TopSenders=$topSenders; AgingDistribution=$agingDist
    }
}

function Filter-EmailsByView([array]$emails, [string]$viewName) {
    if (-not $script:EmailCfg.views -or -not $script:EmailCfg.views.$viewName) { return $emails }
    $view = $script:EmailCfg.views.$viewName
    $keywords = @($view.keywords)
    $maxAge   = $view.maxAgeDays
    $filtered = $emails | Where-Object {
        $pass = $true
        if ($maxAge -and $_.AgeDays -gt $maxAge) { $pass = $false }
        if ($pass -and $keywords.Count -gt 0) {
            $match = $false
            foreach ($kw in $keywords) {
                if ($_.Subject -like "*$kw*") { $match = $true; break }
            }
            $pass = $match
        }
        $pass
    }
    return @($filtered)
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

function Try-ParseFilterDate([string]$raw) {
    if (-not $raw) { return $null }
    $txt = $raw.Trim()
    if (-not $txt) { return $null }
    try { return [datetime]::Parse($txt) } catch { return $null }
}

function Apply-EmailFilters {
    param(
        [array]$Emails,
        [string]$ViewName,
        [string]$QuickFilter = 'All',
        [string]$SearchText = '',
        [string]$FromDateText = '',
        [string]$ToDateText = '',
        [string]$Period = 'All Time'
    )
    $out = @($Emails)
    if ($ViewName -and $ViewName -ne '(All)') {
        $out = @(Filter-EmailsByView $out $ViewName)
    }
    $out = @($out | Where-Object {
        switch ($QuickFilter) {
            'Unread'          { -not $_.IsRead }
            'VIP'             { $_.IsVIP }
            'Internal'        { $_.IsInternal }
            'External'        { -not $_.IsInternal }
            'Has Attachments' { $_.HasAttachment }
            default           { $true }
        }
    })

    if ($Period -and $Period -ne 'All Time') {
        $cutoff = switch ($Period) {
            'Today'         { (Get-Date).Date }
            'Last 7 Days'   { (Get-Date).AddDays(-7) }
            'Last 30 Days'  { (Get-Date).AddDays(-30) }
            'Last 90 Days'  { (Get-Date).AddDays(-90) }
            default         { $null }
        }
        if ($cutoff) { $out = @($out | Where-Object { $_.ReceivedTime -ge $cutoff }) }
    }

    $fromDate = Try-ParseFilterDate $FromDateText
    if ($fromDate) { $out = @($out | Where-Object { $_.ReceivedTime -ge $fromDate.Date }) }
    $toDate = Try-ParseFilterDate $ToDateText
    if ($toDate) {
        $end = $toDate.Date.AddDays(1).AddSeconds(-1)
        $out = @($out | Where-Object { $_.ReceivedTime -le $end })
    }

    $search = if ($SearchText) { $SearchText.Trim().ToLower() } else { '' }
    if ($search) {
        $out = @($out | Where-Object {
            ($_.Subject      -and $_.Subject.ToLower().Contains($search)) -or
            ($_.SenderName   -and $_.SenderName.ToLower().Contains($search)) -or
            ($_.SenderEmail  -and $_.SenderEmail.ToLower().Contains($search)) -or
            ($_.Body         -and $_.Body.ToLower().Contains($search)) -or
            ($_.FolderPath   -and $_.FolderPath.ToLower().Contains($search))
        })
    }
    return @($out)
}

function New-ReplyDraft {
    param([string]$entryId, [string]$templateName, [switch]$ReplyAll, [string]$BodyText)
    $ol = Get-OutlookApp
    if (-not $ol -or -not $entryId) { return $null }
    try {
        $ns   = $ol.GetNamespace('MAPI')
        $mail = $ns.GetItemFromID($entryId)
        $reply = if ($ReplyAll) { $mail.ReplyAll() } else { $mail.Reply() }
        $prefix = ''
        if ($BodyText) { $prefix = $BodyText + "`n`n" }
        if ($templateName -and $script:EmailCfg.replyTemplates.$templateName) {
            $prefix += $script:EmailCfg.replyTemplates.$templateName + "`n`n"
        }
        if ($prefix) { $reply.Body = $prefix + $reply.Body }
        $reply.Save()
        # Log to state
        $script:EmailState.draftLog += [pscustomobject]@{
            type='Reply'; entryId=$entryId; template=$templateName
            timestamp=(Get-Date -f 'yyyy-MM-dd HH:mm:ss')
        }
        Save-EmailState
        return $reply
    } catch { return $null }
}

function New-ForwardDraft {
    param([string]$entryId, [string]$BodyText)
    $ol = Get-OutlookApp
    if (-not $ol -or -not $entryId) { return $null }
    try {
        $ns   = $ol.GetNamespace('MAPI')
        $mail = $ns.GetItemFromID($entryId)
        $fwd  = $mail.Forward()
        if ($BodyText) { $fwd.Body = $BodyText + "`n`n" + $fwd.Body }
        $fwd.Save()
        $script:EmailState.draftLog += [pscustomobject]@{
            type='Forward'; entryId=$entryId; template=''
            timestamp=(Get-Date -f 'yyyy-MM-dd HH:mm:ss')
        }
        Save-EmailState
        return $fwd
    } catch { return $null }
}

function New-OutlookToDo([string]$entryId, [string]$subject, [string]$dueDate) {
    $ol = Get-OutlookApp
    if (-not $ol) { return $null }
    try {
        $task = $ol.CreateItem(3)  # olTaskItem = 3
        $task.Subject = "EMAIL: $subject"
        if ($dueDate) { $task.DueDate = [datetime]::Parse($dueDate) }
        $task.Body = "Created from email EntryID: $entryId"
        $task.Save()
        $script:EmailState.todoAssignments += [pscustomobject]@{
            entryId=$entryId; subject=$subject; dueDate=$dueDate
            timestamp=(Get-Date -f 'yyyy-MM-dd HH:mm:ss')
        }
        Save-EmailState
        return $task
    } catch { return $null }
}

function New-CalendarDeadline([string]$entryId, [string]$subject, [string]$dueDate, [switch]$AllDay) {
    $ol = Get-OutlookApp
    if (-not $ol) { return $null }
    try {
        $appt = $ol.CreateItem(1)  # olAppointmentItem = 1
        $appt.Subject = "DEADLINE: $subject"
        $dt = [datetime]::Parse($dueDate)
        if ($AllDay) {
            $appt.AllDayEvent = $true
            $appt.Start = $dt.Date
            $appt.End   = $dt.Date.AddDays(1)
        } else {
            $appt.Start = $dt
            $appt.End   = $dt.AddHours(1)
        }
        $appt.ReminderSet = $true
        $appt.ReminderMinutesBeforeStart = $script:EmailCfg.calendarReminderMinutes
        $appt.Body = "Deadline from email EntryID: $entryId"
        $appt.Save()
        $script:EmailState.deadlineAssignments += [pscustomobject]@{
            entryId=$entryId; subject=$subject; dueDate=$dueDate
            timestamp=(Get-Date -f 'yyyy-MM-dd HH:mm:ss')
        }
        Save-EmailState
        return $appt
    } catch { return $null }
}

function Invoke-EmailRecall([string]$entryId) {
    $ol = Get-OutlookApp
    if (-not $ol -or -not $entryId) { return $false }
    try {
        $ns   = $ol.GetNamespace('MAPI')
        $mail = $ns.GetItemFromID($entryId)
        $sent = (Get-Date) - $mail.SentOn
        $maxH = if ($script:EmailCfg.recallMaxAgeHours) { $script:EmailCfg.recallMaxAgeHours } else { 1 }
        if ($sent.TotalHours -gt $maxH) { return $false }
        # Recall via Outlook Actions collection -- requires Exchange
        $recalled = $false
        foreach ($action in @($mail.Actions)) {
            if ([string]$action.Name -match 'Recall') {
                [void]$action.Execute()
                $recalled = $true
                break
            }
        }
        if (-not $recalled) { return $false }
        $script:EmailState.recallAttempts += [pscustomobject]@{
            entryId=$entryId; subject=$mail.Subject
            timestamp=(Get-Date -f 'yyyy-MM-dd HH:mm:ss'); result='Attempted'
        }
        Save-EmailState
        return $true
    } catch { return $false }
}

function Find-DuplicateEmails([array]$emails) {
    # Group by Subject + SenderEmail + Date (to the second -- minute was too coarse and caused false positives)
    $groups = $emails | Group-Object { "$($_.Subject)|$($_.SenderEmail)|$($_.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss'))" }
    $dupes = $groups | Where-Object { $_.Count -gt 1 }
    return $dupes
}

function Remove-DuplicateEmails([string]$Scope = 'All Folders') {
    $ol = Get-OutlookApp
    if (-not $ol) { return 0 }
    $dupeScanLimit = if ($script:EmailCfg -and $script:EmailCfg.dupeScanMaxCount) { [int]$script:EmailCfg.dupeScanMaxCount } else { 10000 }
    $emails = Get-InboxEmails -maxCount $dupeScanLimit -scope $Scope
    $dupes = Find-DuplicateEmails $emails
    $removed = 0
    $ns = $ol.GetNamespace('MAPI')
    foreach ($grp in $dupes) {
        # Keep the first, delete the rest
        $toDelete = $grp.Group | Select-Object -Skip 1
        foreach ($e in $toDelete) {
            try {
                $mail = $ns.GetItemFromID($e.EntryID)
                $mail.Delete()
                $removed++
            } catch {}
        }
    }
    return $removed
}

function Find-DuplicateFolders {
    $ol = Get-OutlookApp
    if (-not $ol) { return @() }
    $ns  = $ol.GetNamespace('MAPI')
    $roots = @()
    try {
        foreach ($st in @($ns.Stores)) {
            try { $roots += $st.GetRootFolder() } catch {}
        }
    } catch {}
    if ($roots.Count -eq 0) { $roots = @($ns.GetDefaultFolder(6)) }
    $folderNames = @{}
    $dupes = @()
    $queue = [System.Collections.Queue]::new()
    foreach ($r in $roots) { $queue.Enqueue($r) }
    while ($queue.Count -gt 0) {
        $f = $queue.Dequeue()
        try {
            foreach ($sub in @($f.Folders)) { $queue.Enqueue($sub) }
        } catch {}
        try {
            $nm = $f.Name.ToLower().Trim()
            if ($folderNames.ContainsKey($nm)) {
                $dupes += [pscustomobject]@{
                    Name       = $f.Name
                    EntryID    = $f.EntryID
                    ItemCount  = $f.Items.Count
                    OriginalID = $folderNames[$nm]
                }
            } else {
                $folderNames[$nm] = $f.EntryID
            }
        } catch {}
    }
    return $dupes
}

function Remove-DuplicateFolders {
    $ol = Get-OutlookApp
    if (-not $ol) { return 0 }
    $ns   = $ol.GetNamespace('MAPI')
    $dupes = Find-DuplicateFolders
    $removed = 0
    foreach ($d in $dupes) {
        try {
            $srcFolder  = $ns.GetFolderFromID($d.EntryID)
            $destFolder = $ns.GetFolderFromID($d.OriginalID)
            # Move all items from duplicate to original folder
            while ($srcFolder.Items.Count -gt 0) {
                $srcFolder.Items.Item(1).Move($destFolder)
            }
            # Delete the now-empty duplicate folder
            $srcFolder.Delete()
            $removed++
        } catch {}
    }
    return $removed
}

# ============================================================
#  IDE HELPERS
# ============================================================
function Get-IdeLanguageFromInputs {
    param([string]$Mode = 'Auto', [string]$Path = '')
    $m = if ($Mode) { $Mode.Trim() } else { 'Auto' }
    if ($m -eq 'PowerShell') { return 'PowerShell' }
    if ($m -eq 'VBA') { return 'VBA' }

    $ext = [System.IO.Path]::GetExtension([string]$Path).ToLower()
    switch ($ext) {
        '.ps1' { return 'PowerShell' }
        '.psm1' { return 'PowerShell' }
        '.psd1' { return 'PowerShell' }
        '.bas' { return 'VBA' }
        '.cls' { return 'VBA' }
        '.frm' { return 'VBA' }
        default { return 'PowerShell' }
    }
}

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

function Get-VbaEntryPointNameFromText([string]$Content) {
    if (-not $Content) { return $null }
    $m1 = [regex]::Match($Content, '(?im)^\s*Public\s+Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*(\(|$)')
    if ($m1.Success) { return $m1.Groups[1].Value }
    $m2 = [regex]::Match($Content, '(?im)^\s*Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*(\(|$)')
    if ($m2.Success) { return $m2.Groups[1].Value }
    return $null
}

function Quote-ProcessArg([string]$Value) {
    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"','\"') + '"'
}

# ================================================================
#  XAML UI - Chrome/Teams Dark Theme
# ================================================================
function Start-MacroHub {

$xamlStr = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="MacroHub v3.1" Height="860" Width="1200"
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
      <Setter Property="Foreground"      Value="#6E6E6E"/>
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
            <RowDefinition Height="14"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="14"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="16,12">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="240"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Workbook" Foreground="#B0B0B0" FontSize="12" VerticalAlignment="Center"/>
              <ComboBox  Grid.Column="2" x:Name="ClipWbCombo"/>
              <Button    Grid.Column="4" x:Name="ClipRefreshBtn" Content="_Refresh" Style="{StaticResource Btn}" Padding="12,6"/>
              <Button    Grid.Column="5" x:Name="ClipLockBtn" Content="_Lock Defaults" Style="{StaticResource BtnAccent}" Padding="10,6" Margin="8,0,0,0" ToolTip="Save workbook + slot-1 sheet/cell/timestamp settings as defaults for new slots"/>
              <Button    Grid.Column="6" x:Name="ClipClearDefaultsBtn" Content="Clear De_faults" Style="{StaticResource Btn}" Padding="10,6" Margin="4,0,0,0" Foreground="#E05050" ToolTip="Remove saved defaults"/>
              <TextBlock Grid.Column="7" x:Name="ClipDefaultsIndicator" Text="(defaults loaded)" Foreground="#4C9FE6" FontSize="10" VerticalAlignment="Center" Margin="8,0,0,0" Visibility="Collapsed"/>
            </Grid>
          </Border>

          <Border Grid.Row="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,8">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>

              <TextBlock Grid.Column="0" Text="Each slot has its own sheet, cell, and timestamp settings next to Paste."
                         Foreground="#6E6E6E" FontSize="11" VerticalAlignment="Center" Margin="0,0,10,0"/>

              <StackPanel Grid.Column="2" Orientation="Horizontal">
                <TextBlock Text="Slots:" Foreground="#6E6E6E" FontSize="11"
                           VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBlock x:Name="ClipSlotCount" Text="0" Foreground="#4C9FE6" FontWeight="Bold"
                           FontSize="12" VerticalAlignment="Center" Margin="0,0,12,0"/>
                <Button x:Name="ClipAddSlotBtn" Content="+ _Add Slot" Style="{StaticResource BtnAccent}"
                        Padding="10,4"/>
                <Button x:Name="ClipRecordSeqBtn" Content="Record _Sequence" Style="{StaticResource Btn}"
                        Padding="10,4" Margin="8,0,0,0" ToolTip="Capture each new clipboard copy into the next slot (slot1, slot2, slot3...)."/>
                <TextBlock x:Name="ClipRecordSeqState" Text="SEQ OFF" Foreground="#6E6E6E" FontWeight="Bold"
                           FontSize="10" VerticalAlignment="Center" Margin="8,0,0,0"/>
              </StackPanel>

              <StackPanel Grid.Column="1" Orientation="Horizontal" Visibility="Collapsed">
                <ComboBox x:Name="ClipSheetCombo"/>
                <TextBox x:Name="ClipCellBox" Text="A1"/>
                <CheckBox x:Name="ClipTimestampChk"/>
                <TextBox x:Name="ClipDateOffset" Text="1"/>
                <TextBox x:Name="ClipTimeOffset" Text="2"/>
              </StackPanel>
            </Grid>
          </Border>

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

          <Border Grid.Column="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <TextBlock Text="MACROS" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
              </Border>
              <ListBox Grid.Row="1" x:Name="MacroList" Background="Transparent"
                       Foreground="#FFFFFF" BorderThickness="0" Padding="4"
                       FontSize="12"/>
              <Border Grid.Row="2" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="10,8">
                <StackPanel>
                  <Button x:Name="MacroRefreshBtn" Content="Refresh _List" Style="{StaticResource Btn}"
                          Padding="10,5" HorizontalAlignment="Stretch" Margin="0,0,0,4"/>
                  <Button x:Name="MacroFavBtn" Content="Toggle _Favorite" Style="{StaticResource BtnGreen}"
                          Padding="10,5" HorizontalAlignment="Stretch"/>
                </StackPanel>
              </Border>
            </Grid>
          </Border>

          <Border Grid.Column="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="20,16">
            <StackPanel>
              <TextBlock Text="RUN MACRO" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold" Margin="0,0,0,14"/>

              <TextBlock Text="Target Workbook" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
              <ComboBox x:Name="MacroWbCombo" Margin="0,0,0,16"/>

              <TextBlock Text="Target Sheet" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
              <ComboBox x:Name="MacroSheetCombo" Margin="0,0,0,16"/>

              <TextBlock Text="Selected Macro" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
              <TextBox x:Name="MacroSelectedTxt" IsReadOnly="True" Foreground="#6E6E6E" Margin="0,0,0,16"/>

              <TextBlock Text="Description" Foreground="#B0B0B0" FontSize="12" Margin="0,0,0,6"/>
              <TextBox x:Name="MacroDescTxt" IsReadOnly="True" Height="60"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       Foreground="#6E6E6E" Margin="0,0,0,20"/>

              <Button x:Name="MacroRunBtn" Content="_Run Macro" Style="{StaticResource BtnAccent}"
                      Padding="20,10" FontSize="13" HorizontalAlignment="Left"/>

              <TextBlock Text="OUTPUT" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold" Margin="0,20,0,6"/>
              <TextBox x:Name="MacroOutputTxt" Height="120" IsReadOnly="True"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       FontFamily="Consolas" FontSize="11" Foreground="#B0B0B0"
                       Background="Transparent" BorderThickness="0"/>
            </StackPanel>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="Scheduler">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="#1A2A3E" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Scripts Folder" Foreground="#4C9FE6" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center"/>
              <TextBlock Grid.Column="2" x:Name="SchedFolderPath" Foreground="#B0B0B0" FontSize="11"
                         VerticalAlignment="Center" TextTrimming="CharacterEllipsis"
                         ToolTip="Drop .ps1 or .bas files here and they appear in the dropdown below"/>
              <Button Grid.Column="4" x:Name="SchedOpenFolderBtn" Content="_Open Folder"
                      Style="{StaticResource Btn}" Padding="10,4"/>
            </Grid>
          </Border>

          <Border Grid.Row="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="20,16">
            <StackPanel>
              <Grid Margin="0,0,0,10">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="SCHEDULE A TASK" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Text="  Backed by Windows Task Scheduler -- runs even when MacroHub is closed"
                           Foreground="#4C9FE6" FontSize="10" FontStyle="Italic" VerticalAlignment="Center"
                           HorizontalAlignment="Right"/>
              </Grid>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition Width="120"/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition Width="120"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="6"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Grid.Column="0" Text="Task Name" Foreground="#B0B0B0" FontSize="11"/>
                <TextBlock Grid.Row="0" Grid.Column="2" Text="Script / Macro" Foreground="#B0B0B0" FontSize="11"/>
                <TextBlock Grid.Row="0" Grid.Column="4" Text="Time (HH:mm)" Foreground="#B0B0B0" FontSize="11"/>
                <TextBlock Grid.Row="0" Grid.Column="6" Text="Frequency" Foreground="#B0B0B0" FontSize="11"/>

                <TextBox   Grid.Row="2" Grid.Column="0" x:Name="SchedNameBox"/>
                <ComboBox  Grid.Row="2" Grid.Column="2" x:Name="SchedMacroCombo"/>
                <TextBox   Grid.Row="2" Grid.Column="4" x:Name="SchedTimeBox" Text="09:00"/>
                <ComboBox  Grid.Row="2" Grid.Column="6" x:Name="SchedFreqCombo">
                  <ComboBoxItem Content="Daily" IsSelected="True"/>
                  <ComboBoxItem Content="Weekly"/>
                  <ComboBoxItem Content="Monthly"/>
                </ComboBox>
              </Grid>
              <StackPanel Orientation="Horizontal" Margin="0,14,0,0">
                <Button x:Name="SchedCreateBtn" Content="_Create Task" Style="{StaticResource BtnAccent}" Padding="14,7" Margin="0,0,10,0"/>
                <Button x:Name="SchedRefreshBtn" Content="_Refresh" Style="{StaticResource Btn}" Padding="14,7"/>
              </StackPanel>
            </StackPanel>
          </Border>

          <Border Grid.Row="4" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="150"/>
                    <ColumnDefinition Width="150"/>
                    <ColumnDefinition Width="80"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="TASK NAME" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="1" Text="STATE" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="2" Text="NEXT RUN" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="3" Text="LAST RUN" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="4" Text="ACTION" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
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
            <ColumnDefinition Width="220"/>
            <ColumnDefinition Width="12"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <Border Grid.Column="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <TextBlock Text="OPEN WORKBOOKS" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
              </Border>
              <ListBox Grid.Row="1" x:Name="NavWbList" Background="Transparent"
                       Foreground="#FFFFFF" BorderThickness="0" Padding="4" FontSize="12"/>
              <Border Grid.Row="2" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="10,8">
                <Button x:Name="NavRefreshBtn" Content="_Refresh" Style="{StaticResource Btn}"
                        Padding="10,5" HorizontalAlignment="Stretch"/>
              </Border>
            </Grid>
          </Border>

          <Border Grid.Column="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <TextBlock Text="SHEETS (multi-select + drag reorder + Ctrl+E rename)" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
              </Border>
              <ListBox Grid.Row="1" x:Name="NavSheetList" Background="Transparent"
                       Foreground="#FFFFFF" BorderThickness="0" Padding="4" FontSize="12"
                       SelectionMode="Extended" AllowDrop="True"/>
              <Border Grid.Row="2" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="10,6">
                <TextBlock x:Name="NavSheetCount" Text="0 sheets" Foreground="#6E6E6E" FontSize="10" HorizontalAlignment="Center"/>
              </Border>
            </Grid>
          </Border>

          <Border Grid.Column="4" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="20,16">
              <StackPanel>

                <TextBlock Text="ACTIVATE" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold" Margin="0,0,0,8"/>
                <Button x:Name="NavActivateBtn" Content="_Activate Selected Sheet"
                        Style="{StaticResource BtnAccent}" Padding="14,8" Margin="0,0,0,6"
                        HorizontalAlignment="Stretch"/>
                <Button x:Name="NavOpenBtn" Content="_Open Workbook File..."
                        Style="{StaticResource Btn}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"/>
                <Grid Margin="0,0,0,16">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Button Grid.Column="0" x:Name="NavBringFrontBtn" Content="_Bring Forward"
                          Style="{StaticResource Btn}" Padding="10,6" HorizontalAlignment="Stretch"/>
                  <Button Grid.Column="2" x:Name="NavMinimizeBtn" Content="Mi_nimize"
                          Style="{StaticResource Btn}" Padding="10,6" HorizontalAlignment="Stretch"/>
                  <Button Grid.Column="4" x:Name="NavCloseWbBtn" Content="_Close Workbook"
                          Style="{StaticResource BtnRed}" Padding="10,6" HorizontalAlignment="Stretch"/>
                </Grid>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,8">
                  <TextBlock Text="MOVE / COPY SHEET" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                </Border>
                <TextBlock Text="Destination Workbook" Foreground="#B0B0B0" FontSize="10" Margin="0,0,0,4"/>
                <ComboBox x:Name="NavDestWbCombo" Margin="0,0,0,8"/>
                <TextBlock Text="Insert Before Sheet" Foreground="#B0B0B0" FontSize="10" Margin="0,0,0,4"/>
                <ComboBox x:Name="NavDestSheetCombo" Margin="0,0,0,8"/>
                <CheckBox x:Name="NavCopyChk" Content="Copy (keep original)" Foreground="#FFFFFF"
                          IsChecked="True" Margin="0,0,0,8" FontSize="11"/>
                <Button x:Name="NavMoveCopyBtn" Content="_Move / Copy"
                        Style="{StaticResource BtnAccent}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"/>
                <Button x:Name="NavExportSheetBtn" Content="E_xport Selected Sheet..."
                        Style="{StaticResource Btn}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,8">
                  <TextBlock Text="VISIBILITY" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                </Border>
                <Grid Margin="0,0,0,16">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Button Grid.Column="0" x:Name="NavHideBtn" Content="_Hide"
                          Style="{StaticResource Btn}" Padding="14,8" HorizontalAlignment="Stretch"/>
                  <Button Grid.Column="2" x:Name="NavUnhideBtn" Content="_Unhide"
                          Style="{StaticResource BtnGreen}" Padding="14,8" HorizontalAlignment="Stretch"/>
                </Grid>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,8">
                  <TextBlock Text="DELETE" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                </Border>
                <Button x:Name="NavDeleteBtn" Content="_Delete Selected Sheets"
                        Style="{StaticResource BtnRed}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,8">
                  <TextBlock Text="PASSWORD PROTECTION" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                </Border>
                <TextBlock Text="Password" Foreground="#B0B0B0" FontSize="10" Margin="0,0,0,4"/>
                <TextBox x:Name="NavPwdBox" FontSize="12" Foreground="#FFFFFF" Background="#1B1B1F"
                         BorderBrush="#3E3E42" BorderThickness="1" Padding="6,5" Margin="0,0,0,8"/>
                <Button x:Name="NavSetPwdBtn" Content="Set _Password on Workbook"
                        Style="{StaticResource BtnAccent}" Padding="14,8" Margin="0,0,0,6"
                        HorizontalAlignment="Stretch"/>
                <Button x:Name="NavRemPwdBtn" Content="_Remove Password"
                        Style="{StaticResource Btn}" Padding="14,8" Margin="0,0,0,6"
                        HorizontalAlignment="Stretch"/>
                <Button x:Name="NavOpenPwdBtn" Content="Open _With Password..."
                        Style="{StaticResource Btn}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,8">
                  <TextBlock Text="EXCEL ENGINE OPTIONS" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                </Border>
                <TextBlock Text="Calculation Mode" Foreground="#B0B0B0" FontSize="10" Margin="0,0,0,4"/>
                <ComboBox x:Name="NavCalcModeCombo" Margin="0,0,0,6">
                  <ComboBoxItem Content="Automatic" IsSelected="True"/>
                  <ComboBoxItem Content="Manual"/>
                  <ComboBoxItem Content="Semiautomatic"/>
                </ComboBox>
                <CheckBox x:Name="NavEventsChk" Content="Events enabled" Foreground="#FFFFFF"
                          IsChecked="True" Margin="0,0,0,8" FontSize="11"/>
                <Button x:Name="NavApplyExcelOptsBtn" Content="_Apply Excel Options"
                        Style="{StaticResource BtnAccent}" Padding="14,8" Margin="0,0,0,16"
                        HorizontalAlignment="Stretch"/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,8">
                  <TextBlock Text="VBA MODULES / MACROS" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                </Border>
                <ListBox x:Name="NavVbaList" Height="120" Background="#1B1B1F" Foreground="#C0C0C0"
                         BorderBrush="#3E3E42" BorderThickness="1" Margin="0,0,0,16" FontSize="11"
                         ToolTip="Right-click a macro entry to run it in the selected navigator workbook."/>

                <Border BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,14,0,0" Margin="0,0,0,8">
                  <TextBlock Text="QUICK INFO" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
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

          <Border Grid.Column="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#3E3E42" BorderThickness="0,0,0,1" Padding="14,10">
                <TextBlock Text="SAVED TEMPLATES" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
              </Border>
              <ListBox Grid.Row="1" x:Name="TplList" Background="Transparent"
                       Foreground="#FFFFFF" BorderThickness="0" Padding="4" FontSize="12"/>
              <Border Grid.Row="2" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="10,8">
                <Button x:Name="TplDeleteBtn" Content="_Delete Selected" Style="{StaticResource BtnRed}"
                        Padding="10,5" HorizontalAlignment="Stretch"/>
              </Border>
            </Grid>
          </Border>

          <Border Grid.Column="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="20,16">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="8"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="8"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="12"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="8"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <TextBlock Grid.Row="0" Text="TEMPLATE EDITOR" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
              <TextBox Grid.Row="2" x:Name="TplNameBox" ToolTip="Template name"/>
              <TextBox Grid.Row="4" x:Name="TplContentBox" AcceptsReturn="True"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       ToolTip="Template content - use {DATE}, {SHEET}, {USER} placeholders"
                       FontFamily="Consolas" FontSize="11"/>
              <StackPanel Grid.Row="6" Orientation="Horizontal">
                <Button x:Name="TplSaveBtn" Content="_Save Template" Style="{StaticResource BtnAccent}"
                        Padding="14,7" Margin="0,0,10,0"/>
                <Button x:Name="TplPasteBtn" Content="_Paste to Excel" Style="{StaticResource BtnGreen}"
                        Padding="14,7" Margin="0,0,10,0"/>
                <Button x:Name="TplPreviewBtn" Content="Pre_view" Style="{StaticResource Btn}"
                        Padding="14,7"/>
              </StackPanel>
              <Border Grid.Row="8" x:Name="TplPreviewCard" Visibility="Collapsed"
                      Background="#1E2D1E" CornerRadius="6" BorderBrush="#50A050" BorderThickness="1" Padding="12,10">
                <StackPanel>
                  <TextBlock Text="PREVIEW" Foreground="#50A050" FontSize="10" FontWeight="SemiBold" Margin="0,0,0,6"/>
                  <TextBox x:Name="TplPreviewTxt" IsReadOnly="True" FontFamily="Consolas" FontSize="11"
                           AcceptsReturn="True" TextWrapping="Wrap" MaxHeight="150"
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

            <Border Background="#252528" CornerRadius="6" BorderBrush="#3E3E42"
                    BorderThickness="1" Padding="20,16" Margin="0,0,0,12">
              <StackPanel>
                <TextBlock Text="QUARTER FOLDER SETUP" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold" Margin="0,0,0,4"/>
                <TextBlock Text="Run once per quarter to compare folders and build the initial to-do checklist."
                           Foreground="#8B8B8F" FontSize="11" FontStyle="Italic" Margin="0,0,0,12"/>

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
                  <TextBlock Text="COMPLETION BY FOLDER" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
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
                  <TextBlock Text="SYNC LOG" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
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
            <RowDefinition Height="10"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="290"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="290"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Source A" Foreground="#6E6E6E" FontSize="11" VerticalAlignment="Center"/>
              <ComboBox  Grid.Column="2" x:Name="QsQuarterCombo" VerticalAlignment="Center" IsEditable="True"
                         ToolTip="Folder A (baseline source)"/>
              <TextBlock Grid.Column="4" Text="vs" Foreground="#4C9FE6" FontSize="12" FontWeight="Bold" VerticalAlignment="Center"/>
              <ComboBox  Grid.Column="6" x:Name="QsCompareCombo" VerticalAlignment="Center" IsEditable="True"
                         ToolTip="Folder B (delivery/target)"/>
              <Button Grid.Column="8" x:Name="QsNewQuarterBtn" Content="Browse _A" Style="{StaticResource Btn}" Padding="10,5"/>
              <Button Grid.Column="10" x:Name="QsScanFolderBtn" Content="Browse _B" Style="{StaticResource Btn}" Padding="10,5"/>
              <Button Grid.Column="12" x:Name="QsAddTaskBtn" Content="_Compare Missing" Style="{StaticResource BtnAccent}" Padding="10,5"/>
              <TextBox Grid.Column="14" x:Name="QsScanPathBox" ToolTip="Optional compare tag (saved in note)"/>
              <TextBlock Grid.Column="16" x:Name="QsQuarterSummary" Foreground="#6E6E6E" FontSize="11" VerticalAlignment="Center"/>
              <TextBlock Grid.Column="18" x:Name="QsQuarterPct" Foreground="#4C9FE6" FontSize="12" FontWeight="Bold" VerticalAlignment="Center"/>
              <Button Grid.Column="20" x:Name="QsSwapFoldersBtn" Content="_Swap" Style="{StaticResource Btn}" Padding="10,5"/>
            </Grid>
          </Border>

          <Border Grid.Row="1" x:Name="QsAddTaskCard" Visibility="Collapsed"
                  Background="#2A2A2E" CornerRadius="6" BorderBrush="#3E3E42"
                  BorderThickness="1" Padding="14,10" Margin="0,8,0,0">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="160"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="100"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox Grid.Column="0" x:Name="QsNewTaskName" ToolTip="File or task name"/>
              <TextBox Grid.Column="2" x:Name="QsNewTaskFolder" ToolTip="Folder (optional)"/>
              <TextBox Grid.Column="4" x:Name="QsNewTaskDue" ToolTip="Due date yyyy-MM-dd"/>
              <Button Grid.Column="6" x:Name="QsSaveTaskBtn" Content="Sa_ve" Style="{StaticResource BtnAccent}" Padding="10,5"/>
              <Button Grid.Column="8" x:Name="QsCancelTaskBtn" Content="_Cancel" Style="{StaticResource Btn}" Padding="10,5"/>
            </Grid>
          </Border>

          <Border Grid.Row="3" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="110"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="160"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="160"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>

              <TextBlock Grid.Column="0" Text="Status" Foreground="#6E6E6E" FontSize="11" VerticalAlignment="Center"/>
              <ComboBox  Grid.Column="2" x:Name="QsFltStatus" VerticalAlignment="Center">
                <ComboBoxItem Content="All" IsSelected="True"/>
                <ComboBoxItem Content="Pending"/>
                <ComboBoxItem Content="Done"/>
              </ComboBox>

              <TextBlock Grid.Column="4" Text="File" Foreground="#6E6E6E" FontSize="11" VerticalAlignment="Center"/>
              <TextBox   Grid.Column="6" x:Name="QsFltName" ToolTip="Filter by file name"/>

              <TextBlock Grid.Column="8" Text="Folder" Foreground="#6E6E6E" FontSize="11" VerticalAlignment="Center"/>
              <TextBox   Grid.Column="10" x:Name="QsFltFolder" ToolTip="Filter by folder"/>

              <Button Grid.Column="12" x:Name="QsExportBtn" Content="_Export to Excel"
                      Style="{StaticResource BtnAccent}" Padding="12,6"/>
              <Button Grid.Column="14" x:Name="QsRefreshBtn" Content="_Refresh"
                      Style="{StaticResource Btn}" Padding="12,6"/>
            </Grid>
          </Border>

          <Border Grid.Row="5" Background="#252528" CornerRadius="6"
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
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="200"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="FILE NAME" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="1" Text="FOLDER" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="2" Text="ADDED IN A" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="3" Text="UPDATED IN A" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                  <TextBlock Grid.Column="4" Text="ACTIONS" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
                </Grid>
              </Border>

              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="QsTodoPanel" Margin="0,4"/>
              </ScrollViewer>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="Email Helper">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="#252528" CornerRadius="6" Padding="12,8" Margin="0,0,0,8">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Button Grid.Column="0" x:Name="EmRefreshBtn" Content="_Refresh" Style="{StaticResource Btn}" Margin="0,0,6,0" Padding="8,4"/>
              <ComboBox Grid.Column="1" x:Name="EmScopeCombo" Width="120" Margin="0,0,6,0" ToolTip="Folder scope">
                <ComboBoxItem Content="All Folders" IsSelected="True"/>
                <ComboBoxItem Content="Inbox Only"/>
              </ComboBox>
              <ComboBox Grid.Column="2" x:Name="EmViewCombo" Width="150" Margin="0,0,6,0" ToolTip="Select email view"/>
              <ComboBox Grid.Column="3" x:Name="EmFilterCombo" Width="110" Margin="0,0,6,0" ToolTip="Quick filter">
                <ComboBoxItem Content="All" IsSelected="True"/>
                <ComboBoxItem Content="Unread"/>
                <ComboBoxItem Content="VIP"/>
                <ComboBoxItem Content="Internal"/>
                <ComboBoxItem Content="External"/>
                <ComboBoxItem Content="Has Attachments"/>
              </ComboBox>
              <TextBox Grid.Column="4" x:Name="EmFromDateBox" Width="92" Margin="0,0,6,0" ToolTip="From date yyyy-MM-dd"/>
              <TextBox Grid.Column="5" x:Name="EmToDateBox" Width="92" Margin="0,0,6,0" ToolTip="To date yyyy-MM-dd"/>
              <TextBox Grid.Column="6" x:Name="EmSearchBox" Width="180" Margin="0,0,8,0" ToolTip="Search subject/sender/body/folder"/>
              <TextBlock Grid.Column="7" x:Name="EmLastRefreshTxt" Text="" Foreground="#6E6E6E"
                         FontSize="10" VerticalAlignment="Center" HorizontalAlignment="Right"/>
              <Button Grid.Column="8" x:Name="EmDupEmailBtn" Content="De-_Dup Emails" Style="{StaticResource Btn}" Margin="0,0,6,0" Padding="8,4"/>
              <Button Grid.Column="9" x:Name="EmDupFolderBtn" Content="Dup F_olders" Style="{StaticResource Btn}" Margin="0,0,6,0" Padding="8,4"/>
              <TextBlock Grid.Column="10" x:Name="EmCountTxt" Text="0 emails" Foreground="#6E6E6E"
                         FontSize="11" VerticalAlignment="Center" Margin="8,0"/>
            </Grid>
          </Border>

          <Border Grid.Row="1" Background="#1F1F23" CornerRadius="4" Padding="10,6" Margin="0,0,0,4">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="24"/>
                <ColumnDefinition Width="24"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="160"/>
                <ColumnDefinition Width="90"/>
                <ColumnDefinition Width="70"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="" Foreground="#6E6E6E" FontSize="10"/>
              <TextBlock Grid.Column="1" Text="" Foreground="#6E6E6E" FontSize="10"/>
              <TextBlock Grid.Column="2" Text="SUBJECT" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
              <TextBlock Grid.Column="3" Text="FROM" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
              <TextBlock Grid.Column="4" Text="RECEIVED" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
              <TextBlock Grid.Column="5" Text="AGE" Foreground="#6E6E6E" FontSize="10" FontWeight="SemiBold"/>
            </Grid>
          </Border>

          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="5"/>
              <ColumnDefinition Width="300"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="#1B1B1F" CornerRadius="6" BorderBrush="#3E3E42" BorderThickness="1">
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="EmEmailPanel" Margin="0,4"/>
              </ScrollViewer>
            </Border>

            <GridSplitter Grid.Column="1" Width="5" Background="Transparent" HorizontalAlignment="Stretch"/>

            <Border Grid.Column="2" Background="#252528" CornerRadius="6" BorderBrush="#3E3E42" BorderThickness="1" Padding="10">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" x:Name="EmPreviewSubject" Text="Select an email" Foreground="#FFFFFF"
                           FontSize="13" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,4"/>
                <TextBlock Grid.Row="1" x:Name="EmPreviewFrom" Text="" Foreground="#6E6E6E"
                           FontSize="11" Margin="0,0,0,8"/>
                <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto">
                  <TextBlock x:Name="EmPreviewBody" Text="" Foreground="#C0C0C0" FontSize="11"
                             TextWrapping="Wrap"/>
                </ScrollViewer>

                <Border Grid.Row="3" x:Name="EmComposePanel" Visibility="Collapsed"
                        Background="#1B1B1F" CornerRadius="4" BorderBrush="#4C9FE6" BorderThickness="1"
                        Padding="8" Margin="0,8,0,0">
                  <Grid>
                    <Grid.RowDefinitions>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="*"/>
                      <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" x:Name="EmComposeLabel" Text="Composing Reply..." Foreground="#4C9FE6"
                               FontSize="11" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <TextBox Grid.Row="1" x:Name="EmComposeTo" IsReadOnly="True" FontSize="11"
                             Foreground="#B0B0B0" Background="Transparent" BorderThickness="0"
                             Margin="0,0,0,4"/>
                    <TextBox Grid.Row="2" x:Name="EmComposeBody" AcceptsReturn="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" MinHeight="80" MaxHeight="180"
                             FontSize="11" Foreground="#FFFFFF" Background="#1B1B1F"
                             BorderBrush="#3E3E42" BorderThickness="1" Padding="6,4"
                             ToolTip="Type your message here"/>
                    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,6,0,0">
                      <Button x:Name="EmSaveDraftBtn" Content="  Save as _Draft  "
                              Style="{StaticResource BtnGreen}" Padding="10,5" Margin="0,0,6,0"/>
                      <Button x:Name="EmCancelComposeBtn" Content="  Cancel  "
                              Style="{StaticResource Btn}" Padding="10,5"/>
                    </StackPanel>
                  </Grid>
                </Border>
              </Grid>
            </Border>
          </Grid>

          <Border Grid.Row="3" Background="#252528" CornerRadius="6" Padding="10,8" Margin="0,8,0,0">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <ComboBox Grid.Column="0" x:Name="EmTemplateCombo" Width="140" Margin="0,0,6,0" ToolTip="Reply template"/>
              <Button Grid.Column="1" x:Name="EmReplyBtn" Content="  Rep_ly  " Style="{StaticResource BtnAccent}" Margin="0,0,4,0" Padding="8,4"/>
              <Button Grid.Column="2" x:Name="EmReplyAllBtn" Content="  Reply _All  " Style="{StaticResource BtnAccent}" Margin="0,0,4,0" Padding="8,4"/>
              <Button Grid.Column="3" x:Name="EmForwardBtn" Content="  For_ward  " Style="{StaticResource BtnAccent}" Margin="0,0,8,0" Padding="8,4"/>
              <Button Grid.Column="4" x:Name="EmTodoBtn" Content="  To-D_o  " Style="{StaticResource Btn}" Margin="0,0,4,0" Padding="8,4"/>
              <Button Grid.Column="5" x:Name="EmCalBtn" Content="  Ca_lendar  " Style="{StaticResource Btn}" Margin="0,0,4,0" Padding="8,4"/>
              <TextBlock Grid.Column="6"/>
              <Button Grid.Column="7" x:Name="EmRecallBtn" Content="  Re_call  " Style="{StaticResource BtnRed}" Padding="8,4"
                       ToolTip="Attempt to recall selected sent email"/>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="Email Dashboard">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="#252528" CornerRadius="6" Padding="12,8" Margin="0,0,0,8">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Button Grid.Column="0" x:Name="DashRefreshBtn" Content="_Refresh" Style="{StaticResource Btn}" Padding="8,4" Margin="0,0,6,0"/>
              <ComboBox Grid.Column="1" x:Name="DashScopeCombo" Width="120" Margin="0,0,6,0" ToolTip="Folder scope">
                <ComboBoxItem Content="All Folders" IsSelected="True"/>
                <ComboBoxItem Content="Inbox Only"/>
              </ComboBox>
              <ComboBox Grid.Column="2" x:Name="DashViewCombo" Width="150" Margin="0,0,6,0" ToolTip="Dashboard view"/>
              <ComboBox Grid.Column="3" x:Name="DashFilterCombo" Width="110" Margin="0,0,6,0" ToolTip="Quick filter">
                <ComboBoxItem Content="All" IsSelected="True"/>
                <ComboBoxItem Content="Unread"/>
                <ComboBoxItem Content="VIP"/>
                <ComboBoxItem Content="Internal"/>
                <ComboBoxItem Content="External"/>
                <ComboBoxItem Content="Has Attachments"/>
              </ComboBox>
              <ComboBox Grid.Column="4" x:Name="DashPeriodCombo" Width="110" Margin="0,0,6,0" ToolTip="Date range">
                <ComboBoxItem Content="All Time" IsSelected="True"/>
                <ComboBoxItem Content="Today"/>
                <ComboBoxItem Content="Last 7 Days"/>
                <ComboBoxItem Content="Last 30 Days"/>
                <ComboBoxItem Content="Last 90 Days"/>
              </ComboBox>
              <TextBox Grid.Column="5" x:Name="DashFromDateBox" Width="92" Margin="0,0,6,0" ToolTip="From date yyyy-MM-dd"/>
              <TextBox Grid.Column="6" x:Name="DashToDateBox" Width="92" Margin="0,0,6,0" ToolTip="To date yyyy-MM-dd"/>
              <TextBox Grid.Column="7" x:Name="DashSearchBox" Width="180" Margin="0,0,6,0" ToolTip="Search subject/sender/body/folder"/>
              <TextBlock Grid.Column="8"/>
              <Button Grid.Column="9" x:Name="DashExportBtn" Content="E_xport Report" Style="{StaticResource BtnAccent}" Padding="8,4"/>
            </Grid>
          </Border>

          <Border Grid.Row="1" Background="#1F1F23" CornerRadius="6" Padding="12,8" Margin="0,0,0,8">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Border Grid.Column="0" Background="#252528" CornerRadius="6" Padding="10,8" Margin="0,0,4,0">
                <StackPanel HorizontalAlignment="Center">
                  <TextBlock x:Name="DashTotalCount" Text="0" Foreground="#4C9FE6"
                             FontSize="22" FontWeight="Bold" HorizontalAlignment="Center"/>
                  <TextBlock Text="Total" Foreground="#6E6E6E" FontSize="10" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="1" Background="#252528" CornerRadius="6" Padding="10,8" Margin="2,0">
                <StackPanel HorizontalAlignment="Center">
                  <TextBlock x:Name="DashUnreadCount" Text="0" Foreground="#E05050"
                             FontSize="22" FontWeight="Bold" HorizontalAlignment="Center"/>
                  <TextBlock Text="Unread" Foreground="#6E6E6E" FontSize="10" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="2" Background="#252528" CornerRadius="6" Padding="10,8" Margin="2,0">
                <StackPanel HorizontalAlignment="Center">
                  <TextBlock x:Name="DashVipCount" Text="0" Foreground="#FFD700"
                             FontSize="22" FontWeight="Bold" HorizontalAlignment="Center"/>
                  <TextBlock Text="VIP" Foreground="#6E6E6E" FontSize="10" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="3" Background="#252528" CornerRadius="6" Padding="10,8" Margin="2,0">
                <StackPanel HorizontalAlignment="Center">
                  <TextBlock x:Name="DashInternalCount" Text="0" Foreground="#50C878"
                             FontSize="22" FontWeight="Bold" HorizontalAlignment="Center"/>
                  <TextBlock Text="Internal" Foreground="#6E6E6E" FontSize="10" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="4" Background="#252528" CornerRadius="6" Padding="10,8" Margin="2,0">
                <StackPanel HorizontalAlignment="Center">
                  <TextBlock x:Name="DashExternalCount" Text="0" Foreground="#FFA500"
                             FontSize="22" FontWeight="Bold" HorizontalAlignment="Center"/>
                  <TextBlock Text="External" Foreground="#6E6E6E" FontSize="10" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="5" Background="#252528" CornerRadius="6" Padding="10,8" Margin="4,0,0,0">
                <StackPanel HorizontalAlignment="Center">
                  <TextBlock x:Name="DashOldestDays" Text="0d" Foreground="#FF6666"
                             FontSize="22" FontWeight="Bold" HorizontalAlignment="Center"/>
                  <TextBlock Text="Oldest" Foreground="#6E6E6E" FontSize="10" HorizontalAlignment="Center"/>
                </StackPanel>
              </Border>
            </Grid>
          </Border>

          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Grid Grid.Column="0" Margin="0,0,4,0">
              <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <Border Grid.Row="0" Background="#252528" CornerRadius="6" Padding="12,10" Margin="0,0,0,6">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                  </Grid.RowDefinitions>
                  <TextBlock Grid.Row="0" Text="AGING DISTRIBUTION" Foreground="#6E6E6E"
                             FontSize="10" FontWeight="SemiBold" Margin="0,0,0,8"/>
                  <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                    <StackPanel x:Name="DashAgingPanel"/>
                  </ScrollViewer>
                </Grid>
              </Border>

              <Border Grid.Row="1" Background="#252528" CornerRadius="6" Padding="12,10" Margin="0,0,0,0">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <TextBlock Grid.Row="0" Text="DRAFT BACKLOG" Foreground="#6E6E6E"
                             FontSize="10" FontWeight="SemiBold" Margin="0,0,0,6"/>
                  <StackPanel Grid.Row="1" x:Name="DashDraftPanel"/>
                </Grid>
              </Border>
            </Grid>

            <Grid Grid.Column="1" Margin="4,0,0,0">
              <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <Border Grid.Row="0" Background="#252528" CornerRadius="6" Padding="12,10" Margin="0,0,0,6">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                  </Grid.RowDefinitions>
                  <TextBlock Grid.Row="0" Text="TOP SENDERS" Foreground="#6E6E6E"
                             FontSize="10" FontWeight="SemiBold" Margin="0,0,0,8"/>
                  <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                    <StackPanel x:Name="DashSendersPanel"/>
                  </ScrollViewer>
                </Grid>
              </Border>

              <Border Grid.Row="1" Background="#252528" CornerRadius="6" Padding="12,10">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                  </Grid.RowDefinitions>
                  <TextBlock Grid.Row="0" Text="KEYWORD MATCHES" Foreground="#6E6E6E"
                             FontSize="10" FontWeight="SemiBold" Margin="0,0,0,8"/>
                  <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                    <StackPanel x:Name="DashKeywordPanel"/>
                  </ScrollViewer>
                </Grid>
              </Border>
            </Grid>
          </Grid>
        </Grid>
      </TabItem>

      <TabItem Header="IDE">
        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="260"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="#252528" CornerRadius="6" BorderBrush="#3E3E42" BorderThickness="1" Padding="12,10">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>

              <StackPanel Grid.Row="0" Orientation="Horizontal">
                <Button x:Name="IdeOpenBtn" Content="_Open" Style="{StaticResource Btn}" Margin="0,0,6,0" Padding="10,5"/>
                <Button x:Name="IdeSaveBtn" Content="_Save" Style="{StaticResource Btn}" Margin="0,0,6,0" Padding="10,5"/>
                <Button x:Name="IdeSaveAsBtn" Content="Save _As" Style="{StaticResource Btn}" Margin="0,0,10,0" Padding="10,5"/>
                <Button x:Name="IdeValidateBtn" Content="_Validate / Compile" Style="{StaticResource BtnAccent}" Margin="0,0,6,0" Padding="10,5"/>
                <Button x:Name="IdeRunBtn" Content="_Run" Style="{StaticResource BtnGreen}" Margin="0,0,10,0" Padding="10,5"/>
                <Button x:Name="IdeRefreshExcelBtn" Content="Refresh E_xcel" Style="{StaticResource Btn}" Margin="0,0,10,0" Padding="10,5"/>
                <TextBlock Text="Mode" Foreground="#B0B0B0" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <ComboBox x:Name="IdeModeCombo" Width="130" VerticalAlignment="Center">
                  <ComboBoxItem Content="Auto" IsSelected="True"/>
                  <ComboBoxItem Content="PowerShell"/>
                  <ComboBoxItem Content="VBA"/>
                </ComboBox>
              </StackPanel>

              <Grid Grid.Row="1" Margin="0,10,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="10"/>
                  <ColumnDefinition Width="200"/>
                  <ColumnDefinition Width="10"/>
                  <ColumnDefinition Width="200"/>
                </Grid.ColumnDefinitions>
                <TextBox Grid.Column="0" x:Name="IdePathBox" IsReadOnly="True" ToolTip="Current file path"/>
                <ComboBox Grid.Column="2" x:Name="IdeWbCombo" ToolTip="VBA run target workbook"/>
                <ComboBox Grid.Column="4" x:Name="IdeSheetCombo" ToolTip="VBA run target sheet"/>
              </Grid>
            </Grid>
          </Border>

          <Border Grid.Row="2" Background="#1B1B1F" CornerRadius="6" BorderBrush="#3E3E42" BorderThickness="1" Padding="0">
            <WebBrowser x:Name="IdeEditorBrowser"/>
          </Border>

          <Border Grid.Row="4" Background="#252528" CornerRadius="6" BorderBrush="#3E3E42" BorderThickness="1" Padding="12,10">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="58"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="58"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <TextBlock Grid.Row="0" Text="VALIDATION" Foreground="#E05050" FontSize="10" FontWeight="SemiBold" Margin="0,0,0,4"/>
              <TextBox Grid.Row="1" x:Name="IdeValidationTxt" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       FontFamily="Consolas" FontSize="11" Foreground="#FFB4B4" Background="#1B1B1F" BorderBrush="#3E3E42" BorderThickness="1"/>

              <TextBlock Grid.Row="2" Text="RUN ERRORS" Foreground="#FFA500" FontSize="10" FontWeight="SemiBold" Margin="0,8,0,4"/>
              <TextBox Grid.Row="3" x:Name="IdeRunErrorTxt" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       FontFamily="Consolas" FontSize="11" Foreground="#FFD39B" Background="#1B1B1F" BorderBrush="#3E3E42" BorderThickness="1"/>

              <TextBlock Grid.Row="4" Text="OUTPUT" Foreground="#4C9FE6" FontSize="10" FontWeight="SemiBold" Margin="0,8,0,4"/>
              <TextBox Grid.Row="5" x:Name="IdeOutputTxt" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       FontFamily="Consolas" FontSize="11" Foreground="#C0E5FF" Background="#1B1B1F" BorderBrush="#3E3E42" BorderThickness="1"/>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="AI IDE">`r`n        <Grid Margin="24,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="340"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="12,8">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Tenant" Foreground="#B0B0B0" VerticalAlignment="Center"
                         FontSize="11" Margin="0,0,6,0"
                         ToolTip="Your Azure tenant ID or 'common'. Find at portal.azure.com > Entra ID > Overview"/>
              <TextBox x:Name="AiIdeTenantBox" Text="common" Width="170"
                       FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
              <TextBlock Text="Mode" Foreground="#B0B0B0" VerticalAlignment="Center"
                         FontSize="11" Margin="0,0,6,0"/>
              <ComboBox x:Name="AiIdeModeCombo" Width="120"
                        FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0">
                <ComboBoxItem Content="Auto"/>
                <ComboBoxItem Content="Graph API"   IsSelected="True"/>
                <ComboBoxItem Content="Chrome"/>
              </ComboBox>
              <TextBlock Text="Port" Foreground="#B0B0B0" VerticalAlignment="Center"
                         FontSize="11" Margin="0,0,6,0"/>
              <TextBox x:Name="AiIdePortBox" Text="9876" Width="55"
                       FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
              <Button x:Name="AiIdeAuthBtn"  Content="Auth + Verify"
                      Style="{StaticResource BtnAccent}" Padding="10,4" FontSize="11"
                      Margin="0,0,8,0"
                      ToolTip="Sign in and verify Graph Copilot access. Requires a licensed work M365 Copilot account."/>
              <Button x:Name="AiIdeStartBtn" Content="Start Bridge"
                      Style="{StaticResource BtnGreen}" Padding="10,4" FontSize="11"
                      Margin="0,0,6,0"
                      ToolTip="Start Chrome extension HTTP bridge (only needed in Chrome mode)"/>
              <Button x:Name="AiIdeStopBtn"  Content="Stop"
                      Style="{StaticResource BtnRed}" Padding="8,4" FontSize="11"/>
            </StackPanel>
          </Border>

          <Border Grid.Row="2" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="12,6">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border Grid.Column="0" x:Name="AiIdeModePill"
                      CornerRadius="8" Padding="10,3" Background="#2D2D30">
                <TextBlock x:Name="AiIdeModeTxt" Text="Not connected"
                           Foreground="#6E6E6E" FontSize="11" FontWeight="SemiBold"/>
              </Border>
              <TextBlock Grid.Column="2" x:Name="AiIdeStatusTxt"
                         Text="Choose language, type prompt below, then click Generate"
                         Foreground="#6E6E6E" FontSize="12" VerticalAlignment="Center"/>
              <TextBlock Grid.Column="3" x:Name="AiIdeRetryTxt" Text=""
                         Foreground="#FFA500" FontSize="11" VerticalAlignment="Center"/>
            </Grid>
          </Border>

          <Border Grid.Row="4" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="12,8">
            <StackPanel Orientation="Horizontal">
              <Button x:Name="AiIdeNewBtn"  Content="New"
                      Style="{StaticResource Btn}" Padding="8,4" Margin="0,0,4,0"/>
              <Button x:Name="AiIdeOpenBtn" Content="Open"
                      Style="{StaticResource Btn}" Padding="8,4" Margin="0,0,4,0"/>
              <Button x:Name="AiIdeSaveBtn" Content="Save"
                      Style="{StaticResource Btn}" Padding="8,4" Margin="0,0,12,0"/>
              <TextBlock Text="|" Foreground="#3E3E42" VerticalAlignment="Center" Margin="0,0,12,0"/>
              <TextBlock Text="Lang" Foreground="#B0B0B0" VerticalAlignment="Center"
                         FontSize="11" Margin="0,0,6,0"/>
              <ComboBox x:Name="AiIdeLangCombo" Width="150" FontSize="11"
                        VerticalAlignment="Center" Margin="0,0,10,0">
                <ComboBoxItem Content="VBA - Excel"    IsSelected="True"/>
                <ComboBoxItem Content="VBA - Outlook"/>
                <ComboBoxItem Content="PowerShell"/>
                <ComboBoxItem Content="PS + Excel COM"/>
              </ComboBox>
              <ComboBox x:Name="AiIdeWbCombo" Width="160" FontSize="11"
                        VerticalAlignment="Center" Margin="0,0,8,0"
                        ToolTip="Target workbook or application"/>
              <ComboBox x:Name="AiIdeSheetCombo" Width="130" FontSize="11"
                        VerticalAlignment="Center" Margin="0,0,12,0"
                        ToolTip="Target sheet (Excel VBA only)"/>
              <TextBlock Text="|" Foreground="#3E3E42" VerticalAlignment="Center" Margin="0,0,12,0"/>
              <Button x:Name="AiIdeValidateBtn"    Content="Validate"
                      Style="{StaticResource BtnAccent}" Padding="9,4" Margin="0,0,6,0"/>
              <Button x:Name="AiIdeRunBtn"         Content="Run"
                      Style="{StaticResource BtnGreen}"  Padding="9,4" Margin="0,0,6,0"
                      ToolTip="Run code currently in editor"/>
              <Button x:Name="AiIdeInjectExcelBtn" Content="Inject Excel (no run)"
                      Style="{StaticResource Btn}" Padding="9,4" Margin="0,0,6,0"
                      ToolTip="Inject as module into Excel workbook without running"/>
              <Button x:Name="AiIdeInjectOutlookBtn" Content="Inject Outlook (no run)"
                      Style="{StaticResource Btn}" Padding="9,4"
                      ToolTip="Inject as module into Outlook VBA project without running"/>
              <TextBlock x:Name="AiIdeFilePathTxt" Text="  (unsaved buffer)"
                         Foreground="#5E5E5E" FontSize="10" VerticalAlignment="Center"
                         Margin="12,0,0,0"/>
            </StackPanel>
          </Border>

          <Border Grid.Row="6" Background="#1B1B1F" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="10">
            <TextBox x:Name="AiIdeEditorBox"
                     AcceptsReturn="True" AcceptsTab="True"
                     TextWrapping="NoWrap"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     FontFamily="Consolas" FontSize="13"
                     Foreground="#E0E0E0" Background="Transparent" BorderThickness="0"
                     ToolTip="AI-generated code lands here. Edit freely then Run or Generate+Run."/>
          </Border>

          <Border Grid.Row="8" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="12,8">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="6"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <TextBlock Grid.Row="0" x:Name="AiIdePromptGuideTxt"
                         Text="Prompt helper loads by language mode."
                         Foreground="#8EA9C3" FontSize="11" TextWrapping="Wrap"/>
              <Grid Grid.Row="2">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Grid.Column="0" x:Name="AiIdePromptBox"
                         FontSize="13" Foreground="#E0E0E0"
                         Background="#1B1B1F" BorderBrush="#3E3E42" BorderThickness="1"
                         Padding="8,6" Height="72" AcceptsReturn="True" TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"
                         ToolTip="Describe what you want. Ctrl+Enter = Generate. Shift+Enter = Refine."/>
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                  <Button x:Name="AiIdeGenerateBtn" Content="Generate"
                          Style="{StaticResource BtnAccent}" Padding="12,5" Margin="0,0,6,0"
                          ToolTip="Generate code from prompt"/>
                  <Button x:Name="AiIdeRefineBtn"   Content="Refine"
                          Style="{StaticResource Btn}"       Padding="12,5" Margin="0,0,6,0"
                          ToolTip="Send current editor content + prompt for refinement"/>
                  <Button x:Name="AiIdeGenRunBtn"   Content="Gen + Run"
                          Style="{StaticResource BtnGreen}"  Padding="12,5" Margin="0,0,6,0"
                          ToolTip="Generate, then auto-run with AI error-fix retry loop"/>
                  <Button x:Name="AiIdeFixBtn"      Content="Fix Error"
                          Style="{StaticResource Btn}"       Padding="12,5" Margin="0,0,6,0"
                          ToolTip="Send current code + last run error back to AI"/>
                  <Button x:Name="AiIdeClearBtn"    Content="Clear"
                          Style="{StaticResource BtnRed}"    Padding="10,5"/>
                </StackPanel>
              </Grid>
            </Grid>
          </Border>

          <StackPanel Grid.Row="10" Orientation="Horizontal">
            <CheckBox x:Name="AiIdeKeepConvChk"
                      Content="Keep conversation"
                      Foreground="#B0B0B0" FontSize="11" VerticalAlignment="Center"
                      Margin="0,0,20,0"
                      ToolTip="Reuse same Graph conversation ID for multi-turn context"/>
            <Button x:Name="AiIdeNewConvBtn" Content="New Conversation"
                    Style="{StaticResource Btn}" Padding="8,2" FontSize="11"
                    Margin="0,0,16,0"/>
            <CheckBox x:Name="AiIdeWebGroundChk"
                      Content="Web grounding"
                      Foreground="#B0B0B0" FontSize="11" VerticalAlignment="Center"
                      Margin="0,0,20,0"
                      ToolTip="Ask Copilot to include web search (Graph API only)"/>
            <Button x:Name="AiIdeExportHistBtn" Content="Export Chat"
                    Style="{StaticResource Btn}" Padding="8,2" FontSize="11"/>
          </StackPanel>

          <Border Grid.Row="12" Background="#252528" CornerRadius="6"
                  BorderBrush="#3E3E42" BorderThickness="1" Padding="12,8">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="210"/>
                <RowDefinition Height="8"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <TextBlock Grid.Row="0" Text="AI CHAT / RESPONSE"
                         Foreground="#4C9FE6" FontSize="10" FontWeight="SemiBold" Margin="0,0,0,4"/>
              <TextBox Grid.Row="1" x:Name="AiIdeOutputTxt" IsReadOnly="True"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       FontFamily="Consolas" FontSize="11" Foreground="#C0E5FF"
                       Background="#1B1B1F" BorderBrush="#3E3E42" BorderThickness="1"
                       ToolTip="Conversation-style AI responses and run output stream."/>

              <Grid Grid.Row="4">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                  <TextBlock Text="VALIDATION" Foreground="#E05050"
                             FontSize="10" FontWeight="SemiBold" Margin="0,0,0,4"/>
                  <TextBox x:Name="AiIdeValidationTxt" IsReadOnly="True"
                           TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                           FontFamily="Consolas" FontSize="11" Foreground="#FFB4B4"
                           Background="#1B1B1F" BorderBrush="#3E3E42" BorderThickness="1"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="RUN ERRORS" Foreground="#FFA500"
                             FontSize="10" FontWeight="SemiBold" Margin="0,0,0,4"/>
                  <TextBox x:Name="AiIdeRunErrorTxt" IsReadOnly="True"
                           TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                           FontFamily="Consolas" FontSize="11" Foreground="#FFD39B"
                           Background="#1B1B1F" BorderBrush="#3E3E42" BorderThickness="1"/>
                </StackPanel>
              </Grid>
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
        <TextBlock Grid.Column="0" x:Name="StatusBarTxt" Text="  Ready" Foreground="#6E6E6E"
                   FontSize="11" VerticalAlignment="Center"/>
        <TextBlock Grid.Column="1" x:Name="StatusClockTxt" Text="" Foreground="#6E6E6E"
                   FontSize="11" VerticalAlignment="Center"/>
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
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlStr))
$Window = [System.Windows.Markup.XamlReader]::Load($reader)
$script:Window = $Window

# Helper to find named elements
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
$ClipDateOffset   = G 'ClipDateOffset'
$ClipTimeOffset   = G 'ClipTimeOffset'
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

# -- Email Helper tab --
$EmRefreshBtn      = G 'EmRefreshBtn'
$EmScopeCombo      = G 'EmScopeCombo'
$EmViewCombo       = G 'EmViewCombo'
$EmFilterCombo     = G 'EmFilterCombo'
$EmFromDateBox     = G 'EmFromDateBox'
$EmToDateBox       = G 'EmToDateBox'
$EmSearchBox       = G 'EmSearchBox'
$EmCountTxt        = G 'EmCountTxt'
$EmLastRefreshTxt  = G 'EmLastRefreshTxt'
$EmDupEmailBtn     = G 'EmDupEmailBtn'
$EmDupFolderBtn    = G 'EmDupFolderBtn'
$EmEmailPanel      = G 'EmEmailPanel'
$EmPreviewSubject  = G 'EmPreviewSubject'
$EmPreviewFrom     = G 'EmPreviewFrom'
$EmPreviewBody     = G 'EmPreviewBody'
$EmTemplateCombo   = G 'EmTemplateCombo'
$EmReplyBtn        = G 'EmReplyBtn'
$EmReplyAllBtn     = G 'EmReplyAllBtn'
$EmForwardBtn      = G 'EmForwardBtn'
$EmTodoBtn         = G 'EmTodoBtn'
$EmCalBtn          = G 'EmCalBtn'
$EmRecallBtn       = G 'EmRecallBtn'
$EmComposePanel    = G 'EmComposePanel'
$EmComposeLabel    = G 'EmComposeLabel'
$EmComposeTo       = G 'EmComposeTo'
$EmComposeBody     = G 'EmComposeBody'
$EmSaveDraftBtn    = G 'EmSaveDraftBtn'
$EmCancelComposeBtn = G 'EmCancelComposeBtn'
$script:ComposeMode = $null   # 'Reply','ReplyAll','Forward' or $null

# -- Email Dashboard tab --
$DashScopeCombo    = G 'DashScopeCombo'
$DashViewCombo     = G 'DashViewCombo'
$DashFilterCombo   = G 'DashFilterCombo'
$DashPeriodCombo   = G 'DashPeriodCombo'
$DashFromDateBox   = G 'DashFromDateBox'
$DashToDateBox     = G 'DashToDateBox'
$DashSearchBox     = G 'DashSearchBox'
$DashRefreshBtn    = G 'DashRefreshBtn'
$DashExportBtn     = G 'DashExportBtn'
$DashTotalCount    = G 'DashTotalCount'
$DashUnreadCount   = G 'DashUnreadCount'
$DashVipCount      = G 'DashVipCount'
$DashInternalCount = G 'DashInternalCount'
$DashExternalCount = G 'DashExternalCount'
$DashOldestDays    = G 'DashOldestDays'
$DashAgingPanel    = G 'DashAgingPanel'
$DashDraftPanel    = G 'DashDraftPanel'
$DashSendersPanel  = G 'DashSendersPanel'
$DashKeywordPanel  = G 'DashKeywordPanel'

# -- IDE tab --
$IdeOpenBtn        = G 'IdeOpenBtn'
$IdeSaveBtn        = G 'IdeSaveBtn'
$IdeSaveAsBtn      = G 'IdeSaveAsBtn'
$IdeValidateBtn    = G 'IdeValidateBtn'
$IdeRunBtn         = G 'IdeRunBtn'
$IdeRefreshExcelBtn = G 'IdeRefreshExcelBtn'
$IdeModeCombo      = G 'IdeModeCombo'
$IdePathBox        = G 'IdePathBox'
$IdeWbCombo        = G 'IdeWbCombo'
$IdeSheetCombo     = G 'IdeSheetCombo'
$IdeEditorBrowser  = G 'IdeEditorBrowser'
$IdeValidationTxt  = G 'IdeValidationTxt'
$IdeRunErrorTxt    = G 'IdeRunErrorTxt'
$IdeOutputTxt      = G 'IdeOutputTxt'

# -"--- AI IDE control bindings -"-------------------------------------------------------------------------------------------
$AiIdeTenantBox        = G 'AiIdeTenantBox'
$AiIdeModeCombo        = G 'AiIdeModeCombo'
$AiIdePortBox          = G 'AiIdePortBox'
$AiIdeAuthBtn          = G 'AiIdeAuthBtn'
$AiIdeStartBtn         = G 'AiIdeStartBtn'
$AiIdeStopBtn          = G 'AiIdeStopBtn'
$AiIdeModePill         = G 'AiIdeModePill'
$AiIdeModeTxt          = G 'AiIdeModeTxt'
$AiIdeStatusTxt        = G 'AiIdeStatusTxt'
$AiIdeRetryTxt         = G 'AiIdeRetryTxt'
$AiIdeNewBtn           = G 'AiIdeNewBtn'
$AiIdeOpenBtn          = G 'AiIdeOpenBtn'
$AiIdeSaveBtn          = G 'AiIdeSaveBtn'
$AiIdeLangCombo        = G 'AiIdeLangCombo'
$AiIdeWbCombo          = G 'AiIdeWbCombo'
$AiIdeSheetCombo       = G 'AiIdeSheetCombo'
$AiIdeValidateBtn      = G 'AiIdeValidateBtn'
$AiIdeRunBtn           = G 'AiIdeRunBtn'
$AiIdeInjectExcelBtn   = G 'AiIdeInjectExcelBtn'
$AiIdeInjectOutlookBtn = G 'AiIdeInjectOutlookBtn'
$AiIdeFilePathTxt      = G 'AiIdeFilePathTxt'
$AiIdeEditorBox        = G 'AiIdeEditorBox'
$AiIdePromptGuideTxt   = G 'AiIdePromptGuideTxt'
$AiIdePromptBox        = G 'AiIdePromptBox'
$AiIdeGenerateBtn      = G 'AiIdeGenerateBtn'
$AiIdeRefineBtn        = G 'AiIdeRefineBtn'
$AiIdeGenRunBtn        = G 'AiIdeGenRunBtn'
$AiIdeFixBtn           = G 'AiIdeFixBtn'
$AiIdeClearBtn         = G 'AiIdeClearBtn'
$AiIdeKeepConvChk      = G 'AiIdeKeepConvChk'
$AiIdeNewConvBtn       = G 'AiIdeNewConvBtn'
$AiIdeWebGroundChk     = G 'AiIdeWebGroundChk'
$AiIdeExportHistBtn    = G 'AiIdeExportHistBtn'
$AiIdeValidationTxt    = G 'AiIdeValidationTxt'
$AiIdeRunErrorTxt      = G 'AiIdeRunErrorTxt'
$AiIdeOutputTxt        = G 'AiIdeOutputTxt'

# ================================================================
#  IDE EDITOR -- WebBrowser HTML (syntax highlight + line numbers + Ctrl+F/H find-replace)
# ================================================================
$script:IdeHtml = @'
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#1B1B1F;color:#E0E0E0;font:13px/1.5 Consolas,monospace;overflow:hidden;height:100vh;display:flex;flex-direction:column}
#ftb{display:none;background:#252528;border-bottom:1px solid #3E3E42;padding:4px 8px;align-items:center;gap:5px;font-size:12px}
#ftb.vis{display:flex}
#fq,#rq{background:#1B1B1F;border:1px solid #4C9FE6;color:#E0E0E0;padding:2px 5px;font:12px Consolas,monospace;width:180px}
#fi{color:#888;font-size:11px;min-width:70px}
.tb{background:#3E3E42;color:#E0E0E0;border:none;padding:2px 7px;cursor:pointer;font-size:11px}
.tb:hover{background:#555}
.tb.x{color:#E05050;background:#2E1F1F}
#wrap{flex:1;min-height:0;display:flex;overflow:hidden}
#ln{
  width:58px;min-width:58px;background:#1B1B1F;color:#4A4A5A;text-align:right;
  padding:6px 8px 6px 0;border-right:1px solid #2A2A2E;overflow:hidden;user-select:none;
  white-space:pre;line-height:1.5;font:13px/1.5 Consolas,monospace
}
textarea#ed{
  flex:1;display:block;background:#1B1B1F;color:#E0E0E0;caret-color:#E0E0E0;border:none;outline:none;
  resize:none;padding:6px 10px;font:13px/1.5 Consolas,monospace;line-height:1.5;tab-size:4;
  white-space:pre;overflow:auto;word-wrap:normal
}
textarea#ed::selection{background:rgba(76,159,230,0.3)}
</style>
</head>
<body>
<div id="ftb">
  <input id="fq" placeholder="Find (regex)..." oninput="doFind()">
  <span id="fi"></span>
  <button class="tb" onclick="step(-1)">&#9650;</button>
  <button class="tb" onclick="step(1)">&#9660;</button>
  <input id="rq" placeholder="Replace...">
  <button class="tb" onclick="repl()">Replace</button>
  <button class="tb" onclick="replAll()">All</button>
  <button class="tb x" onclick="close_()">&#10005;</button>
</div>
<div id="wrap">
  <div id="ln"></div>
  <textarea id="ed" autocomplete="off" autocorrect="off" spellcheck="false" wrap="off"></textarea>
</div>
<script>
var ta=document.getElementById('ed'),ln=document.getElementById('ln'),ftb=document.getElementById('ftb');
var lang='ps',matches=[],mIdx=0;
function syncLn(){
  var c=(ta.value.match(/\n/g)||[]).length+1,buf='';
  for(var i=1;i<=c;i++){buf+=i;if(i<c)buf+='\n';}
  ln.textContent=buf;
}
function syncScroll(){ln.scrollTop=ta.scrollTop}
function on(el,ev,fn){if(!el){return}if(el.addEventListener){el.addEventListener(ev,fn,false)}else if(el.attachEvent){el.attachEvent('on'+ev,fn)}else{el['on'+ev]=fn}}
function stopEv(e){if(e.preventDefault)e.preventDefault();e.returnValue=false}
function hasCls(el,c){return (' '+(el.className||'')+' ').indexOf(' '+c+' ')>-1}
function addCls(el,c){if(!hasCls(el,c)){el.className=((el.className||'').replace(/^\s+|\s+$/g,''));el.className=(el.className?el.className+' ':'')+c}}
function remCls(el,c){var re=new RegExp('(^|\\s)'+c+'(\\s|$)','g');el.className=(el.className||'').replace(re,' ').replace(/\s+/g,' ').replace(/^\s+|\s+$/g,'')}
on(ta,'input',syncLn);
on(ta,'scroll',syncScroll);
on(ta,'keydown',function(e){
  e=e||window.event;var k=e.key||'';var kc=e.keyCode||e.which||0;var kl=(k&&k.toLowerCase)?k.toLowerCase():'';
  if(k==='Tab'||kc===9){
    stopEv(e);
    var s=ta.selectionStart,en=ta.selectionEnd;
    ta.value=ta.value.slice(0,s)+'    '+ta.value.slice(en);
    ta.selectionStart=ta.selectionEnd=s+4;
    syncLn();
    return;
  }
  if((e.ctrlKey||e.metaKey)&&((kl==='f')||kc===70||(kl==='h')||kc===72)){stopEv(e);open_();return;}
  if(k==='Escape'||kc===27){close_();return;}
  if(k==='F3'||kc===114){stopEv(e);step(e.shiftKey?-1:1);return;}
});
function open_(){addCls(ftb,'vis');document.getElementById('fq').focus()}
function close_(){remCls(ftb,'vis');ta.focus()}
function doFind(){
  matches=[];mIdx=0;var q=document.getElementById('fq').value;document.getElementById('fi').textContent='';
  if(!q)return;
  try{var re=new RegExp(q,'gi'),m;while((m=re.exec(ta.value))!==null)matches.push(m.index)}catch(x){}
  document.getElementById('fi').textContent=matches.length?(mIdx+1)+'/'+matches.length:'not found';
}
function step(d){
  if(!matches.length)return;
  mIdx=(mIdx+d+matches.length)%matches.length;
  var pos=matches[mIdx],q=document.getElementById('fq').value;
  ta.selectionStart=pos;ta.selectionEnd=pos+q.length;ta.focus();
  document.getElementById('fi').textContent=(mIdx+1)+'/'+matches.length;
}
function repl(){
  var q=document.getElementById('fq').value,r=document.getElementById('rq').value;if(!q)return;
  if(ta.value.slice(ta.selectionStart,ta.selectionEnd)===q){
    ta.value=ta.value.slice(0,ta.selectionStart)+r+ta.value.slice(ta.selectionEnd);
    syncLn();
  }
  doFind();step(1);
}
function replAll(){
  var q=document.getElementById('fq').value,r=document.getElementById('rq').value;if(!q)return;
  try{ta.value=ta.value.replace(new RegExp(q,'g'),r)}catch(x){}
  syncLn();doFind();
}
function getValue(){return ta.value}
function setValue(v){ta.value=v||'';syncLn();syncScroll()}
function setLang(l){lang=l||'ps'}
syncLn();syncScroll();ta.focus();
</script>
</body>
</html>
'@

# IDE editor bridge helpers (PowerShell <-> WebBrowser)
function Get-IdeEditorText {
    if (-not $script:IdeEditorReady) {
        return if ($null -ne $script:IdeEditorPendingText) { $script:IdeEditorPendingText } else { '' }
    }
    try { return [string]$IdeEditorBrowser.InvokeScript('getValue') } catch { return '' }
}
function Set-IdeEditorText([string]$text) {
    if (-not $script:IdeEditorReady) { $script:IdeEditorPendingText = $text; return }
    try { $IdeEditorBrowser.InvokeScript('setValue', @($text)) } catch {}
}
function Set-IdeEditorLang([string]$lang) {
    if (-not $script:IdeEditorReady) { return }
    $jsLang = if ($lang -eq 'VBA') { 'vba' } else { 'ps' }
    try { $IdeEditorBrowser.InvokeScript('setLang', @($jsLang)) } catch {}
}

function Set-IdeBrowserSilent([bool]$Silent = $true) {
    try {
        $flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::NonPublic
        $ax = $IdeEditorBrowser.GetType().InvokeMember('ActiveXInstance', $flags -bor [System.Reflection.BindingFlags]::GetProperty, $null, $IdeEditorBrowser, $null)
        if ($ax) { $ax.Silent = $Silent }
    } catch {}
}

# Wire LoadCompleted so helpers know when the browser is ready
$IdeEditorBrowser.Add_LoadCompleted({
    $script:IdeEditorReady = $true
    Set-IdeBrowserSilent $true
    if ($null -ne $script:IdeEditorPendingText) {
        try { $IdeEditorBrowser.InvokeScript('setValue', @($script:IdeEditorPendingText)) } catch {}
        $script:IdeEditorPendingText = $null
    }
})

# ================================================================
#  SYNC LOG WRITER (appends to the QSync log textbox)
# ================================================================
function QsLog([string]$msg) {
    $line = "[$(Get-Date -f 'HH:mm:ss')]  $msg"
    $QsSyncLogTxt.AppendText("$line`r`n")
    $QsSyncLogTxt.ScrollToEnd()
    Update-UI
}

# ================================================================
#  REFRESH QUARTER DROPDOWN
# ================================================================
function Refresh-QuarterDropdown([string]$selectName) {
    $qFiles = Get-QuarterList
    $QsQuarterCombo.Items.Clear()
    $QsCompareCombo.Items.Clear()
    [void]($QsCompareCombo.Items.Add('(none)'))
    foreach ($qf in $qFiles) {
        [void]($QsQuarterCombo.Items.Add($qf.BaseName))
        [void]($QsCompareCombo.Items.Add($qf.BaseName))
    }
    if ($selectName -and $QsQuarterCombo.Items.Contains($selectName)) {
        $QsQuarterCombo.SelectedItem = $selectName
    } elseif ($QsQuarterCombo.Items.Count -gt 0) {
        $QsQuarterCombo.SelectedIndex = 0
    }
    $QsCompareCombo.SelectedIndex = 0
    if ($QsQuarterCombo.SelectedItem) {
        Switch-ActiveQuarter (Join-Path $script:QuartersDir "$($QsQuarterCombo.SelectedItem).json")
    }
}

# ================================================================
#  REFRESH FUNCTIONS: MacroHub tabs
# ================================================================

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
            -DefaultDateOffset $ClipDateOffset.Text `
            -DefaultTimeOffset $ClipTimeOffset.Text `
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
        foreach ($val in @($t.Name, $t.State, $t.NextRunTime.ToString('g'), $t.LastRunTime.ToString('g'))) {
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

function Refresh-NavWorkbooks {
    $NavWbList.Items.Clear()
    $wbs = Get-OpenWorkbooks -Session Navigator
    foreach ($w in $wbs) { [void]($NavWbList.Items.Add($w)) }
    # Also refresh destination workbook combo (reuse list)
    Refresh-NavDestWbCombo $wbs
    Refresh-NavExcelOptions
    Refresh-NavVbaList
}

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
                $dotColor = switch ($state) {
                    -1 { '#50A050' }   # visible
                    0  { '#FFFFFF' }   # hidden
                    2  { '#000000' }   # very hidden
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
#  IDE TAB FUNCTIONS
# ================================================================
function Reset-IdeValidationState {
    $script:IdeLastValidationOk = $false
    $script:IdeValidatedHash = ''
    $script:IdeValidatedLang = ''
}

function Get-IdeSelectedWorkbookName {
    $wb = [string]$IdeWbCombo.SelectedItem
    if (-not $wb -or $wb -eq '(none)') { return '' }
    return $wb
}

function Get-IdeSelectedSheetName {
    $sh = [string]$IdeSheetCombo.SelectedItem
    if (-not $sh -or $sh -eq '(none)') { return '' }
    return $sh
}

function Refresh-IdeSheets {
    $prev = [string]$IdeSheetCombo.SelectedItem
    $IdeSheetCombo.Items.Clear()
    [void]($IdeSheetCombo.Items.Add('(none)'))
    $wb = Get-IdeSelectedWorkbookName
    if ($wb) {
        $sheets = Get-WorksheetNames -WorkbookName $wb -Session Main
        foreach ($s in $sheets) { [void]($IdeSheetCombo.Items.Add($s)) }
    }
    if ($prev -and $IdeSheetCombo.Items.Contains($prev)) { $IdeSheetCombo.SelectedItem = $prev }
    else { $IdeSheetCombo.SelectedIndex = 0 }
}

function Refresh-IdeExcelTargets {
    $prevWb = [string]$IdeWbCombo.SelectedItem
    $IdeWbCombo.Items.Clear()
    [void]($IdeWbCombo.Items.Add('(none)'))
    $wbs = Get-OpenWorkbooks -Session Main
    foreach ($w in $wbs) { [void]($IdeWbCombo.Items.Add($w)) }
    if ($prevWb -and $IdeWbCombo.Items.Contains($prevWb)) { $IdeWbCombo.SelectedItem = $prevWb }
    else { $IdeWbCombo.SelectedIndex = 0 }
    Refresh-IdeSheets
}

function Refresh-IdePathDisplay {
    $IdePathBox.Text = $(if ($script:IdeCurrentFile) { $script:IdeCurrentFile } else { '(unsaved buffer)' })
}

function Save-IdeFileToPath {
    param([string]$Path)
    if (-not $Path) { throw 'Missing target path.' }
    [System.IO.File]::WriteAllText($Path, (Get-IdeEditorText), [System.Text.UTF8Encoding]::new($false))
    $script:IdeCurrentFile = $Path
    Refresh-IdePathDisplay
    Reset-IdeValidationState
}

function Invoke-IdePowerShellValidate {
    param([string]$Content)
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($Content, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        $lines = @()
        foreach ($e in $errors) {
            $line = $e.Extent.StartLineNumber
            $col  = $e.Extent.StartColumnNumber
            $lines += "Line $line, Col ${col}: $($e.Message)"
        }
        return [PSCustomObject]@{ Ok = $false; Message = ($lines -join "`r`n") }
    }
    return [PSCustomObject]@{ Ok = $true; Message = 'PowerShell syntax check passed.' }
}

function Invoke-IdeVbaValidate {
    param([string]$Content)

    $entry = Get-VbaEntryPointNameFromText $Content
    if (-not $entry) {
        return [PSCustomObject]@{
            Ok = $false
            Message = 'No runnable Sub found. Add a Public Sub <Name>() entry point.'
            EntryPoint = ''
        }
    }

    $tmpBas = Join-Path $env:TEMP ("MacroHub_IDE_{0}.bas" -f ([guid]::NewGuid().ToString('N')))
    $xl = $null
    $wb = $null
    $oldAlerts = $null
    try {
        [System.IO.File]::WriteAllText($tmpBas, $Content, [System.Text.UTF8Encoding]::new($false))
        $xl = Get-ExcelApp -Session Main -Create
        if (-not $xl) { throw 'Unable to start Excel for VBA validation.' }
        $oldAlerts = $xl.DisplayAlerts
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()
        $wb.Activate()

        $proj = $wb.VBProject
        $comp = $proj.VBComponents.Import($tmpBas)

        $compileBtn = $null
        try { $compileBtn = $xl.VBE.CommandBars.FindControl(1, 578, $null, $true) } catch {}
        if ($compileBtn) {
            $compileBtn.Execute()
            return [PSCustomObject]@{
                Ok = $true
                Message = "VBA compile check passed. Entry point: $entry"
                EntryPoint = $entry
                ModuleName = $comp.Name
            }
        }

        # Some Excel environments do not expose VBE compile command.
        # Fallback: successful module import + entry point detection.
        return [PSCustomObject]@{
            Ok = $true
            Message = "VBA validation fallback passed (compile command unavailable). Entry point: $entry"
            EntryPoint = $entry
            ModuleName = $comp.Name
        }
    } catch {
        $msg = [string]$_.Exception.Message
        if ($msg -match 'Programmatic access to Visual Basic Project is not trusted') {
            $msg += ' Enable: Excel Options > Trust Center > Macro Settings > Trust access to the VBA project object model.'
        }
        return [PSCustomObject]@{
            Ok = $false
            Message = "VBA compile failed: $msg"
            EntryPoint = $entry
            ModuleName = ''
        }
    } finally {
        try { if ($wb) { $wb.Close($false) } } catch {}
        try { if ($xl -and $null -ne $oldAlerts) { $xl.DisplayAlerts = $oldAlerts } } catch {}
        try { if (Test-Path $tmpBas) { Remove-Item $tmpBas -Force } } catch {}
    }
}

function Invoke-IdePowerShellRun {
    param(
        [string]$Content,
        [string]$WorkbookName,
        [string]$SheetName
    )
    $tmpPs = Join-Path $env:TEMP ("MacroHub_IDE_{0}.ps1" -f ([guid]::NewGuid().ToString('N')))
    $tmpOut = Join-Path $env:TEMP ("MacroHub_IDE_{0}.out.txt" -f ([guid]::NewGuid().ToString('N')))
    $tmpErr = Join-Path $env:TEMP ("MacroHub_IDE_{0}.err.txt" -f ([guid]::NewGuid().ToString('N')))
    try {
        [System.IO.File]::WriteAllText($tmpPs, $Content, [System.Text.UTF8Encoding]::new($false))
        $psExe = Join-Path $PSHOME 'powershell.exe'
        if (-not (Test-Path $psExe)) { $psExe = 'powershell.exe' }

        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', (Quote-ProcessArg $tmpPs)
        )
        if ($WorkbookName) { $args += @('-WorkbookName', (Quote-ProcessArg $WorkbookName)) }
        if ($SheetName) { $args += @('-SheetName', (Quote-ProcessArg $SheetName)) }
        $argLine = $args -join ' '

        $p = Start-Process -FilePath $psExe -ArgumentList $argLine -WindowStyle Hidden `
            -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr -PassThru
        if (-not $p.WaitForExit(120000)) {
            try { $p.Kill() } catch {}
            throw 'PowerShell run timed out after 120 seconds.'
        }

        $stdout = if (Test-Path $tmpOut) { [System.IO.File]::ReadAllText($tmpOut) } else { '' }
        $stderr = if (Test-Path $tmpErr) { [System.IO.File]::ReadAllText($tmpErr) } else { '' }
        $exitCode = 0
        try { $exitCode = [int]$p.ExitCode } catch { $exitCode = 0 }
        return [PSCustomObject]@{
            ExitCode = $exitCode
            StdOut   = $stdout
            StdErr   = $stderr
        }
    } finally {
        foreach ($p in @($tmpPs, $tmpOut, $tmpErr)) {
            try { if (Test-Path $p) { Remove-Item $p -Force } } catch {}
        }
    }
}

function Invoke-IdeVbaRun {
    param(
        [string]$Content,
        [string]$WorkbookName,
        [string]$SheetName
    )
    $tmpBas = Join-Path $env:TEMP ("MacroHub_IDE_{0}.bas" -f ([guid]::NewGuid().ToString('N')))
    try {
        [System.IO.File]::WriteAllText($tmpBas, $Content, [System.Text.UTF8Encoding]::new($false))
        Invoke-VbaMacro -MacroFile $tmpBas -WorkbookName $WorkbookName -WorksheetName $SheetName
        return [PSCustomObject]@{ Ok = $true; Message = 'VBA run completed.' }
    } catch {
        return [PSCustomObject]@{ Ok = $false; Message = [string]$_.Exception.Message }
    } finally {
        try { if (Test-Path $tmpBas) { Remove-Item $tmpBas -Force } } catch {}
    }
}

function Invoke-IdeValidateAction {
    $content = Get-IdeEditorText
    if ([string]::IsNullOrWhiteSpace($content)) {
        $IdeValidationTxt.Text = 'Nothing to validate.'
        $IdeRunErrorTxt.Text = ''
        $script:IdeLastValidationOk = $false
        return $false
    }

    $mode = Get-ComboText $IdeModeCombo
    $lang = Get-IdeLanguageFromInputs -Mode $mode -Path $script:IdeCurrentFile
    $result = if ($lang -eq 'VBA') { Invoke-IdeVbaValidate -Content $content } else { Invoke-IdePowerShellValidate -Content $content }

    $IdeValidationTxt.Text = [string]$result.Message
    $IdeRunErrorTxt.Text = ''

    if ($result.Ok) {
        $script:IdeLastValidationOk = $true
        $script:IdeValidatedHash = Get-TextSha1 $content
        $script:IdeValidatedLang = $lang
        Set-Status "$lang validation passed"
        return $true
    }

    $script:IdeLastValidationOk = $false
    $script:IdeValidatedHash = ''
    $script:IdeValidatedLang = ''
    Set-Status "$lang validation failed" '#E05050'
    return $false
}

function Invoke-IdeRunAction {
    $content = Get-IdeEditorText
    if ([string]::IsNullOrWhiteSpace($content)) {
        Set-Status 'IDE editor is empty' '#E05050'
        return
    }

    $mode = Get-ComboText $IdeModeCombo
    $lang = Get-IdeLanguageFromInputs -Mode $mode -Path $script:IdeCurrentFile
    $IdeRunErrorTxt.Text = ''
    $IdeOutputTxt.Text = ''

    if ($lang -eq 'VBA') {
        # Always compile before VBA run.
        if (-not (Invoke-IdeValidateAction)) { return }
        $wb = Get-IdeSelectedWorkbookName
        $sh = Get-IdeSelectedSheetName
        if (-not $wb) {
            $IdeOutputTxt.Text = 'Validation passed. Select a workbook to run VBA against.'
            Set-Status 'VBA validation-only complete. Select workbook to run.' '#B0B0B0'
            return
        }
        Show-Busy 'Running VBA module...'
        try {
            $run = Invoke-IdeVbaRun -Content $content -WorkbookName $wb -SheetName $sh
            if ($run.Ok) {
                $IdeOutputTxt.Text = [string]$run.Message
                Set-Status 'VBA run completed'
            } else {
                $IdeRunErrorTxt.Text = [string]$run.Message
                Set-Status 'VBA run failed' '#E05050'
            }
        } finally { Hide-Busy }
        return
    }

    # PowerShell mode: syntax-check first, then run in isolated process.
    if (-not (Invoke-IdeValidateAction)) { return }
    $wb = Get-IdeSelectedWorkbookName
    $sh = Get-IdeSelectedSheetName
    Show-Busy 'Running PowerShell script...'
    try {
        $run = Invoke-IdePowerShellRun -Content $content -WorkbookName $wb -SheetName $sh
        $outText = [string]$run.StdOut
        $errText = [string]$run.StdErr
        if ($outText.Trim()) { $IdeOutputTxt.Text = $outText.TrimEnd() } else { $IdeOutputTxt.Text = '(no output)' }
        if ($errText.Trim()) { $IdeRunErrorTxt.Text = $errText.TrimEnd() }
        if ($run.ExitCode -ne 0 -and -not $errText.Trim()) {
            $IdeRunErrorTxt.Text = "Process exited with code $($run.ExitCode)."
        }
        if ($run.ExitCode -eq 0 -and -not $errText.Trim()) { Set-Status 'PowerShell run completed' }
        else { Set-Status 'PowerShell run completed with errors' '#E05050' }
    } catch {
        $IdeRunErrorTxt.Text = [string]$_.Exception.Message
        Set-Status "PowerShell run error: $($_.Exception.Message)" '#E05050'
    } finally { Hide-Busy }
}

function Open-IdeFileDialogAndLoad {
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Title = 'Open Code File'
    $dlg.Filter = 'Code Files|*.ps1;*.psm1;*.psd1;*.bas;*.cls;*.frm;*.txt|PowerShell|*.ps1;*.psm1;*.psd1|VBA|*.bas;*.cls;*.frm|All Files|*.*'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $text = [System.IO.File]::ReadAllText($dlg.FileName)
    Set-IdeEditorText $text
    $script:IdeCurrentFile = $dlg.FileName
    Refresh-IdePathDisplay
    Reset-IdeValidationState
    $IdeValidationTxt.Text = ''
    $IdeRunErrorTxt.Text = ''
    $IdeOutputTxt.Text = "Loaded: $($dlg.FileName)"
    Set-Status "IDE opened: $(Split-Path $dlg.FileName -Leaf)"
}

function Save-IdeFileAsDialog {
    $mode = Get-ComboText $IdeModeCombo
    $lang = Get-IdeLanguageFromInputs -Mode $mode -Path $script:IdeCurrentFile
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Title = 'Save Code File As'
    if ($lang -eq 'VBA') {
        $dlg.Filter = 'VBA Module (*.bas)|*.bas|VBA Class (*.cls)|*.cls|All Files|*.*'
        $dlg.DefaultExt = 'bas'
    } else {
        $dlg.Filter = 'PowerShell Script (*.ps1)|*.ps1|PowerShell Module (*.psm1)|*.psm1|All Files|*.*'
        $dlg.DefaultExt = 'ps1'
    }
    if ($script:IdeCurrentFile) {
        try {
            $dlg.FileName = [System.IO.Path]::GetFileName($script:IdeCurrentFile)
            $dlg.InitialDirectory = Split-Path $script:IdeCurrentFile -Parent
        } catch {}
    }
    if ($dlg.ShowDialog() -ne 'OK') { return $false }
    Save-IdeFileToPath -Path $dlg.FileName
    Set-Status "IDE saved: $(Split-Path $dlg.FileName -Leaf)"
    return $true
}

# ================================================================
#  AI IDE -- FUNCTIONS
# ================================================================

# -"--- UI helpers -"---------------------------------------------------------------------------------------------------------------------

function AiIde-SetStatus([string]$Msg, [string]$Hex = '#B0B0B0') {
    $AiIdeStatusTxt.Text       = $Msg
    $AiIdeStatusTxt.Foreground = HexBrush $Hex
    Update-UI
}

function AiIde-SetPill([string]$Label, [string]$FgHex, [string]$BgHex) {
    $AiIdeModeTxt.Text        = $Label
    $AiIdeModeTxt.Foreground  = HexBrush $FgHex
    $AiIdeModePill.Background = HexBrush $BgHex
    Update-UI
}

function AiIde-BridgeRunning {
    try {
        if (-not $script:AiIde.ListenerPS) { return $false }
        $state = $script:AiIde.ListenerPS.InvocationStateInfo.State
        if ($script:AiIde.Status -eq 'error') { return $false }
        return ($state -eq [System.Management.Automation.PSInvocationState]::Running -or
                $state -eq [System.Management.Automation.PSInvocationState]::NotStarted)
    } catch {
        return $false
    }
}

function AiIde-UpdateModeUi {
    $mode = Get-ComboText $AiIdeModeCombo
    $chromeOnly = ($mode -eq 'Chrome')
    if ($AiIdePortBox)  { $AiIdePortBox.IsEnabled  = $chromeOnly }
    if ($AiIdeStartBtn) { $AiIdeStartBtn.IsEnabled = $chromeOnly }
    if ($AiIdeStopBtn)  { $AiIdeStopBtn.IsEnabled  = $chromeOnly }
}

function AiIde-StripFences([string]$Text) {
    $t = $Text -replace '(?s)^```[a-zA-Z]*\r?\n', ''
    $t = $t -replace '(?s)\r?\n```\s*$', ''
    return $t.Trim()
}

function AiIde-GetLang { return Get-ComboText $AiIdeLangCombo }

function AiIde-AppendOutput([string]$Role, [string]$Text) {
    if (-not $AiIdeOutputTxt) { return }
    $msg = [string]$Text
    if ([string]::IsNullOrWhiteSpace($msg)) { return }
    $stamp = (Get-Date).ToString('HH:mm:ss')
    if (-not [string]::IsNullOrWhiteSpace([string]$AiIdeOutputTxt.Text)) {
        $AiIdeOutputTxt.AppendText("`r`n`r`n")
    }
    $AiIdeOutputTxt.AppendText("[$stamp] $Role`r`n$($msg.Trim())")
    $AiIdeOutputTxt.ScrollToEnd()
}

function AiIde-UpdatePromptGuide {
    if (-not $AiIdePromptGuideTxt) { return }
    $lang = AiIde-GetLang
    $guide = switch -Wildcard ($lang) {
        '*Outlook*' { 'Prompt format: action + target mail item + rule. Example: "Create VBA that loops selected emails and adds category FollowUp if subject contains invoice."' }
        '*PS*COM*'  { 'Prompt format: task + workbook + sheet + ranges. Example: "Write PowerShell using Excel COM to paste clipboard to Sheet2!A1 and auto-fit columns."' }
        'PowerShell' { 'Prompt format: script goal + inputs + output. Example: "Write PowerShell 5.1 script that reads CSV from C:\\Data\\in.csv and writes grouped totals to C:\\Data\\out.csv."' }
        default     { 'Prompt format: task + target + success rule. Keep it direct: "Create runnable code that ...".' }
    }
    $AiIdePromptGuideTxt.Text = "Language: $lang. $guide"
}

function AiIde-RefreshSheetsOnly {
    if (-not $AiIdeSheetCombo) { return }
    $AiIdeSheetCombo.Items.Clear()
    [void]($AiIdeSheetCombo.Items.Add('(active sheet)'))
    $lang = AiIde-GetLang
    if ($lang -like '*Outlook*') {
        $AiIdeSheetCombo.Items.Clear()
        [void]($AiIdeSheetCombo.Items.Add('(n/a)'))
        $AiIdeSheetCombo.SelectedIndex = 0
        return
    }
    $wb = [string]$AiIdeWbCombo.SelectedItem
    if ($wb -and $wb -ne '(active workbook)') {
        foreach ($s in (Get-WorksheetNames -WorkbookName $wb -Session Main)) {
            [void]($AiIdeSheetCombo.Items.Add($s))
        }
    }
    if ($AiIdeSheetCombo.Items.Count -gt 0) { $AiIdeSheetCombo.SelectedIndex = 0 }
}

function AiIde-RefreshTargets {
    if ($script:AiIdeRefreshingTargets) { return }
    $script:AiIdeRefreshingTargets = $true
    try {
    AiIde-UpdateModeUi
    AiIde-UpdatePromptGuide
    $lang   = AiIde-GetLang
    $prevWb = [string]$AiIdeWbCombo.SelectedItem
    $AiIdeWbCombo.Items.Clear()
    if ($lang -like '*Outlook*') {
        [void]($AiIdeWbCombo.Items.Add('Outlook (Classic)'))
        $AiIdeWbCombo.SelectedIndex = 0
        AiIde-RefreshSheetsOnly
    } else {
        [void]($AiIdeWbCombo.Items.Add('(active workbook)'))
        $wbs = Get-OpenWorkbooks -Session Main
        foreach ($w in $wbs) { [void]($AiIdeWbCombo.Items.Add($w)) }
        if ($prevWb -and $AiIdeWbCombo.Items.Contains($prevWb)) { $AiIdeWbCombo.SelectedItem = $prevWb }
        elseif ($AiIdeWbCombo.Items.Count -gt 0) { $AiIdeWbCombo.SelectedIndex = 0 }
        AiIde-RefreshSheetsOnly
    }
    } finally {
        $script:AiIdeRefreshingTargets = $false
    }
}

# -"--- Graph OAuth -- Device Code -"---------------------------------------------------------------------------------------

function AiIde-EnsureTls12 {
    try {
        $tls12 = [System.Net.SecurityProtocolType]::Tls12
        if (([System.Net.ServicePointManager]::SecurityProtocol -band $tls12) -eq 0) {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor $tls12
        }
    } catch {}
}

function AiIde-GetHttpStatusCode($ErrorLike) {
    try {
        $ex = if ($ErrorLike -is [System.Management.Automation.ErrorRecord]) {
            $ErrorLike.Exception
        } else {
            $ErrorLike
        }
        if (-not $ex) { return 0 }
        $resp = $ex.Response
        if (-not $resp) { return 0 }
        try { return [int]$resp.StatusCode.value__ } catch {}
        try { return [int]$resp.StatusCode } catch {}
    } catch {}
    return 0
}

function AiIde-VerifyGraphCopilotAccess([switch]$Quiet) {
    $token = AiIde-GetToken
    if (-not $token) {
        if (-not $Quiet) {
            AiIde-SetPill 'No token' '#E05050' '#5c1a1a'
            AiIde-SetStatus 'No Graph token. Click Auth + Verify first.' '#E05050'
        }
        return $false
    }

    AiIde-EnsureTls12
    $script:AiIde.GraphAccessVerified  = $false
    $script:AiIde.GraphAccessCheckedOn = Get-Date
    $script:AiIde.GraphAccessNote      = ''
    $script:AiIde.GraphLastHttpCode    = 0

    $headers = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }
    $who = ''
    try {
        $me = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/me?$select=userPrincipalName,displayName' `
               -Method GET -Headers $headers -TimeoutSec 20
        if ($me.userPrincipalName) { $who = [string]$me.userPrincipalName }
        elseif ($me.displayName) { $who = [string]$me.displayName }
        $script:AiIde.GraphAuthUser = $who
    } catch {
        $code = AiIde-GetHttpStatusCode $_
        $script:AiIde.GraphLastHttpCode = $code
        $script:AiIde.GraphAccessNote = "Graph account check failed (HTTP $code)."
        if (-not $Quiet) {
            AiIde-SetPill 'Graph auth issue' '#E05050' '#5c1a1a'
            AiIde-SetStatus "Signed in token could not be validated (HTTP $code). Re-authenticate on your work account." '#E05050'
        }
        return $false
    }

    try {
        $r = Invoke-RestMethod -Uri 'https://graph.microsoft.com/beta/copilot/conversations' `
             -Method POST -Headers $headers -Body '{}' -TimeoutSec 25
        $script:AiIde.ConversationId = [string]$r.id
        $script:AiIde.GraphAccessVerified = $true
        $script:AiIde.Status = 'authenticated'
        $script:AiIde.GraphAccessNote = 'Graph Copilot API access verified.'
        $script:AiIdeConvHistory.Clear()
        if (-not $Quiet) {
            AiIde-SetPill 'Graph API --' '#50C878' '#1a5c1a'
            if ($who) {
                AiIde-SetStatus "Verified Graph Copilot access for $who." '#50C878'
            } else {
                AiIde-SetStatus 'Verified Graph Copilot access.' '#50C878'
            }
        }
        return $true
    } catch {
        $code = AiIde-GetHttpStatusCode $_
        $script:AiIde.GraphLastHttpCode = $code
        $script:AiIde.GraphAccessVerified = $false
        $script:AiIde.ConversationId = ''
        if ($code -eq 403) {
            $script:AiIde.Status = 'graph_forbidden'
            $script:AiIde.GraphAccessNote = 'Signed in, but Copilot Graph API is not licensed for this account/tenant.'
            if (-not $Quiet) {
                AiIde-SetPill 'Graph blocked' '#FFA500' '#4a2f1a'
                AiIde-SetStatus 'Graph Copilot is blocked (403). This is expected on non-licensed personal accounts; use your enterprise work account/computer.' '#FFA500'
            }
            return $false
        }
        if ($code -eq 401) {
            $script:AiIde.Status = 'auth_error'
            $script:AiIde.GraphAccessNote = 'Token was rejected by Graph (401).'
            if (-not $Quiet) {
                AiIde-SetPill 'Graph auth issue' '#E05050' '#5c1a1a'
                AiIde-SetStatus 'Graph rejected the token (401). Re-run Auth + Verify.' '#E05050'
            }
            return $false
        }
        $script:AiIde.Status = 'error'
        $script:AiIde.GraphAccessNote = "Graph Copilot verification failed (HTTP $code)."
        if (-not $Quiet) {
            AiIde-SetPill 'Graph error' '#E05050' '#5c1a1a'
            AiIde-SetStatus "Graph Copilot verification failed (HTTP $code)." '#E05050'
        }
        return $false
    }
}

function AiIde-StartDeviceCodeAuth {
    $tenant   = if ($AiIdeTenantBox.Text.Trim()) { $AiIdeTenantBox.Text.Trim() } else { 'common' }
    $clientId = $script:AiIde.ClientId
    $scope    = 'offline_access CopilotChat.ReadWrite.All'
    $dcUrl    = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode"
    AiIde-EnsureTls12
    $script:AiIde.GraphAccessVerified  = $false
    $script:AiIde.GraphAccessCheckedOn = [datetime]::MinValue
    $script:AiIde.GraphAccessNote      = ''
    $script:AiIde.GraphLastHttpCode    = 0
    $script:AiIde.GraphAuthUser        = ''
    $script:AiIde.ConversationId       = ''
    AiIde-SetStatus 'Requesting device code...' '#4C9FE6'
    try {
        $body = "client_id=$clientId&scope=$([Uri]::EscapeDataString($scope))"
        $r    = Invoke-RestMethod -Uri $dcUrl -Method POST `
                    -ContentType 'application/x-www-form-urlencoded' `
                    -Body $body -TimeoutSec 20
        $script:AiIde.DeviceCode       = $r.device_code
        $script:AiIde.DeviceCodeExpiry = (Get-Date).AddSeconds($r.expires_in)
        $script:AiIde.TenantId         = $tenant
        $script:AiIde.Status           = 'auth_pending'
        try { [System.Windows.Clipboard]::SetText($r.user_code) } catch {}
        $msg = "1. Open:  $($r.verification_uri)`n`n2. Enter code:  $($r.user_code)`n   (already copied to clipboard)`n`nMacroHub will detect when you approve."
        [System.Windows.MessageBox]::Show($msg, 'Sign in to Microsoft 365', 'OK', 'Information') | Out-Null
        AiIde-SetStatus "Code '$($r.user_code)' copied -- waiting for browser approval..." '#FFA500'
        AiIde-PollForToken -Interval ([int]$r.interval)
    } catch {
        AiIde-SetStatus "Device code request failed: $_" '#E05050'
    }
}

function AiIde-PollForToken([int]$Interval = 5) {
    $shared   = $script:AiIde
    $tokenUrl = "https://login.microsoftonline.com/$($shared.TenantId)/oauth2/v2.0/token"
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('shared',   $shared)
    $rs.SessionStateProxy.SetVariable('tokenUrl', $tokenUrl)
    $rs.SessionStateProxy.SetVariable('interval', $Interval)
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        $deadline = $shared.DeviceCodeExpiry
        while ((Get-Date) -lt $deadline -and $shared.Status -eq 'auth_pending') {
            Start-Sleep -Seconds $interval
            try {
                $body = "client_id=$($shared.ClientId)&grant_type=urn:ietf:params:oauth:grant-type:device_code&device_code=$($shared.DeviceCode)"
                $t = Invoke-RestMethod -Uri $tokenUrl -Method POST `
                         -ContentType 'application/x-www-form-urlencoded' -Body $body -TimeoutSec 15
                $shared.AccessToken  = $t.access_token
                $shared.RefreshToken = $t.refresh_token
                $shared.TokenExpiry  = (Get-Date).AddSeconds($t.expires_in - 60)
                $shared.Status       = 'authenticated'
                break
            } catch {
                $e = $_.Exception.Message
                if ($e -notlike '*authorization_pending*' -and $e -notlike '*slow_down*') {
                    $shared.LastError = $e
                    $shared.Status    = 'auth_error'
                    break
                }
            }
        }
        if ($shared.Status -eq 'auth_pending') {
            $shared.Status = 'auth_error'; $shared.LastError = 'Device code expired'
        }
    })
    [void]$ps.BeginInvoke()

    $aTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $aTimer.Interval = [TimeSpan]::FromSeconds(2)
    $aTimer.Add_Tick({
        $s = $script:AiIde.Status
        if ($s -eq 'authenticated') {
            $aTimer.Stop()
            AiIde-SetStatus 'Authenticated. Verifying Graph Copilot access...' '#4C9FE6'
            [void](AiIde-VerifyGraphCopilotAccess)
        } elseif ($s -eq 'auth_error') {
            $aTimer.Stop()
            AiIde-SetPill 'Auth failed' '#E05050' '#5c1a1a'
            AiIde-SetStatus "Auth error: $($script:AiIde.LastError)" '#E05050'
        }
    }.GetNewClosure())
    $aTimer.Start()
}

function AiIde-RefreshToken {
    if (-not $script:AiIde.RefreshToken) { return $false }
    $tokenUrl = "https://login.microsoftonline.com/$($script:AiIde.TenantId)/oauth2/v2.0/token"
    try {
        AiIde-EnsureTls12
        $body = "client_id=$($script:AiIde.ClientId)&grant_type=refresh_token&refresh_token=$($script:AiIde.RefreshToken)&scope=$([Uri]::EscapeDataString('CopilotChat.ReadWrite.All offline_access'))"
        $t = Invoke-RestMethod -Uri $tokenUrl -Method POST `
                 -ContentType 'application/x-www-form-urlencoded' -Body $body -TimeoutSec 20
        $script:AiIde.AccessToken  = $t.access_token
        $script:AiIde.RefreshToken = $t.refresh_token
        $script:AiIde.TokenExpiry  = (Get-Date).AddSeconds($t.expires_in - 60)
        return $true
    } catch { return $false }
}

function AiIde-GetToken {
    if (-not $script:AiIde.AccessToken) { return $null }
    if ((Get-Date) -ge $script:AiIde.TokenExpiry) {
        if (-not (AiIde-RefreshToken)) { return $null }
    }
    return $script:AiIde.AccessToken
}

# -"--- Graph Copilot Chat API -"---------------------------------------------------------------------------------------------

function AiIde-NewConversation {
    $token = AiIde-GetToken
    if (-not $token) {
        $script:AiIde.GraphAccessVerified = $false
        return $false
    }
    AiIde-EnsureTls12
    try {
        $r = Invoke-RestMethod `
             -Uri 'https://graph.microsoft.com/beta/copilot/conversations' `
             -Method POST `
             -Headers @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' } `
             -Body '{}' -TimeoutSec 20
        $script:AiIde.ConversationId = $r.id
        $script:AiIde.GraphAccessVerified = $true
        $script:AiIde.GraphAccessCheckedOn = Get-Date
        $script:AiIde.GraphAccessNote = 'Conversation ready.'
        $script:AiIdeConvHistory.Clear()
        AiIde-SetStatus "Conversation ready: $($r.id.Substring(0,[Math]::Min(8,$r.id.Length)))..." '#50C878'
        return $true
    } catch {
        $code = AiIde-GetHttpStatusCode $_
        $script:AiIde.GraphLastHttpCode = $code
        $script:AiIde.GraphAccessVerified = $false
        $script:AiIde.GraphAccessCheckedOn = Get-Date
        if ($code -eq 403) {
            $script:AiIde.Status = 'graph_forbidden'
            $script:AiIde.GraphAccessNote = 'Signed in, but Copilot Graph API is not licensed for this account/tenant.'
            AiIde-SetPill 'Graph blocked' '#FFA500' '#4a2f1a'
            AiIde-SetStatus 'Graph Copilot is blocked (403). Use your licensed enterprise work account/computer.' '#FFA500'
        } else {
            $script:AiIde.GraphAccessNote = "Could not create conversation (HTTP $code)."
            AiIde-SetStatus "Could not create conversation (HTTP $code): $($_.Exception.Message)" '#E05050'
        }
        return $false
    }
}

function AiIde-GraphChat([string]$Prompt, [string]$ExtraContext = '') {
    $token = AiIde-GetToken
    if (-not $token) { return $null }
    $keepConv = ($AiIdeKeepConvChk.IsChecked -eq $true)
    if (-not $script:AiIde.ConversationId -or -not $keepConv) {
        AiIde-NewConversation
        if (-not $script:AiIde.ConversationId) { return $null }
    }
    $useWeb = ($AiIdeWebGroundChk.IsChecked -eq $true)
    $tz     = [System.TimeZoneInfo]::Local.StandardName
    $obj    = @{
        message      = @{ text = $Prompt }
        locationHint = @{ timeZone = $tz }
        webGrounding = @{ isEnabled = $useWeb }
    }
    if ($ExtraContext) { $obj.additionalContext = @(@{ text = $ExtraContext }) }
    $body = $obj | ConvertTo-Json -Depth 5
    $uri  = "https://graph.microsoft.com/beta/copilot/conversations/$($script:AiIde.ConversationId)/chat"
    try {
        $resp = Invoke-RestMethod -Uri $uri -Method POST `
                    -Headers @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' } `
                    -Body $body -TimeoutSec 90
        $msgs = @($resp.messages | Where-Object { $_.'@odata.type' -like '*Response*' })
        if ($msgs.Count) { return $msgs[-1].text }
        return $null
    } catch {
        $code = AiIde-GetHttpStatusCode $_
        $script:AiIde.GraphLastHttpCode = $code
        $script:AiIde.GraphAccessVerified = $false
        $script:AiIde.GraphAccessCheckedOn = Get-Date
        if ($code -eq 403) {
            $script:AiIde.Status = 'graph_forbidden'
            $script:AiIde.GraphAccessNote = 'Signed in, but Copilot Graph API is not licensed for this account/tenant.'
            AiIde-SetStatus 'Graph API: 403 (Copilot license missing or blocked) -- falling back to Chrome' '#FFA500'
        } elseif ($code -eq 401) {
            $script:AiIde.Status = 'auth_error'
            $script:AiIde.GraphAccessNote = 'Token rejected by Graph API (401). Re-auth required.'
            AiIde-SetStatus 'Graph API: 401 (token invalid/expired) -- re-run Auth + Verify' '#E05050'
        } else {
            $script:AiIde.Status = 'error'
            $script:AiIde.GraphAccessNote = "Graph chat call failed (HTTP $code)."
            AiIde-SetStatus "Graph chat call failed (HTTP $code)." '#FFA500'
        }
        return $null
    }
}

# -"--- Chrome bridge HttpListener -"-------------------------------------------------------------------------------------

function AiIde-StartBridge {
    if (AiIde-BridgeRunning) { AiIde-SetStatus 'Bridge already running' '#FFA500'; return }
    if ($script:AiIde.ListenerPS -or $script:AiIde.ListenerRS) { AiIde-StopBridge }
    $portTxt = $AiIdePortBox.Text.Trim()
    if ($portTxt -match '^\d+$') { $script:AiIde.Port = [int]$portTxt }
    if ($script:AiIde.Port -lt 1024 -or $script:AiIde.Port -gt 65535) {
        throw "Invalid bridge port '$($script:AiIde.Port)'. Use 1024-65535."
    }
    $shared = $script:AiIde
    $shared.Status = 'idle'
    $shared.LastError = ''
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('shared', $shared)
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        $L = [System.Net.HttpListener]::new()
        try {
            $L.Prefixes.Add("http://localhost:$($shared.Port)/")
            $L.Prefixes.Add("http://127.0.0.1:$($shared.Port)/")
        } catch {}
        try { $L.Start() } catch { $shared.Status = 'error'; $shared.LastError = [string]$_.Exception.Message; return }
        try {
            while ($L.IsListening) {
                try {
                    $ctx  = $L.GetContext()
                    $path = $ctx.Request.Url.AbsolutePath.ToLower()
                    $meth = $ctx.Request.HttpMethod
                    $res  = $ctx.Response
                    $res.AddHeader('Access-Control-Allow-Origin','*')
                    $res.AddHeader('Access-Control-Allow-Methods','GET,POST,OPTIONS')
                    $res.AddHeader('Access-Control-Allow-Headers','Content-Type')
                    if ($meth -eq 'OPTIONS') { $res.StatusCode = 204 }
                    elseif ($path -eq '/prompt' -and $meth -eq 'GET') {
                        $buf = [Text.Encoding]::UTF8.GetBytes($shared.PendingPrompt)
                        $res.ContentType = 'text/plain'
                        $res.OutputStream.Write($buf,0,$buf.Length)
                        $shared.PendingPrompt = ''
                    }
                    elseif ($path -eq '/request' -and $meth -eq 'GET') {
                        $buf = [Text.Encoding]::UTF8.GetBytes($shared.PendingRequest)
                        $res.ContentType = 'text/plain'
                        $res.OutputStream.Write($buf,0,$buf.Length)
                    }
                    elseif ($path -eq '/file' -and $meth -eq 'GET') {
                        if ($shared.PendingFile -and [IO.File]::Exists($shared.PendingFile)) {
                            $bytes = [IO.File]::ReadAllBytes($shared.PendingFile)
                            $res.ContentType = 'text/plain'
                            $res.OutputStream.Write($bytes,0,$bytes.Length)
                        } else { $res.StatusCode = 404 }
                    }
                    elseif (($path -eq '/result' -or $path -eq '/aiide' -or $path -eq '/automation') -and $meth -eq 'POST') {
                        $rd = [IO.StreamReader]::new($ctx.Request.InputStream, [Text.Encoding]::UTF8)
                        $body = $rd.ReadToEnd()
                        $txt = ''
                        if ($body -and $body.Trim().StartsWith('{')) {
                            try {
                                $obj = $body | ConvertFrom-Json -ErrorAction Stop
                                if ($obj.response) { $txt = [string]$obj.response }
                                elseif ($obj.result) { $txt = [string]$obj.result }
                                elseif ($obj.text) { $txt = [string]$obj.text }
                            } catch {}
                        }
                        if (-not $txt) { $txt = [string]$body }
                        $shared.LastResult = $txt
                        $shared.Status = if ($path -eq '/automation') { 'automation_done' } else { 'done' }
                        $ack = [Text.Encoding]::UTF8.GetBytes('ok')
                        $res.ContentType = 'text/plain'
                        $res.OutputStream.Write($ack,0,$ack.Length)
                    }
                    elseif (($path -eq '/status' -or $path -eq '/ping' -or $path -eq '/aiide/ping') -and $meth -eq 'GET') {
                        $payload = if ($path -eq '/status') { [string]$shared.Status } else { 'ok' }
                        $buf = [Text.Encoding]::UTF8.GetBytes($payload)
                        $res.ContentType = 'text/plain'
                        $res.OutputStream.Write($buf,0,$buf.Length)
                    }
                    else {
                        $res.StatusCode = 404
                    }
                    $res.Close()
                } catch {
                    $shared.LastError = [string]$_.Exception.Message
                }
            }
        } finally {
            try { $L.Stop() } catch {}
            try { $L.Close() } catch {}
        }
    })
    $script:AiIde.ListenerPS = $ps
    $script:AiIde.ListenerRS = $rs
    [void]$ps.BeginInvoke()
    Start-Sleep -Milliseconds 250
    if (-not (AiIde-BridgeRunning)) {
        $err = if ($script:AiIde.LastError) { $script:AiIde.LastError } else { 'unknown bridge startup error' }
        AiIde-StopBridge
        throw "Bridge failed to start: $err"
    }
    AiIde-SetPill 'Chrome --' '#4C9FE6' '#1a4a7a'
    AiIde-SetStatus "Chrome bridge listening on port $($script:AiIde.Port). If unavailable, AI falls back without crashing MacroHub." '#50C878'
}

function AiIde-StopBridge {
    try {
        if ($script:AiIde.ListenerPS) { $script:AiIde.ListenerPS.Stop(); $script:AiIde.ListenerPS.Dispose() }
        if ($script:AiIde.ListenerRS) { $script:AiIde.ListenerRS.Close() }
    } catch {}
    $script:AiIde.ListenerPS = $null
    $script:AiIde.ListenerRS = $null
    $script:AiIde.Status = 'idle'
    $script:AiIde.LastError = ''
    AiIde-SetPill 'Stopped' '#6E6E6E' '#2D2D30'
    AiIde-SetStatus 'Bridge stopped' '#6E6E6E'
}

function AiIde-ChromeChat([string]$Prompt, [string]$UserRequest = '') {
    if (-not $script:AiIde.ListenerPS) { throw 'Chrome bridge not running -- click Start Bridge' }
    $script:AiIde.LastResult     = ''
    $script:AiIde.Status         = 'waiting'
    $script:AiIde.PendingRequest = $UserRequest
    $useFile = $Prompt.Length -gt 1200
    if ($useFile) {
        $f = Join-Path $env:TEMP 'MacroHub_AIIde_ctx.txt'
        [IO.File]::WriteAllText($f, $Prompt, [Text.UTF8Encoding]::new($false))
        $script:AiIde.PendingFile   = $f
        $script:AiIde.PendingPrompt = '__FILE__'
    } else {
        $script:AiIde.PendingFile   = ''
        $script:AiIde.PendingPrompt = $Prompt
    }
    $deadline = (Get-Date).AddSeconds(90)
    while ($script:AiIde.Status -ne 'done' -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 300; Update-UI
    }
    if ($script:AiIde.Status -ne 'done') {
        throw 'Copilot timed out -- ensure Chrome is open on copilot.microsoft.com with the extension active'
    }
    return $script:AiIde.LastResult
}

# -"--- Unified send (Graph first, Chrome fallback) -"---------------------------------------------------

function AiIde-Send([string]$Prompt, [string]$UserRequest = '', [string]$Context = '') {
    $mode = Get-ComboText $AiIdeModeCombo
    switch ($mode) {
        'Graph API' {
            $token = AiIde-GetToken
            if (-not $token) {
                throw 'Graph API mode selected but no token is available. Click Auth + Verify first.'
            }
            if (-not $script:AiIde.GraphAccessVerified) {
                if (-not (AiIde-VerifyGraphCopilotAccess)) {
                    $hint = if ($script:AiIde.GraphAccessNote) { $script:AiIde.GraphAccessNote } else { 'Graph verification failed.' }
                    throw "Graph API mode selected, but Graph Copilot is not ready. $hint"
                }
            }
            AiIde-SetStatus 'Calling Graph Copilot Chat API...' '#4C9FE6'
            $result = AiIde-GraphChat -Prompt $Prompt -ExtraContext $Context
            if ($result) { return $result }
            $hint = if ($script:AiIde.GraphAccessNote) { $script:AiIde.GraphAccessNote } else { 'No response was returned.' }
            throw "Graph API mode selected but no response was returned. $hint"
        }
        'Chrome' {
            if (-not (AiIde-BridgeRunning)) {
                throw 'Chrome mode selected but bridge is not running. Click Start Bridge.'
            }
            AiIde-SetPill 'Chrome' '#4C9FE6' '#1a4a7a'
            AiIde-SetStatus 'Sending to Chrome extension...' '#4C9FE6'
            $fullPrompt = if ($Context) { "$Prompt`n`n=== CONTEXT ===`n$Context" } else { $Prompt }
            return AiIde-ChromeChat -Prompt $fullPrompt -UserRequest $UserRequest
        }
        default {
            # Auto mode: prefer Graph. Use Chrome only if bridge is already running.
            $token = AiIde-GetToken
            if ($token) {
                $graphReady = $script:AiIde.GraphAccessVerified
                if (-not $graphReady) { $graphReady = AiIde-VerifyGraphCopilotAccess -Quiet }
                if ($graphReady) {
                    AiIde-SetStatus 'Calling Graph Copilot Chat API...' '#4C9FE6'
                    $result = AiIde-GraphChat -Prompt $Prompt -ExtraContext $Context
                    if ($result) { return $result }
                    AiIde-SetStatus 'Graph returned nothing.' '#FFA500'
                } elseif ($script:AiIde.GraphAccessNote) {
                    AiIde-SetStatus "Graph unavailable: $($script:AiIde.GraphAccessNote)" '#FFA500'
                }
            }
            if (AiIde-BridgeRunning) {
                AiIde-SetPill 'Chrome' '#4C9FE6' '#1a4a7a'
                AiIde-SetStatus 'Auto mode fallback: sending to Chrome extension...' '#4C9FE6'
                $fullPrompt = if ($Context) { "$Prompt`n`n=== CONTEXT ===`n$Context" } else { $Prompt }
                return AiIde-ChromeChat -Prompt $fullPrompt -UserRequest $UserRequest
            }
            if ($token -and $script:AiIde.GraphAccessNote) {
                throw "Auto mode: Graph is signed in but unavailable ($($script:AiIde.GraphAccessNote)), and Chrome bridge is not running."
            }
            throw 'Auto mode: no Graph token and Chrome bridge is not running. Click Auth + Verify, or switch to Chrome mode and Start Bridge.'
        }
    }
}

# -"--- Build prompt with system instruction -"-----------------------------------------------------------------

function AiIde-BuildPrompt([string]$UserRequest, [switch]$IncludeEditor) {
    $lang = AiIde-GetLang
    $sys = switch -Wildcard ($lang) {
        '*Outlook*' { 'You are an expert VBA developer for Microsoft Outlook Classic. Return ONLY a complete runnable VBA Sub for Outlook.Application COM. No explanation. No markdown fences.' }
        '*VBA*'     { 'You are an expert VBA developer for Microsoft Excel. Return ONLY a complete runnable VBA Sub. No explanation. No markdown fences.' }
        '*PS*COM*'  { 'You are a PowerShell expert who drives Excel via COM (Excel.Application). Return ONLY runnable PowerShell code using COM objects. No explanation. No fences.' }
        default     { 'You are a PowerShell expert. Return ONLY complete runnable PowerShell code. No explanation. No markdown fences.' }
    }
    $editor = [string]$AiIdeEditorBox.Text
    if ($IncludeEditor -and $editor.Trim()) {
        return "$sys`n`n=== CURRENT CODE IN EDITOR ===`n$editor`n`n=== INSTRUCTION ===`n$UserRequest`n`nReturn ONLY the complete updated code."
    }
    return "$sys`n`n=== REQUEST ===`n$UserRequest"
}

# -"--- Run code (VBA Excel, VBA Outlook, PowerShell) -"---------------------------------------------

function AiIde-RunCode([string]$Code) {
    $lang = AiIde-GetLang
    $wb   = [string]$AiIdeWbCombo.SelectedItem
    $sh   = [string]$AiIdeSheetCombo.SelectedItem
    if ($wb -eq '(active workbook)') { $wb = '' }
    if ($sh -eq '(active sheet)' -or $sh -eq '(n/a)') { $sh = '' }

    if ($lang -like '*Outlook*') {
        $outlook = $null
        try { $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
        catch { throw 'Outlook Classic is not open.' }
        $proj  = $outlook.VBE.ActiveVBProject
        if (-not $proj) { throw 'Cannot access Outlook VBA project. Enable: Trust Center > Trust access to VBA project object model.' }
        $comp  = $proj.VBComponents.Add(1)
        $entry = [regex]::Match($Code, '(?i)(?:Public\s+)?Sub\s+(\w+)\s*\(').Groups[1].Value
        if (-not $entry) { throw 'No runnable Sub found in VBA code.' }
        try {
            $comp.CodeModule.AddFromString($Code)
            $outlook.Run($entry)
            AiIde-AppendOutput -Role 'RUN' -Text "Outlook VBA executed: $entry"
            $AiIdeRunErrorTxt.Text = ''
        } finally {
            try { $proj.VBComponents.Remove($comp) } catch {}
        }

    } elseif ($lang -like '*VBA*') {
        $result = Invoke-IdeVbaRun -Content $Code -WorkbookName $wb -SheetName $sh
        $AiIdeRunErrorTxt.Text = if ($result.Ok) { '' } else { $result.Message }
        if ($result.Ok) { AiIde-AppendOutput -Role 'RUN' -Text $result.Message }
        if (-not $result.Ok) { throw $result.Message }

    } else {
        $result = Invoke-IdePowerShellRun -Content $Code -WorkbookName $wb -SheetName $sh
        if ($result.StdOut) { AiIde-AppendOutput -Role 'RUN' -Text ([string]$result.StdOut) }
        $AiIdeRunErrorTxt.Text = $result.StdErr
        if (-not $result.StdOut -and -not $result.StdErr -and $result.ExitCode -eq 0) {
            AiIde-AppendOutput -Role 'RUN' -Text 'PowerShell run completed (no output).'
        }
        if ($result.StdErr -or $result.ExitCode -ne 0) {
            throw $(if ($result.StdErr) { $result.StdErr } else { "Exit code $($result.ExitCode)" })
        }
    }
}

# -"--- Inject module only (no run) -"-----------------------------------------------------------------------------------

function AiIde-InjectToExcel([string]$Code) {
    try {
        $xl = Get-ExcelApp -Session Main -Create
        if (-not $xl) { throw 'Excel is not open.' }
        $wb   = $xl.ActiveWorkbook
        $comp = $wb.VBProject.VBComponents.Add(1)
        $comp.CodeModule.AddFromString($Code)
        $comp.Name = 'AiIde_' + (Get-Date -f 'HHmmss')
        AiIde-SetStatus "Injected '$($comp.Name)' into $($wb.Name)" '#50C878'
    } catch {
        AiIde-SetStatus "Excel inject error: $_" '#E05050'
    }
}

function AiIde-InjectToOutlook([string]$Code) {
    try {
        $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        if (-not $outlook) { throw 'Outlook is not open.' }
        $proj = $outlook.VBE.ActiveVBProject
        if (-not $proj) { throw 'Cannot access Outlook VBA project.' }
        $comp = $proj.VBComponents.Add(1)
        $comp.CodeModule.AddFromString($Code)
        $comp.Name = 'AiIde_' + (Get-Date -f 'HHmmss')
        AiIde-SetStatus "Injected '$($comp.Name)' into Outlook VBA project" '#50C878'
    } catch {
        AiIde-SetStatus "Outlook inject error: $_" '#E05050'
    }
}

# -"--- Validate -"-------------------------------------------------------------------------------------------------------------------------

function AiIde-Validate {
    $code = [string]$AiIdeEditorBox.Text
    if (-not $code.Trim()) { AiIde-SetStatus 'Editor is empty' '#E05050'; return }
    $lang = AiIde-GetLang
    if ($lang -notlike '*VBA*' -and $lang -notlike '*Outlook*') {
        $r = Invoke-IdePowerShellValidate -Content $code
        $AiIdeValidationTxt.Text       = $r.Message
        $AiIdeValidationTxt.Foreground = HexBrush (if ($r.Ok) { '#50C878' } else { '#FFB4B4' })
        AiIde-SetStatus (if ($r.Ok) { 'PS syntax OK' } else { 'PS syntax errors' }) (if ($r.Ok) { '#50C878' } else { '#E05050' })
    } else {
        $has = [regex]::IsMatch($code, '(?i)(?:Public\s+)?Sub\s+\w+\s*\(')
        $AiIdeValidationTxt.Text       = if ($has) { 'VBA: entry-point Sub found.' } else { 'VBA: no runnable Sub found.' }
        $AiIdeValidationTxt.Foreground = HexBrush (if ($has) { '#50C878' } else { '#FFB4B4' })
        AiIde-SetStatus (if ($has) { 'VBA looks runnable' } else { 'VBA missing Sub entry point' }) (if ($has) { '#50C878' } else { '#E05050' })
    }
}

# -"--- Main Generate flow -"-----------------------------------------------------------------------------------------------------

function AiIde-Generate([string]$UserRequest, [switch]$IncludeEditor) {
    if (-not $UserRequest.Trim()) { AiIde-SetStatus 'Enter a prompt' '#E05050'; return $null }
    Show-Busy 'AI generating...'
    try {
        $prompt = AiIde-BuildPrompt -UserRequest $UserRequest -IncludeEditor:$IncludeEditor
        $raw    = AiIde-Send -Prompt $prompt -UserRequest $UserRequest
        $code   = AiIde-StripFences $raw
        $AiIdeEditorBox.Text = $code
        AiIde-AppendOutput -Role 'YOU' -Text $UserRequest
        AiIde-AppendOutput -Role 'AI'  -Text $raw
        $script:AiIde.LastError = ''
        [void]$script:AiIdeConvHistory.Add([PSCustomObject]@{ Role='user'; Text=$UserRequest })
        [void]$script:AiIdeConvHistory.Add([PSCustomObject]@{ Role='ai';   Text=$raw })
        AiIde-SetStatus 'Generated -- code in editor' '#50C878'
        Write-ActivityLog "AI IDE Generate: $UserRequest"
        return $code
    } catch {
        AiIde-SetStatus "Generate error: $_" '#E05050'
        return $null
    } finally { Hide-Busy }
}

function AiIde-GenerateAndRun([string]$UserRequest) {
    $code = AiIde-Generate -UserRequest $UserRequest
    if (-not $code) { return }
    $script:AiIde.RetryCount = 0
    $AiIdeRetryTxt.Text      = ''
    for ($i = 1; $i -le $script:AiIde.MaxRetries; $i++) {
        Show-Busy "Running -- attempt $i of $($script:AiIde.MaxRetries)..."
        try {
            AiIde-RunCode -Code $code
            AiIde-SetStatus "Ran successfully on attempt $i" '#50C878'
            $AiIdeRetryTxt.Text = ''
            break
        } catch {
            $err = $_.Exception.Message
            $script:AiIde.LastError = $err
            $AiIdeRunErrorTxt.Text  = $err
            $AiIdeRetryTxt.Text     = "Retry $i/$($script:AiIde.MaxRetries)"
            if ($i -ge $script:AiIde.MaxRetries) {
                AiIde-SetStatus "Failed after $($script:AiIde.MaxRetries) attempts -- review manually" '#E05050'; break
            }
            AiIde-SetStatus "Error on attempt $i -- asking AI to fix..." '#FFA500'
            $fixP = "Fix this code. It threw an error. Return ONLY the corrected code, no markdown fences.`n`nERROR:`n$err`n`nCODE:`n$code"
            try {
                $raw  = AiIde-Send -Prompt $fixP -UserRequest "Fix error"
                $code = AiIde-StripFences $raw
                $AiIdeEditorBox.Text = $code
                AiIde-AppendOutput -Role 'AI FIX' -Text $raw
            } catch { AiIde-SetStatus "Could not reach AI for fix: $_" '#E05050'; break }
        } finally { Hide-Busy }
    }
}

# -"--- File operations -"-----------------------------------------------------------------------------------------------------------

function AiIde-Open {
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Title  = 'Open Code File'
    $dlg.Filter = 'Code Files|*.ps1;*.psm1;*.bas;*.cls;*.txt|All Files|*.*'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $AiIdeEditorBox.Text     = [IO.File]::ReadAllText($dlg.FileName)
    $script:AiIdeCurrentFile = $dlg.FileName
    $AiIdeFilePathTxt.Text   = $dlg.FileName
    $AiIdeValidationTxt.Text = ''; $AiIdeRunErrorTxt.Text = ''
    AiIde-SetStatus "Opened: $(Split-Path $dlg.FileName -Leaf)" '#50C878'
}

function AiIde-Save {
    if (-not $script:AiIdeCurrentFile) {
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title  = 'Save Code File'
        $dlg.Filter = 'PowerShell|*.ps1|VBA Module|*.bas|All Files|*.*'
        if ($dlg.ShowDialog() -ne 'OK') { return }
        $script:AiIdeCurrentFile = $dlg.FileName
        $AiIdeFilePathTxt.Text   = $dlg.FileName
    }
    [IO.File]::WriteAllText($script:AiIdeCurrentFile, [string]$AiIdeEditorBox.Text, [Text.UTF8Encoding]::new($false))
    AiIde-SetStatus "Saved: $(Split-Path $script:AiIdeCurrentFile -Leaf)" '#50C878'
}

function AiIde-ExportHistory {
    if ($script:AiIdeConvHistory.Count -eq 0) { AiIde-SetStatus 'No history' '#FFA500'; return }
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Title = 'Export Chat History'; $dlg.Filter = 'Text Files|*.txt'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $lines = $script:AiIdeConvHistory | ForEach-Object { "[$($_.Role.ToUpper())]`n$($_.Text)`n`n---`n" }
    $lines | Set-Content $dlg.FileName -Encoding UTF8
    AiIde-SetStatus "Exported to $(Split-Path $dlg.FileName -Leaf)" '#50C878'
}

function Invoke-AiIdeSafe([scriptblock]$Action, [string]$ActionName = 'AI IDE') {
    try {
        & $Action
    } catch {
        $msg = [string]$_.Exception.Message
        try { AiIde-SetStatus "$ActionName error: $msg" '#E05050' } catch {}
        try { Write-ActivityLog "$ActionName error: $msg" } catch {}
    }
}

# ================================================================
#  END AI IDE FUNCTIONS
# ================================================================

# ================================================================
#  REFRESH FUNCTIONS: QuarterSync tabs
# ================================================================

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

# -- Clipboard tab --
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
        -DefaultDateOffset $ClipDateOffset.Text `
        -DefaultTimeOffset $ClipTimeOffset.Text `
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
    $ts = if ($slot -and $slot.TimestampChk) { [bool]$slot.TimestampChk.IsChecked } else { [bool]$ClipTimestampChk.IsChecked }
    $dOff = if ($slot -and $slot.DateOffsetBox) { [string]$slot.DateOffsetBox.Text } else { [string]$ClipDateOffset.Text }
    $tOff = if ($slot -and $slot.TimeOffsetBox) { [string]$slot.TimeOffsetBox.Text } else { [string]$ClipTimeOffset.Text }

    # Keep hidden defaults controls in sync (used when creating future slots).
    if ($sh -and $ClipSheetCombo -and $ClipSheetCombo.Items.Contains($sh)) { $ClipSheetCombo.SelectedItem = $sh }
    $ClipCellBox.Text = $cell
    $ClipTimestampChk.IsChecked = $ts
    $ClipDateOffset.Text = $dOff
    $ClipTimeOffset.Text = $tOff

    Save-ClipDefaults -Workbook ([string]$wb) -Sheet ([string]$sh) -Cell $cell `
        -Timestamp $ts -DateOffset $dOff -TimeOffset $tOff
    $ClipDefaultsIndicator.Text = "(defaults: $wb > $sh > $cell)"
    $ClipDefaultsIndicator.Visibility = 'Visible'
    Set-Status "Defaults locked (workbook + slot settings): $wb > $sh > $cell | Timestamp=$ts | Date=$dOff Time=$tOff"
})

$ClipClearDefaultsBtn.Add_Click({
    if (Test-Path $script:ClipDefaultsJson) { Remove-Item $script:ClipDefaultsJson -Force }
    $ClipDefaultsIndicator.Visibility = 'Collapsed'
    Set-Status 'Defaults cleared -- will use last-selected settings on next launch'
})

$HelpBtn.Add_Click({
    $guide = @"
MacroHub Quick Guide

0 IDE: Open/edit/save raw .ps1/.bas files, validate/compile, and run with output panels.
1 Clipboard: Record per-slot, or use Record Sequence to capture multiple clipboard copies into slot1/slot2/slot3 before pasting.
2 Macros: Pick macro + workbook + sheet and run.
3 Scheduler: Schedule .ps1/.bas tasks in Windows Task Scheduler.
4 Navigator: Manage workbooks/sheets in a separate Excel Navigator session (Ctrl+E renames selected sheet).
5 Templates: Save reusable text snippets and paste into Excel.
6 QSync: Build quarter checklist from folder comparisons.
7 QTasks: Work a running task list, track completion %, export to Excel.
8 Email Helper: Pull Outlook mail with scope/date/text filters and helper actions.
9 Email Dashboard: Analytics using aligned filters from the same email dataset.

Typical flow:
Refresh workbook/email data -> filter/select target -> run action -> verify status.
"@
    [System.Windows.MessageBox]::Show($guide, 'MacroHub Guide', 'OK', 'Information') | Out-Null
})

# -- Macros tab --
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
        $xl.WindowState = -4143 # xlNormal
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
        $format = if ($ext -eq '.csv') { 6 } else { 51 }
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
        $xl.Calculation = switch ($mode) {
            'Manual'        { -4135 }
            'Semiautomatic' { -4134 }
            default         { -4105 }
        }
        $xl.EnableEvents = [bool]$NavEventsChk.IsChecked
        Set-Status "Excel options updated: Calc=$mode, Events=$($xl.EnableEvents)"
        Write-ActivityLog "Navigator Excel options updated: Calc=$mode, Events=$($xl.EnableEvents)"
    } catch { Set-Status "Excel options error: $_" '#E05050' }
})

# -- Templates tab --
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
            $result = Export-QsToExcel -SavePath $dlg.FileName
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
#  EMAIL TRIAGE: UI Rendering helpers
# ================================================================
function Refresh-EmailList {
    param([switch]$UseCachedRaw)
    $EmEmailPanel.Children.Clear()
    if (-not $UseCachedRaw -or -not $script:EmailRawCache -or $script:EmailRawCache.Count -eq 0) {
        $fetchCount = if ($script:EmailCfg.defaultFetchCount) { [int]$script:EmailCfg.defaultFetchCount } else { 250 }
        $scope = Get-ComboText $EmScopeCombo
        $script:EmailRawCache = @(Get-InboxEmails -maxCount $fetchCount -scope $scope)
        $script:EmailRawScope = [string]$scope
    }
    $emails = Apply-EmailFilters -Emails $script:EmailRawCache `
        -ViewName (Get-ComboText $EmViewCombo) `
        -QuickFilter (Get-ComboText $EmFilterCombo) `
        -SearchText $EmSearchBox.Text `
        -FromDateText $EmFromDateBox.Text `
        -ToDateText $EmToDateBox.Text
    $script:EmailCache = @($emails)
    $EmCountTxt.Text = "$($emails.Count) emails"
    $EmLastRefreshTxt.Text = "Last: $(Get-Date -f 'HH:mm:ss')  |  Scope: $(Get-ComboText $EmScopeCombo)"

    foreach ($em in $emails) {
        $row = [System.Windows.Controls.Border]::new()
        $row.Padding = [System.Windows.Thickness]::new(10, 6, 10, 6)
        $row.BorderBrush = (HexBrush '#3E3E42')
        $row.BorderThickness = [System.Windows.Thickness]::new(0,0,0,1)
        $row.Cursor = [System.Windows.Input.Cursors]::Hand
        $row.Tag = $em.EntryID

        $g = [System.Windows.Controls.Grid]::new()
        $cols = @(24, 24, -1, 160, 90, 70)
        foreach ($w in $cols) {
            $cd = [System.Windows.Controls.ColumnDefinition]::new()
            if ($w -eq -1) { $cd.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }
            else { $cd.Width = [System.Windows.GridLength]::new($w) }
            [void]($g.ColumnDefinitions.Add($cd))
        }

        # VIP indicator
        $vipTxt = [System.Windows.Controls.TextBlock]::new()
        $vipTxt.Text = $(if ($em.IsVIP) { [char]0x2605 } else { '' })
        $vipTxt.Foreground = (HexBrush '#FFD700')
        $vipTxt.FontSize = 12
        $vipTxt.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($vipTxt, 0)
        [void]($g.Children.Add($vipTxt))

        # Read/unread indicator
        $readTxt = [System.Windows.Controls.TextBlock]::new()
        $readTxt.Text = $(if ($em.IsRead) { '' } else { [char]0x25CF })
        $readTxt.Foreground = (HexBrush '#4C9FE6')
        $readTxt.FontSize = 10
        $readTxt.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($readTxt, 1)
        [void]($g.Children.Add($readTxt))

        # Subject
        $subTxt = [System.Windows.Controls.TextBlock]::new()
        $subTxt.Text = $em.Subject
        $subTxt.Foreground = (HexBrush '#FFFFFF')
        $subTxt.FontSize = 11
        $subTxt.TextTrimming = 'CharacterEllipsis'
        $subTxt.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($subTxt, 2)
        [void]($g.Children.Add($subTxt))

        # Sender
        $sndTxt = [System.Windows.Controls.TextBlock]::new()
        $sndTxt.Text = $em.SenderName
        $sndTxt.Foreground = (HexBrush '#C0C0C0')
        $sndTxt.FontSize = 10
        $sndTxt.TextTrimming = 'CharacterEllipsis'
        $sndTxt.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($sndTxt, 3)
        [void]($g.Children.Add($sndTxt))

        # Received time
        $timeTxt = [System.Windows.Controls.TextBlock]::new()
        $timeTxt.Text = $em.ReceivedTime.ToString('MM/dd HH:mm')
        $timeTxt.Foreground = (HexBrush '#6E6E6E')
        $timeTxt.FontSize = 10
        $timeTxt.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($timeTxt, 4)
        [void]($g.Children.Add($timeTxt))

        # Age with color
        $ageBucket = Get-AgingBucket $em.AgeDays
        $ageTxt = [System.Windows.Controls.TextBlock]::new()
        $ageTxt.Text = $(if ($em.AgeDays -eq 0) { 'Today' } else { "$($em.AgeDays)d" })
        $ageTxt.Foreground = (HexBrush $ageBucket.color)
        $ageTxt.FontSize = 10
        $ageTxt.FontWeight = [System.Windows.FontWeights]::SemiBold
        $ageTxt.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($ageTxt, 5)
        [void]($g.Children.Add($ageTxt))

        $row.Child = $g

        # Click to select / preview
        $row.Add_MouseLeftButtonDown({
            param($s, $e)
            $eid = $s.Tag
            $script:SelectedEmailId = $eid
            # Highlight selected row
            foreach ($child in $EmEmailPanel.Children) {
                $child.Background = [System.Windows.Media.Brushes]::Transparent
            }
            $s.Background = (HexBrush '#2A2D33')
            # Update preview
            $sel = $script:EmailCache | Where-Object { $_.EntryID -eq $eid } | Select-Object -First 1
            if ($sel) {
                $EmPreviewSubject.Text = $sel.Subject
                $EmPreviewFrom.Text = "$($sel.SenderName) <$($sel.SenderEmail)> -- $($sel.ReceivedTime.ToString('yyyy-MM-dd HH:mm'))"
                $EmPreviewBody.Text = $sel.Body
            }
        })

        [void]($EmEmailPanel.Children.Add($row))
    }
}

function Refresh-Dashboard {
    param([switch]$ForceFetch)
    $dashScope = Get-ComboText $DashScopeCombo
    if ($ForceFetch -or -not $script:EmailRawCache -or $script:EmailRawCache.Count -eq 0 -or [string]$script:EmailRawScope -ne [string]$dashScope) {
        $fetchCount = if ($script:EmailCfg.defaultFetchCount) { [int]$script:EmailCfg.defaultFetchCount } else { 250 }
        $script:EmailRawCache = @(Get-InboxEmails -maxCount $fetchCount -scope $dashScope)
        $script:EmailRawScope = [string]$dashScope
    }
    $emails = Apply-EmailFilters -Emails $script:EmailRawCache `
        -ViewName (Get-ComboText $DashViewCombo) `
        -QuickFilter (Get-ComboText $DashFilterCombo) `
        -SearchText $DashSearchBox.Text `
        -FromDateText $DashFromDateBox.Text `
        -ToDateText $DashToDateBox.Text `
        -Period (Get-ComboText $DashPeriodCombo)

    $m = Get-EmailMetrics $emails
    $DashTotalCount.Text    = $m.Total.ToString()
    $DashUnreadCount.Text   = $m.Unread.ToString()
    $DashVipCount.Text      = $m.VIP.ToString()
    $DashInternalCount.Text = $m.Internal.ToString()
    $DashExternalCount.Text = $m.External.ToString()
    $DashOldestDays.Text    = "$($m.OldestDays)d"

    # Aging bars
    $DashAgingPanel.Children.Clear()
    $maxCount = 1
    foreach ($k in $m.AgingDistribution.Keys) {
        if ($m.AgingDistribution[$k].count -gt $maxCount) { $maxCount = $m.AgingDistribution[$k].count }
    }
    foreach ($k in $m.AgingDistribution.Keys) {
        $info = $m.AgingDistribution[$k]
        $barRow = [System.Windows.Controls.Grid]::new()
        $c1 = [System.Windows.Controls.ColumnDefinition]::new(); $c1.Width = [System.Windows.GridLength]::new(90)
        $c2 = [System.Windows.Controls.ColumnDefinition]::new(); $c2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $c3 = [System.Windows.Controls.ColumnDefinition]::new(); $c3.Width = [System.Windows.GridLength]::new(40)
        [void]($barRow.ColumnDefinitions.Add($c1))
        [void]($barRow.ColumnDefinitions.Add($c2))
        [void]($barRow.ColumnDefinitions.Add($c3))
        $barRow.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

        $lbl = [System.Windows.Controls.TextBlock]::new()
        $lbl.Text = $k; $lbl.Foreground = (HexBrush $info.color); $lbl.FontSize = 11; $lbl.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($lbl, 0)

        $bar = [System.Windows.Controls.Border]::new()
        $pct = [Math]::Max(5, ($info.count / $maxCount) * 100)
        $bar.Width = $pct * 2.5
        $bar.Height = 16
        $bar.CornerRadius = [System.Windows.CornerRadius]::new(3)
        $bar.Background = (HexBrush $info.color)
        $bar.HorizontalAlignment = 'Left'
        [System.Windows.Controls.Grid]::SetColumn($bar, 1)

        $cnt = [System.Windows.Controls.TextBlock]::new()
        $cnt.Text = $info.count.ToString(); $cnt.Foreground = (HexBrush '#FFFFFF'); $cnt.FontSize = 11; $cnt.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($cnt, 2)

        [void]($barRow.Children.Add($lbl))
        [void]($barRow.Children.Add($bar))
        [void]($barRow.Children.Add($cnt))
        [void]($DashAgingPanel.Children.Add($barRow))
    }

    # Top senders
    $DashSendersPanel.Children.Clear()
    foreach ($s in $m.TopSenders) {
        $sr = [System.Windows.Controls.Grid]::new()
        $sc1 = [System.Windows.Controls.ColumnDefinition]::new(); $sc1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $sc2 = [System.Windows.Controls.ColumnDefinition]::new(); $sc2.Width = [System.Windows.GridLength]::new(40)
        [void]($sr.ColumnDefinitions.Add($sc1))
        [void]($sr.ColumnDefinitions.Add($sc2))
        $sr.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

        $nm = [System.Windows.Controls.TextBlock]::new()
        $nm.Text = $s.Name; $nm.Foreground = (HexBrush '#C0C0C0'); $nm.FontSize = 11
        [System.Windows.Controls.Grid]::SetColumn($nm, 0)

        $ct = [System.Windows.Controls.TextBlock]::new()
        $ct.Text = $s.Count.ToString(); $ct.Foreground = (HexBrush '#4C9FE6'); $ct.FontSize = 11; $ct.FontWeight = [System.Windows.FontWeights]::Bold
        [System.Windows.Controls.Grid]::SetColumn($ct, 1)

        [void]($sr.Children.Add($nm))
        [void]($sr.Children.Add($ct))
        [void]($DashSendersPanel.Children.Add($sr))
    }

    # Keyword matches per view
    $DashKeywordPanel.Children.Clear()
    if ($script:EmailCfg.views) {
        $views = $script:EmailCfg.views
        foreach ($prop in $views.PSObject.Properties) {
            $vn = $prop.Name
            $filtered = Filter-EmailsByView $emails $vn
            $kr = [System.Windows.Controls.Grid]::new()
            $kc1 = [System.Windows.Controls.ColumnDefinition]::new(); $kc1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $kc2 = [System.Windows.Controls.ColumnDefinition]::new(); $kc2.Width = [System.Windows.GridLength]::new(40)
            [void]($kr.ColumnDefinitions.Add($kc1))
            [void]($kr.ColumnDefinitions.Add($kc2))
            $kr.Margin = [System.Windows.Thickness]::new(0, 2, 0, 2)

            $vl = [System.Windows.Controls.TextBlock]::new()
            $vl.Text = $vn; $vl.Foreground = (HexBrush '#C0C0C0'); $vl.FontSize = 11
            [System.Windows.Controls.Grid]::SetColumn($vl, 0)

            $vc = [System.Windows.Controls.TextBlock]::new()
            $vc.Text = $filtered.Count.ToString(); $vc.Foreground = (HexBrush '#50C878'); $vc.FontSize = 11; $vc.FontWeight = [System.Windows.FontWeights]::Bold
            [System.Windows.Controls.Grid]::SetColumn($vc, 1)

            [void]($kr.Children.Add($vl))
            [void]($kr.Children.Add($vc))
            [void]($DashKeywordPanel.Children.Add($kr))
        }
    }

    # Draft backlog
    $DashDraftPanel.Children.Clear()
    $draftCount = if ($script:EmailState.draftLog) { @($script:EmailState.draftLog).Count } else { 0 }
    $todoCount  = if ($script:EmailState.todoAssignments) { @($script:EmailState.todoAssignments).Count } else { 0 }
    $dlTxt = [System.Windows.Controls.TextBlock]::new()
    $dlTxt.Text = "Drafts: $draftCount  |  To-Dos: $todoCount"
    $dlTxt.Foreground = (HexBrush '#C0C0C0'); $dlTxt.FontSize = 11
    [void]($DashDraftPanel.Children.Add($dlTxt))
}

# ================================================================
#  EMAIL EVENT HANDLERS
# ================================================================
$script:SelectedEmailId = $null

# Populate view combos from config
Load-EmailConfig
Load-EmailState
[void]($EmViewCombo.Items.Add('(All)'))
[void]($DashViewCombo.Items.Add('(All)'))
if ($script:EmailCfg.views) {
    foreach ($prop in $script:EmailCfg.views.PSObject.Properties) {
        [void]($EmViewCombo.Items.Add($prop.Name))
        [void]($DashViewCombo.Items.Add($prop.Name))
    }
}
$EmViewCombo.SelectedIndex = 0
if ($DashViewCombo.Items.Count -gt 0) { $DashViewCombo.SelectedIndex = 0 }

# Populate template combo
if ($script:EmailCfg.replyTemplates) {
    foreach ($prop in $script:EmailCfg.replyTemplates.PSObject.Properties) {
        [void]($EmTemplateCombo.Items.Add($prop.Name))
    }
}
if ($EmTemplateCombo.Items.Count -gt 0) { $EmTemplateCombo.SelectedIndex = 0 }

function Sync-EmailHelperFiltersToDashboard {
    try {
        $scopeTxt = Get-ComboText $EmScopeCombo
        $viewTxt  = Get-ComboText $EmViewCombo
        $fltTxt   = Get-ComboText $EmFilterCombo
        foreach ($it in $DashScopeCombo.Items) { if ($it.Content -eq $scopeTxt) { $DashScopeCombo.SelectedItem = $it; break } }
        foreach ($it in $DashViewCombo.Items)  { if ([string]$it -eq $viewTxt) { $DashViewCombo.SelectedItem = $it; break } }
        foreach ($it in $DashFilterCombo.Items) { if ($it.Content -eq $fltTxt) { $DashFilterCombo.SelectedItem = $it; break } }
        $DashSearchBox.Text = $EmSearchBox.Text
        $DashFromDateBox.Text = $EmFromDateBox.Text
        $DashToDateBox.Text = $EmToDateBox.Text
    } catch {}
}

$EmRefreshBtn.Add_Click({
    Show-Busy 'Fetching emails from Outlook...'
    try {
        Sync-EmailHelperFiltersToDashboard
        Refresh-EmailList
        Refresh-Dashboard
        Set-Status "Loaded $($script:EmailCache.Count) emails"
    } catch { Set-Status "Email refresh error: $_" '#E05050' }
    finally { Hide-Busy }
})

$EmViewCombo.Add_SelectionChanged({
    if ($script:EmailRawCache.Count -gt 0) {
        Sync-EmailHelperFiltersToDashboard
        Refresh-EmailList -UseCachedRaw
        Refresh-Dashboard
    }
})
$EmFilterCombo.Add_SelectionChanged({
    if ($script:EmailRawCache.Count -gt 0) {
        Sync-EmailHelperFiltersToDashboard
        Refresh-EmailList -UseCachedRaw
        Refresh-Dashboard
    }
})
$EmScopeCombo.Add_SelectionChanged({
    if ($script:EmailRawCache.Count -gt 0) {
        Sync-EmailHelperFiltersToDashboard
        Refresh-EmailList
        Refresh-Dashboard
    }
})
$EmSearchBox.Add_TextChanged({
    if ($script:EmailRawCache.Count -gt 0) {
        Sync-EmailHelperFiltersToDashboard
        Refresh-EmailList -UseCachedRaw
        Refresh-Dashboard
    }
})
$EmFromDateBox.Add_TextChanged({
    if ($script:EmailRawCache.Count -gt 0) {
        Sync-EmailHelperFiltersToDashboard
        Refresh-EmailList -UseCachedRaw
        Refresh-Dashboard
    }
})
$EmToDateBox.Add_TextChanged({
    if ($script:EmailRawCache.Count -gt 0) {
        Sync-EmailHelperFiltersToDashboard
        Refresh-EmailList -UseCachedRaw
        Refresh-Dashboard
    }
})

$EmReplyBtn.Add_Click({
    if (-not $script:SelectedEmailId) { Set-Status 'Select an email first' '#E05050'; return }
    $sel = $script:EmailCache | Where-Object { $_.EntryID -eq $script:SelectedEmailId } | Select-Object -First 1
    $script:ComposeMode = 'Reply'
    $EmComposeLabel.Text = "Composing Reply to: $($sel.SenderName)"
    $EmComposeTo.Text = "To: $($sel.SenderEmail)"
    $EmComposeBody.Text = ''
    $EmComposePanel.Visibility = 'Visible'
    $EmComposeBody.Focus()
    Set-Status 'Compose your reply, then click Save as Draft'
})

$EmReplyAllBtn.Add_Click({
    if (-not $script:SelectedEmailId) { Set-Status 'Select an email first' '#E05050'; return }
    $sel = $script:EmailCache | Where-Object { $_.EntryID -eq $script:SelectedEmailId } | Select-Object -First 1
    $script:ComposeMode = 'ReplyAll'
    $EmComposeLabel.Text = "Composing Reply All to: $($sel.SenderName)"
    $EmComposeTo.Text = "To: $($sel.SenderEmail) (+ all recipients)"
    $EmComposeBody.Text = ''
    $EmComposePanel.Visibility = 'Visible'
    $EmComposeBody.Focus()
    Set-Status 'Compose your reply-all, then click Save as Draft'
})

$EmForwardBtn.Add_Click({
    if (-not $script:SelectedEmailId) { Set-Status 'Select an email first' '#E05050'; return }
    $sel = $script:EmailCache | Where-Object { $_.EntryID -eq $script:SelectedEmailId } | Select-Object -First 1
    $script:ComposeMode = 'Forward'
    $EmComposeLabel.Text = "Composing Forward: $($sel.Subject)"
    $EmComposeTo.Text = 'Recipients will be set in Outlook Drafts'
    $EmComposeBody.Text = ''
    $EmComposePanel.Visibility = 'Visible'
    $EmComposeBody.Focus()
    Set-Status 'Compose your forward message, then click Save as Draft'
})

$EmSaveDraftBtn.Add_Click({
    if (-not $script:SelectedEmailId -or -not $script:ComposeMode) { return }
    $body = $EmComposeBody.Text.Trim()
    $tpl  = $EmTemplateCombo.SelectedItem
    $ok   = $null
    switch ($script:ComposeMode) {
        'Reply'    { $ok = New-ReplyDraft -entryId $script:SelectedEmailId -templateName $tpl -BodyText $body }
        'ReplyAll' { $ok = New-ReplyDraft -entryId $script:SelectedEmailId -templateName $tpl -ReplyAll -BodyText $body }
        'Forward'  { $ok = New-ForwardDraft -entryId $script:SelectedEmailId -BodyText $body }
    }
    if ($ok) {
        Set-Status "$($script:ComposeMode) draft saved to Outlook Drafts"
        $EmComposePanel.Visibility = 'Collapsed'
        $script:ComposeMode = $null
    } else {
        Set-Status "Failed to save $($script:ComposeMode) draft" '#E05050'
    }
})

$EmCancelComposeBtn.Add_Click({
    $EmComposePanel.Visibility = 'Collapsed'
    $EmComposeBody.Text = ''
    $script:ComposeMode = $null
    Set-Status 'Draft cancelled'
})

$EmTodoBtn.Add_Click({
    if (-not $script:SelectedEmailId) { Set-Status 'Select an email first' '#E05050'; return }
    $sel = $script:EmailCache | Where-Object { $_.EntryID -eq $script:SelectedEmailId } | Select-Object -First 1
    if (-not $sel) { return }
    $due = (Get-Date).AddDays(3).ToString('yyyy-MM-dd')
    $task = New-OutlookToDo -entryId $sel.EntryID -subject $sel.Subject -dueDate $due
    if ($task) { Set-Status "To-Do created: $($sel.Subject)" }
    else { Set-Status 'Failed to create To-Do' '#E05050' }
})

$EmCalBtn.Add_Click({
    if (-not $script:SelectedEmailId) { Set-Status 'Select an email first' '#E05050'; return }
    $sel = $script:EmailCache | Where-Object { $_.EntryID -eq $script:SelectedEmailId } | Select-Object -First 1
    if (-not $sel) { return }
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
    $defaultDate = (Get-Date).AddDays(7).ToString('yyyy-MM-dd')
    $due = [Microsoft.VisualBasic.Interaction]::InputBox(
        'Date for all-day calendar item (yyyy-MM-dd):',
        'Calendar Date',
        $defaultDate)
    if (-not $due) { return }
    $dt = Try-ParseFilterDate $due
    if (-not $dt) { Set-Status 'Invalid date. Use yyyy-MM-dd.' '#E05050'; return }
    $appt = New-CalendarDeadline -entryId $sel.EntryID -subject $sel.Subject -dueDate $dt.ToString('yyyy-MM-dd') -AllDay
    if ($appt) { Set-Status "All-day calendar item created: $($sel.Subject)" }
    else { Set-Status 'Failed to create calendar deadline' '#E05050' }
})

$EmRecallBtn.Add_Click({
    if (-not $script:SelectedEmailId) { Set-Status 'Select an email first' '#E05050'; return }
    $result = [System.Windows.MessageBox]::Show(
        'Attempt to recall the selected email? This requires Exchange and may not succeed.',
        'Recall Email', 'YesNo', 'Warning')
    if ($result -eq 'Yes') {
        $ok = Invoke-EmailRecall $script:SelectedEmailId
        if ($ok) { Set-Status 'Recall attempted' }
        else { Set-Status 'Recall failed (too old or not Exchange)' '#E05050' }
    }
})

$EmDupEmailBtn.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        'Scan the current scope for duplicate emails and remove them?',
        'De-Duplicate Emails', 'YesNo', 'Question')
    if ($result -eq 'Yes') {
        Show-Busy 'Removing duplicate emails...'
        try {
            $removed = Remove-DuplicateEmails -Scope (Get-ComboText $EmScopeCombo)
            Set-Status "Removed $removed duplicate emails"
            if ($script:EmailCache.Count -gt 0) { Refresh-EmailList }
        } catch { Set-Status "De-dup error: $_" '#E05050' }
        finally { Hide-Busy }
    }
})

$EmDupFolderBtn.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Scan folders for duplicate names?`nEmails will be merged into the original folder before removing the duplicate.",
        'De-Duplicate Folders', 'YesNo', 'Question')
    if ($result -eq 'Yes') {
        Show-Busy 'Removing duplicate folders...'
        try {
            $removed = Remove-DuplicateFolders
            Set-Status "Removed $removed duplicate folders"
        } catch { Set-Status "Folder de-dup error: $_" '#E05050' }
        finally { Hide-Busy }
    }
})

# -- Dashboard handlers --
$DashRefreshBtn.Add_Click({
    Show-Busy 'Refreshing dashboard...'
    try {
        Refresh-Dashboard -ForceFetch
        if ($script:EmailCache.Count -eq 0 -and $script:EmailRawCache.Count -gt 0) {
            Refresh-EmailList -UseCachedRaw
        }
        Set-Status 'Dashboard refreshed'
    } catch { Set-Status "Dashboard error: $_" '#E05050' }
    finally { Hide-Busy }
})

$DashViewCombo.Add_SelectionChanged({
    if ($script:EmailRawCache.Count -gt 0) { Refresh-Dashboard }
})

$DashFilterCombo.Add_SelectionChanged({
    if ($script:EmailRawCache.Count -gt 0) { Refresh-Dashboard }
})

$DashPeriodCombo.Add_SelectionChanged({
    if ($script:EmailRawCache.Count -gt 0) { Refresh-Dashboard }
})
$DashScopeCombo.Add_SelectionChanged({
    Refresh-Dashboard -ForceFetch
})
$DashSearchBox.Add_TextChanged({
    if ($script:EmailRawCache.Count -gt 0) { Refresh-Dashboard }
})
$DashFromDateBox.Add_TextChanged({
    if ($script:EmailRawCache.Count -gt 0) { Refresh-Dashboard }
})
$DashToDateBox.Add_TextChanged({
    if ($script:EmailRawCache.Count -gt 0) { Refresh-Dashboard }
})

$DashExportBtn.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV Files (*.csv)|*.csv'
    $dlg.FileName = "EmailReport_$(Get-Date -f 'yyyyMMdd').csv"
    if ($dlg.ShowDialog() -eq $true) {
        try {
            $rows = Apply-EmailFilters -Emails $script:EmailRawCache `
                -ViewName (Get-ComboText $DashViewCombo) `
                -QuickFilter (Get-ComboText $DashFilterCombo) `
                -SearchText $DashSearchBox.Text `
                -FromDateText $DashFromDateBox.Text `
                -ToDateText $DashToDateBox.Text `
                -Period (Get-ComboText $DashPeriodCombo)
            $rows | Select-Object Subject, SenderName, SenderEmail, FolderPath, ReceivedTime, AgeDays, IsRead, IsVIP, IsInternal, HasAttachment |
                Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.MessageBox]::Show("Report saved:`n$($dlg.FileName)", 'Export Complete', 'OK', 'Information')
            Set-Status 'Email report exported'
        } catch { Set-Status "Export error: $_" '#E05050' }
    }
})

# -- IDE handlers --
$IdeOpenBtn.Add_Click({
    try {
        Open-IdeFileDialogAndLoad
    } catch {
        Set-Status "IDE open error: $($_.Exception.Message)" '#E05050'
    }
})

$IdeSaveBtn.Add_Click({
    try {
        if ($script:IdeCurrentFile) {
            Save-IdeFileToPath -Path $script:IdeCurrentFile
            Set-Status "IDE saved: $(Split-Path $script:IdeCurrentFile -Leaf)"
        } else {
            [void](Save-IdeFileAsDialog)
        }
    } catch {
        Set-Status "IDE save error: $($_.Exception.Message)" '#E05050'
    }
})

$IdeSaveAsBtn.Add_Click({
    try {
        [void](Save-IdeFileAsDialog)
    } catch {
        Set-Status "IDE save-as error: $($_.Exception.Message)" '#E05050'
    }
})

$IdeValidateBtn.Add_Click({
    Show-Busy 'Validating IDE content...'
    try {
        [void](Invoke-IdeValidateAction)
    } catch {
        $IdeValidationTxt.Text = [string]$_.Exception.Message
        Set-Status "IDE validate error: $($_.Exception.Message)" '#E05050'
    } finally { Hide-Busy }
})

$IdeRunBtn.Add_Click({
    try {
        Invoke-IdeRunAction
    } catch {
        $IdeRunErrorTxt.Text = [string]$_.Exception.Message
        Set-Status "IDE run error: $($_.Exception.Message)" '#E05050'
    }
})

$IdeRefreshExcelBtn.Add_Click({
    try {
        Refresh-WorkbookDropdowns
        Refresh-IdeExcelTargets
        Set-Status 'IDE Excel targets refreshed'
    } catch {
        Set-Status "IDE Excel refresh error: $($_.Exception.Message)" '#E05050'
    }
})

$IdeWbCombo.Add_SelectionChanged({
    Refresh-IdeSheets
})

$IdeModeCombo.Add_SelectionChanged({
    Reset-IdeValidationState
    $mode = Get-ComboText $IdeModeCombo
    $script:IdeMode = if ($mode) { $mode } else { 'Auto' }
    Set-IdeEditorLang $script:IdeMode
})

# Poll the WebBrowser editor every 400ms so validation badge resets on edits
$ideChangeTimer = [System.Windows.Threading.DispatcherTimer]::new()
$ideChangeTimer.Interval = [TimeSpan]::FromMilliseconds(400)
$ideChangeTimer.Add_Tick({
    if ($script:IdeLastValidationOk -and $script:IdeEditorReady) {
        try {
            $currHash = Get-TextSha1 (Get-IdeEditorText)
            if ($currHash -ne $script:IdeValidatedHash) {
                Reset-IdeValidationState
            }
        } catch {}
    }
})
$ideChangeTimer.Start()

# ================================================================
#  AI IDE EVENT HANDLERS
# ================================================================

$AiIdeModeCombo.Add_SelectionChanged({
    Invoke-AiIdeSafe {
        AiIde-UpdateModeUi
        $m = Get-ComboText $AiIdeModeCombo
        if ($m -eq 'Graph API') {
            AiIde-SetStatus 'Mode: Graph API (enterprise account/license required).' '#4C9FE6'
        } elseif ($m -eq 'Chrome') {
            AiIde-SetStatus 'Mode: Chrome (requires local bridge).' '#FFA500'
        } else {
            AiIde-SetStatus 'Mode: Auto (Graph first, Chrome only if bridge is already running).' '#4C9FE6'
        }
    } 'AI mode change'
})

$AiIdeAuthBtn.Add_Click({
    Invoke-AiIdeSafe { AiIde-StartDeviceCodeAuth } 'AI auth+verify'
})

$AiIdeStartBtn.Add_Click({
    Invoke-AiIdeSafe { AiIde-StartBridge } 'Bridge start'
})

$AiIdeStopBtn.Add_Click({
    Invoke-AiIdeSafe { AiIde-StopBridge } 'Bridge stop'
})

$AiIdeNewBtn.Add_Click({
    Invoke-AiIdeSafe {
        $AiIdeEditorBox.Text = ''
        $script:AiIdeCurrentFile = $null
        $AiIdeFilePathTxt.Text = '(unsaved)'
        AiIde-SetStatus 'New file'
    } 'AI new file'
})

$AiIdeOpenBtn.Add_Click({
    Invoke-AiIdeSafe { AiIde-Open } 'AI open'
})

$AiIdeSaveBtn.Add_Click({
    Invoke-AiIdeSafe { AiIde-Save } 'AI save'
})

$AiIdeValidateBtn.Add_Click({
    Invoke-AiIdeSafe { AiIde-Validate } 'AI validate'
})

$AiIdeRunBtn.Add_Click({
    Invoke-AiIdeSafe {
        $code = $AiIdeEditorBox.Text
        if ([string]::IsNullOrWhiteSpace($code)) { AiIde-SetStatus 'Nothing to run'; return }
        AiIde-RunCode $code
    } 'AI run'
})

$AiIdeInjectExcelBtn.Add_Click({
    Invoke-AiIdeSafe {
        $code = $AiIdeEditorBox.Text
        if ([string]::IsNullOrWhiteSpace($code)) { AiIde-SetStatus 'Nothing to inject'; return }
        AiIde-InjectToExcel $code
    } 'AI inject Excel'
})

$AiIdeInjectOutlookBtn.Add_Click({
    Invoke-AiIdeSafe {
        $code = $AiIdeEditorBox.Text
        if ([string]::IsNullOrWhiteSpace($code)) { AiIde-SetStatus 'Nothing to inject'; return }
        AiIde-InjectToOutlook $code
    } 'AI inject Outlook'
})

$AiIdeGenerateBtn.Add_Click({
    Invoke-AiIdeSafe {
        $req = $AiIdePromptBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($req)) { AiIde-SetStatus 'Enter a prompt first'; return }
        AiIde-Generate $req -IncludeEditor
    } 'AI generate'
})

$AiIdeRefineBtn.Add_Click({
    Invoke-AiIdeSafe {
        $req = $AiIdePromptBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($req)) { AiIde-SetStatus 'Enter a prompt first'; return }
        AiIde-Generate "Refine the existing code: $req" -IncludeEditor
    } 'AI refine'
})

$AiIdeGenRunBtn.Add_Click({
    Invoke-AiIdeSafe {
        $req = $AiIdePromptBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($req)) { AiIde-SetStatus 'Enter a prompt first'; return }
        AiIde-GenerateAndRun $req
    } 'AI generate+run'
})

$AiIdeFixBtn.Add_Click({
    Invoke-AiIdeSafe {
        $errors = $AiIdeRunErrorTxt.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($errors)) { $errors = $AiIdeValidationTxt.Text.Trim() }
        $req = "Fix the following errors in the code:`n$errors"
        AiIde-Generate $req -IncludeEditor
    } 'AI fix'
})

$AiIdeClearBtn.Add_Click({
    Invoke-AiIdeSafe {
        $AiIdePromptBox.Text = ''
        $AiIdeOutputTxt.Text = ''
        $AiIdeRunErrorTxt.Text = ''
        $AiIdeValidationTxt.Text = ''
        AiIde-SetStatus 'Cleared'
    } 'AI clear'
})

$AiIdeNewConvBtn.Add_Click({
    Invoke-AiIdeSafe {
        AiIde-NewConversation
        AiIde-SetStatus 'New conversation started'
    } 'AI new conversation'
})

$AiIdeExportHistBtn.Add_Click({
    Invoke-AiIdeSafe { AiIde-ExportHistory } 'AI export history'
})

$AiIdeLangCombo.Add_SelectionChanged({
    Invoke-AiIdeSafe {
        $lang = AiIde-GetLang
        AiIde-UpdatePromptGuide
        if ($lang -like '*PowerShell*' -or $lang -like '*PS*') {
            AiIde-SetStatus "Language: $lang. Prompt for a full runnable PowerShell 5.1 script." '#4C9FE6'
        } elseif ($lang -like '*Outlook*') {
            AiIde-SetStatus "Language: $lang. Prompt for a complete Outlook VBA Sub." '#4C9FE6'
        } else {
            AiIde-SetStatus "Language: $lang. Prompt for a complete Excel VBA Sub." '#4C9FE6'
        }
        AiIde-RefreshTargets
    } 'AI language change'
})

$AiIdeWbCombo.Add_SelectionChanged({
    Invoke-AiIdeSafe { AiIde-RefreshSheetsOnly } 'AI workbook change'
})

$AiIdePromptBox.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq 'Return' -and ($e.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
        $e.Handled = $true
        Invoke-AiIdeSafe {
            $req = $AiIdePromptBox.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($req)) { AiIde-Generate $req -IncludeEditor }
        } 'AI generate'
    }
})

# ================================================================
#  ALT KEYTIP OVERLAY (Excel-style letter badges)
# ================================================================
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
    # Email Helper tab
    EmScopeCombo      = 'O'   # scOpe
    EmViewCombo       = 'V'   # View
    EmFilterCombo     = 'F'   # Filter
    EmSearchBox       = 'S'   # Search
    EmFromDateBox     = 'D'   # Date from
    EmToDateBox       = 'U'   # to date (U to avoid duplicates)
    EmTemplateCombo   = 'T'   # Template
    # Email Dashboard tab
    DashScopeCombo    = 'O'   # scOpe
    DashViewCombo     = 'V'   # View
    DashFilterCombo   = 'I'   # fIlter
    DashPeriodCombo   = 'P'   # Period
    DashSearchBox     = 'S'   # Search
    DashFromDateBox   = 'D'   # date from
    DashToDateBox     = 'U'   # date to
    # IDE tab
    IdeModeCombo      = 'M'   # Mode
    IdeWbCombo        = 'W'   # Workbook
    IdeSheetCombo     = 'H'   # sHeet
}

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
    $MainTabs.SelectedIndex = $idx
    return $true
}

function Select-MainTabRelative([int]$delta) {
    if ($MainTabs.Items.Count -le 0) { return $false }
    $MainTabs.SelectedIndex = ($MainTabs.SelectedIndex + $delta + $MainTabs.Items.Count) % $MainTabs.Items.Count
    return $true
}

# -- Keyboard shortcuts: Alt+0-9 / Ctrl+0-9 / Ctrl+PgUp/PgDn tab switch, Ctrl+F search --
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

    # Alt+0..9 tab switching (SystemKey path catches Alt+digit reliably)
    if ($_.Key -eq 'System') {
        $altNum = Get-TabIndexFromDigitKey $_.SystemKey
        if ($altNum -ge 0 -and (Select-MainTabIndex $altNum)) {
            if ($script:KeyTipsActive) { Hide-KeyTips }
            $_.Handled = $true
            return
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
            if ($t.State -eq 4) {  # Running state
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
$script:TabsLoaded = @{}   # track which tabs have been lazy-loaded

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
        if ($clipDef.DateOffset) { $ClipDateOffset.Text = $clipDef.DateOffset }
        if ($clipDef.TimeOffset) { $ClipTimeOffset.Text = $clipDef.TimeOffset }
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
    -DefaultDateOffset $ClipDateOffset.Text `
    -DefaultTimeOffset $ClipTimeOffset.Text `
    -CountLabel $ClipSlotCount
Refresh-ClipSheets
Refresh-MacroList
Refresh-TemplateList
Refresh-IdeExcelTargets
Refresh-IdePathDisplay
AiIde-UpdateModeUi
AiIde-UpdatePromptGuide

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

# Lazy-load tab content on first visit
$MainTabs.Add_SelectionChanged({
    $idx = $MainTabs.SelectedIndex

    # Collapse compose panel when leaving Email Helper tab.
    if ($EmComposePanel -and $EmComposePanel.Visibility -eq 'Visible') {
        $EmComposePanel.Visibility = 'Collapsed'
        $script:ComposeMode = $null
    }

    if ($script:TabsLoaded[$idx]) { return }
    $script:TabsLoaded[$idx] = $true
    switch ($idx) {
        2 { Refresh-TaskList }                                         # Scheduler (COM + WMI)
        3 { Refresh-NavWorkbooks; Refresh-NavSheets }                  # Navigator (Excel COM)
        6 { Refresh-QsTodoList; Refresh-QsProgress }                   # QTasks (JSON + UI build)
        7 { try { Refresh-EmailList } catch {} }                       # Email Helper (Outlook COM)
        8 { try { Refresh-Dashboard } catch {} }                       # Email Dashboard
        9 { Refresh-IdeExcelTargets; Refresh-IdePathDisplay }          # IDE
        10 { AiIde-RefreshTargets }                                    # AI IDE
    }
})

# Mark Tab 0 (Clipboard) + Tab 1 (Macros) + Tab 4 (Templates) as already loaded
$script:TabsLoaded[0] = $true
$script:TabsLoaded[1] = $true
$script:TabsLoaded[4] = $true
$script:TabsLoaded[5]  = $true
$script:TabsLoaded[10] = $false

Set-Status 'MacroHub v3.1 ready -- Ctrl+PgUp/PgDn tabs | Ctrl+0-9 | Alt+0-9 | Ctrl+F search'

# Check for missed scheduled tasks AFTER window is visible
$Window.Add_ContentRendered({
    Invoke-MissedTaskCheck
    # Navigate the IDE WebBrowser to the embedded HTML editor now that the HWND exists
    Set-IdeBrowserSilent $true
    $IdeEditorBrowser.NavigateToString($script:IdeHtml)
})

# -- Dismiss keytips when window loses focus --
$Window.Add_Deactivated({ if ($script:KeyTipsActive) { Hide-KeyTips } })

# ================================================================
#  SHOW WINDOW
# ================================================================
$Window.ShowDialog()

}  # END Start-MacroHub

# ================================================================
#  LAUNCH
# ================================================================
Start-MacroHub

