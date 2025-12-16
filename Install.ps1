#Requires -Version 5.1
<#
.SYNOPSIS
    MacroHub v3.0 Installer - run this ONCE per machine to create a desktop shortcut.

.DESCRIPTION
    Pure PowerShell only. No VBS, no executables, no admin rights.

    Creates a desktop shortcut that launches MacroHub.ps1 directly
    from this OneDrive folder using:
        powershell.exe -WindowStyle Hidden -File "MacroHub.ps1"

    The -WindowStyle Hidden suppresses the black console window while
    the WPF GUI window still appears normally.

    USAGE:
        Right-click Install.ps1 -> "Run with PowerShell"

    Safe to run multiple times (just overwrites the shortcut).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '=======================================' -ForegroundColor Cyan
Write-Host '  MacroHub v3.0 Installer' -ForegroundColor Cyan
Write-Host '=======================================' -ForegroundColor Cyan
Write-Host ''

# ============================================================
#  STEP 1: VERIFY WE ARE IN THE RIGHT FOLDER
# ============================================================
$HubRoot   = $PSScriptRoot
$HubScript = Join-Path $HubRoot 'MacroHub.ps1'

Write-Host "Hub folder : $HubRoot" -ForegroundColor Gray
Write-Host "Main script: $HubScript" -ForegroundColor Gray
Write-Host ''

if (-not (Test-Path $HubScript)) {
    Write-Host 'ERROR: MacroHub.ps1 was not found in the same folder as Install.ps1.' -ForegroundColor Red
    Write-Host 'Make sure both files are in the same MacroHub folder, then run Install.ps1 again.' -ForegroundColor Yellow
    Read-Host 'Press Enter to exit'
    exit 1
}

# ============================================================
#  STEP 2: CREATE MACROS SUBFOLDER
# ============================================================
$MacroFolder = Join-Path $HubRoot 'Macros'
if (-not (Test-Path $MacroFolder)) {
    New-Item -ItemType Directory -Force -Path $MacroFolder | Out-Null
    Write-Host '[OK] Created Macros folder.' -ForegroundColor Green
} else {
    Write-Host '[OK] Macros folder already exists.' -ForegroundColor Green
}

# ============================================================
#  STEP 3: SET EXECUTION POLICY (CurrentUser scope - no admin)
# ============================================================
try {
    $current = Get-ExecutionPolicy -Scope CurrentUser
    if ($current -eq 'Restricted' -or $current -eq 'Undefined') {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host '[OK] Execution policy set to RemoteSigned for your user account.' -ForegroundColor Green
    } else {
        Write-Host "[OK] Execution policy already set: $current" -ForegroundColor Green
    }
} catch {
    Write-Host "[WARN] Could not update execution policy: $_" -ForegroundColor Yellow
    Write-Host '       If MacroHub will not open, try right-clicking MacroHub.ps1 and choosing Run with PowerShell.' -ForegroundColor Yellow
}

# ============================================================
#  STEP 4: CREATE DESKTOP SHORTCUT (pure PowerShell, no VBS)
#
#  The shortcut runs:
#      powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File "MacroHub.ps1"
#
#  -WindowStyle Hidden hides the black console window.
#  The WPF window created by ShowDialog() is a separate window
#  and appears normally regardless of -WindowStyle Hidden.
# ============================================================

# Resolve desktop path (handles OneDrive-redirected desktops on Windows 11)
$desktopPath = [System.Environment]::GetFolderPath('Desktop')
if (-not (Test-Path $desktopPath)) {
    # OneDrive-redirected desktop fallback
    $desktopPath = Join-Path $env:OneDrive 'Desktop'
}
$shortcutPath = Join-Path $desktopPath 'MacroHub.lnk'

try {
    $shell    = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)

    # Target: powershell.exe with flags to suppress console and run our script
    $shortcut.TargetPath       = 'powershell.exe'
    $shortcut.Arguments        = "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$HubScript`""
    $shortcut.WorkingDirectory = $HubRoot
    $shortcut.Description      = 'MacroHub v3.0 - Office Productivity Super App'
    $shortcut.WindowStyle      = 1   # 1 = Normal window (for the WPF popup itself)

    # Use custom icon if present, otherwise PowerShell default
    $icoPath = Join-Path $HubRoot 'MacroHub.ico'
    if (Test-Path $icoPath) {
        $shortcut.IconLocation = "$icoPath, 0"
    } else {
        $shortcut.IconLocation = 'powershell.exe, 0'
    }

    $shortcut.Save()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null

    Write-Host "[OK] Desktop shortcut created: $shortcutPath" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Could not create shortcut: $_" -ForegroundColor Red
    Write-Host '       You can still launch MacroHub by right-clicking MacroHub.ps1 and choosing Run with PowerShell.' -ForegroundColor Yellow
}

# ============================================================
#  STEP 5: UNBLOCK ALL FILES (OneDrive / download safety)
# ============================================================
try {
    Get-ChildItem $HubRoot -Recurse | Unblock-File -ErrorAction SilentlyContinue
    Write-Host '[OK] Unblocked all MacroHub files.' -ForegroundColor Green
} catch {
    Write-Host "[WARN] Could not unblock some files: $_" -ForegroundColor Yellow
}

# ============================================================
#  STEP 6: VERIFY
# ============================================================
$ok = (Test-Path $shortcutPath) -and (Test-Path $HubScript)
if ($ok) {
    Write-Host ''
    Write-Host '=======================================' -ForegroundColor Cyan
    Write-Host '  Installation complete!' -ForegroundColor Green
    Write-Host '=======================================' -ForegroundColor Cyan
} else {
    Write-Host ''
    Write-Host '[WARN] Something may not have installed correctly. Check the messages above.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host '  What to do next:' -ForegroundColor White
Write-Host ''
Write-Host '  1. Double-click the MacroHub shortcut on your desktop.' -ForegroundColor White
Write-Host ''
Write-Host '  2. Add macro files (.bas or .ps1) to the Macros folder:' -ForegroundColor White
Write-Host "     $MacroFolder" -ForegroundColor Gray
Write-Host ''
Write-Host '  3. Share this OneDrive folder link with colleagues.' -ForegroundColor White
Write-Host '     Each person runs Install.ps1 once, then uses the shortcut.' -ForegroundColor White
Write-Host ''
Write-Host '  Updates: Edit MacroHub.ps1 on OneDrive. Everyone gets' -ForegroundColor White
Write-Host '  the changes automatically on their next launch.' -ForegroundColor White
Write-Host ''

Read-Host 'Press Enter to close'
