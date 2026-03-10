#Requires -Version 5.1
<#
.SYNOPSIS
    MacroHub v3.1 Full Setup - installs the app, Chrome extension, and AI IDE dependencies.

.DESCRIPTION
    Run this script ONCE per machine (right-click -> Run with PowerShell).
    No admin rights required.

    What it does:
      1. Rebuilds MacroHub.ps1 and Chrome extension files from export TXT files (if needed)
      2. Verifies MacroHub.ps1 exists alongside this script
      3. Creates required subfolders (Macros, ChromeExtension, quarters, etc.)
      4. Sets execution policy to RemoteSigned (CurrentUser)
      5. Creates a desktop shortcut
      6. Unblocks all files (OneDrive download safety)
      7. Verifies Chrome extension files and exports setup/code TXT docs
      8. Verifies the AI IDE HttpListener port 9876 is available
      9. Writes text-only setup package to F:\ flash drive (if present)
      10. Runs a quick syntax check on MacroHub.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# --- Banner -----------------------------------------------------------------
Write-Host ''
Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host '  MacroHub v3.1 Full Setup (AI IDE + Chrome Extension)' -ForegroundColor Cyan
Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host ''

$HubRoot   = $PSScriptRoot
$HubScript = Join-Path $HubRoot 'MacroHub.ps1'
$ChromeExtDir = Join-Path $HubRoot 'ChromeExtension'
$SetupScriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }

Write-Host "Hub folder : $HubRoot" -ForegroundColor Gray
Write-Host "Main script: $HubScript" -ForegroundColor Gray
Write-Host ''

function Read-NormalizedTextFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        $txt = [System.IO.File]::ReadAllText($Path)
        if ($null -eq $txt) { return $null }
        if ($txt.Length -gt 0 -and [int][char]$txt[0] -eq 0xFEFF) {
            $txt = $txt.Substring(1)
        }
        # Allow running directly from markdown/codefenced exports.
        $txt = $txt -replace '^\s*```[a-zA-Z0-9_-]*\s*\r?\n', ''
        $txt = $txt -replace '\r?\n\s*```\s*$', ''
        return $txt
    } catch {
        return $null
    }
}

function Try-RebuildMacroHubFromTextExports {
    param([string]$RootPath, [string]$TargetScriptPath)

    $candidates = @(
        (Join-Path $RootPath 'for export\MacroHub_Code.txt'),
        (Join-Path $RootPath 'MacroHub_Code.txt')
    )
    foreach ($cand in $candidates) {
        $raw = Read-NormalizedTextFile $cand
        if (-not $raw) { continue }
        if ($raw -notmatch '(?m)^\s*#Requires\s+-Version\s+5\.1') { continue }
        if ($raw -notmatch '(?m)^\s*function\s+Start-MacroHub\s*\{') { continue }
        [System.IO.File]::WriteAllText($TargetScriptPath, $raw, [System.Text.UTF8Encoding]::new($false))
        Write-Host "[OK] Rebuilt MacroHub.ps1 from export file: $cand" -ForegroundColor Green
        return $true
    }
    return $false
}

function Parse-ChromeExtensionCodeExport {
    param([string]$RawText)

    $map = @{}
    if (-not $RawText) { return $map }
    $lines = $RawText -split "\r?\n"
    $current = ''
    $buffer = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        if ($line -match '^FILE:\s*(.+)$') {
            if ($current) {
                $content = ($buffer -join "`r`n").Trim("`r","`n")
                $map[$current] = $content
                $buffer.Clear()
            }
            $current = $Matches[1].Trim()
            continue
        }
        if (-not $current) { continue }
        if ($line -match '^={20,}$') { continue }
        $buffer.Add($line)
    }
    if ($current) {
        $content = ($buffer -join "`r`n").Trim("`r","`n")
        $map[$current] = $content
    }
    return $map
}

function Try-RebuildChromeExtensionFromTextExports {
    param([string]$RootPath, [string]$ExtDir)

    $required = @('manifest.json','background.js','bridge.js','popup.html','popup.js')
    $candidates = @(
        (Join-Path $RootPath 'for export\ChromeExtension_Code.txt'),
        (Join-Path $RootPath 'ChromeExtension_Code.txt')
    )
    foreach ($cand in $candidates) {
        $raw = Read-NormalizedTextFile $cand
        if (-not $raw) { continue }
        $parts = Parse-ChromeExtensionCodeExport -RawText $raw
        $hasAll = $true
        foreach ($name in $required) {
            if (-not $parts.ContainsKey($name) -or [string]::IsNullOrWhiteSpace([string]$parts[$name])) {
                $hasAll = $false
                break
            }
        }
        if (-not $hasAll) { continue }
        if (-not (Test-Path $ExtDir)) {
            New-Item -ItemType Directory -Path $ExtDir -Force | Out-Null
        }
        foreach ($name in $required) {
            [System.IO.File]::WriteAllText((Join-Path $ExtDir $name), [string]$parts[$name], [System.Text.UTF8Encoding]::new($false))
        }
        Write-Host "[OK] Rebuilt Chrome extension files from export file: $cand" -ForegroundColor Green
        return $true
    }
    return $false
}

# --- STEP 0: Bootstrap from TXT exports (zero manual formatting) ------------
Write-Host '[INFO] Bootstrap check: rebuilding files from TXT exports if needed...' -ForegroundColor Cyan

if (-not (Test-Path $HubScript)) {
    if (-not (Try-RebuildMacroHubFromTextExports -RootPath $HubRoot -TargetScriptPath $HubScript)) {
        Write-Host '[WARN] MacroHub.ps1 missing and no valid MacroHub_Code.txt export was found.' -ForegroundColor Yellow
    }
}

$requiredExt = @('manifest.json','background.js','bridge.js','popup.html','popup.js')
$missingExt = @()
foreach ($name in $requiredExt) {
    if (-not (Test-Path (Join-Path $ChromeExtDir $name))) { $missingExt += $name }
}
if ($missingExt.Count -gt 0) {
    [void](Try-RebuildChromeExtensionFromTextExports -RootPath $HubRoot -ExtDir $ChromeExtDir)
}

function Write-ChromeExtensionExports {
    param(
        [string]$RootPath,
        [string]$ExtDir
    )

    $exportDir = Join-Path $RootPath 'for export'
    if (-not (Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }

    $required = @('manifest.json', 'background.js', 'bridge.js', 'popup.html', 'popup.js')
    $missing = @()
    foreach ($name in $required) {
        if (-not (Test-Path (Join-Path $ExtDir $name))) { $missing += $name }
    }

    $guidePath = Join-Path $exportDir 'ChromeExtension_Setup_Guide.txt'
    $codePath  = Join-Path $exportDir 'ChromeExtension_Code.txt'

    $guide = @"
MacroHub Chrome Extension Setup Guide
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

Purpose
- Connects Chrome AI pages with MacroHub AI IDE bridge.
- Lets AI IDE send prompts in Chrome mode and receive responses safely.

Files expected in ChromeExtension folder
- manifest.json
- background.js
- bridge.js
- popup.html
- popup.js

Chrome install (Developer Mode)
1. Open Chrome and go to chrome://extensions/
2. Enable Developer mode (top right).
3. Click Load unpacked.
4. Select this folder:
   $ExtDir
5. Pin MacroHub Bridge extension (optional but recommended).

MacroHub-side setup
1. Open MacroHub.
2. Go to AI IDE tab.
3. Select Mode = Chrome (or Auto).
4. Click Start Bridge.
5. Confirm status shows bridge listening on port 9876.

Quick verification
1. Open copilot.microsoft.com in the same Chrome profile.
2. Open MacroHub Bridge popup and click Check Connection.
3. In AI IDE, click Generate (Chrome mode) for a short prompt.
4. Confirm AI response appears in AI IDE output panel.

Failure behavior
- If Chrome extension is missing/offline, MacroHub remains running.
- Errors stay in AI IDE status/output only.
- You can switch to Graph API mode at any time.

Troubleshooting
- If connection fails: close any app using port 9876 or change AI IDE port.
- Reload extension after updates from chrome://extensions/.
- Keep a Copilot tab open when using Chrome mode prompt injection.
"@
    [System.IO.File]::WriteAllText($guidePath, $guide, [System.Text.UTF8Encoding]::new($false))

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('MacroHub Chrome Extension - Full Source Export')
    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine('')
    foreach ($name in $required) {
        $fp = Join-Path $ExtDir $name
        [void]$sb.AppendLine(('=' * 86))
        [void]$sb.AppendLine("FILE: $name")
        [void]$sb.AppendLine(('=' * 86))
        if (Test-Path $fp) {
            [void]$sb.AppendLine([System.IO.File]::ReadAllText($fp))
        } else {
            [void]$sb.AppendLine("[MISSING] $name")
        }
        [void]$sb.AppendLine('')
    }
    [System.IO.File]::WriteAllText($codePath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))

    return [PSCustomObject]@{
        GuidePath = $guidePath
        CodePath  = $codePath
        Missing   = $missing
    }
}

function Write-AiIdeEnterpriseGuide {
    param([string]$RootPath)

    $exportDir = Join-Path $RootPath 'for export'
    if (-not (Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }

    $guidePath = Join-Path $exportDir 'AI_IDE_Enterprise_Setup.txt'
    $guide = @"
MacroHub AI IDE - Enterprise Work Computer Setup
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

Why this matters
- Graph Copilot API requires a licensed enterprise/work Microsoft 365 Copilot account.
- Personal/unlicensed accounts commonly return HTTP 403 (expected behavior).

Work-computer first-run checklist
1. Use your company-managed computer/profile.
2. Confirm you can sign in to Microsoft 365 with your work account.
3. Open MacroHub -> AI IDE tab.
4. Set Tenant (or keep 'common' if your org allows it).
5. Click: Auth + Verify.
6. Complete the device-code sign in using your work account.
7. Wait for status: Verified Graph Copilot access.

If Graph is blocked
- "Graph Copilot is blocked (403)" means the signed-in account/tenant is not licensed for Graph Copilot.
- On personal machines this is expected.
- On work machines, contact your M365 admin for Copilot licensing/Graph access.

Chrome fallback mode (optional)
1. In AI IDE set Mode = Chrome (or Auto).
2. Click Start Bridge.
3. Load the MacroHub extension in chrome://extensions (Developer mode -> Load unpacked).
4. Keep a copilot.microsoft.com tab open in the same Chrome profile.

Recommended mode
- Work computer with enterprise license: Graph API (or Auto).
- Personal computer without license: Chrome mode only.
"@
    [System.IO.File]::WriteAllText($guidePath, $guide, [System.Text.UTF8Encoding]::new($false))
    return $guidePath
}

function Write-CoreScriptExports {
    param(
        [string]$RootPath,
        [string]$HubScriptPath,
        [string]$SetupPath
    )

    $exportDir = Join-Path $RootPath 'for export'
    if (-not (Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }

    $utf8 = [System.Text.UTF8Encoding]::new($false)
    $hubTxtPath = Join-Path $exportDir 'MacroHub_Code.txt'
    $setupTxtPath = Join-Path $exportDir 'MacroHub_Setup.txt'

    if (Test-Path $HubScriptPath) {
        [System.IO.File]::WriteAllText($hubTxtPath, [System.IO.File]::ReadAllText($HubScriptPath), $utf8)
    }
    if (Test-Path $SetupPath) {
        [System.IO.File]::WriteAllText($setupTxtPath, [System.IO.File]::ReadAllText($SetupPath), $utf8)
    }

    return [PSCustomObject]@{
        HubCodePath   = $hubTxtPath
        SetupTextPath = $setupTxtPath
    }
}

# --- STEP 1: Verify MacroHub.ps1 --------------------------------------------
if (-not (Test-Path $HubScript)) {
    Write-Host '[FAIL] MacroHub.ps1 not found in this folder.' -ForegroundColor Red
    Write-Host '       Provide MacroHub.ps1, or provide MacroHub_Code.txt (root or for export\\) so setup can rebuild it.' -ForegroundColor Yellow
    Read-Host 'Press Enter to exit'
    exit 1
}
Write-Host '[OK] MacroHub.ps1 found.' -ForegroundColor Green

# --- STEP 2: Create required subfolders -------------------------------------
$subFolders = @(
    'Macros',
    'Macros\Modular',
    'Macros\Modular\Data',
    'Macros\Modular\Format',
    'Macros\Modular\Export',
    'Macros\Modular\Ops',
    'Macros\Modular\Quality',
    'Macros\Modular\VBA',
    'Macros\EmailTemplates',
    'ChromeExtension',
    'quarters'
)
foreach ($sub in $subFolders) {
    $full = Join-Path $HubRoot $sub
    if (-not (Test-Path $full)) {
        New-Item -ItemType Directory -Force -Path $full | Out-Null
        Write-Host "[OK] Created: $sub" -ForegroundColor Green
    }
}

# --- STEP 3: Execution policy ------------------------------------------------
try {
    $current = Get-ExecutionPolicy -Scope CurrentUser
    if ($current -eq 'Restricted' -or $current -eq 'Undefined') {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host '[OK] Execution policy set to RemoteSigned.' -ForegroundColor Green
    } else {
        Write-Host "[OK] Execution policy already: $current" -ForegroundColor Green
    }
} catch {
    Write-Host "[WARN] Could not set execution policy: $_" -ForegroundColor Yellow
}

# --- STEP 4: Desktop shortcut ------------------------------------------------
$desktopPath = [System.Environment]::GetFolderPath('Desktop')
if (-not (Test-Path $desktopPath)) {
    $desktopPath = Join-Path $env:OneDrive 'Desktop'
}
$shortcutPath = Join-Path $desktopPath 'MacroHub.lnk'

try {
    $shell    = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath       = 'powershell.exe'
    $shortcut.Arguments        = "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$HubScript`""
    $shortcut.WorkingDirectory = $HubRoot
    $shortcut.Description      = 'MacroHub v3.1 - Office Productivity + AI IDE'
    $shortcut.WindowStyle      = 1
    $icoPath = Join-Path $HubRoot 'MacroHub.ico'
    if (Test-Path $icoPath) { $shortcut.IconLocation = "$icoPath, 0" }
    else { $shortcut.IconLocation = 'powershell.exe, 0' }
    $shortcut.Save()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
    Write-Host "[OK] Desktop shortcut created: $shortcutPath" -ForegroundColor Green
} catch {
    Write-Host "[WARN] Could not create shortcut: $_" -ForegroundColor Yellow
}

# --- STEP 5: Unblock all files -----------------------------------------------
try {
    Get-ChildItem $HubRoot -Recurse | Unblock-File -ErrorAction SilentlyContinue
    Write-Host '[OK] Unblocked all MacroHub files.' -ForegroundColor Green
} catch {
    Write-Host "[WARN] Unblock step: $_" -ForegroundColor Yellow
}

# --- STEP 6: Chrome extension verification + TXT exports --------------------
$manifestFile = Join-Path $ChromeExtDir 'manifest.json'
if (Test-Path $manifestFile) {
    Write-Host ''
    Write-Host '[INFO] Chrome extension folder:' -ForegroundColor Cyan
    Write-Host "       $ChromeExtDir" -ForegroundColor Gray
    try {
        $null = Get-Content $manifestFile -Raw | ConvertFrom-Json -ErrorAction Stop
        Write-Host '[OK] manifest.json is valid JSON.' -ForegroundColor Green
    } catch {
        Write-Host "[WARN] manifest.json parse issue: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    $exports = Write-ChromeExtensionExports -RootPath $HubRoot -ExtDir $ChromeExtDir
    if ($exports.Missing.Count -gt 0) {
        Write-Host "[WARN] Missing extension files: $($exports.Missing -join ', ')" -ForegroundColor Yellow
    } else {
        Write-Host '[OK] All required extension files are present.' -ForegroundColor Green
    }
    Write-Host "[OK] Wrote extension setup guide: $($exports.GuidePath)" -ForegroundColor Green
    Write-Host "[OK] Wrote extension source export: $($exports.CodePath)" -ForegroundColor Green

    Write-Host ''
    Write-Host '  Chrome Developer Mode install:' -ForegroundColor White
    Write-Host '  1. Open chrome://extensions/' -ForegroundColor White
    Write-Host '  2. Enable Developer mode.' -ForegroundColor White
    Write-Host '  3. Click Load unpacked.' -ForegroundColor White
    Write-Host "  4. Select: $ChromeExtDir" -ForegroundColor White
    Write-Host ''
} else {
    Write-Host "[WARN] Chrome extension files not found in $ChromeExtDir." -ForegroundColor Yellow
    Write-Host '       Provide ChromeExtension_Code.txt (root or for export\\) so setup can rebuild the extension files automatically.' -ForegroundColor Yellow
}

# --- STEP 6b: AI IDE enterprise setup guide ----------------------------------
try {
    $aiGuide = Write-AiIdeEnterpriseGuide -RootPath $HubRoot
    Write-Host "[OK] Wrote AI IDE enterprise setup guide: $aiGuide" -ForegroundColor Green
} catch {
    Write-Host "[WARN] Could not write AI IDE setup guide: $_" -ForegroundColor Yellow
}

# --- STEP 6c: Core script TXT exports ---------------------------------------
try {
    $coreExports = Write-CoreScriptExports -RootPath $HubRoot -HubScriptPath $HubScript -SetupPath $SetupScriptPath
    Write-Host "[OK] Wrote core code export: $($coreExports.HubCodePath)" -ForegroundColor Green
    Write-Host "[OK] Wrote setup export: $($coreExports.SetupTextPath)" -ForegroundColor Green
} catch {
    Write-Host "[WARN] Could not write core script exports: $_" -ForegroundColor Yellow
}

# --- STEP 7: Check AI IDE port 9876 -----------------------------------------
try {
    $tcpConn = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()
    $usedPorts = $tcpConn.GetActiveTcpListeners() | Select-Object -ExpandProperty Port
    if ($usedPorts -contains 9876) {
        Write-Host '[WARN] Port 9876 is already in use. The AI IDE Chrome bridge may conflict.' -ForegroundColor Yellow
        Write-Host '       Change the port in MacroHub AI IDE -> Port box before starting the bridge.' -ForegroundColor Yellow
    } else {
        Write-Host '[OK] Port 9876 is available for the AI IDE Chrome bridge.' -ForegroundColor Green
    }
} catch {
    Write-Host "[INFO] Could not check port 9876: $_" -ForegroundColor Gray
}

# --- STEP 7b: Graph endpoint reachability check ------------------------------
Write-Host ''
Write-Host '[INFO] Checking Graph/Auth endpoint reachability (HTTPS 443)...' -ForegroundColor Cyan
try {
    try {
        $tls12 = [System.Net.SecurityProtocolType]::Tls12
        if (([System.Net.ServicePointManager]::SecurityProtocol -band $tls12) -eq 0) {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor $tls12
        }
    } catch {}
    foreach ($hostName in @('login.microsoftonline.com', 'graph.microsoft.com')) {
        try {
            $req = [System.Net.HttpWebRequest]::Create("https://$hostName/")
            $req.Method = 'HEAD'
            $req.Timeout = 5000
            $resp = $req.GetResponse()
            try { $resp.Close() } catch {}
            Write-Host "[OK] Reachable: $hostName" -ForegroundColor Green
        } catch {
            $httpCode = 0
            try { $httpCode = [int]$_.Exception.Response.StatusCode.value__ } catch {}
            if ($httpCode -eq 0) {
                $m = [string]$_.Exception.Message
                if ($m -match '\((\d{3})\)') { $httpCode = [int]$Matches[1] }
            }
            if ($httpCode -in @(401,403,405)) {
                Write-Host "[OK] Reachable: $hostName (HTTP $httpCode response)" -ForegroundColor Green
            } else {
                Write-Host "[WARN] Could not reach $hostName over HTTPS (443): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
} catch {
    Write-Host "[INFO] Endpoint check skipped: $($_.Exception.Message)" -ForegroundColor Gray
}

# --- STEP 8: Flash drive text-only package (F:\) -----------------------------
$flashDrive = 'F:\'
if (Test-Path $flashDrive) {
    Write-Host ''
    Write-Host "[INFO] Flash drive detected at $flashDrive. Writing text-only setup package..." -ForegroundColor Cyan
    try {
        $localExport = Join-Path $HubRoot 'for export'
        $txtFiles = @()
        if (Test-Path $localExport) {
            $txtFiles = @(Get-ChildItem -Path $localExport -File -Filter '*.txt' -ErrorAction SilentlyContinue)
        }
        if ($txtFiles.Count -eq 0) {
            throw "No .txt exports found in $localExport"
        }

        foreach ($folderName in @('MacroHub', 'MacroHub_Export')) {
            $target = Join-Path $flashDrive $folderName
            if (-not (Test-Path $target)) {
                New-Item -ItemType Directory -Path $target -Force | Out-Null
            }

            # Enforce text-only package in MacroHub folders on flash drive.
            Get-ChildItem -Path $target -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -ne '.txt' } |
                ForEach-Object { Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue }
            Get-ChildItem -Path $target -Recurse -Directory -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending |
                ForEach-Object {
                    if ((Get-ChildItem -Path $_.FullName -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
                        Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
                    }
                }

            foreach ($txt in $txtFiles) {
                Copy-Item -Path $txt.FullName -Destination (Join-Path $target $txt.Name) -Force -ErrorAction SilentlyContinue
            }
            Write-Host "[OK] Text files synced to $target" -ForegroundColor Green
        }
    } catch {
        Write-Host "[WARN] Flash drive text package sync failed: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "[INFO] Flash drive F:\ not detected. Skipping flash text package sync." -ForegroundColor Gray
}

# --- STEP 9: Syntax check MacroHub.ps1 --------------------------------------
Write-Host ''
Write-Host '[INFO] Running PowerShell syntax check on MacroHub.ps1...' -ForegroundColor Cyan
try {
    $src    = [System.IO.File]::ReadAllText($HubScript)
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$null, [ref]$errors)
    if ($errors.Count -eq 0) {
        Write-Host '[OK] MacroHub.ps1 syntax is clean - no parse errors.' -ForegroundColor Green
    } else {
        Write-Host "[WARN] $($errors.Count) parse error(s) found:" -ForegroundColor Yellow
        foreach ($e in $errors) {
            Write-Host "  Line $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "[WARN] Syntax check failed: $_" -ForegroundColor Yellow
}

# --- Done --------------------------------------------------------------------
Write-Host ''
Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host '  Setup complete!' -ForegroundColor Green
Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host '  Next steps:' -ForegroundColor White
Write-Host '  1. Double-click the MacroHub shortcut on your desktop.' -ForegroundColor White
Write-Host '  2. On your work computer: open AI IDE, set Tenant, click Auth + Verify, and sign in with your enterprise account.' -ForegroundColor White
Write-Host '  3. Load the ChromeExtension in Chrome Developer Mode (see above).' -ForegroundColor White
Write-Host '  4. Read for export\\ChromeExtension_Setup_Guide.txt for full extension setup.' -ForegroundColor White
Write-Host '  5. Read for export\\AI_IDE_Enterprise_Setup.txt for Graph licensing + troubleshooting.' -ForegroundColor White
Write-Host '  6. Add macro files (.bas or .ps1) to the Macros folder.' -ForegroundColor White
Write-Host ''

Read-Host 'Press Enter to close'
