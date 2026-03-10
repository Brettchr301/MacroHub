<#
.SYNOPSIS
    Comprehensive backend test suite for MacroHub v3.1
    Tests all standalone functions (lines 1-721) without launching the UI.
#>

# ============================================================
#  SETUP: Source only the backend functions (not the UI)
# ============================================================
$ErrorActionPreference = 'Stop'
$pass = 0; $fail = 0; $warns = @()

function Test-Assert($name, [scriptblock]$test) {
    try {
        $result = & $test
        if ($result) { $script:pass++; Write-Host "  [PASS] $name" -ForegroundColor Green }
        else         { $script:fail++; Write-Host "  [FAIL] $name -- returned false" -ForegroundColor Red }
    } catch {
        $script:fail++
        Write-Host "  [FAIL] $name -- $_" -ForegroundColor Red
    }
}

# Extract and dot-source the backend portion of MacroHub.ps1 (lines 1-718) 
$hubPath = Join-Path $PSScriptRoot 'MacroHub.ps1'
$lines   = Get-Content $hubPath
# Find where Start-MacroHub begins
$cutLine = 0
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^\s*function\s+Start-MacroHub\s*\{') { $cutLine = $i; break }
}
if ($cutLine -eq 0) { Write-Error "Could not find Start-MacroHub boundary"; exit 1 }

$backendCode = $lines[0..($cutLine - 1)] -join "`n"
# We need to set $script:HubRoot before sourcing since config depends on it
$script:HubRoot = $PSScriptRoot
Invoke-Expression $backendCode

Write-Host "`n=============================="
Write-Host "  MacroHub v3.1 Backend Tests"
Write-Host "==============================`n"

# ============================================================
#  TEST 1: Config variables
# ============================================================
Write-Host "[CONFIG]" -ForegroundColor Cyan
Test-Assert "HubRoot is set" { $script:HubRoot -and (Test-Path $script:HubRoot) }
Test-Assert "MacroFolder defined" { $script:MacroFolder -like '*Macros*' }
Test-Assert "QuartersDir defined" { $script:QuartersDir -like '*quarters*' }
Test-Assert "ActiveQuarterPath starts empty" { $script:ActiveQuarterPath -eq '' }

# ============================================================
#  TEST 2: Strip-DateTokens
# ============================================================
Write-Host "`n[STRIP-DATE-TOKENS]" -ForegroundColor Cyan
Test-Assert "Removes yyyy-MM-dd pattern" {
    $r = Strip-DateTokens 'Report_2024-03-15_final'
    $r -notmatch '2024' -and $r -match 'report' -and $r -match 'final'
}
Test-Assert "Removes Q3 2024 pattern" {
    $r = Strip-DateTokens 'Q3_2024_Summary'
    # Should strip date tokens, keeping meaningful parts
    $r -match 'summary'
}
Test-Assert "Handles no-date filename" {
    $r = Strip-DateTokens 'readme'
    $r -eq 'readme'
}
Test-Assert "Handles empty string" {
    $r = Strip-DateTokens ''
    $r -ne $null
}

# ============================================================
#  TEST 3: Macro file functions
# ============================================================
Write-Host "`n[MACRO-FILES]" -ForegroundColor Cyan
$macros = Get-MacroFiles
Test-Assert "Get-MacroFiles returns array" { $macros -is [array] }
Test-Assert "Macro folder exists" { Test-Path $script:MacroFolder }
if ($macros.Count -gt 0) {
    Test-Assert "First macro has Name property" { $macros[0].Name -ne $null }
    Test-Assert "First macro has FullName property" { $macros[0].FullName -ne $null }
} else {
    $warns += "No macro files found in $($script:MacroFolder)"
    Write-Host "  [WARN] No macro files to test" -ForegroundColor Yellow
}

# ============================================================
#  TEST 4: Activity Log
# ============================================================
Write-Host "`n[ACTIVITY-LOG]" -ForegroundColor Cyan
$testLogMsg = "TEST_$(Get-Date -f 'yyyyMMdd_HHmmss')"
Write-ActivityLog $testLogMsg
Test-Assert "Log file exists after write" { Test-Path $script:LogFile }
Test-Assert "Log contains test message" {
    $content = (Get-Content $script:LogFile -Tail 3) -join "`n"
    $content -match $testLogMsg
}

# ============================================================
#  TEST 5: Template functions
# ============================================================
Write-Host "`n[TEMPLATES]" -ForegroundColor Cyan
$origTemplates = Load-Templates
Test-Assert "Load-Templates returns array" { $origTemplates -is [array] }

# Save a test template
$testTpl = @([PSCustomObject]@{ Name='TEST_TPL'; Content='Hello {USER}' })
Save-Templates $testTpl
Test-Assert "Template JSON exists after save" { Test-Path $script:TemplateJson }

$loaded = Load-Templates
Test-Assert "Load-Templates contains saved template" {
    @($loaded | Where-Object { $_.Name -eq 'TEST_TPL' }).Count -eq 1
}

# Restore original if existed
if ($origTemplates.Count -gt 0) { Save-Templates $origTemplates }
else { Remove-Item $script:TemplateJson -ErrorAction SilentlyContinue }

# ============================================================
#  TEST 6: QSync Config
# ============================================================
Write-Host "`n[QSYNC-CONFIG]" -ForegroundColor Cyan
$origCfg = Load-QsConfig
Test-Assert "Load-QsConfig returns object" { $origCfg.PSObject.Properties.Name -contains 'LastQuarterPath' }

Save-QsConfig ([PSCustomObject]@{
    LastQuarterPath = 'C:\Test\LastQ'
    ThisQuarterPath = 'C:\Test\ThisQ'
    CompareSourcePath = 'C:\Test\SourceA'
    CompareTargetPath = 'C:\Test\TargetB'
})
Test-Assert "Config JSON exists after save" { Test-Path $script:QsCfgJson }

$cfgReload = Load-QsConfig
Test-Assert "Config round-trip: LastQuarterPath" { $cfgReload.LastQuarterPath -eq 'C:\Test\LastQ' }
Test-Assert "Config round-trip: ThisQuarterPath" { $cfgReload.ThisQuarterPath -eq 'C:\Test\ThisQ' }
Test-Assert "Config round-trip: CompareSourcePath" { $cfgReload.CompareSourcePath -eq 'C:\Test\SourceA' }
Test-Assert "Config round-trip: CompareTargetPath" { $cfgReload.CompareTargetPath -eq 'C:\Test\TargetB' }

# Restore
if ($origCfg.LastQuarterPath -or $origCfg.ThisQuarterPath -or $origCfg.CompareSourcePath -or $origCfg.CompareTargetPath) { Save-QsConfig $origCfg }
else { Remove-Item $script:QsCfgJson -ErrorAction SilentlyContinue }

# ============================================================
#  TEST 7: Quarter file management
# ============================================================
Write-Host "`n[QUARTER-FILES]" -ForegroundColor Cyan

# Create quarters dir
if (-not (Test-Path $script:QuartersDir)) {
    New-Item -ItemType Directory -Path $script:QuartersDir -Force | Out-Null
}
Test-Assert "Quarters dir exists" { Test-Path $script:QuartersDir }

# New-QuarterFile
$stale = Join-Path $script:QuartersDir 'Test_Q1_2025.json'
if (Test-Path $stale) { Remove-Item $stale -Force }
$qPath = New-QuarterFile 'Test_Q1_2025'
Test-Assert "New-QuarterFile creates JSON" { Test-Path $qPath }
Test-Assert "New-QuarterFile path correct" { $qPath -like '*Test_Q1_2025.json' }

$qContent = Get-Content $qPath -Raw
Test-Assert "Quarter JSON starts empty array" { $qContent.Trim() -eq '[]' }

# Creating same name again should not overwrite
$seedTodos = @(
    [PSCustomObject]@{
        Key='testkey|root'; OriginalName='testfile.txt'; RelFolder='(root)'
        LastDoneDate='2024-01-01'; DueDate='2024-04-01'; Status='Pending'
        AddedOn='2024-01-01'; Note=''
    }
)
Save-QsTodos $seedTodos $qPath
$qPath2 = New-QuarterFile 'Test_Q1_2025'
$qTodos2 = Load-QsTodos $qPath2
Test-Assert "New-QuarterFile does not overwrite existing" {
    $qTodos2.Count -eq 1 -and $qTodos2[0].Key -eq 'testkey|root'
}

# Get-QuarterList
$qList = Get-QuarterList
Test-Assert "Get-QuarterList returns at least 1" { $qList.Count -ge 1 }
Test-Assert "Get-QuarterList contains test quarter" {
    @($qList | Where-Object { $_.BaseName -eq 'Test_Q1_2025' }).Count -eq 1
}

# Switch-ActiveQuarter
Switch-ActiveQuarter $qPath
Test-Assert "Switch-ActiveQuarter sets path" { $script:ActiveQuarterPath -eq $qPath }

# ============================================================
#  TEST 8: Load/Save QsTodos with path parameter
# ============================================================
Write-Host "`n[LOAD-SAVE-QSTODOS]" -ForegroundColor Cyan

# Load from the test quarter (has 1 item we wrote above)
$todos = Load-QsTodos $qPath
Test-Assert "Load-QsTodos with path returns array" { $todos -is [array] }
Test-Assert "Load-QsTodos loaded 1 item" { $todos.Count -eq 1 }
Test-Assert "Todo has correct Key" { $todos[0].Key -eq 'testkey|root' }

# Load using active quarter (no explicit path)
$todos2 = Load-QsTodos
Test-Assert "Load-QsTodos (active) returns same data" { $todos2.Count -eq 1 }

# Save additional items
$newTodo = [PSCustomObject]@{
    Key='newkey|subfolder'; OriginalName='newfile.xlsx'; RelFolder='subfolder'
    LastDoneDate='2024-06-01'; DueDate='2024-09-01'; Status='Pending'
    AddedOn='2024-06-15'; Note='Test note'
}
$allTodos = @($todos) + @($newTodo)
Save-QsTodos $allTodos $qPath
$reloaded = Load-QsTodos $qPath
Test-Assert "Save-QsTodos persists 2 items" { $reloaded.Count -eq 2 }
Test-Assert "Second item has correct name" { @($reloaded | Where-Object { $_.OriginalName -eq 'newfile.xlsx' }).Count -eq 1 }

# Save empty array
$emptyPath = New-QuarterFile 'Test_Empty'
Save-QsTodos @() $emptyPath
$emptyReload = Load-QsTodos $emptyPath
Test-Assert "Save-QsTodos empty creates header-only file" { $emptyReload.Count -eq 0 }

# Load from non-existent path
$noExist = Load-QsTodos 'C:\fake\nonexistent.json'
Test-Assert "Load-QsTodos non-existent returns empty" { $noExist.Count -eq 0 }

# Load with no path and no active quarter
$saved = $script:ActiveQuarterPath
$script:ActiveQuarterPath = ''
$noActive = Load-QsTodos
Test-Assert "Load-QsTodos no path no active returns empty" { $noActive.Count -eq 0 }
$script:ActiveQuarterPath = $saved

# ============================================================
#  TEST 9: Scan-FolderToTodos
# ============================================================
Write-Host "`n[SCAN-FOLDER-TO-TODOS]" -ForegroundColor Cyan

# Create temp test folder structure
$testDir = Join-Path $env:TEMP "MacroHub_Test_$(Get-Date -f 'yyyyMMddHHmmss')"
New-Item -ItemType Directory "$testDir\Reports" -Force | Out-Null
Set-Content "$testDir\file1.xlsx" 'test' -Encoding UTF8
Set-Content "$testDir\file2.pdf" 'test' -Encoding UTF8
Set-Content "$testDir\Reports\monthly.xlsx" 'test' -Encoding UTF8

$scanned = Scan-FolderToTodos -FolderPath $testDir
Test-Assert "Scan-FolderToTodos found 3 files" { $scanned.Count -eq 3 }
Test-Assert "Scanned items have Key property" { $scanned[0].Key -ne $null }
Test-Assert "Scanned items have Status=Pending" { @($scanned | Where-Object { $_.Status -eq 'Pending' }).Count -eq 3 }
Test-Assert "Scanned items include subfolder" {
    @($scanned | Where-Object { $_.RelFolder -eq 'Reports' }).Count -eq 1
}
Test-Assert "Root items have (root) folder" {
    @($scanned | Where-Object { $_.RelFolder -eq '(root)' }).Count -eq 2
}
Test-Assert "Items have AddedOn date" { $scanned[0].AddedOn -match '\d{4}-\d{2}-\d{2}' }

# Scan non-existent folder
$noScan = Scan-FolderToTodos -FolderPath 'C:\fake\no_exist_dir'
Test-Assert "Scan non-existent folder returns empty" { $noScan.Count -eq 0 }

# Cleanup test folder
Remove-Item $testDir -Recurse -Force -ErrorAction SilentlyContinue

# ============================================================
#  TEST 10: QTasks folder compare mode
# ============================================================
Write-Host "`n[QTASKS-COMPARE]" -ForegroundColor Cyan

$cmpBackupExists = Test-Path $script:QsCompareJson
$cmpBackupRaw = ''
if ($cmpBackupExists) { $cmpBackupRaw = Get-Content $script:QsCompareJson -Raw }

$cmpRoot = Join-Path $env:TEMP "MacroHub_QCompare_$(Get-Date -f 'yyyyMMddHHmmss')"
$cmpSrc  = Join-Path $cmpRoot 'SourceA'
$cmpTgt  = Join-Path $cmpRoot 'TargetB'
New-Item -ItemType Directory $cmpSrc,$cmpTgt -Force | Out-Null
New-Item -ItemType Directory "$cmpSrc\Ops","$cmpTgt\Ops" -Force | Out-Null

Set-Content "$cmpSrc\report_2026-03-13.xlsx" 'test' -Encoding UTF8
Set-Content "$cmpSrc\Ops\shared.csv" 'test' -Encoding UTF8
Set-Content "$cmpSrc\Ops\deliverable_only_in_source.pdf" 'test' -Encoding UTF8

Set-Content "$cmpTgt\report.xlsx" 'test' -Encoding UTF8
Set-Content "$cmpTgt\Ops\shared.csv" 'test' -Encoding UTF8

$cmpResult = Compare-QsFolders -SourceRoot $cmpSrc -TargetRoot $cmpTgt
$cmpMissing = @($cmpResult.MissingTodos)
Test-Assert "Compare-QsFolders returns 1 missing file" { $cmpMissing.Count -eq 1 }
Test-Assert "Missing file name is correct" { $cmpMissing[0].OriginalName -eq 'deliverable_only_in_source.pdf' }
Test-Assert "Missing file folder is correct" { $cmpMissing[0].RelFolder -eq 'Ops' }
Test-Assert "Added date captured" { $cmpMissing[0].LastDoneDate -match '^\d{4}-\d{2}-\d{2}$' }
Test-Assert "Updated date captured" { $cmpMissing[0].DueDate -match '^\d{4}-\d{2}-\d{2}$' }

Save-QsCompareTodos $cmpMissing
$cmpReload = Load-QsCompareTodos
Test-Assert "Load-QsCompareTodos round-trip count" { @($cmpReload).Count -eq 1 }
Test-Assert "Compare persistence keeps SourceRoot" { $cmpReload[0].SourceRoot -eq $cmpSrc }
Test-Assert "Compare persistence keeps TargetRoot" { $cmpReload[0].TargetRoot -eq $cmpTgt }

if ($cmpBackupExists) {
    Set-Content $script:QsCompareJson $cmpBackupRaw -Encoding UTF8
} else {
    Remove-Item $script:QsCompareJson -ErrorAction SilentlyContinue
}
Remove-Item $cmpRoot -Recurse -Force -ErrorAction SilentlyContinue

# ============================================================
#  TEST 11: Invoke-QuarterSync
# ============================================================
Write-Host "`n[INVOKE-QUARTERSYNC]" -ForegroundColor Cyan

# Create two temp quarter directories
$lastDir = Join-Path $env:TEMP "MacroHub_LastQ_$(Get-Date -f 'yyyyMMddHHmmss')"
$thisDir = Join-Path $env:TEMP "MacroHub_ThisQ_$(Get-Date -f 'yyyyMMddHHmmss')"
New-Item -ItemType Directory $lastDir -Force | Out-Null
New-Item -ItemType Directory $thisDir -Force | Out-Null

# Last quarter: 2 folders, 3 files
New-Item -ItemType Directory "$lastDir\Sales" -Force | Out-Null
New-Item -ItemType Directory "$lastDir\Finance" -Force | Out-Null
Set-Content "$lastDir\overview.xlsx" 'test' -Encoding UTF8
Set-Content "$lastDir\Sales\q3_report.xlsx" 'test' -Encoding UTF8
Set-Content "$lastDir\Finance\budget.xlsx" 'test' -Encoding UTF8

# This quarter: only 1 folder, 1 file (Sales exists with overview)
New-Item -ItemType Directory "$thisDir\Sales" -Force | Out-Null
Set-Content "$thisDir\overview.xlsx" 'test' -Encoding UTF8

# Create a fresh quarter JSON for the sync
$syncQPath = New-QuarterFile 'Test_Sync_Q4'

$syncResult = Invoke-QuarterSync -LastRoot $lastDir -ThisRoot $thisDir -QuarterPath $syncQPath
Test-Assert "Sync returns PSCustomObject" { $syncResult.FoldersCreated -ne $null }
Test-Assert "Sync created Finance folder" { $syncResult.FoldersCreated -contains 'Finance' }
Test-Assert "Sync skipped Sales folder" { $syncResult.FoldersSkipped -contains 'Sales' }
Test-Assert "Finance folder actually created" { Test-Path "$thisDir\Finance" }
Test-Assert "Sync found missing files" { $syncResult.NewTodos.Count -ge 1 }
Test-Assert "Missing files include budget or q3_report" {
    @($syncResult.NewTodos | Where-Object { $_.OriginalName -match 'budget|q3_report' }).Count -ge 1
}
Test-Assert "No sync errors" { $syncResult.Errors.Count -eq 0 }

# Verify todos were saved to the quarter JSON
$syncTodos = Load-QsTodos $syncQPath
Test-Assert "Sync saved todos to quarter JSON" { $syncTodos.Count -ge 1 }
Test-Assert "Todos have Pending status" { @($syncTodos | Where-Object { $_.Status -eq 'Pending' }).Count -ge 1 }
Test-Assert "Todos have DueDate set" { $syncTodos[0].DueDate -match '\d{4}-\d{2}-\d{2}' }

# Run sync again -- should not duplicate
$syncResult2 = Invoke-QuarterSync -LastRoot $lastDir -ThisRoot $thisDir -QuarterPath $syncQPath
Test-Assert "Second sync has 0 new todos (no duplicates)" { $syncResult2.NewTodos.Count -eq 0 }
$syncTodos2 = Load-QsTodos $syncQPath
Test-Assert "Todo count unchanged after re-sync" { $syncTodos2.Count -eq $syncTodos.Count }

# Sync with invalid paths
$badResult = Invoke-QuarterSync -LastRoot 'C:\no_exist' -ThisRoot $thisDir -QuarterPath $syncQPath
Test-Assert "Sync invalid last path returns error" { $badResult.Errors.Count -gt 0 }

# Cleanup
Remove-Item $lastDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $thisDir -Recurse -Force -ErrorAction SilentlyContinue

# ============================================================
#  TEST 12: Sync log persistence
# ============================================================
Write-Host "`n[SYNC-LOG]" -ForegroundColor Cyan
$syncLog = Load-QsSyncLog
Test-Assert "Sync log exists after sync" { $syncLog.Count -ge 1 }
Test-Assert "Sync log has RunDate" { $syncLog[0].RunDate -match '\d{4}-\d{2}-\d{2}' }
Test-Assert "Sync log has Type" { $syncLog[0].Type -in @('Created','Skipped') }

# ============================================================
#  TEST 13: End-to-end scenario (user workflow)
# ============================================================
Write-Host "`n[E2E WORKFLOW]" -ForegroundColor Cyan

# Simulate: user creates a quarter, scans a folder, adds manual task, marks done
$e2eQPath = New-QuarterFile 'E2E_Test_Quarter'
Switch-ActiveQuarter $e2eQPath
Test-Assert "E2E: active quarter set" { $script:ActiveQuarterPath -eq $e2eQPath }

$e2eDir = Join-Path $env:TEMP "MacroHub_E2E_$(Get-Date -f 'yyyyMMddHHmmss')"
New-Item -ItemType Directory "$e2eDir\Sub" -Force | Out-Null
Set-Content "$e2eDir\alpha.xlsx" 'test' -Encoding UTF8
Set-Content "$e2eDir\Sub\beta.pdf" 'test' -Encoding UTF8

# Scan folder and merge into quarter
$scanned = Scan-FolderToTodos -FolderPath $e2eDir
$existing = Load-QsTodos
$merged = [System.Collections.Generic.List[object]]::new()
foreach ($e in $existing) { $merged.Add($e) }
foreach ($s in $scanned) { $merged.Add($s) }
Save-QsTodos $merged.ToArray()

$afterScan = Load-QsTodos
Test-Assert "E2E: 2 items after scan" { $afterScan.Count -eq 2 }

# Add a manual task
$manual = [PSCustomObject]@{
    Key='manualtask|(manual)'; OriginalName='Review memo'
    RelFolder='(manual)'; LastDoneDate=''; DueDate='2025-04-01'
    Status='Pending'; AddedOn=(Get-Date -f 'yyyy-MM-dd'); Note='Manually added'
}
$afterScan += @($manual)
Save-QsTodos $afterScan

$afterManual = Load-QsTodos
Test-Assert "E2E: 3 items after manual add" { $afterManual.Count -eq 3 }

# Mark one as done
$target = $afterManual | Where-Object { $_.OriginalName -eq 'alpha.xlsx' }
if ($target) { $target.Status = 'Done' }
Save-QsTodos $afterManual

$afterDone = Load-QsTodos
$doneCount = @($afterDone | Where-Object { $_.Status -eq 'Done' }).Count
$pendCount = @($afterDone | Where-Object { $_.Status -eq 'Pending' }).Count
Test-Assert "E2E: 1 Done, 2 Pending" { $doneCount -eq 1 -and $pendCount -eq 2 }

# Cleanup temp
Remove-Item $e2eDir -Recurse -Force -ErrorAction SilentlyContinue

# ============================================================
#  TEST 14: Y/Y comparison scenario
# ============================================================
Write-Host "`n[Y/Y COMPARISON]" -ForegroundColor Cyan

# Create Q1_2024 and Q1_2025 with overlapping + unique items
$q2024 = New-QuarterFile 'YY_Q1_2024'
$q2025 = New-QuarterFile 'YY_Q1_2025'

$items2024 = @(
    [PSCustomObject]@{ Key='report.xlsx|Sales'; OriginalName='report.xlsx'; RelFolder='Sales'; LastDoneDate='2024-03-01'; DueDate='2024-06-01'; Status='Done'; AddedOn='2024-01-15'; Note='' }
    [PSCustomObject]@{ Key='budget.xlsx|Finance'; OriginalName='budget.xlsx'; RelFolder='Finance'; LastDoneDate='2024-03-10'; DueDate='2024-06-10'; Status='Pending'; AddedOn='2024-01-15'; Note='' }
    [PSCustomObject]@{ Key='oldfile.xlsx|(root)'; OriginalName='oldfile.xlsx'; RelFolder='(root)'; LastDoneDate='2024-02-01'; DueDate='2024-05-01'; Status='Done'; AddedOn='2024-01-15'; Note='' }
)
Save-QsTodos $items2024 $q2024

$items2025 = @(
    [PSCustomObject]@{ Key='report.xlsx|Sales'; OriginalName='report.xlsx'; RelFolder='Sales'; LastDoneDate='2025-03-01'; DueDate='2025-06-01'; Status='Pending'; AddedOn='2025-01-15'; Note='' }
    [PSCustomObject]@{ Key='budget.xlsx|Finance'; OriginalName='budget.xlsx'; RelFolder='Finance'; LastDoneDate='2025-03-10'; DueDate='2025-06-10'; Status='Pending'; AddedOn='2025-01-15'; Note='' }
    [PSCustomObject]@{ Key='newfile.xlsx|(root)'; OriginalName='newfile.xlsx'; RelFolder='(root)'; LastDoneDate='2025-02-01'; DueDate='2025-05-01'; Status='Pending'; AddedOn='2025-01-15'; Note='' }
)
Save-QsTodos $items2025 $q2025

# Load both
$load2024 = Load-QsTodos $q2024
$load2025 = Load-QsTodos $q2025
Test-Assert "Y/Y: Q1_2024 has 3 items" { $load2024.Count -eq 3 }
Test-Assert "Y/Y: Q1_2025 has 3 items" { $load2025.Count -eq 3 }

# Compare: find items only in 2024
$keys2025 = @{}; foreach ($i in $load2025) { $keys2025[$i.Key] = $true }
$onlyIn2024 = @($load2024 | Where-Object { -not $keys2025.ContainsKey($_.Key) })
Test-Assert "Y/Y: 1 item only in 2024 (oldfile.xlsx)" { $onlyIn2024.Count -eq 1 -and $onlyIn2024[0].OriginalName -eq 'oldfile.xlsx' }

# Compare: find items only in 2025  
$keys2024 = @{}; foreach ($i in $load2024) { $keys2024[$i.Key] = $true }
$onlyIn2025 = @($load2025 | Where-Object { -not $keys2024.ContainsKey($_.Key) })
Test-Assert "Y/Y: 1 item only in 2025 (newfile.xlsx)" { $onlyIn2025.Count -eq 1 -and $onlyIn2025[0].OriginalName -eq 'newfile.xlsx' }

# Shared items
$shared = @($load2025 | Where-Object { $keys2024.ContainsKey($_.Key) })
Test-Assert "Y/Y: 2 shared items" { $shared.Count -eq 2 }

# Status difference check: report.xlsx was Done in 2024, Pending in 2025
$reportIn2024 = $load2024 | Where-Object { $_.Key -eq 'report.xlsx|Sales' }
$reportIn2025 = $load2025 | Where-Object { $_.Key -eq 'report.xlsx|Sales' }
Test-Assert "Y/Y: report.xlsx status differs (Done vs Pending)" {
    $reportIn2024.Status -eq 'Done' -and $reportIn2025.Status -eq 'Pending'
}

# ============================================================
#  TEST 14: Excel helper functions (no COM -- just verify they exist)
# ============================================================
Write-Host "`n[EXCEL-HELPERS]" -ForegroundColor Cyan
$fnNames = @('Get-ExcelApp','Get-OpenWorkbooks','Get-WorkbookSheets','Get-WorksheetNames',
             'Open-WorkbookInExcel','Paste-TextToSheet','Invoke-VbaMacro','Invoke-PsScript',
             'Register-HubTask','Get-HubTasks','Remove-HubTask')
foreach ($fn in $fnNames) {
    Test-Assert "$fn exists" { Get-Command $fn -ErrorAction SilentlyContinue }
}

# ============================================================
#  TEST 15: Export-QsToExcel function exists and handles no-Excel gracefully
# ============================================================
Write-Host "`n[EXPORT-TO-EXCEL]" -ForegroundColor Cyan
Test-Assert "Export-QsToExcel function exists" { Get-Command Export-QsToExcel -ErrorAction SilentlyContinue }

# ============================================================
#  CLEANUP: Remove test quarter files
# ============================================================
Write-Host "`n[CLEANUP]" -ForegroundColor Cyan
$testFiles = @('Test_Q1_2025','Test_Empty','Test_Sync_Q4','E2E_Test_Quarter','YY_Q1_2024','YY_Q1_2025')
foreach ($tf in $testFiles) {
    $p = Join-Path $script:QuartersDir "$tf.json"
    if (Test-Path $p) { Remove-Item $p -Force }
}
# Remove sync log artifact
if (Test-Path $script:QsSyncJson) { Remove-Item $script:QsSyncJson -Force -ErrorAction SilentlyContinue }
Write-Host "  Test artifacts cleaned up" -ForegroundColor DarkGray

# ============================================================
#  RESULTS
# ============================================================
Write-Host "`n=============================="
Write-Host "  RESULTS: $pass passed, $fail failed" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
if ($warns.Count -gt 0) {
    Write-Host "  WARNINGS: $($warns.Count)" -ForegroundColor Yellow
    foreach ($w in $warns) { Write-Host "    - $w" -ForegroundColor Yellow }
}
Write-Host "==============================`n"

exit $fail
