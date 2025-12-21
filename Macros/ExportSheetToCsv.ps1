# Exports the active sheet from the target workbook to a CSV in the same folder as the workbook.
# The CSV is named:  <WorkbookName>_<SheetName>_<YYYYMMDD>.csv
# Useful for sending data extracts without sharing the full Excel file.

param(
    [string]$WorkbookName = ''
)

$ErrorActionPreference = 'Stop'

# ---- Connect to Excel ----
try {
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    throw 'Excel is not open.'
}

$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
$ws = $wb.ActiveSheet

# ---- Build output path ----
$wbDir    = Split-Path $wb.FullName -Parent
$safeName = ($wb.Name -replace '\.xlsx?$', '') + '_' + $ws.Name + '_' + (Get-Date -Format 'yyyyMMdd')
$csvPath  = Join-Path $wbDir "$safeName.csv"

# ---- Read used range and write CSV ----
$used = $ws.UsedRange
$rows = $used.Rows.Count
$cols = $used.Columns.Count

$lines = @()
for ($r = 1; $r -le $rows; $r++) {
    $cells = @()
    for ($c = 1; $c -le $cols; $c++) {
        $val = $used.Cells.Item($r, $c).Text
        # Wrap in quotes if the value contains a comma or quote
        if ($val -match '[,"]') {
            $val = '"' + ($val -replace '"', '""') + '"'
        }
        $cells += $val
    }
    $lines += ($cells -join ',')
}

$lines | Set-Content $csvPath -Encoding UTF8

Write-Output "Exported $rows row(s) to: $csvPath"

# Friendly popup via Windows shell (no Excel dialog dependency)
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show(
    "Export complete!`n`n$csvPath",
    "MacroHub - Export", 'OK', 'Information') | Out-Null
