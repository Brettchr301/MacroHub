# MacroHub Macro: DeduplicateRows
# Removes duplicate rows from the active sheet's used range, keeping the first occurrence.
# Optionally specify which columns to key on (1-based comma-separated string, e.g. "1,3")

param(
    [string]$TargetWorkbook = '',
    [string]$TargetSheet    = '',
    [string]$KeyColumns     = ''   # e.g. "1,2" — blank = all columns
)

$xl = $null
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { Write-Error "Excel is not open."; return }

$wb = if ($TargetWorkbook) {
    $xl.Workbooks | Where-Object { $_.Name -eq $TargetWorkbook } | Select-Object -First 1
} else { $xl.ActiveWorkbook }
if (-not $wb) { Write-Error "Workbook not found."; return }

$sh = if ($TargetSheet) {
    $wb.Sheets | Where-Object { $_.Name -eq $TargetSheet } | Select-Object -First 1
} else { $wb.ActiveSheet }
if (-not $sh) { Write-Error "Sheet not found."; return }

$ur = $sh.UsedRange
$rowCount = $ur.Rows.Count
$colCount  = $ur.Columns.Count
if ($rowCount -lt 2) { Write-Output "No data rows to deduplicate."; return }

# Build key column index array (0-based from used range start)
$keyCols = if ($KeyColumns) {
    $KeyColumns -split ',' | ForEach-Object { [int]$_.Trim() - 1 }
} else {
    0..($colCount - 1)
}

# Read all data into memory
$data = @()
for ($r = 1; $r -le $rowCount; $r++) {
    $row = @()
    for ($c = 1; $c -le $colCount; $c++) {
        $row += [string]$ur.Cells.Item($r, $c).Value2
    }
    $data += , $row
}

$seen    = [System.Collections.Generic.HashSet[string]]::new()
$toDelete = [System.Collections.Generic.List[int]]::new()

# Row 0 = header; start checking from row index 1
for ($r = 1; $r -lt $data.Count; $r++) {
    $key = ($keyCols | ForEach-Object { $data[$r][$_] }) -join '|'
    if (-not $seen.Add($key)) { $toDelete.Add($r) }
}

# Delete in reverse order to preserve indices
$toDelete = $toDelete | Sort-Object -Descending
foreach ($r in $toDelete) {
    # r is 0-based index into $data; row in sheet = ur.Row + r
    $sheetRow = $ur.Row + $r
    $sh.Rows.Item($sheetRow).Delete() | Out-Null
}

Write-Output "Removed $($toDelete.Count) duplicate row(s). $($rowCount - 1 - $toDelete.Count) unique data rows remain."
