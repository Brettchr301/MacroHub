# MacroHub Macro: CreatePivotTable
# Creates a PivotTable from the active sheet's used range on a new sheet named "Pivot"
# Requires: Active Excel workbook with data in Sheet1 (or active sheet)

param(
    [string]$TargetWorkbook = '',
    [string]$SourceSheet    = '',
    [string]$RowField       = '',   # Column header to use as row labels (blank = first column)
    [string]$ValueField     = '',   # Column header to sum (blank = second column)
    [string]$PivotSheetName = 'Pivot'
)

$xl = $null
try {
    $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    Write-Error "Excel is not open."; return
}

$wb = if ($TargetWorkbook) {
    try { $xl.Workbooks | Where-Object { $_.Name -eq $TargetWorkbook } | Select-Object -First 1 }
    catch { $null }
} else { $xl.ActiveWorkbook }

if (-not $wb) { Write-Error "Could not find workbook."; return }

$srcSh = if ($SourceSheet) {
    try { $wb.Sheets | Where-Object { $_.Name -eq $SourceSheet } | Select-Object -First 1 } catch { $null }
} else { $wb.ActiveSheet }

if (-not $srcSh) { Write-Error "Source sheet not found."; return }

$usedRange = $srcSh.UsedRange
if ($usedRange.Rows.Count -lt 2) { Write-Error "No data found on source sheet."; return }

# Determine row and value fields from headers if not supplied
$headers = @()
for ($c = 1; $c -le $usedRange.Columns.Count; $c++) {
    $headers += [string]$usedRange.Cells.Item(1, $c).Value2
}
if (-not $RowField)   { $RowField   = $headers[0] }
if (-not $ValueField) { $ValueField = if ($headers.Count -ge 2) { $headers[1] } else { $headers[0] } }

# Remove existing pivot sheet if present
$existing = $wb.Sheets | Where-Object { $_.Name -eq $PivotSheetName }
if ($existing) { $existing.Delete() }

# Add pivot sheet at end
$pvSh = $wb.Sheets.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
$pvSh.Name = $PivotSheetName

# Build the pivot cache source address
$srcAddr = "'$($srcSh.Name)'!" + $usedRange.Address($true, $true, 1)   # xlA1

$pc = $wb.PivotCaches().Create(
    1,          # xlDatabase
    $srcAddr,
    6           # xlPivotTableVersion15
)

$pt = $pc.CreatePivotTable(
    $pvSh.Range("A3"),
    'MacroHubPivot'
)

# Add row field
$rf = $pt.PivotFields($RowField)
$rf.Orientation = 1   # xlRowField
$rf.Position     = 1

# Add value field
$vf = $pt.PivotFields($ValueField)
$vf.Orientation = 4   # xlDataField
try { $vf.Function = -4157 }  catch {}  # xlSum; ignore if non-numeric

$pvSh.Columns.AutoFit()
$pvSh.Activate()

Write-Output "PivotTable created on sheet '$PivotSheetName' (Rows: $RowField, Values: $ValueField)"
