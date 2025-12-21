param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$ur = $ws.UsedRange
$rows = [Math]::Max(1, [int]$ur.Rows.Count)
$cols = [Math]::Max(1, [int]$ur.Columns.Count)
$changed = 0
for ($r = 1; $r -le $rows; $r++) {
    for ($c = 1; $c -le $cols; $c++) {
        $cell = $ur.Cells.Item($r, $c)
        $val = $cell.Value2
        if ($val -is [double] -or $val -is [decimal] -or $val -is [int]) {
            $fmt = [string]$cell.NumberFormat
            $fmtLow = $fmt.ToLower()
            if ($fmtLow -notmatch '%' -and $fmtLow -notmatch 'y{2,4}|d{1,4}|m{1,4}|h{1,2}|s{1,2}') {
                $cell.NumberFormat = '#,##0.00'
                $changed++
            }
        }
    }
}
Write-Output ("Applied standard numeric format to {0} cells on '{1}'." -f $changed, $ws.Name)