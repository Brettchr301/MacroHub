param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$ur = $ws.UsedRange
$rows = [Math]::Max(1, [int]$ur.Rows.Count)
$cols = [Math]::Max(1, [int]$ur.Columns.Count)
$conv = 0
for ($r = 1; $r -le $rows; $r++) {
    for ($c = 1; $c -le $cols; $c++) {
        $cell = $ur.Cells.Item($r, $c)
        if ($cell.HasFormula) { continue }
        $v = $cell.Value2
        if ($v -is [string]) {
            $txt = $v.Trim()
            if (-not $txt) { continue }
            $dt = $null
            try { $dt = [datetime]::Parse($txt) } catch { $dt = $null }
            if ($dt) {
                $cell.Value2 = $dt
                $cell.NumberFormat = 'mm/dd/yyyy'
                $conv++
            }
        }
    }
}
Write-Output ("Normalized {0} date cells on '{1}'." -f $conv, $ws.Name)