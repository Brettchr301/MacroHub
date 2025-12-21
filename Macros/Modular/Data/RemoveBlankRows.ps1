param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$ur = $ws.UsedRange
$firstRow = [int]$ur.Row
$firstCol = [int]$ur.Column
$lastRow = $firstRow + [int]$ur.Rows.Count - 1
$lastCol = $firstCol + [int]$ur.Columns.Count - 1
$deleted = 0
for ($ri = $lastRow; $ri -ge $firstRow; $ri--) {
    $rowRange = $ws.Range($ws.Cells.Item($ri, $firstCol), $ws.Cells.Item($ri, $lastCol))
    $cnt = [int]$xl.WorksheetFunction.CountA($rowRange)
    if ($cnt -eq 0) {
        [void]$rowRange.EntireRow.Delete()
        $deleted++
    }
}
Write-Output ("Deleted {0} blank rows on '{1}'." -f $deleted, $ws.Name)