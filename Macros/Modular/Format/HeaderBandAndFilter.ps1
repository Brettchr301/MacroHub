param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$ur = $ws.UsedRange
$lastCol = [Math]::Max(1, [int]$ur.Columns.Count)
$hdr = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $lastCol))
$hdr.Font.Bold = $true
$hdr.Interior.Color = 15773696
$hdr.WrapText = $false
if (-not $ws.AutoFilterMode) {
    [void]$hdr.AutoFilter()
}
Write-Output ("Styled header and applied filter on '{0}'." -f $ws.Name)