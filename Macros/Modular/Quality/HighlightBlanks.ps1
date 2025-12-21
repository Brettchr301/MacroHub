param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$ur = $ws.UsedRange
$blank = $null
try { $blank = $ur.SpecialCells(4) } catch { $blank = $null }  # xlCellTypeBlanks
if ($blank) {
    $blank.Interior.Color = 65535
    Write-Output ("Highlighted blank cells on '{0}'." -f $ws.Name)
} else {
    Write-Output ("No blank cells found on '{0}'." -f $ws.Name)
}