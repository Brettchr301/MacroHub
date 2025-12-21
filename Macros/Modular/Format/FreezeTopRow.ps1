param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$ws.Activate() | Out-Null
$xl.ActiveWindow.SplitRow = 1
$xl.ActiveWindow.SplitColumn = 0
$xl.ActiveWindow.FreezePanes = $true
Write-Output ("Froze top row on '{0}'." -f $ws.Name)