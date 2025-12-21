param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
[void]$wb.Activate()
[void]$xl.CalculateFullRebuild()
Write-Output ("Recalculated workbook '{0}'." -f $wb.Name)