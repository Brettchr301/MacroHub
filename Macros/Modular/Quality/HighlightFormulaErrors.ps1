param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$ur = $ws.UsedRange
$err = $null
try { $err = $ur.SpecialCells(-4123, 16) } catch { $err = $null }  # formulas with errors
if ($err) {
    $err.Interior.Color = 255
    Write-Output ("Highlighted formula errors on '{0}'." -f $ws.Name)
} else {
    Write-Output ("No formula errors found on '{0}'." -f $ws.Name)
}