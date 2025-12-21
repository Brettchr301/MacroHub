param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$folder = if ($wb.Path) { $wb.Path } else { [Environment]::GetFolderPath('MyDocuments') }
$safeSheet = ($ws.Name -replace '[\\/:*?""<>|]', '_')
$file = '{0}_{1}.csv' -f $safeSheet, (Get-Date -Format 'yyyyMMdd_HHmmss')
$dest = Join-Path $folder $file
$prevAlerts = $xl.DisplayAlerts
try {
    $xl.DisplayAlerts = $false
    [void]$ws.Copy()
    $tmp = $xl.ActiveWorkbook
    [void]$tmp.SaveAs($dest, 6)  # xlCSV
    [void]$tmp.Close($false)
} finally {
    $xl.DisplayAlerts = $prevAlerts
}
Write-Output ("Exported '{0}' to CSV: {1}" -f $ws.Name, $dest)