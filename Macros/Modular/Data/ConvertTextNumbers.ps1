param([string]$WorkbookName, [string]$SheetName)
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = if ($WorkbookName) { $xl.Workbooks.Item($WorkbookName) } else { $xl.ActiveWorkbook }
if (-not $wb) { throw 'No workbook is open.' }
$ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $wb.ActiveSheet }
if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
$ur = $ws.UsedRange
$rows = [Math]::Max(1, [int]$ur.Rows.Count)
$cols = [Math]::Max(1, [int]$ur.Columns.Count)
$styles = [System.Globalization.NumberStyles]::Float -bor [System.Globalization.NumberStyles]::AllowThousands
$conv = 0
for ($r = 1; $r -le $rows; $r++) {
    for ($c = 1; $c -le $cols; $c++) {
        $cell = $ur.Cells.Item($r, $c)
        if ($cell.HasFormula) { continue }
        $v = $cell.Value2
        if ($v -is [string]) {
            $txt = $v.Trim()
            if ($txt -match '%' -or $txt -match '^\d{1,4}[/-]\d{1,2}[/-]\d{1,4}$') { continue }
            $n = 0.0
            if ([double]::TryParse($txt, $styles, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$n) -or
                [double]::TryParse($txt, $styles, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
                $cell.Value2 = $n
                $conv++
            }
        }
    }
}
Write-Output ("Converted {0} numeric text cells on '{1}'." -f $conv, $ws.Name)