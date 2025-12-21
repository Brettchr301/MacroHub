# MacroHub Macro: ExportRangeToHtml
# Exports a specified range (or used range) to a standalone HTML table file.

param(
    [string]$TargetWorkbook = '',
    [string]$TargetSheet    = '',
    [string]$TargetRange    = '',    # blank = used range
    [string]$OutputPath     = "$env:USERPROFILE\Desktop\ExportedRange.html",
    [string]$TableTitle     = ''
)

$xl = $null
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { Write-Error "Excel is not open."; return }

$wb = if ($TargetWorkbook) {
    $xl.Workbooks | Where-Object { $_.Name -eq $TargetWorkbook } | Select-Object -First 1
} else { $xl.ActiveWorkbook }
if (-not $wb) { Write-Error "Workbook not found."; return }

$sh = if ($TargetSheet) {
    $wb.Sheets | Where-Object { $_.Name -eq $TargetSheet } | Select-Object -First 1
} else { $wb.ActiveSheet }
if (-not $sh) { Write-Error "Sheet not found."; return }

$rng = if ($TargetRange) { $sh.Range($TargetRange) } else { $sh.UsedRange }

$nRows = $rng.Rows.Count
$nCols = $rng.Columns.Count

$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8">')
[void]$sb.AppendLine('<style>body{font-family:Calibri,Arial,sans-serif;font-size:13px}')
[void]$sb.AppendLine('table{border-collapse:collapse;width:100%}')
[void]$sb.AppendLine('th{background:#2E75B6;color:#fff;padding:6px 10px;text-align:left}')
[void]$sb.AppendLine('td{padding:5px 10px;border-bottom:1px solid #ddd}')
[void]$sb.AppendLine('tr:nth-child(even){background:#f2f2f2}')
[void]$sb.AppendLine('</style></head><body>')

$title = if ($TableTitle) { $TableTitle } elseif ($sh.Name) { $sh.Name } else { 'Export' }
[void]$sb.AppendLine("<h2>$([System.Web.HttpUtility]::HtmlEncode($title))</h2>")
[void]$sb.AppendLine('<table><thead><tr>')

# Header row
for ($c = 1; $c -le $nCols; $c++) {
    $val = [string]$rng.Cells.Item(1, $c).Text
    [void]$sb.AppendLine("<th>$([System.Web.HttpUtility]::HtmlEncode($val))</th>")
}
[void]$sb.AppendLine('</tr></thead><tbody>')

# Data rows
for ($r = 2; $r -le $nRows; $r++) {
    [void]$sb.Append('<tr>')
    for ($c = 1; $c -le $nCols; $c++) {
        $val = [string]$rng.Cells.Item($r, $c).Text
        [void]$sb.Append("<td>$([System.Web.HttpUtility]::HtmlEncode($val))</td>")
    }
    [void]$sb.AppendLine('</tr>')
}

[void]$sb.AppendLine("</tbody></table><p><small>Exported $(Get-Date -Format 'yyyy-MM-dd HH:mm') from $($wb.Name)</small></p>")
[void]$sb.AppendLine('</body></html>')

[System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
Write-Output "HTML export saved to: $OutputPath"
