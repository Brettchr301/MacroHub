# MacroHub Macro: ConditionalFormatTopN
# Applies a green-fill conditional format to the top N values in a specified column.
# Also adds a red-fill for the bottom N values.

param(
    [string]$TargetWorkbook = '',
    [string]$TargetSheet    = '',
    [int]$Column            = 1,    # 1-based column index containing numeric data
    [int]$TopN              = 5,
    [string]$TopColor       = '#C6EFCE',   # light green
    [string]$BottomColor    = '#FFC7CE'    # light red/pink
)

function Hex-ToColor([string]$hex) {
    $hex = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($hex.Substring(0,2),16)
    $g = [Convert]::ToInt32($hex.Substring(2,2),16)
    $b = [Convert]::ToInt32($hex.Substring(4,2),16)
    # Excel RGB = R + G*256 + B*65536
    return $r + ($g * 256) + ($b * 65536)
}

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

$ur = $sh.UsedRange
$firstDataRow = $ur.Row + 1      # skip header
$lastDataRow  = $ur.Row + $ur.Rows.Count - 1
$col          = $ur.Column + $Column - 1

$rng = $sh.Range(
    $sh.Cells.Item($firstDataRow, $col),
    $sh.Cells.Item($lastDataRow,  $col)
)

# Clear existing conditional formats on this range
$rng.FormatConditions.Delete()

# Top N — green
$topFmt = $rng.FormatConditions.AddTop10()
$topFmt.TopBottom = 0    # xlTop10Top
$topFmt.Rank      = $TopN
$topFmt.Interior.Color = Hex-ToColor $TopColor

# Bottom N — red
$botFmt = $rng.FormatConditions.AddTop10()
$botFmt.TopBottom = 1    # xlTop10Bottom
$botFmt.Rank      = $TopN
$botFmt.Interior.Color = Hex-ToColor $BottomColor

Write-Output "Conditional formatting applied: top/bottom $TopN highlighted in column $Column."
