# MacroHub Macro: SummaryReport
# Generates a Summary sheet showing column-by-column: Count, Sum, Average, Min, Max
# for every numeric column in the active sheet's used range.

param(
    [string]$TargetWorkbook  = '',
    [string]$TargetSheet     = '',
    [string]$SummarySheetName = 'Summary'
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

$ur     = $sh.UsedRange
$nCols  = $ur.Columns.Count
$nRows  = $ur.Rows.Count
$srcCol = $ur.Column
$srcRow = $ur.Row

# Remove existing summary sheet
$existing = $wb.Sheets | Where-Object { $_.Name -eq $SummarySheetName }
if ($existing) { $existing.Delete() }

$sumSh = $wb.Sheets.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
$sumSh.Name = $SummarySheetName

# Header row
$headers = @('Column', 'Count', 'Sum', 'Average', 'Min', 'Max')
for ($h = 0; $h -lt $headers.Count; $h++) {
    $cell = $sumSh.Cells.Item(1, $h + 1)
    $cell.Value2 = $headers[$h]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 4166697   # steel blue
    $cell.Font.Color = 16777215      # white
}

$outRow = 2
for ($c = 1; $c -le $nCols; $c++) {
    $header = [string]$ur.Cells.Item(1, $c).Value2
    # Check if column is numeric by sampling first data cell
    $sample = $ur.Cells.Item(2, $c).Value2
    if ($nRows -lt 2 -or $sample -isnot [double] -and $sample -isnot [int] -and $null -eq ($sample -as [double])) { continue }

    # Build column address for SUBTOTAL formulas
    $colLetter = [char](64 + ($srcCol + $c - 1))
    $dataStart = "$colLetter$($srcRow + 1)"
    $dataEnd   = "$colLetter$($srcRow + $nRows - 1)"

    $sumSh.Cells.Item($outRow, 1).Value2 = $header
    $sumSh.Cells.Item($outRow, 2).Formula = "=COUNTA($($sh.Name)!$dataStart`:$dataEnd)"
    $sumSh.Cells.Item($outRow, 3).Formula = "=SUM($($sh.Name)!$dataStart`:$dataEnd)"
    $sumSh.Cells.Item($outRow, 4).Formula = "=IFERROR(AVERAGE($($sh.Name)!$dataStart`:$dataEnd),""N/A"")"
    $sumSh.Cells.Item($outRow, 5).Formula = "=MIN($($sh.Name)!$dataStart`:$dataEnd)"
    $sumSh.Cells.Item($outRow, 6).Formula = "=MAX($($sh.Name)!$dataStart`:$dataEnd)"

    # Alternate row shading
    if ($outRow % 2 -eq 0) {
        $sumSh.Rows.Item($outRow).Interior.Color = 15921906   # light gray
    }
    $outRow++
}

$sumSh.Columns.AutoFit()
$sumSh.Activate()

Write-Output "Summary report created on sheet '$SummarySheetName' with $($outRow - 2) numeric column(s)."
