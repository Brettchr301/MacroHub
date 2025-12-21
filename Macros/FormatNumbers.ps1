# Formats all numeric cells in the active sheet: integers get commas, decimals get 2 decimal places.
# This is a PowerShell macro. MacroHub runs it and passes the target workbook name.
# You can call the target workbook via the $WorkbookName parameter that MacroHub injects.
#
# HOW IT WORKS:
#   MacroHub calls:  .\FormatNumbers.ps1 -WorkbookName "MyReport.xlsx"
#   The script connects to Excel via COM and formats the active sheet.

param(
    [string]$WorkbookName = ''    # Injected by MacroHub - the workbook selected in the UI
)

$ErrorActionPreference = 'Stop'

# ---- Connect to Excel ----
$xl = $null
try {
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    throw "Excel is not open. Please open your workbook in Excel first."
}

# ---- Locate the workbook ----
$wb = $null
if ($WorkbookName) {
    try { $wb = $xl.Workbooks.Item($WorkbookName) }
    catch { throw "Workbook '$WorkbookName' is not open in Excel." }
} else {
    $wb = $xl.ActiveWorkbook
}

$ws = $wb.ActiveSheet

Write-Output "Formatting numbers on: $($wb.Name) > $($ws.Name)"

# ---- Loop through used range and apply number formats ----
$formatted = 0
$usedRange  = $ws.UsedRange

for ($r = 1; $r -le $usedRange.Rows.Count; $r++) {
    for ($c = 1; $c -le $usedRange.Columns.Count; $c++) {
        $cell = $usedRange.Cells.Item($r, $c)
        $val  = $cell.Value2

        # Only touch cells that contain a number (not text, not empty)
        if ($null -ne $val -and $val -is [double]) {
            if ($val -eq [Math]::Floor($val)) {
                # Whole number -> comma-separated, no decimals
                $cell.NumberFormat = '#,##0'
            } else {
                # Decimal -> comma-separated, 2 decimal places
                $cell.NumberFormat = '#,##0.00'
            }
            $formatted++
        }
    }
}

Write-Output "Done. Formatted $formatted numeric cell(s)."

# ---- Show a friendly completion message via Excel ----
$xl.StatusBar = "MacroHub: Formatted $formatted cells on $($ws.Name)"

# (The status bar message clears automatically when Excel next updates)
