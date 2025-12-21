# MacroHub Macro: DataValidationDropdown
# Adds a dropdown (list) data validation to a specified range.
# List items can be supplied directly or read from a named range / sheet range.

param(
    [string]$TargetWorkbook = '',
    [string]$TargetSheet    = '',
    [string]$TargetRange    = 'B2:B100',   # range to receive validation
    [string]$ListItems      = '',           # comma-separated: "Yes,No,N/A"
    [string]$ListRange      = '',           # sheet-qualified range: "Lists!A2:A10"
    [string]$ErrorMessage   = 'Please select a value from the list.',
    [string]$ErrorTitle     = 'Invalid Entry'
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

$rng = $sh.Range($TargetRange)
$rng.Validation.Delete()

$formula = if ($ListRange) {
    "=$ListRange"
} elseif ($ListItems) {
    '"' + ($ListItems -replace '"', '""') + '"'
} else {
    Write-Error "Provide either -ListItems or -ListRange."; return
}

# 3 = xlValidateList, 1 = xlValidAlertStop, 1 = xlBetween
$rng.Validation.Add(3, 1, 1, $formula)
$rng.Validation.IgnoreBlank   = $true
$rng.Validation.InCellDropdown = $true
$rng.Validation.ErrorMessage   = $ErrorMessage
$rng.Validation.ErrorTitle     = $ErrorTitle
$rng.Validation.ShowErrorMessage = $true

Write-Output "Dropdown validation applied to '$TargetRange'."
