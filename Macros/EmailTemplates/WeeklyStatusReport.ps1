# MacroHub Email Macro: WeeklyStatusReport
# Creates a weekly status report email from a simple hashtable of sections.

param(
    [string]$To          = '',
    [string]$Cc          = '',
    [string]$TeamName    = 'Team',
    [string]$Accomplishments = '',   # bullet items, one per line
    [string]$InProgress      = '',
    [string]$Blockers         = '',
    [string]$NextWeek         = ''
)

$ol = $null
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
catch {
    try { $ol = New-Object -ComObject 'Outlook.Application' }
    catch { Write-Error "Could not start Outlook."; return }
}

$weekStart = (Get-Date).AddDays(-(([int](Get-Date).DayOfWeek) - 1))
$weekEnd   = $weekStart.AddDays(4)
$weekLabel = "$($weekStart.ToString('MMM d')) - $($weekEnd.ToString('MMM d, yyyy'))"
$subject   = "Weekly Status Report — $weekLabel"

function Format-Bullets([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "  • (none)`n" }
    ($text.Trim() -split '\r?\n' | ForEach-Object { "  • $($_.Trim())" }) -join "`n"
}

$body = @"
Hi $TeamName,

Please find below my status report for the week of $weekLabel.

ACCOMPLISHMENTS
$(Format-Bullets $Accomplishments)

IN PROGRESS
$(Format-Bullets $InProgress)

BLOCKERS / RISKS
$(Format-Bullets $Blockers)

NEXT WEEK
$(Format-Bullets $NextWeek)

Let me know if you have any questions.

Best regards,
$env:USERNAME
"@

$mail = $ol.CreateItem(0)
if ($To) { $mail.To = $To }
if ($Cc) { $mail.CC = $Cc }
$mail.Subject = $subject
$mail.Body    = $body
$mail.Display()
Write-Output "Weekly status report email opened in Outlook."
