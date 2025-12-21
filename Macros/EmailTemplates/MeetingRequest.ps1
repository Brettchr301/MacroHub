# MacroHub Email Macro: MeetingRequest
# Opens a new Outlook meeting request pre-filled with the supplied parameters.

param(
    [string]$To          = '',
    [string]$Subject     = 'Meeting Request',
    [string]$Body        = '',
    [string]$Location    = '',
    [DateTime]$StartTime = [DateTime]::Today.AddDays(1).AddHours(10),
    [int]$DurationMins   = 60
)

$ol = $null
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
catch {
    try { $ol = New-Object -ComObject 'Outlook.Application' }
    catch { Write-Error "Could not start Outlook."; return }
}

# 26 = olAppointmentItem
$appt = $ol.CreateItem(26)
$appt.MeetingStatus = 1   # olMeeting
if ($To)       { $appt.Recipients.Add($To) | Out-Null }
$appt.Subject  = $Subject
$appt.Location = $Location
$appt.Start    = $StartTime
$appt.End      = $StartTime.AddMinutes($DurationMins)
$appt.Body     = if ($Body) { $Body } else {
@"
Hi,

I'd like to schedule a meeting to discuss:

[Topic]

Please confirm your availability.

Best regards,
$env:USERNAME
"@
}
$appt.Display()
Write-Output "Meeting request opened in Outlook."
