# MacroHub Email Macro: FollowUpEmail
# Opens a pre-filled follow-up email in Outlook.

param(
    [string]$To       = '',
    [string]$Subject  = 'Following Up',
    [string]$Context  = '',   # brief description of what you're following up on
    [string]$OrigDate = ''    # date of original email/meeting (optional)
)

$ol = $null
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
catch {
    try { $ol = New-Object -ComObject 'Outlook.Application' }
    catch { Write-Error "Could not start Outlook."; return }
}

$dateRef = if ($OrigDate) { " sent on $OrigDate" } else { '' }
$contextRef = if ($Context) { $Context } else { '[original topic / request]' }

$bodyText = @"
Hi,

I wanted to follow up on my earlier message$dateRef regarding $contextRef.

Please let me know if you need any additional information or if there's anything I can help with.

Thank you for your time.

Best regards,
$env:USERNAME
"@

$mail = $ol.CreateItem(0)   # olMailItem
if ($To) { $mail.To = $To }
$mail.Subject  = $Subject
$mail.Body     = $bodyText
$mail.Display()
Write-Output "Follow-up email opened in Outlook."
