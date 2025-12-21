# MacroHub Email Macro: SendWithAttachment
# Composes and optionally sends an email with one or more file attachments.
# Pass -Send to send immediately; otherwise the draft is displayed for review.

param(
    [string]$To          = '',
    [string]$Cc          = '',
    [string]$Subject     = '',
    [string]$Body        = '',
    [string[]]$Attachments = @(),   # array of file paths
    [switch]$Send
)

$ol = $null
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
catch {
    try { $ol = New-Object -ComObject 'Outlook.Application' }
    catch { Write-Error "Could not start Outlook."; return }
}

$mail = $ol.CreateItem(0)
if ($To)      { $mail.To      = $To }
if ($Cc)      { $mail.CC      = $Cc }
if ($Subject) { $mail.Subject = $Subject }
$mail.Body = if ($Body) { $Body } else {
"Hi,`n`nPlease find the attached file(s).`n`nBest regards,`n$env:USERNAME"
}

foreach ($path in $Attachments) {
    if (Test-Path $path) {
        $mail.Attachments.Add($path) | Out-Null
    } else {
        Write-Warning "Attachment not found, skipped: $path"
    }
}

if ($Send) {
    $mail.Send()
    Write-Output "Email sent to $To with $($Attachments.Count) attachment(s)."
} else {
    $mail.Display()
    Write-Output "Email draft opened with $($Attachments.Count) attachment(s)."
}
