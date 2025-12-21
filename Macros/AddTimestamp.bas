Attribute VB_Name = "AddTimestamp"
' Writes today's date and time into cell A1 of the active sheet.
' Useful as a "last updated" stamp on reports.

Option Explicit

Public Sub AddTimestamp()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim stampCell As Range
    Set stampCell = ws.Range("A1")

    ' Write a formatted date/time string
    stampCell.Value  = "Last updated: " & Format(Now(), "yyyy-mm-dd  hh:mm AM/PM")
    stampCell.Font.Italic = True
    stampCell.Font.Color  = RGB(128, 128, 128)   ' Grey so it's subtle

    MsgBox "Timestamp added to " & ws.Name & "!A1", vbInformation, "MacroHub"
End Sub
