Attribute VB_Name = "ProtectUnlockSheet"
' MacroHub VBA Macro: ProtectUnlockSheet
' Toggles sheet protection on the active sheet.
' If the sheet is protected it will prompt for the password to unprotect.
' If unprotected it will prompt for a password (or blank for no password) then protect.
Option Explicit

Sub ProtectUnlockSheet()
    Dim ws As Worksheet
    Dim pwd As String

    Set ws = ActiveSheet

    If ws.ProtectContents Then
        pwd = InputBox("Sheet '" & ws.Name & "' is protected." & vbCrLf & _
                       "Enter password to unprotect (leave blank if none):", _
                       "Unprotect Sheet")
        On Error Resume Next
        ws.Unprotect Password:=pwd
        If ws.ProtectContents Then
            MsgBox "Incorrect password or sheet could not be unprotected.", vbCritical
        Else
            MsgBox "Sheet '" & ws.Name & "' is now unprotected.", vbInformation
        End If
        On Error GoTo 0
    Else
        pwd = InputBox("Enter a password to protect sheet '" & ws.Name & _
                       "'." & vbCrLf & "(Leave blank for no password):", _
                       "Protect Sheet")
        ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
        MsgBox "Sheet '" & ws.Name & "' is now protected.", vbInformation
    End If
End Sub
