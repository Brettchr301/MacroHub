' LaunchMacroHub.vba
' ---------------------------------------------------------------
' Drop this module into any Excel workbook (Alt+F11 -> Insert Module)
' to get a one-click launcher for MacroHub.
'
' CONFIGURATION — edit the constants below to suit your setup:
' ---------------------------------------------------------------

' Full path to MacroHub.ps1.
' Leave as "" to auto-detect (same folder as this workbook).
Private Const MACROHUB_PATH As String = ""

' Name of this workbook (used in the error message only).
Private Const WORKBOOK_NAME As String = "MacroHub Launcher"

' ---------------------------------------------------------------
' Tab visibility — set to False to hide a tab at launch.
' Tabs set to False are passed to MacroHub via -HideTabs.
' ---------------------------------------------------------------
Private Const SHOW_CLIPBOARD  As Boolean = True
Private Const SHOW_MACROS     As Boolean = True
Private Const SHOW_SCHEDULER  As Boolean = True
Private Const SHOW_NAVIGATOR  As Boolean = True
Private Const SHOW_TEMPLATES  As Boolean = True
Private Const SHOW_QSYNC      As Boolean = True
Private Const SHOW_QTASKS     As Boolean = True
Private Const SHOW_FILEINDEX  As Boolean = True
Private Const SHOW_AUDITOR    As Boolean = True

' ---------------------------------------------------------------
' LAUNCHER — no changes needed below this line.
' ---------------------------------------------------------------

Sub LaunchMacroHub()
    Dim scriptPath As String

    ' Resolve path: use constant if set, otherwise same folder as workbook.
    If MACROHUB_PATH <> "" Then
        scriptPath = MACROHUB_PATH
    Else
        scriptPath = ThisWorkbook.Path & "\MacroHub.ps1"
    End If

    ' Verify the script exists before launching.
    If Dir(scriptPath) = "" Then
        MsgBox "MacroHub.ps1 not found at:" & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "Update the MACROHUB_PATH constant at the top of LaunchMacroHub.vba, or " & _
               "place MacroHub.ps1 in the same folder as " & WORKBOOK_NAME & ".", _
               vbExclamation, "MacroHub Launcher"
        Exit Sub
    End If

    ' Build -HideTabs argument from False toggles.
    Dim hideTabs As String
    hideTabs = ""
    If Not SHOW_CLIPBOARD Then hideTabs = hideTabs & "Clipboard,"
    If Not SHOW_MACROS     Then hideTabs = hideTabs & "Macros,"
    If Not SHOW_SCHEDULER  Then hideTabs = hideTabs & "Scheduler,"
    If Not SHOW_NAVIGATOR  Then hideTabs = hideTabs & "Navigator,"
    If Not SHOW_TEMPLATES  Then hideTabs = hideTabs & "Templates,"
    If Not SHOW_QSYNC      Then hideTabs = hideTabs & "QSync,"
    If Not SHOW_QTASKS     Then hideTabs = hideTabs & "QTasks,"
    If Not SHOW_FILEINDEX  Then hideTabs = hideTabs & "File Index,"
    If Not SHOW_AUDITOR    Then hideTabs = hideTabs & "Formula Auditor,"
    ' Strip trailing comma.
    If Len(hideTabs) > 0 Then hideTabs = Left(hideTabs, Len(hideTabs) - 1)

    ' Build the powershell command.
    Dim psArgs As String
    psArgs = "-WindowStyle Hidden -File """ & scriptPath & """"
    If Len(hideTabs) > 0 Then
        psArgs = psArgs & " -HideTabs """ & hideTabs & """"
    End If

    ' Launch with -WindowStyle Hidden so no console appears.
    ' No -ExecutionPolicy flag — relies on the user's existing policy.
    Shell "powershell.exe " & psArgs, vbHide
End Sub
