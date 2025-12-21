Attribute VB_Name = "FormatReportHeader"
' Formats the top header row: bold white text on a blue background.
' Drop this file in the MacroHub\Macros folder. It will appear in the Macros tab.
' Works on whichever workbook is selected in the "Target Workbook" dropdown.

Option Explicit

Public Sub FormatReportHeader()
    ' This is the Sub MacroHub will find and run automatically.
    ' No parameters needed - works on the active worksheet.

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim headerRow As Range
    Set headerRow = ws.Rows(1)

    ' ---- Style the header row ----
    With headerRow
        .Font.Bold      = True
        .Font.Color     = RGB(255, 255, 255)   ' White text
        .Interior.Color = RGB(31, 73, 125)     ' Dark navy blue background
        .Font.Size      = 11
        .RowHeight      = 22
    End With

    ' ---- Auto-fit columns so headers are fully visible ----
    ws.UsedRange.Columns.AutoFit

    ' ---- Freeze the header row so it stays visible when scrolling ----
    ActiveWindow.FreezePanes = False            ' Reset first
    ws.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True

    MsgBox "Header formatted on sheet: " & ws.Name & Chr(13) & _
           "Columns auto-fitted and header row frozen.", _
           vbInformation, "MacroHub - Format Complete"
End Sub
