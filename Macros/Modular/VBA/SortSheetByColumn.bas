Attribute VB_Name = "SortSheetByColumn"
' MacroHub VBA Macro: SortSheetByColumn
' Sorts the active sheet's used range by a user-chosen column, ascending or descending.
Option Explicit

Sub SortSheetByColumn()
    Dim ws As Worksheet
    Dim ur As Range
    Dim colInput As String
    Dim colNum As Long
    Dim direction As String
    Dim xlDir As XlSortOrder

    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    If ur.Rows.Count < 2 Then
        MsgBox "No data to sort.", vbExclamation
        Exit Sub
    End If

    colInput = InputBox("Enter column number to sort by (1 = first column):", "Sort Column", "1")
    If colInput = "" Then Exit Sub
    If Not IsNumeric(colInput) Then MsgBox "Invalid column number.": Exit Sub
    colNum = CLng(colInput)
    If colNum < 1 Or colNum > ur.Columns.Count Then
        MsgBox "Column number out of range (1-" & ur.Columns.Count & ").": Exit Sub
    End If

    direction = InputBox("Sort direction? Enter A for Ascending or D for Descending:", "Sort Direction", "A")
    If UCase(Trim(direction)) = "D" Then
        xlDir = xlDescending
    Else
        xlDir = xlAscending
    End If

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ur.Columns(colNum), SortOn:=xlSortOnValues, Order:=xlDir
        .SetRange ur
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    Dim dirLabel As String
    dirLabel = IIf(xlDir = xlAscending, "Ascending", "Descending")
    MsgBox "Sorted by column " & colNum & " (" & dirLabel & ").", vbInformation
End Sub
