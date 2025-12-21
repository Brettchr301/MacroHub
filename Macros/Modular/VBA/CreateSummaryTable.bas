Attribute VB_Name = "CreateSummaryTable"
' MacroHub VBA Macro: CreateSummaryTable
' Inserts a formatted summary table at the end of the active sheet
' showing totals/averages for every numeric column.
Option Explicit

Sub CreateSummaryTable()
    Dim ws As Worksheet
    Dim ur As Range
    Dim lastRow As Long, lastCol As Long, firstRow As Long, firstCol As Long
    Dim sumRow As Long, avgRow As Long, labelCol As Long
    Dim c As Long
    Dim cellVal As Variant
    Dim isNumericCol As Boolean

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    firstRow = ur.Row
    firstCol = ur.Column
    lastRow  = firstRow + ur.Rows.Count - 1
    lastCol  = firstCol + ur.Columns.Count - 1

    If ur.Rows.Count < 2 Then
        MsgBox "No data found on active sheet.", vbExclamation
        Exit Sub
    End If

    ' Insert 3 blank rows below data then write labels
    sumRow = lastRow + 2
    avgRow = lastRow + 3
    labelCol = firstCol

    ws.Cells(sumRow, labelCol).Value = "TOTAL"
    ws.Cells(avgRow, labelCol).Value = "AVERAGE"

    ' Style label cells
    With ws.Range(ws.Cells(sumRow, labelCol), ws.Cells(avgRow, labelCol))
        .Font.Bold = True
        .Interior.Color = RGB(46, 117, 182)
        .Font.Color = RGB(255, 255, 255)
    End With

    For c = firstCol + 1 To lastCol
        ' Check if column is numeric by testing row 2 (first data row)
        cellVal = ws.Cells(firstRow + 1, c).Value
        isNumericCol = IsNumeric(cellVal) And Not IsEmpty(cellVal)

        If isNumericCol Then
            Dim dataRange As String
            dataRange = ws.Cells(firstRow + 1, c).Address & ":" & ws.Cells(lastRow, c).Address

            ws.Cells(sumRow, c).Formula = "=SUM(" & dataRange & ")"
            ws.Cells(avgRow, c).Formula = "=IFERROR(AVERAGE(" & dataRange & "),""N/A"")"

            ' Format: number with commas, 2 decimals
            ws.Cells(sumRow, c).NumberFormat = "#,##0.00"
            ws.Cells(avgRow, c).NumberFormat = "#,##0.00"
        Else
            ws.Cells(sumRow, c).Value = "-"
            ws.Cells(avgRow, c).Value = "-"
        End If

        ' Shading
        ws.Cells(sumRow, c).Interior.Color = RGB(189, 215, 238)
        ws.Cells(avgRow, c).Interior.Color = RGB(221, 235, 247)
    Next c

    ws.Columns.AutoFit
    MsgBox "Summary rows added at rows " & sumRow & " and " & avgRow & ".", vbInformation
End Sub
