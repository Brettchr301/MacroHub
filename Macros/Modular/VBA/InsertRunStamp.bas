Attribute VB_Name = "InsertRunStamp"
Public Sub InsertRunStamp()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").Value = "Run Stamp"
    ws.Range("B1").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub