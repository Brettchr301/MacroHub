Attribute VB_Name = "ClearActiveFilters"
Public Sub ClearActiveFilters()
    On Error Resume Next
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    ActiveSheet.AutoFilterMode = True
End Sub