Attribute VB_Name = "TrimSelection"
Sub TrimSelection()
Attribute TrimSelection.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim cell As Range
    For Each cell In Selection
        cell.Value = Trim(cell.Value)
    Next cell

End Sub
