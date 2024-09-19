Sub ToggleWrapText()
    Dim cell As Range
    Dim wrapState As Boolean
    
    If Selection.Cells.Count = 0 Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    
    ' Get the wrap text state of the first cell
    wrapState = Selection.Cells(1).WrapText
    
    ' Toggle wrap text for all cells in the selection
    For Each cell In Selection
        cell.WrapText = Not wrapState
    Next cell
    
End Sub
