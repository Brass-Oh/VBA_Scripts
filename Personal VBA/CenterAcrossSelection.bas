Sub ImprovedCenterAcrossSelection()
    Dim rng As Range
    Dim cell As Range
    Dim firstCell As Range
    
    ' Check if there's a selection
    If Selection.Cells.Count = 0 Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    
    ' Loop through each area in the selection (in case of non-contiguous selection)
    For Each rng In Selection.Areas
        ' Loop through each row in the selection
        For Each row In rng.Rows
            Set firstCell = row.Cells(1)
            
            ' If there's only one cell in the row, just center it
            If row.Cells.Count = 1 Then
                firstCell.HorizontalAlignment = xlCenter
            Else
                ' Apply "Center Across Selection" to the entire row
                row.HorizontalAlignment = xlCenterAcrossSelection
                
                ' If the first cell is empty, find the first non-empty cell and move its content
                If Len(Trim(firstCell.Value)) = 0 Then
                    For Each cell In row.Cells
                        If Len(Trim(cell.Value)) > 0 Then
                            firstCell.Value = cell.Value
                            cell.ClearContents
                            Exit For
                        End If
                    Next cell
                End If
            End If
        Next row
    Next rng

End Sub
