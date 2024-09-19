Sub RemoveAllFormatting()
    Dim rng As Range
    Dim cell As Range
    Dim shape As shape
    
    ' Check if there's a selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    
    ' Set the range to work with
    Set rng = Selection
    
    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False
    
    ' Remove all cell formatting
    rng.ClearFormats
    
    ' Clear contents (optional, comment out if you want to keep cell values)
    ' rng.ClearContents
    
    ' Remove hyperlinks
    For Each cell In rng
        If cell.Hyperlinks.Count > 0 Then
            cell.Hyperlinks.Delete
        End If
    Next cell
    
    ' Remove comments
    rng.ClearComments
    
    ' Unmerge cells
    rng.UnMerge
    
    ' Remove conditional formatting
    rng.FormatConditions.Delete
    
    ' Remove data validation
    rng.Validation.Delete
    
    ' Remove all shapes (including charts) that intersect with the range
    For Each shape In ActiveSheet.Shapes
        If Not Intersect(rng, Range(shape.TopLeftCell, shape.BottomRightCell)) Is Nothing Then
            shape.Delete
        End If
    Next shape
    
    ' Reset column widths to standard
    rng.ColumnWidth = 8.43 ' Standard Excel column width
    
    ' Reset row heights to standard
    rng.RowHeight = 15 ' Standard Excel row height
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
End Sub
