Sub RemoveAllHyperlinks()
    Dim rng As Range
    Dim cell As Range
    Dim hyperlinkCount As Long
    
    ' Check if there's a selection
    If TypeName(Selection) = "Range" Then
        Set rng = Selection
    Else
        ' If no selection, use the entire used range
        Set rng = ActiveSheet.UsedRange
    End If
    
    ' Remove hyperlinks
    For Each cell In rng
        If cell.Hyperlinks.Count > 0 Then
            cell.Hyperlinks.Delete
            hyperlinkCount = hyperlinkCount + 1
        End If
    Next cell
    
    MsgBox hyperlinkCount & " hyperlink(s) have been removed.", vbInformation
End Sub
