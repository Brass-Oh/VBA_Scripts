Sub AdvancedFindAndReplace()
    Dim findWhat As String, replaceWith As String
    Dim rng As Range, cell As Range
    Dim caseSensitive As Boolean, wholeCell As Boolean
    Dim replacementCount As Long
    
    ' Get user input
    findWhat = InputBox("Enter the text to find:")
    If findWhat = "" Then Exit Sub
    
    replaceWith = InputBox("Enter the replacement text:")
    caseSensitive = (MsgBox("Case sensitive?", vbYesNo) = vbYes)
    wholeCell = (MsgBox("Match whole cell only?", vbYesNo) = vbYes)
    
    ' Set range to search
    If TypeName(Selection) = "Range" Then
        Set rng = Selection
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    
    ' Perform find and replace
    For Each cell In rng
        Dim cellValue As String
        cellValue = cell.Value
        
        If Not caseSensitive Then
            If wholeCell Then
                ' For whole cell match, we need to preserve the original case
                ' We'll do the case-insensitive comparison separately
                If LCase(cellValue) = LCase(findWhat) Then
                    cell.Value = replaceWith
                    replacementCount = replacementCount + 1
                End If
            Else
                cellValue = LCase(cellValue)
                findWhat = LCase(findWhat)
                If InStr(cellValue, findWhat) > 0 Then
                    cell.Value = Replace(cell.Value, findWhat, replaceWith, , , vbTextCompare)
                    replacementCount = replacementCount + 1
                End If
            End If
        Else
            If wholeCell Then
                If cellValue = findWhat Then
                    cell.Value = replaceWith
                    replacementCount = replacementCount + 1
                End If
            Else
                If InStr(cellValue, findWhat) > 0 Then
                    cell.Value = Replace(cell.Value, findWhat, replaceWith)
                    replacementCount = replacementCount + 1
                End If
            End If
        End If
    Next cell
    
    MsgBox replacementCount & " replacements made.", vbInformation
End Sub
