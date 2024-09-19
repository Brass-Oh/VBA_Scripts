Sub CopyPasteAsValues()
    On Error GoTo ErrorHandler
    
    If Selection.Cells.Count = 0 Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    
    ' Store the selected range
    Dim selectedRange As Range
    Set selectedRange = Selection
    
    ' Copy the selection
    selectedRange.Copy
    
    ' Paste as values in the same location
    selectedRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Clear the clipboard
    Application.CutCopyMode = False
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub
