Sub RemoveColorFormatting()
    Dim cell As Range
    Dim countCellsChanged As Long
    Dim selectedRange As Range
    
    ' Check if a range is selected
    If TypeName(Selection) = "Range" Then
        Set selectedRange = Selection
    Else
        MsgBox "Please select a range before running this macro.", vbExclamation
        Exit Sub
    End If
    
    ' Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    
    ' Loop through each cell in the selected range
    For Each cell In selectedRange
        ' Check if the cell has any color formatting (fill or font)
        If cell.Interior.ColorIndex <> xlNone Or _
           cell.Font.ColorIndex <> xlAutomatic Then
            
            ' Remove fill color
            cell.Interior.ColorIndex = xlNone
            
            ' Remove font color
            cell.Font.ColorIndex = xlAutomatic
            
            ' Increment the counter
            countCellsChanged = countCellsChanged + 1
        End If
    Next cell
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    ' Display a message with the number of cells changed
    MsgBox countCellsChanged & " cells have had their color formatting removed.", vbInformation
End Sub