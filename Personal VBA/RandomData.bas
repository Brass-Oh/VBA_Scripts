Sub InsertRandomTestData()
    Dim cell As Range
    Dim rng As Range
    
    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    
    ' Set the random number generator
    Randomize
    
    ' Loop through each cell in the selection
    For Each cell In Selection
        ' Generate a random integer between 1 and 1000
        cell.Value = Int((1000 * Rnd) + 1)
    Next cell
    
    MsgBox "Random test data inserted successfully!", vbInformation
End Sub
