Option Explicit

Function SmartDateTransformer(inputText As String) As Date
    Dim result As Date
    Dim parsedDate As Date
    
    ' Handle special cases
    Select Case LCase(Trim(inputText))
        Case "today", "now"
            result = Date
        Case "tomorrow"
            result = Date + 1
        Case "yesterday"
            result = Date - 1
        Case Else
            ' Try to parse the date using various methods
            If TryParseDate(inputText, parsedDate) Then
                result = parsedDate
            Else
                ' If all parsing attempts fail, raise an error
                Err.Raise 1000, , "Unable to parse date: " & inputText
            End If
    End Select
    
    SmartDateTransformer = result
End Function

Function TryParseDate(inputText As String, ByRef outputDate As Date) As Boolean
    On Error Resume Next
    
    ' Try standard date conversion
    outputDate = CDate(inputText)
    If Err.Number = 0 Then
        TryParseDate = True
        Exit Function
    End If
    
    ' Try to parse dates like "Monday May 3rd"
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "(\w+)\s+(\w+)\s+(\d+)(?:st|nd|rd|th)?"
    regex.IgnoreCase = True
    
    Dim matches As Object
    Set matches = regex.Execute(inputText)
    
    If matches.Count > 0 Then
        Dim dayOfWeek As String, month As String, day As String
        dayOfWeek = matches(0).SubMatches(0)
        month = matches(0).SubMatches(1)
        day = matches(0).SubMatches(2)
        
        outputDate = DateValue(month & " " & day & ", " & year(Date))
        
        ' Adjust the year if the resulting date is in the past
        If outputDate < Date Then
            outputDate = DateAdd("yyyy", 1, outputDate)
        End If
        
        TryParseDate = True
        Exit Function
    End If
    
    ' If all attempts fail, return False
    TryParseDate = False
End Function

Sub TransformDatesInRange()
    Dim rng As Range
    Dim cell As Range
    Dim transformedDate As Date
    Dim errorCount As Long
    Dim successCount As Long
    
    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    
    Set rng = Selection
    
    ' Transform dates in the range
    For Each cell In rng
        If cell.Value <> "" Then
            On Error Resume Next
            transformedDate = SmartDateTransformer(cell.Value)
            
            If Err.Number = 0 Then
                cell.Value = transformedDate
                cell.NumberFormat = "yyyy-mm-dd"
                successCount = successCount + 1
            Else
                cell.Interior.Color = RGB(255, 200, 200)  ' Light red for errors
                errorCount = errorCount + 1
            End If
            On Error GoTo 0
        End If
    Next cell
    
    ' Report results
    MsgBox "Transformation complete." & vbNewLine & _
           successCount & " dates transformed successfully." & vbNewLine & _
           errorCount & " errors encountered (highlighted in red).", vbInformation
End Sub

