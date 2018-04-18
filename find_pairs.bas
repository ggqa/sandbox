Attribute VB_Name = "find_pairs"
Option Explicit
Function check(eachRow As Range, userDigit As Integer) As Boolean
    
    Dim digit As Variant
    Dim result As Boolean
    Dim allRowDigits As Variant
    
    result = False
    For Each digit In eachRow
        If userDigit = digit Then
            digit.Interior.Color = rgbChartreuse
            result = True
            Exit For
        End If
    Next digit
    
    check = result
    
End Function

Function gotoRows(currentCell As Range, numOne As Integer, numTwo As Integer) As Integer
    
    Dim eachRow As Range
    Dim rowDigits As Object

    
    If currentCell <> "" Then
        For Each eachRow In Range(currentCell, currentCell.End(xlDown))
            Set rowDigits = Range(eachRow, eachRow.End(xlToRight))
            If Not (check(rowDigits, numOne) And check(rowDigits, numTwo)) Then
                With rowDigits
                    .ClearFormats
                    .Font.Size = 13
                    .HorizontalAlignment = xlCenter
                End With
            End If
        Next eachRow
        
        gotoRows = 1
    Else:
        gotoRows = 0
    End If

End Function

Sub findPairs()

    Dim currentCell As Range
    Dim counter As Integer
    Dim firstDigit As Integer
    Dim secondDigit As Integer
    
    Set currentCell = Application.InputBox(Prompt:="Enter a cell", Type:=8)
    firstDigit = Application.InputBox(Prompt:="Enter first digit", Type:=1)
    secondDigit = Application.InputBox(Prompt:="Enter second digit", Type:=1)
    
    Application.ScreenUpdating = False
    
    counter = 4
    Do Until gotoRows(currentCell, firstDigit, secondDigit) = 0
        Set currentCell = Range("A2").Offset(0, counter)
        currentCell.Select
        counter = counter + 4
    Loop
    
    Application.ScreenUpdating = True
    
End Sub
