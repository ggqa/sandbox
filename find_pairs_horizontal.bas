Attribute VB_Name = "find_pairs_horizontal"
Option Explicit
Sub findPairs()

    Dim currentCell As Range
    Dim counter As Integer
    Dim firstNum As Integer, secondNum As Integer
    
    firstNum = Application.InputBox(Prompt:="Enter first number:", Type:=1)
    secondNum = Application.InputBox(Prompt:="Enter second number:", Type:=1)
    
    Set currentCell = Range("A2")
    counter = 4
    Do Until gotoRows(currentCell, firstNum, secondNum) = 0
        Set currentCell = Range("A2").Offset(0, counter)
        currentCell.Select
        counter = counter + 4
    Loop
    
    
End Sub

Function gotoRows(currentCell As Range, first As Integer, second As Integer) As Integer
    
    Dim eachRow As Range
    Dim rowDigits As Variant
    Dim var1 As Variant, var2 As Variant

    
    If currentCell <> "" Then
        For Each eachRow In Range(currentCell, currentCell.End(xlDown))
            rowDigits = Range(eachRow, eachRow.End(xlToRight))
            var1 = Application.Match(first, rowDigits, 0)
            var2 = Application.Match(second, rowDigits, 0)
            
            If Not IsError(var1) And Not IsError(var2) Then
                eachRow.Offset(0, var1 - 1).Interior.Color = rgbChartreuse
                eachRow.Offset(0, var2 - 1).Interior.Color = rgbChartreuse
            End If
        Next eachRow
        
        gotoRows = 1
    Else:
        gotoRows = 0
    End If

End Function
