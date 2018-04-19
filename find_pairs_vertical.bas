Attribute VB_Name = "find_pairs_vertical"
Option Explicit

Sub findPairsVertical()

    Dim iRowL As Long, iRow As Long, digits As Variant
    Dim firstNum As Integer, secondNum As Integer
    Dim var1 As Variant, var2 As Variant
    
    firstNum = Application.InputBox(Prompt:="Enter first digit:", Type:=1)
    secondNum = Application.InputBox(Prompt:="Enter first digit:", Type:=1)
    
    iRowL = Cells(Rows.count, 1).End(xlUp).Row
    
    For iRow = 1 To iRowL
        digits = Range(Cells(iRow, 1), Cells(iRow, 3))
        var1 = Application.Match(firstNum, digits, 0)
        var2 = Application.Match(secondNum, digits, 0)
        
        If Not IsError(var1) And Not IsError(var2) Then
            Cells(iRow, var1).Interior.Color = rgbChartreuse
            Cells(iRow, var2).Interior.Color = rgbChartreuse
        End If
    Next iRow
    
End Sub
