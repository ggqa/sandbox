Attribute VB_Name = "find_duplicate"
Option Explicit
Function bubbleSort(arr As Variant) As String
    Dim temp As Variant
    Dim tempArr As Variant
    Dim i As Integer, j As Integer
    
    tempArr = arr
    
    For i = 1 To 3
        For j = 1 To (3 - i)
            If tempArr(1, j) > tempArr(1, j + 1) Then
                temp = tempArr(1, j)
                tempArr(1, j) = tempArr(1, j + 1)
                tempArr(1, j + 1) = temp
            End If
        Next j
    Next i
    
    bubbleSort = tempArr(1, 1) & tempArr(1, 2) & tempArr(1, 3)
    
End Function

Function compareDigits(arr1 As Variant, arr2 As Variant) As Integer
    
    If bubbleSort(arr1) = bubbleSort(arr2) Then
        compareDigits = 1
    ElseIf arr1(1, 1) = "" Or arr1(1, 1) = "n/a" Then
        compareDigits = -1
    Else:
        compareDigits = 0
    End If
    
End Function

Sub findDup()

    Dim group1 As Range
    Dim group2 As Range
    Dim rowGroup1 As Range
    Dim rowGroup2 As Range
    Dim counter As Integer
    Dim rowGroup1Digits As Range
    Dim rowGroup2Digits As Range
    Dim result As Integer
    
    Set group1 = Application.InputBox(Prompt:="Choose a cell: ", Type:=8)
    
    For Each rowGroup1 In Range(group1, group1(27, 1))
        Set rowGroup1Digits = Range(rowGroup1, rowGroup1.Cells(1, 3))
        Set group2 = group1.Offset(0, 4)
        For Each rowGroup2 In Range(group2, group2(27, 1))
            Set rowGroup2Digits = Range(rowGroup2, rowGroup2.Cells(1, 3))
            result = compareDigits(rowGroup1Digits, rowGroup2Digits)
            If result = 1 Then
                rowGroup2Digits.Interior.Color = rgbChartreuse
                rowGroup1Digits.Interior.Color = rgbChartreuse
            Else
                If result = -1 Then
                    Exit For
                End If
            End If
        Next
    Next
End Sub


