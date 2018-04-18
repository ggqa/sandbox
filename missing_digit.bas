Attribute VB_Name = "missing_digit"
Option Explicit

Sub findMissing()

    Dim nextRow As Integer
    Dim counter As Integer
    Dim digit As Variant
    Dim dict As Scripting.Dictionary
    Dim missingDigits As String

    Set dict = New Scripting.Dictionary
    nextRow = ActiveCell.Row
    counter = 1
    
    Do Until IsEmpty(ActiveCell)
        Cells(nextRow, ActiveCell.Column).Select
        Range(ActiveCell, ActiveCell.Cells(1, 3)).Select
        For Each digit In Selection
            If digit.Value <> "" Then
                If dict.Exists(digit.Value) Then
                    dict(digit.Value) = dict(digit.Value) + 1
                Else:
                    dict.Add digit.Value, counter
                End If
            Else
                Exit For
            End If
        Next digit
        nextRow = nextRow + 4
    Loop
    
    For counter = 0 To 9
        If Not dict.Exists(counter) Then
            missingDigits = missingDigits & counter & ", "
        End If
    Next counter
    
    MsgBox Left(missingDigits, Len(missingDigits) - 2)

End Sub
