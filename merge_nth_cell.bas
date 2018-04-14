Attribute VB_Name = "merge_nth_cell"
Option Explicit

Sub mergeCells()

    Dim counter As Integer
    
    For counter = 1 To 1000 Step 4
        Range(Cells(1, counter), Cells(1, counter + 2)).Merge
    Next
    
End Sub
