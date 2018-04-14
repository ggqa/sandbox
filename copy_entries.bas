Attribute VB_Name = "copy_entries"
Option Explicit

Sub copyEntries()
    
    Dim rCounter As Long, cCounter As Long
    Dim destination As Range
    
    Worksheets("Sheet2").Activate
    ActiveSheet.Range("A1").Select
    
    cCounter = 1
    Do Until ActiveCell.Value = ""
        rCounter = rCounter + 1
        
        Set destination = Worksheets("Sheet5").Cells(rCounter + 1, cCounter)
        Range(ActiveCell, ActiveCell.End(xlToRight)).Copy destination
        
        If rCounter = 21 Then
            rCounter = 0
            cCounter = cCounter + 5
        End If
        
        ActiveCell.Offset(1, 0).Select
    Loop
    
    
End Sub
    


