Attribute VB_Name = "clearTesting"
Option Explicit

Sub clearTesting()

Dim answer  As Integer


answer = MsgBox("Are you sure to clear all the testing information ", vbQuestion + vbYesNo)
    If answer = vbYes Then
    
        With testRoster
            .Range(.Cells(3, 1), .Cells(.Rows.Count, "K")).ClearContents
            .Range(.Cells(3, 1), .Cells(.Rows.Count, "K")).Interior.ColorIndex = 0
            
        End With
        
        
    Else
        Exit Sub

    End If


End Sub
