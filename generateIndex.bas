Attribute VB_Name = "generateIndex"
Public Function generateIndex(ByVal c As String) As Long
    Dim o As New Dictionary
    Dim rng As Range
    Dim last_row As Long
    Dim out As Long
    With empList
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        Set rng = .Range("A2:A" & last_row)
    End With
    
    Dim letter As String
    Dim idx As Long
    For Each cell In rng
        letter = Left(cell.value, 1)
        If Not o.Exists(UCase(letter)) Then
            o(letter) = 1
        Else
            o(letter) = o(letter) + 1
        End If
    Next cell
    out = o(UCase(c))
    
    generateIndex = out
End Function

Sub main()
    MsgBox generateIndex("c")
End Sub
