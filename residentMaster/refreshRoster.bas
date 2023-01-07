Attribute VB_Name = "refreshRoster"

Sub importBirthday(birthday As Variant)

    Dim i As Long
    Dim start As Long
    ResidentInfo.Cells.ClearContents
    
    start = 1
    If Not util.isArrayEmpty(birthday) Then
        
        For i = LBound(birthday, 2) To UBound(birthday, 2)
            With ResidentInfo
                .Cells(start + i, 1).value = birthday(0, i)
                .Cells(start + i, 2).value = birthday(1, i)
                
            End With
        Next i
    
    End If
    
    
End Sub
