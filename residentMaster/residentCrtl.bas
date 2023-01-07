Attribute VB_Name = "residentCrtl"
Sub importName(ByVal wingsName As String)
   
    Dim rDb As New residentDb
    Dim names As Variant
    Dim util As New testUtil
    Dim last_row As Long
    
    names = rDb.getResidentName(wingsName)
    
    
    Dim i As Long
    Dim j As Long
    Dim start As Long
    Dim idx As Long
    
    
    With residentList
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("A2:A" & .Rows.Count).ClearContents
        .Range("B2:B" & .Rows.Count).ClearContents
        .Range("A1").value = "residentName"
        .Range("D3").value = wingsName
    End With
    
    
    start = 2
    idx = 0
    If Not util.isArrayEmpty(names) Then
        importBirthday names

        For i = LBound(names, 2) To UBound(names, 2)
            
                residentList.Cells(start + idx, 1).value = names(0, i)
                residentList.Cells(start + idx, 2).value = names(1, i)
                idx = idx + 1
        
        Next i
    
    End If
    
End Sub

Sub importBirthday(ByVal birthday As Variant)

    Dim i As Long
    Dim start As Long
    Dim util As New testUtil
    ResidentInfo.Cells.ClearContents
    
    start = 1
    If Not util.isArrayEmpty(birthday) Then
        
        For i = LBound(birthday, 2) To UBound(birthday, 2)
            With ResidentInfo
                .Cells(start + i, 1).value = birthday(0, i)
                .Cells(start + i, 2).value = birthday(2, i)
                
            End With
        Next i
    
    End If
    
    
End Sub


Function compareName(ByVal name As String, group As String) As Boolean:
    Dim firstLetter As String
    Dim nameFirst As String
    Dim lastLetter As String
    
    
    firstLetter = Left(group, 1)
    lastLetter = Right(group, 1)
    
    nameFirst = Left(UCase(CStr(name)), 1)

    
    Dim firstResult As Integer
    Dim lastResult As Integer
    
    firstResult = StrComp(nameFirst, firstLetter, vbTextCompare)
    lastResult = StrComp(nameFirst, lastLetter, vbTextCompare)
    
    If firstResult >= 0 And lastResult <= 0 Then
        compareName = True
    Else
        compareName = False
    End If
    
    
End Function

Public Sub checkinSelectedCell()
        Dim select_rng As Range
        
        Set select_rng = Selection
        Dim idx As Long
        
        For Each cell In select_rng
            Debug.Print cell.Row
        Next cell
End Sub
