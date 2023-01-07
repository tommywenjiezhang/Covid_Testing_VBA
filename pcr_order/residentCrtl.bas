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
        .Range("A1").Value = "residentName"
    End With
    
    
    start = 2
    idx = 0
    If Not util.isArrayEmpty(names) Then

        For i = LBound(names, 2) To UBound(names, 2)
            
                residentList.Cells(start + idx, 1).Value = names(0, i)
                idx = idx + 1
        
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
