Attribute VB_Name = "generateIndex"
Public Function generateIndex(ByVal c As String) As Long
    Dim o As New Dictionary
    Dim last_row As Long
    Dim out As Long
    
    
    Dim db As New testDb
    Dim names As Variant
    Dim util As New testUtil
    Dim wb_sht As Worksheet, wb As Workbook, rng As Range
    
   
    Set wb = Workbooks.Add
    Set wb_sht = wb.Sheets(1)
   
    
    names = db.getEmpName()
    Dim i As Long
    Dim j As Long
    Dim start As Long
    
    start = 2
    If Not util.isArrayEmpty(names) Then
                With wb_sht
                    .Range("A2:B" & 4000).ClearContents
                    .Range("A2:B" & 4000).Interior.ColorIndex = xlNone
                     For i = LBound(names, 2) To UBound(names, 2)
                                .Cells(start + i, 1).value = names(0, i)
                                .Cells(start + i, 2).value = names(1, i)
                        Next i
                    last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
                    Set rng = .Range("A2:B" & last_row)
                End With
               
    End If
    
  

    
    
    out = getEmpID(c, 1, rng)
    
    generateIndex = out

    wb.Close savechanges:=False
    
End Function



Function getEmpID(ByVal letter As String, num As Long, ByRef rng As Range) As Long
    Dim look_val As Variant
    Dim id As String
    
    id = letter & CStr(num)
    look_val = Application.VLookup(id, rng, 1, 0)
    
    If IsError(look_val) Then
        getEmpID = 1
    Else
        getEmpID = getEmpID(letter, num + 1, rng) + 1
    End If

End Function

