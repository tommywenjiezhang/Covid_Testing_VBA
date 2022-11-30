Attribute VB_Name = "importTest"
Option Explicit



Sub importTest()
    Dim result As Variant
    Dim db As New testDb
    Dim todayDate As Date
    Dim tenDayAgo As Date
    Dim util As New testUtil
    
    todayDate = DateAdd("d", 1, Date)
    tenDayAgo = DateAdd("d", -7, Date)
    
    testImport.Cells.ClearContents
    
    Dim i As Long
    Dim j As Long
    
    Dim start As Long
    start = 1
    result = db.getTestHistory(tenDayAgo, todayDate)
    If Not util.isArrayEmpty(result) Then
        For j = LBound(result, 2) To UBound(result, 2)
            With testImport
                .Cells(start + j, 1).value = result(0, j)
                .Cells(start + j, 2).value = result(1, j)
                .Cells(start + j, 3).value = result(2, j)
            End With
        Next j
    End If
    
    
    
    
End Sub

Sub importNoTestList()
    Dim result As Variant
    Dim db As New testDb
    Dim util As New testUtil
    
    Dim i As Long
    Dim j As Long
    
    noTest.Cells.ClearContents
    Dim start As Long
    start = 1
    result = db.getNoTestList()
    If Not util.isArrayEmpty(result) Then
        For j = LBound(result, 2) To UBound(result, 2)
            With noTest
                .Cells(start + j, 1).value = result(1, j)
                .Cells(start + j, 2).value = result(2, j)
                .Cells(start + j, 3).value = format(result(3, j), "mm/dd/yyyy")
            End With
        Next j
    End If
    
End Sub

Sub populate_testing()
    Dim last_row As Long
    Dim idx As Long
    Dim lookup_result As Variant
    Dim lookup_value As String
    Dim days As Long
    Dim message
    Dim testFrequency As Long
    Dim rapid_look_up_range As Variant, pcr_look_up_range As Variant, pcr_lookup_result As Variant, rapid_lookup_result As Variant
    Dim result_test As Variant
    Dim util As New testUtil
    
    testFrequency = CLng(empList.Range("F2").value)
    
    empList.Unprotect
    
     pcr_look_up_range = getFilterRange("PCR")
     rapid_look_up_range = getFilterRange("RAPID")
    
    
    With empList
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range(.Cells(2, 1), .Cells(last_row, 4)).Interior.ColorIndex = xlNone
        .Range("C2:D" & last_row).ClearContents
    End With
    For idx = 2 To last_row
        If Not empList.Cells(idx, 1).value = "" Then
            lookup_value = empList.Cells(idx, 1).value
            If util.isArrayEmpty(pcr_look_up_range) Or util.isArrayEmpty(rapid_look_up_range) Then
                If util.isArrayEmpty(pcr_look_up_range) Then
                    empList.Cells(idx, 4).value = "Test Not Found"
                End If
                If util.isArrayEmpty(rapid_look_up_range) Then
                    empList.Cells(idx, 3).value = "Test Not Found"
                End If
            Else
                pcr_lookup_result = Application.VLookup(lookup_value, pcr_look_up_range, 2, False)
                rapid_lookup_result = Application.VLookup(lookup_value, rapid_look_up_range, 2, False)
                If Not IsError(pcr_lookup_result) Then
                    empList.Cells(idx, 3).value = CDate(pcr_lookup_result)
                    empList.Cells(idx, 3).NumberFormat = "dddd, mm/dd/yy"
                Else
                     empList.Cells(idx, 3).value = "Test Not Found"
                     empList.Cells(idx, 3).Interior.color = RGB(255, 69, 0)
                End If
                If Not IsError(rapid_lookup_result) Then
                    empList.Cells(idx, 4).value = CDate(rapid_lookup_result)
                    empList.Cells(idx, 4).NumberFormat = "dddd, mm/dd/yy"
                Else
                    empList.Cells(idx, 4).Interior.color = RGB(255, 69, 0)
                     empList.Cells(idx, 4).value = "Test Not Found"
                End If
            End If
        End If
    Next idx
    
    empList.Protect
    
End Sub


Function getFilterRange(ByVal filterType As String) As Variant
    Dim last_row As Long
    Dim filter_rng As Range
    
    Dim row_rng As Range
    
    
    With testImport
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        If IsEmpty(.Range("A1")) Then
            Exit Function
        End If
        .Range("A1").AutoFilter Field:=3, Criteria1:=filterType
        getFilterRange = Arr_Visible_Cells()
        
    End With
    
   
     
End Function


Function Arr_Visible_Cells() As Variant
Dim rRow As Range
Dim aArr() As Variant
Dim i As Long
Dim lCount As Long
Dim CellCount As Variant
Dim Range_To_Get As Variant

CellCount = testImport.UsedRange.Rows.Count
Range_To_Get = testImport.UsedRange.Address

ReDim aArr(1 To 3, 1 To CellCount)

lCount = 1
i = 1
For Each rRow In testImport.Range(Range_To_Get)
    If lCount = 4 Then
    i = i + 1
    lCount = 1
    End If
        If rRow.Rows.Hidden = False Then
         aArr(lCount, i) = rRow
        Else
         GoTo Devo:
        End If
lCount = lCount + 1
Devo:
Next

ReDim Preserve aArr(1 To 3, 1 To i)

Arr_Visible_Cells = Application.Transpose(aArr)
End Function

Sub refreshNoTest()
    Dim idx As Long
    Dim look_rng As Range
    Dim last_row As Integer
    
    Set look_rng = noTest.UsedRange
    
    With empList
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Unprotect
    End With
    
    Dim lookup_value As Variant
    Dim lookup_result As Variant
    
    For idx = 2 To last_row
        If Not empList.Cells(idx, 1).value = "" Then
            lookup_value = empList.Cells(idx, 1).value
            lookup_result = Application.VLookup(lookup_value, look_rng, 3, False)
            
            If Not IsError(lookup_result) Then
                If CDate(lookup_result) >= Date Then
                    With Sheets("EMPLOYEE")
                        .Range("A" & idx & ":D" & idx).Interior.color = RGB(255, 0, 0)
                    End With
                End If
            End If
        End If
    Next idx
    
    empList.Protect

End Sub

Sub Import_test_main()
    importTest
    Call refreshRoster.refreshRoster
    importNoTestList
    refreshNoTest
    Call refreshRoster.importBirthday
    Call refreshRoster.importVaccine
    Call refreshRoster.lookupVaccine
    populate_testing
End Sub
