Attribute VB_Name = "updateTestResult"


Sub main()
    testRoster.Activate
    updateTestResult
    MsgBox "Testing Result updated"
End Sub

Sub updateTestResult()
    Dim pos_arr() As Variant, pos_count As Long
    
    Dim lastRow As Long
    Dim idx As Long
    Dim residentID As String
    Dim result As String
    Dim pos_exe_str As String
    
    Dim message As String
    With testRoster
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        pos_count = 0
        ReDim Preserve pos_arr(pos_count)
        For idx = 3 To lastRow
            If IsEmpty(.Cells(idx, "L")) Then
                .Cells(idx, "L").Interior.color = RGB(255, 255, 102)
                message = "Some result not filled, please fill out the result and export again"
            Else
                If Not IsEmpty(.Cells(idx, "A")) Then
                    residentID = Trim(.Cells(idx, 1).value)
                   
                    result = UCase(Left(.Cells(idx, "L").value, 1))
                    If UCase(result) = "P" Then
                        ReDim Preserve pos_arr(0 To pos_count)
                        pos_arr(pos_count) = residentID
                        pos_count = pos_count + 1
                    End If
            
                End If
                End If
        Next idx
        pos_exe_str = Join(pos_arr, ",")
        Debug.Print "update_resident_test.exe " & "--update --l " & pos_exe_str
        Call run_exe.run_exe("update_resident_test.exe " & "--update --l " & pos_exe_str)
    End With
    
    If Not message = "" Then
        MsgBox message
        testRoster.Activate
    End If
    
End Sub



Function IsvalidDate(ByVal datestr As String) As Boolean
        If Not IsError(CDate(datestr)) _
                    And format(CDate(datestr), "mm-dd-yyyy") = format(Date, "mm-dd-yyyy") Then
                    IsvalidDate = True
        Else
                IsvalidDate = False
        
        End If
End Function


Function checkIfbothTest(ByVal str As String) As Boolean
    If InStr(str, "&") > 0 Then
        checkIfbothTest = True
    Else
        checkIfbothTest = False
        
    End If
    
End Function

Sub prefillTesting()
    Dim last_row As Long
    Dim idx As Long
    Dim db As New testDb
    With testRoster
        last_row = .Cells(Rows.Count, "A").End(xlUp).Row
        For idx = 3 To last_row
            If IsEmpty(.Cells(idx, "G")) Then
                .Cells(idx, "G").value = "N"
            End If
        Next idx
        .Cells(last_row, "A").EntireRow.Select
    End With
End Sub

