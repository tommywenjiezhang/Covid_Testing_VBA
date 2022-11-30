Attribute VB_Name = "updateTestResult"


Sub main()
    testRoster.Activate
    updateTestResult
    MsgBox "Testing Result updated"
End Sub

Sub updateTestResult()
    Dim db As New testDb
    
    Dim lastRow As Long
    Dim idx As Long
    Dim empID As String
    Dim result As String
    
    Dim message As String
    With testRoster
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For idx = 3 To lastRow
            If IsEmpty(.Cells(idx, "G")) Then
                .Cells(idx, "G").Interior.color = RGB(255, 255, 102)
                message = "Some result not filled, please fill out the result and export again"
            Else
                If Not IsEmpty(.Cells(idx, "E")) Then
                    empID = Trim(.Cells(idx, 1).value)
                    If IsvalidDate(.Cells(idx, 3)) Then
                        result = UCase(Left(.Cells(idx, "G").value, 1))
                        db.updateTestResult empID, Now, "RAPID", result
                        db.updateTestResult empID, Now, "PCR", result
                        db.updateTestResult empID, Now, .Cells(idx, "E").value, result
                    Else
                          .Cells(idx, "G").Interior.color = RGB(255, 0, 0)
                          .Cells(idx, "G").value = "Not Today"
                    End If
                End If
            End If
        Next idx
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
