Attribute VB_Name = "deleteRow"
Sub deleteRow()
    Dim util As New testUtil
    Dim selectRow  As Long
    Dim answer As Long
    Dim empID As String
    Dim db As New testDb
    Dim empName As String
    
    
    If util.InRange(ActiveCell, empList.Range("A2:B1000")) Then
        selectRow = ActiveCell.Row
        empName = empList.Cells(selectRow, 2).value
        
        answer = MsgBox("Are you sure to delete " & empName, vbQuestion + vbYesNo + vbDefaultButton1, "Delete Employee")
        If answer = vbYes Then
            With empList
                .Unprotect
                empID = .Cells(selectRow, 1).value
                .Rows(selectRow).Delete
                .Protect
            End With
            db.deleteEmployee empID
        Else
            Exit Sub
        End If
        

    Else
        MsgBox "Select on employee name you wish to delete"
    End If
End Sub

Sub deleteTesting()
    
    Dim util As New testUtil
    Dim selectRow  As Long
    Dim answer As Long
    Dim empID As String
    Dim db As New testDb
    Dim empName As String
    Dim typeOfTesting As String
    Dim testingDate As String
    
    
    
    If util.InRange(ActiveCell, testRoster.Range("A2:B1000")) Then
        selectRow = ActiveCell.Row
        empID = testRoster.Cells(selectRow, 1).value
        empName = testRoster.Cells(selectRow, 2).value
        typeOfTesting = testRoster.Cells(selectRow, 5).value
        testingDate = testRoster.Cells(selectRow, 3).value
        
        answer = MsgBox("Are you sure to delete " & typeOfTesting & " test for" & empName, vbQuestion + vbYesNo + vbDefaultButton1, "Delete Employee Testing")
        If answer = vbYes Then
            If Not IsError(CDate(testingDate)) Then
                db.deleteTesting empID, CDate(testingDate), typeOfTesting
                With testRoster
                    .Rows(selectRow).Delete
                End With
            Else
                MsgBox "Test cannot be deleted"
            End If
        Else
            Exit Sub
        End If
        

    Else
        MsgBox "Select on employee name you wish to delete"
    End If

End Sub
