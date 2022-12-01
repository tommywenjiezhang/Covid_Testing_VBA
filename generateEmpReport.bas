Attribute VB_Name = "generateEmpReport"
Sub generateReport()
    Dim new_wb  As Workbook
    Dim new_sht As Worksheet

    Dim db As New testDb
    Dim result As Variant
    Dim util As New testUtil
    Dim empID As String
    Dim path As String
    
    path = ThisWorkbook.path
    
    
    If util.InRange(ActiveCell, empList.Range("A2:B1000")) Then
        empQryfrm.Show
    Else
    
        MsgBox "Wrong Area, Please select a name from employee list"
    
    End If
    





        
End Sub


Sub ThisWorkbookName()
MsgBox ThisWorkbook.name
End Sub
