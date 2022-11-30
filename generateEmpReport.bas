Attribute VB_Name = "generateEmpReport"
Sub generateReport()
    Dim new_wb  As Workbook
    Dim new_sht As Worksheet

    Dim db As New testDb
    Dim result As Variant
    Dim util As New testUtil
    Dim empName As String
    Dim path As String
    
    path = ThisWorkbook.path
    
    
    
    
    
    
    If util.InRange(ActiveCell, empList.Range("B2:B1000")) Then
        empName = Trim(ActiveCell.value)
            
        
        result = db.getTestedByEmp(empName)

    
        If util.isArrayEmpty(result) Then
            MsgBox "No Testing found"
        Else
            Dim i As Long
            Set new_wb = Workbooks.Add
            Set new_sht = new_wb.Sheets(1)
            
            Dim start As Long
            start = 3
            new_sht.Cells(1, 1).value = empName & "'s Test history"
            new_sht.Cells(2, 1).value = "Name"
            new_sht.Cells(2, 2).value = "Test Date"
            new_sht.Cells(2, 3).value = "Type of Test"
            
            For i = LBound(result, 2) To UBound(result, 2)
                With new_sht
                    .Cells(start + i, 1).value = result(0, i)
                    .Cells(start + i, 2).value = result(1, i)
                    .Cells(start + i, 3).value = result(2, i)
                    .Cells.EntireColumn.AutoFit
                End With
                
            Next i
            
            new_wb.SaveAs filename:=path & "\" & "Testing History for " & empName & ".xlsx"
        End If
        
        
    Else
    
        MsgBox "Wrong Area, Please select a name from employee list"
    
    End If
    





        
End Sub


Sub ThisWorkbookName()
MsgBox ThisWorkbook.name
End Sub
