Attribute VB_Name = "refreshRoster"
Sub refreshRoster()
Dim db As New testDb
Dim names As Variant
Dim util As New testUtil
Dim last_row As Long



names = db.getEmpName()
Dim i As Long
Dim j As Long
Dim start As Long

start = 2
If Not util.isArrayEmpty(names) Then
    empList.Unprotect
    With empList
        .Range("A2:E" & 4000).ClearContents
        .Range("A2:E" & 4000).Interior.ColorIndex = xlNone
        
    End With
    For i = LBound(names, 2) To UBound(names, 2)
        empList.Cells(start + i, 1).value = names(0, i)
        empList.Cells(start + i, 2).value = names(1, i)
    
    Next i
    empList.Protect
End If



End Sub


Sub importBirthday()
    Dim db As New testDb
    Dim birthday As Variant
    Dim util As New testUtil
    
    birthday = db.getEmpBirthday()
    Dim i As Long
    Dim start As Long
    empBirthday.Cells.ClearContents
    
    start = 1
    If Not util.isArrayEmpty(birthday) Then
        
        For i = LBound(birthday, 2) To UBound(birthday, 2)
            With empBirthday
                .Cells(start + i, 1).value = birthday(0, i)
                .Cells(start + i, 2).value = birthday(1, i)
                
            End With
        Next i
    
    End If
    
    
End Sub


Sub importVaccine()
    Dim db As New testDb
    Dim vaccines As Variant
    Dim util As New testUtil
    Dim lastRow As Long
    
    vaccines = db.getVaccinated()
    Dim i As Long
    Dim start As Long
    empVaccine.Cells.ClearContents
    
    start = 1
    
    If Not util.isArrayEmpty(vaccines) Then
    
        For i = LBound(vaccines, 2) To UBound(vaccines, 2)
            With empVaccine
                .Cells(start + i, 1).value = vaccines(0, i)
                .Cells(start + i, 2).value = vaccines(1, i)
                
            End With
        Next i
        
    With empVaccine
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("A1:B" & lastRow).RemoveDuplicates Columns:=Array(1, 2)
    End With
    
    End If
    
    
End Sub

Sub addVaccine()
    Dim db As New testDb
    Dim empName As String
    Dim util As New testUtil
    
        
    If ActiveCell.value = "" Then
        MsgBox "No Person selected exiting........."
        Exit Sub
    Else
        If Not (util.InRange(ActiveCell, empList.Range("B2:B1000"))) Then
            MsgBox "Selecting Wrong Area please select under empolyee name........."
            Exit Sub
        Else
            empName = ActiveCell.value
            db.insertVaccine (empName)
            MsgBox "Employee successfully add to vaccination list"
        End If
    End If
    
    
    
    
End Sub



Sub updateVaccine()
    Dim last_row  As Long
    Dim empID As String
    Dim vaccineType As String
    Dim vaccine_rng As Range
    Dim db As New testDb
    Dim o As Object
    

    With empList
        .Unprotect
        last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        '.Range("E2:E" & last_row).FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(TRIM([@[Employee ID]]),empVaccine!R1C1     :R2000C2,1,FALSE)),""No Vaccine"",""vaccinated"")"
        
        
            For i = 2 To last_row
                If CStr(.Cells(i, "E").value) <> "" Then
                    vaccineType = CStr(.Cells(i, "E").value)
                    empID = CStr(.Cells(i, "A").value)
                    Debug.Print empID & "Update to " & vaccineType
                    db.updateVaccineType empID, vaccineType
                End If
            Next i
            
            
        .Protect
        MsgBox CStr(last_row) & " VACCINE RECORD UPDATED"
        
    End With
End Sub



Sub lookupVaccine()
    
    Dim last_row  As Long
    Dim vaccine_rng As Range
    

    With empList
        .Unprotect
        last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set vaccine_rng = .Range(.Cells(2, "E"), .Cells(last_row, "E"))
        
        .Range("E2:E" & last_row).Interior.ColorIndex = 0
        
        .Range("E2:E" & last_row).FormulaR1C1 = _
        "=VLOOKUP(TRIM(RC[-3]),empVaccine!R[-1]C[-4]:R[5810]C[-3],2,FALSE)"
        
        vaccine_rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""1st Booster"""
        vaccine_rng.FormatConditions(vaccine_rng.FormatConditions.Count).SetFirstPriority
        With vaccine_rng.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.399945066682943
        End With
        
        vaccine_rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""2nd Booster"""
        With vaccine_rng.FormatConditions(2).Interior
            .color = RGB(173, 255, 47)
        End With
        .Protect
    End With
    
End Sub




