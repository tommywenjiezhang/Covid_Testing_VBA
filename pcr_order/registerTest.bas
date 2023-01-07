Attribute VB_Name = "registerTest"
Sub registerTest()
    
    Dim sht As Worksheet
    Dim export_rng As Range
    Dim last_row As Long
    Dim temp_wb As Workbook
    Dim myCSVFileName As String
    Dim testCode As Long
    Dim diagnosisCode As String
    Dim util As New testUtil
    Dim workdir As String
    Dim m_folders As New TestExport
    
    
    
    'the current worksheet
    Set sht = ActiveSheet
    
    
    If m_folders.FolderExists(m_folders.full_path) Then
        
    
    End If
    
    'getting the path for export
    
    workdir = Environ("USERPROFILE") & "\Desktop"
    
    myCSVFileName = m_folders.full_path & "\export_list" & ".csv"
    
    
    If Not sht.CodeName = "empList" And Not sht.CodeName = "residentList" Then
        MsgBox "Select the wrong sheet"
        Exit Sub
    End If
    
    
    With sht
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        Set export_rng = .Range("A1:C" & last_row)
    End With
    
    
    If sht.Range("F1").Value = "" Or sht.Range("F2").Value = "" Or sht.Range("F3").Value = "" Then
        sht.Range("F1:F2").Select
        With Selection
            .Interior.Color = RGB(255, 255, 0)
            .BorderAround _
        LineStyle:=xlContinuous, _
        Weight:=xlThick
        End With
        
        MsgBox "Please enter the test Code and diagnosis code"
        
        Exit Sub
        
        
    Else
    
        sht.Range("B2:C" & sht.Rows.Count).ClearContents
        testCode = CLng(sht.Range("F1").Value)
        diagnosisCode = sht.Range("F2").Value
        sht.Range("B2:B" & last_row).Value = sht.Range("F3").Value
        sht.Range("C2:C" & last_row).Value = format(sht.Range("F4").Value, "hh:mm")
        
    End If
    
    Set temp_wb = Workbooks.Add(1)
    export_rng.Copy
    Dim temp_last As Long
    Dim temp_sht  As Worksheet
    Dim userName As String
    Dim password As String
    
    Set temp_sht = temp_wb.Sheets(1)
    
    With temp_sht
        .Range("A1").PasteSpecial xlPasteValues
            temp_last = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Columns("B").NumberFormat = "mm/dd/yyyy"
        .Columns("C").NumberFormat = "hh:mm"
        .Range("B1").Value = "Collection Date"
        .Range("C1").Value = "Collection Time"
        .Range("D1").Value = "Test Code"
        .Range("E1").Value = "Diagnosis Code"
        .Range("D2:D" & temp_last).Value = testCode
        .Range("E2:E" & temp_last).Value = diagnosisCode
    End With
    
    
    With temp_wb
        On Error GoTo file_not_save
        .SaveAs fileName:=myCSVFileName, FileFormat:=xlCSV, CreateBackup:=False
        .Close
    End With
    
    If sht.CodeName = "empList" Then
        userName = passwordsht.Range("B1").Value
        password = passwordsht.Range("B2").Value
        Call callPython.callPython(userName, password)
    Else
        userName = passwordsht.Range("B3").Value
        password = passwordsht.Range("B4").Value
        Call callPython.callPython(userName, password)
    End If
    
done:
    Exit Sub
file_not_save:
    MsgBox "File not Saved"
    Exit Sub




End Sub
