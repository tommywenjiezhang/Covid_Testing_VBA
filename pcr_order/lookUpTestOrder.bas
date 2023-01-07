Attribute VB_Name = "lookUpTestOrder"
Sub lookUpTestOrder(ByVal shtName)
    
    Dim filepath As String
    Dim temp_wb As Workbook
    Dim main_wb As Workbook
    Dim main_sht As Worksheet
    Dim temp_sht As Worksheet
    Dim lookup_rng As Range
    
    Dim fso As Object
    Dim path As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.getDriveName(ThisWorkbook.path)
    
    Set main_wb = ThisWorkbook
    
    filepath = path & "\programs\automateTesting\RegisterReport.csv"
    
    
    Set temp_wb = Workbooks.Open(fileName:=filepath)
    
    
    Set temp_sht = temp_wb.Sheets(1)
    
    
    Set lookup_rng = temp_sht.UsedRange
    
    Dim irow As Integer
    Dim last_row As Integer
    Dim found_String As Long
    
    
    With main_wb.Sheets(shtName)
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("A2:A" & last_row).Interior.ColorIndex = 0
        For irow = 1 To last_row
             If Not IsEmpty(Cells(irow, 1)) Then
                Var = Application.VLookup(.Cells(irow, 1).Value, lookup_rng, 2, False)
                If Not IsError(Var) Then
                    If CStr(Var) = "Failed to Order Test" Then
                        .Cells(irow, 1).Interior.Color = RGB(255, 0, 0)
                    End If
                    found_String = InStr(Var, "Ordered")
                    If found_String <> 0 Then
                        .Cells(irow, 1).Interior.Color = RGB(124, 252, 0)
                    End If
                End If
             End If
        Next irow
    End With
    
    exportToWord temp_wb.Sheets(1)
    main_wb.Activate
    temp_wb.Close
End Sub



Sub exportToWord(ByRef sht As Worksheet)

    Set obj = CreateObject("Word.Application")
    obj.Visible = True
    Set newObj = obj.Documents.Add
    
    With sht
        .UsedRange.AutoFilter Field:=2, Criteria1:="Failed to Order Test"
        .UsedRange.Copy
        newObj.Range.Paste
    End With
    
    Application.CutCopyMode = False
    obj.Activate
End Sub




Sub lookup_main()
    lookUpTestOrder ActiveSheet.name
End Sub
