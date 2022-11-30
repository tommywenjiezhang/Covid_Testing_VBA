Attribute VB_Name = "exportPDF"

Sub exportPDF(Optional ByVal input_date As Date)
    Dim execute_str  As String
    Dim todayDate As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If Not IsNull(input_date) Then
        todayDate = format(input_date, "MM/DD/YYYY")
    Else
        todayDate = format(Now, "MM/DD/YYYY")
    End If

    execute_str = "pdfReport -csv --start " & todayDate
    
    Call run_exe.run_exe(execute_str)
    
    
End Sub



Sub exportPDFByWeek(ByVal startDate As Date, endDate As Date)
    Dim execute_str  As String
    Dim startDateStr As String, endDateStr As String
    
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If Not IsNull(startDate) And Not IsNull(endDate) Then
        startDateStr = format(startDate, "MM/DD/YYYY")
        endDateStr = format(endDate, "MM/DD/YYYY")
    End If

    execute_str = "pdfReport -csv --start " & startDateStr & " --end  " & endDateStr
    
    Call run_exe.run_exe(execute_str)
    
    
End Sub


