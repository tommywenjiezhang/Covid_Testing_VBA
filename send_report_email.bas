Attribute VB_Name = "send_report_email"
Sub sendemail(receiver As String, attachment_path As String)
    Dim execute_str  As String
    Dim todayDate As String
    
    todayDate = format(Now, "YYYY-MM-DD")
    
        execute_str = "D:\programs\compile_email\dist\send_email.exe" & " --receiver " & Chr(34) & receiver & Chr(34) & _
        " --attachment " & Chr(34) & attachment_path & Chr(34)

        Debug.Print execute_str
        obj = Shell(execute_str, vbMinimizedFocus)
    
    
End Sub
