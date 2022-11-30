Attribute VB_Name = "run_exe"
Sub run_exe(exe_str As String)
    Dim execute_str  As String
    Dim execute_folder_path As String

    execute_folder_path = Environ("USERPROFILE") & "\Covid_Testing\"
    
        execute_str = execute_folder_path & exe_str
        Debug.Print execute_str
        obj = Shell(execute_str, vbMinimizedFocus)
End Sub

