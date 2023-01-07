Attribute VB_Name = "callPython"
Sub callPython(ByVal userName As String, ByVal password As String)
    Dim util As New testUtil
    Dim path As String
    Dim exe_str As String
    path = util.getDriveName()
    
    exe_str = path & "\programs\python\python -i " & path & "\programs\automateTesting\main.py " & _
                "--username " & userName & " --password " & password
    Debug.Print exe_str
    obj = Shell(exe_str, vbMinimizedFocus)
End Sub

