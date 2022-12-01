Attribute VB_Name = "backDb"
Public Sub CreateBackup()
    Dim Source As String
    Dim Target As String
    Dim export As New TestExport
    Dim a As Integer
    Dim objFSO As Object
    Dim path As String
    path = "C:\Users\menlo\OneDrive\Documents"
    Source = Environ("USERPROFILE") & "\Desktop" & "\Testingdb.accdb"
    If Not export.FolderExists(path & "\testingdb_backup") Then
        export.FolderCreate (path & "\testingdb_backup")
        path = path & "\testingdb_backup"
    Else
        path = path & "\testingdb_backup"
    End If
    Target = path & "\Testing_BackupDB "
    Target = Target & format(Now(), "mm-dd") & ".accdb"
    
    a = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    a = objFSO.CopyFile(Source, Target, True)
    Set objFSO = Nothing
    ' create the backup
    
End Sub
