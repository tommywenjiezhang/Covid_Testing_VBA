VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "empDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conn As ADODB.Connection
Private dbPath As String
Private tableName As String





Private Sub Class_Initialize()
    Dim util As New testUtil
    Dim path As String
    path = util.getDriveName()
    dbPath = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & path & "\empList.accdb;"
                    
End Sub


Public Function getEmpName() As Variant

Set conn = New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim query As String
Dim names As Variant
Dim util As New testUtil
Dim fullname As Variant


query = "SELECT empName from empList order by empName asc"

conn.Open dbPath

Set rs = conn.Execute(query)

If Not rs.EOF Or Not rs.BOF Then
    rs.MoveFirst
    names = rs.GetRows
End If


rs.Close
conn.Close

Set rs = Nothing
Set conn = Nothing

getEmpName = names
End Function


Public Function insertEmpName(ByVal empName As String)
    Set conn = New ADODB.Connection
    Dim query As String
    
    query = "INSERT INTO empList" & _
    " (empName)" & _
    " VALUES( " & _
    "'" & empName & "')"
    
    With conn
    .Open dbPath
    .Execute query
    .Close
    End With
    
    Set conn = Nothing
End Function


Public Function updateEmpName(ByVal empName As String, newName As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "UPDATE empList SET empName = '" & _
             newName & "'" & _
             " Where empName like '" & empName & "'"
    With conn
        .Open dbPath
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function

Function deleteEmployee(ByVal empName As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "Delete * from empList where empName like '" & empName & "'"
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function





