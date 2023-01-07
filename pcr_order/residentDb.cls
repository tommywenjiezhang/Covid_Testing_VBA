VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "residentDb"
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
                    "Data Source=" & Environ("USERPROFILE") & "\Desktop" & "\ResidentDb.accdb;"
                    
End Sub


Public Function getResidentName(ByVal wingName As String) As Variant

Set conn = New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim query As String
Dim names As Variant
Dim util As New testUtil
Dim fullname As Variant




query = "SELECT residentName from resident where wings = '" & wingName & "' order by residentName asc"

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

getResidentName = names
End Function




Public Function insertResidentName(ByVal residentName As String, wings As String)
    Set conn = New ADODB.Connection
    Dim query As String
    
   
        query = "INSERT INTO resident" & _
        " (residentName, wings)" & _
        " VALUES('" & Replace(residentName, "'", "''") & "'," & _
        "'" & wings & "')"
    
    With conn
    .Open dbPath
    .Execute query
    .Close
    End With
    
    Set conn = Nothing
End Function

Function deleteResident(ByVal residentName As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "Delete * from resident where residentName = '" & residentName & "'"
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function

Function deleteResidentByWing(ByVal wingName As String)
    
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "Delete * from resident where wings = '" & wingName & "'"
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing


End Function

Public Function updateResidentName(ByVal residentName As String, newName As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "UPDATE resident SET residentName = '" & _
             newName & "'" & _
             " Where residentName like '" & residentName & "'"
    With conn
        .Open dbPath
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function





