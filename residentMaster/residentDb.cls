VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "residentDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conn  As Object

Private dbPath As String
Private TableName As String





Private Sub Class_Initialize()
    Dim util As New testUtil
    Dim path As String
    dbPath = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & Environ("USERPROFILE") & "\Desktop" & "\ResidentDb.accdb;"
                    
End Sub


Public Function getResidentName(ByVal wingName As String) As Variant

Dim conn As Object
Set conn = CreateObject("ADODB.Connection")

Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

Dim query As String
Dim names As Variant
Dim util As New testUtil
Dim fullname As Variant




query = "SELECT residentID, residentName, DOB from resident where wings = '" & wingName & "' order by residentName asc"

Set conn = CreateObject("ADODB.Connection")

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


Public Function insertResidents(ByVal residentName As String, birthday As Date, wingName As String, Optional roomNumber As String = "")
    
    Dim qb As New QueryInsert
    Dim query As String
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    
    qb.TableName = "resident"
    qb.InsertColumn "residentName", residentName
    qb.InsertColumn "DOB", format(birthday, "\#yyyy-mm-dd hh:nn:ss\#")
    qb.InsertColumn "wings", wingName
    qb.InsertColumn "roomNum", roomNumber
    
    With conn
    .Open dbPath
    query = qb.GetInsertCommandText
    Debug.Print query
    .Execute query
    .Close
    End With
   
End Function


Public Function insertTesting(ByVal residentID As String, residentName As String, _
                                        timeTested As Date, residentWings As String, symptom As Boolean, _
                                        typeOfTest As String, lotNumber As String, ByVal expirationDate As Date, testKind As String, testReason As String)
                 Dim qb As New QueryInsert
                Dim query As String
                Dim expirationDateStr As String
                
                Dim conn As Object
                Set conn = CreateObject("ADODB.Connection")
                
                qb.TableName = "resident_testing"
                qb.InsertColumn "ResidentID", residentID
                qb.InsertColumn "residentName", residentName
                
                If symptom = True Then
                    qb.InsertColumn "symptom", CStr(-1)
                Else
                    qb.InsertColumn "symptom", CStr(0)
                End If
                
                qb.InsertColumn "timeTested", format(timeTested, "\#yyyy-mm-dd hh:nn:ss\#")
                qb.InsertColumn "wings", residentWings
                qb.InsertColumn "typeOfTest", typeOfTest
                qb.InsertColumn "lotNumber", lotNumber
                qb.InsertColumn "expirationDate", format(expirationDate, "\#yyyy-mm-dd hh:nn:ss\#")
                qb.InsertColumn "TestKind", testKind
                qb.InsertColumn "testReason", testReason
                
                With conn
                .Open dbPath
                query = qb.GetInsertCommandText
                .Execute query
                .Close
                End With
                

    

End Function



Public Function insertResidentName(ByVal residentName As String, wings As String)
    Dim query As String
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
   
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

Function deleteResident(ByVal residentID As String)
    Dim conn As Object
    Dim query As String
    
    Set conn = CreateObject("ADODB.Connection")
    query = "Delete * from resident where residentID = '" & residentID & "'"
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function

Function deleteResidentByWing(ByVal wingName As String)
    Dim conn As Object
    
    Set conn = CreateObject("ADODB.Connection")
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
    Dim conn As Object
     Set conn = CreateObject("ADODB.Connection")
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





