VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "testDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conn As ADODB.Connection
Private dbPath As String
Private tableName As String





Private Sub Class_Initialize()
    dbPath = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & Environ("USERPROFILE") & "\Desktop" & "\Testingdb.accdb;"
                    
End Sub

Public Function insertTesting(empID As String, name As String, timeIn As Date, symptom As Boolean, typeOfTest As String)
    Set conn = New ADODB.Connection
    
    With conn
    .Open dbPath
    .Execute "INSERT INTO Testing" & _
        " (empID,empName,timeTested,symptom,typeOfTest)" & _
        " VALUES('" & empID & "'," & _
        "'" & name & "'," & _
        format(timeIn, "\#yyyy-mm-dd hh:nn:ss\#") & "," & _
        CStr(symptom) & "," & _
        "'" & typeOfTest & "')"
    .Close
    End With
    
    Set conn = Nothing
    
End Function

Public Function updateBirthday(birthday As Date, empName As String)
       Set conn = New ADODB.Connection
    
    With conn
    .Open dbPath
    .Execute "UPDATE empList SET empList.DOB = " & _
        format(birthday, "\#yyyy-mm-dd\#") & _
        "Where empList.empName like '" & empName & "'"
    .Close
    End With
    
    Set conn = Nothing
End Function



Public Function getEmpName() As Variant

Set conn = New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim query As String
Dim names As Variant
Dim util As New testUtil
Dim fullName As Variant




query = "SELECT empID, empName from empList order by empName asc"

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
Public Function getEmpBirthday() As Variant
    Set conn = New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    Dim query As String
    Dim birthday As Variant
    Dim util As New testUtil




query = "SELECT empID, DOB from empList Where DOB IS NOT NULL order by empName asc"


conn.Open dbPath

Set rs = conn.Execute(query)

If Not rs.EOF Or Not rs.BOF Then
    rs.MoveFirst
    birthday = rs.GetRows
End If




rs.Close
conn.Close

Set rs = Nothing
Set conn = Nothing

getEmpBirthday = birthday


End Function

Public Function insertEmpName(ByVal empName As String, empID As String, Optional dob As Date)
    Set conn = New ADODB.Connection
    Dim query As String
    
    If Not IsError(CDate(dob)) Then
        query = "INSERT INTO empList" & _
        " (empID,empName, DOB)" & _
        " VALUES('" & empID & "'," & _
        "'" & empName & "'," & _
        format(dob, "\#yyyy-mm-dd\#") & ")"
        
    Else
        query = "INSERT INTO empList" & _
        " (empID,empName)" & _
        " VALUES('" & empID & "'," & _
        "'" & empName & "')"
    End If
    
    With conn
    .Open dbPath
    .Execute query
    .Close
    End With
    
    Set conn = Nothing
End Function

Public Function insertNoTest(ByVal empID As String, ByVal empName As String, restartDate As Date)
    Set conn = New ADODB.Connection
    Dim query As String
    
    
        query = "INSERT INTO NoTest" & _
        " (empID,empName,restartTestDate)" & _
        " VALUES('" & empID & "'," & _
        "'" & empName & "'," & _
        format(restartDate, "\#yyyy-mm-dd\#") & _
        ")"
        
    Debug.Print query
 
    With conn
    .Open dbPath
    .Execute query
    .Close
    End With
    
    Set conn = Nothing
End Function



Public Function getTestHistory(startDate As Date, endDate As Date, Optional agg As Boolean = False, Optional visitor As Boolean = False) As Variant
    Set conn = New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim name As String
    Dim util As New testUtil
    Dim checkInDate As String
    Dim output As Variant
    Dim result As Variant
    
    Dim query As String
    
    query = "Select empID, timeTested, typeOfTest, empName, result from testing where " & _
    "timeTested >= " & format(startDate, "\#yyyy-mm-dd\#") & " and timeTested <= " & format(endDate, "\#yyyy-mm-dd\#") & _
    " order by timeTested desc"
    
    If agg Then
    
    query = "Select empList.empID, empList.empName, empList.DOB, Testing.timeTested, Testing.typeOfTest, Testing.result" & _
    " FROM empList " & _
    " LEFT JOIN Testing ON empList.empID = Testing.empID " & _
    " where Testing.timeTested <= " & format(endDate, "\#yyyy-mm-dd\#") & _
    " and Testing.timeTested >= " & format(startDate, "\#yyyy-mm-dd\#")

    End If
    
    
    If visitor Then
        query = "SELECT empList.empID, [Testing].[timeTested] AS timeTested, Testing.typeOfTest AS typeOfTest, empList.empName, 'EMPLOYEE' as Category, Testing.result" & _
        " FROM empList " & _
        " LEFT JOIN Testing ON empList.empID = Testing.empID " & _
        " where Testing.timeTested <= " & format(endDate, "\#yyyy-mm-dd\#") & _
        " and Testing.timeTested >= " & format(startDate, "\#yyyy-mm-dd\#") & _
        " UNION " & _
        "Select ID, timeTested, typeOfTest,visitorName, 'VISITOR' as Category, result from visitorTesting where " & _
        "timeTested >= " & format(startDate, "\#yyyy-mm-dd\#") & " and timeTested < " & format(endDate, "\#yyyy-mm-dd\#") & _
         " order by timeTested desc"
        
    
    End If
    
    
    
    Debug.Print query
    
    conn.Open dbPath

    Set rs = conn.Execute(query)
    
    
    
    If Not rs.BOF Or Not rs.EOF Then
        rs.MoveFirst
        result = rs.GetRows
    End If
    
    
    
    
    rs.Close
    conn.Close
    getTestHistory = result
    
    Set rs = Nothing
    Set conn = Nothing
    
End Function
Public Function getTests() As Variant
    Set conn = New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim name As String
    Dim checkInDate As String
    Dim intRow As Long
    Dim output As Variant
    
    Dim fDate  As String

    intRow = 0
    
    Dim query As String
    
    query = "Select empName, timeIn from Tested order by timeIn desc"
    
    conn.Open dbPath

    Set rs = conn.Execute(query)
    
    If Not rs.BOF Or Not rs.EOF Then
    
        rs.MoveFirst
        output = rs.GetRows
        
    End If
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing
    getTests = output
End Function


Public Function getTestedByEmp(empName As String, Optional ByVal typeOfTest As String = "") As Variant
    Set conn = New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim name As String
    Dim checkInDate As String
    Dim intRow As Long
    Dim output As Variant
    
    Dim fDate  As String

    intRow = 0
    
    Dim query As String
    
    If Not typeOfTest = "" Then
    
    query = "Select empName, timeTested, typeOfTest from Testing where empName like '" & empName & "' and typeOfTest like" & _
    "'" & typeOfTest & "'" & "order by timeTested desc"
    
    Else
        query = "Select empName, timeTested, typeOfTest from Testing where empName like '" & empName & "'order by timeTested desc"
    End If
    
    conn.Open dbPath

    Set rs = conn.Execute(query)
    
    If Not rs.BOF Or Not rs.EOF Then
    
        rs.MoveFirst
        output = rs.GetRows
        
    End If
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing
    getTestedByEmp = output
End Function


Function updateTestResult(ByVal empID As String, testDate As Date, testType As String, testResult As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "UPDATE Testing SET result = '" & _
             testResult & "'" & _
             " Where empID like '" & empID & "'" & _
             " and typeOfTest like '" & testType & "'" & _
             " and Format(timeTested,'yyyy-mm-dd') = '" & _
             format(testDate, "yyyy-mm-dd") & "'"
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function

Function getNoTestList()
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim output As Variant
    Dim util As New testUtil
    

    intRow = 0
    
    Dim query As String
    
    query = "Select * FROM NoTest"
    
    conn.Open dbPath

    Set rs = conn.Execute(query)
    
    If Not rs.BOF Or Not rs.EOF Then
    
        rs.MoveFirst
        output = rs.GetRows
        
    End If
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing
    getNoTestList = output
 
End Function


Function deleteEmployee(ByVal empID As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "Delete * from empList where empID = '" & empID & "'"
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function


Function deleteNoTest(ByVal empID As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "Delete * from NoTest where empID = '" & empID & "'"
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function

Function deleteTesting(ByVal empID As String, testDate As Date, testType As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "Delete * from Testing " & _
             " Where empID like '" & empID & "'" & _
             " and typeOfTest like '" & testType & "'" & _
             " and timeTested = " & _
             format(testDate, "\#yyyy-mm-dd hh:nn:ss\#")
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function

Function getVaccinated()
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim output As Variant
    Dim util As New testUtil

    intRow = 0
    
    Dim query As String
    
    query = "Select empName, vaccine_type FROM vaccinate order by empName asc"
    
    conn.Open dbPath

    Set rs = conn.Execute(query)
    
    If Not rs.BOF Or Not rs.EOF Then
    
        rs.MoveFirst
        output = rs.GetRows
        
    End If
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing
    getVaccinated = output
 
End Function


Public Function updateVaccineType(ByVal empID As String, vaccineType As String)
    Dim conn As New ADODB.Connection
    Dim query As String
    query = "UPDATE vaccinate SET vaccine_type = '" & _
             vaccineType & "'" & _
             " Where empID like '" & empID & "'"
    With conn
        .Open dbPath
        Debug.Print query
        .Execute query
        .Close
    End With
    
    Set conn = Nothing
End Function


Public Function insertVaccine(ByVal empName As String)
    Set conn = New ADODB.Connection
    Dim query As String
    
    
        query = "INSERT INTO vaccinate" & _
        " (empName, Vaccinated)" & _
        " VALUES('" & empName & "'," & _
        "" & str(True) & _
        ")"
        
    Debug.Print query
 
    With conn
    .Open dbPath
    .Execute query
    .Close
    End With
    
    Set conn = Nothing
End Function


Public Function getWeeklyMissingTest(Optional typeOfTest As String = "", Optional startDate As Date, Optional endDate As Date) As Variant
    Set conn = New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim name As String
    Dim checkInDate As String
    Dim intRow As Long
    Dim output As Variant
    Dim util As New testUtil
    
    Dim fDate  As String

    intRow = 0
    
    Dim query As String
    
    If typeOfTest = "" Then
    
        query = "Select empName, timeTested, typeOfTest from Testing order by timeTested desc"
    
    Else
     
        query = "SELECT empList.empID, empList.empName, DateDiff ('yyyy', empList.DOB,now()) as age, t.timeTested, v.type_of_vac FROM (empList" & _
                        " LEFT JOIN (Select * from Testing where Testing.timeTested Between " & _
                        format(startDate, "\#yyyy-mm-dd\#") & _
                        " and " & _
                        format(endDate, "\#yyyy-mm-dd\#") & _
                        " and typeOfTest Like " & _
                        "'" & typeOfTest & "') t" & _
                        " on empList.empID = t.empID) left join (select empID, min(vaccine_type) as type_of_vac FROM vaccinate" & _
                        " group by empID) as v on" & _
                        " empList.empID = v.empID order by empList.empName"
    
    End If
    
    Debug.Print query
    
    conn.Open dbPath

    Set rs = conn.Execute(query)
    
    If Not rs.BOF Or Not rs.EOF Then
    
        rs.MoveFirst
        output = rs.GetRows
        
    End If
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing
    getWeeklyMissingTest = output
End Function



Function getAggragateWeeklyTest(typeOfTest As String, startDate As Date, endDate As Date)
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim output As Variant
    Dim util As New testUtil

    intRow = 0
    
    Dim query As String
    
    query = "Select empName, count(empID) as NumberOfTest from" & _
    "(SELECT empID, empName, timeTested,typeOfTest from Testing where typeOfTest = " & _
    "'" & typeOfTest & "'" & _
    " AND timeTested between " & _
    format(startDate, "\#yyyy-mm-dd\#") & " and " & format(endDate, "\#yyyy-mm-dd\#") & _
    ") Group BY empName"
    
    Debug.Print query
    conn.Open dbPath

    Set rs = conn.Execute(query)
    
    If Not rs.BOF Or Not rs.EOF Then
    
        rs.MoveFirst
        output = rs.GetRows
        
    End If
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing
    getAggragateWeeklyTest = output
 
End Function



