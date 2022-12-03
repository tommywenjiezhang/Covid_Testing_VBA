VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} customDayRng 
   Caption         =   "UserForm1"
   ClientHeight    =   7032
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8310.001
   OleObjectBlob   =   "customDayRng.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "customDayRng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addTestDatebtn_Click()
    Dim testDateStr As String
    Dim testDate As Date
    
    If Not Me.inputDateTxt.value = "" Then
        testDateStr = validationHelper.birthdayExtract(Me.inputDateTxt.value)
        If Not IsError(CDate(testDateStr)) Then
           testDate = CDate(testDateStr)
            Me.dateLstBox.AddItem format(testDate, "mm/dd/yyyy")
            Me.inputDateTxt.value = ""
        Else
            MsgBox "Date you enter does not match the format mm/dd/yyyy"
        End If
    Else
        Me.inputDateTxt.BackColor = RGB(255, 255, 0)
        MsgBox "Please enter a date"
    End If
End Sub



Private Sub dateLstBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim x As Integer
       If KeyCode = 8 Then
            For x = 0 To Me.dateLstBox.ListCount - 1
             If Me.dateLstBox.Selected(x) = True Then
                    Me.dateLstBox.RemoveItem (x)
             End If
          Next x
    End If
End Sub

Private Sub inputDateTxt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim testDateStr As String
    Dim testDate As Date
    
     If Len(Me.inputDateTxt.value) > 1 And KeyCode = 13 Then
    
    If Not Me.inputDateTxt.value = "" Then
        testDateStr = validationHelper.birthdayExtract(Me.inputDateTxt.value)
        If Not IsError(CDate(startDateStr)) Then
            testDate = CDate(testDateStr)
            Me.dateLstBox.AddItem format(testDate, "mm/dd/yyyy")
            Me.inputDateTxt.value = ""
        Else
            MsgBox "Date you enter does not match the format mm/dd/yyyy"
        End If
    Else
        Me.inputDateTxt.BackColor = RGB(255, 255, 0)
        MsgBox "Please enter a date"
    End If
    End If
    
End Sub




Private Sub quitBtn_Click()
    Unload Me
End Sub

Private Sub submitBtn_Click()
    Dim day_rng_str As String
    Dim Size As Integer
    Dim exe_str As String
    Size = Me.dateLstBox.ListCount - 1
    ReDim ListBoxContents(0 To Size) As String
    
    Dim i As Integer

    For i = 0 To Size
        ListBoxContents(i) = Me.dateLstBox.List(i)
    Next i
    
    day_rng_str = Join(ListBoxContents, ",")
    
    If Not Me.reportTypeCbo.value = "" Then
        
        Select Case Me.reportTypeCbo.value
            Case "By Employees"
                exe_str = "custom_reports  --date " & Chr(34) & day_rng_str & Chr(34) & " --report_type " & "EMP_BY_DAY"
                Debug.Print exe_str
                
                
        End Select
    
    Else
        Me.reportTypeCbo.BackColor = RGB(255, 255, 0)
        MsgBox "Please select a report type"
    
    End If
End Sub



Private Sub UserForm_Initialize()
    With Me.reportTypeCbo
            .AddItem "By Employees"
            .AddItem "By Department"
    End With
End Sub
