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



Private Sub BTN_MoveSelectedLeft_Click()
    Call moveSigle(Me.colLstBox, Me.rowLstbox)
End Sub

Private Sub BTN_MoveSelectedRight_Click()
    Call moveSigle(Me.rowLstbox, Me.colLstBox)
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



Function getLstboxvalue(xlstBox As Object) As String


    Size = Me.dateLstBox.ListCount - 1
    ReDim ListBoxContents(0 To Size) As String
    
    Dim I As Integer

    For I = 0 To Size
        ListBoxContents(I) = xlstBox.List(I)
    Next I
    
    return_str = Join(ListBoxContents, ",")
    getLstboxvalue = return_str
End Function


Private Sub quitBtn_Click()
    Unload Me
End Sub

Private Sub reportTypeCbo_Change()
    If Me.reportTypeCbo.value = "Custom Report" Then
        With Me.MultiPage1
         .Pages(1).Visible = True
        .value = 1
        End With
        
        With Me.rowLstbox
            .AddItem "empID"
            .AddItem "empName"
            .AddItem "DOB"
        End With
        
        With Me.colLstBox
            .AddItem "timeTested"
            .AddItem "typeOfTest"
            .AddItem "result"
        End With
         Me.rowLstbox.MultiSelect = fmMultiSelectMulti
        Me.colLstBox.MultiSelect = fmMultiSelectMulti
    End If
End Sub

Private Sub submitBtn_Click()
    Dim day_rng_str As String
    Dim Size As Integer
    Dim exe_str As String
    Size = Me.dateLstBox.ListCount - 1
    ReDim ListBoxContents(0 To Size) As String
    
    Dim I As Integer

    For I = 0 To Size
        ListBoxContents(I) = Me.dateLstBox.List(I)
    Next I
    
    day_rng_str = Join(ListBoxContents, ",")
    
    If Not Me.reportTypeCbo.value = "" Then
        
        Select Case Me.reportTypeCbo.value
            Case "By Employees"
                exe_str = "custom_reports  --date " & Chr(34) & day_rng_str & Chr(34) & " --report_type " & "EMP_BY_DAY"
                Debug.Print exe_str
            Case "Custom Report"
                Dim rows_str As String
                Dim col_str As String
                
                row_str = " --rows " & getLstboxvalue(Me.rowLstbox)
                col_str = " --columns " & getLstboxvalue(Me.colLstBox)
                
                exe_str = "custom_reports  --date " & Chr(34) & day_rng_str & Chr(34) & " --report_type " & "CUSTOM"
                exe_str = exe_str + row_str + col_str
                Debug.Print exe_str
                
                
        End Select
    
    Else
        Me.reportTypeCbo.BackColor = RGB(255, 255, 0)
        MsgBox "Please select a report type"
    
    End If
End Sub


Sub moveSigle(xListBox1 As Object, xListBox2 As Object)
    Dim I As Long
    For I = 0 To xListBox1.ListCount - 1
        If I = xListBox1.ListCount Then Exit Sub
        If xListBox1.Selected(I) = True Then
            xListBox2.AddItem xListBox1.List(I)
            xListBox1.RemoveItem I
            I = I - 1
        End If
    Next
End Sub


Private Sub UserForm_Initialize()
    With Me.reportTypeCbo
            .AddItem "By Employees"
            .AddItem "By Department"
            .AddItem "Custom Report"
    End With
    With Me.MultiPage1
        .Pages(1).Visible = False
    End With
End Sub
