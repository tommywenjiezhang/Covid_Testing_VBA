VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} customDayRng 
   Caption         =   "UserForm1"
   ClientHeight    =   3996
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8388.001
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
        If Not IsError(CDate(startDateStr)) Then
            Me.dateLstBox.AddItem testDateStr
        Else
            MsgBox "Date you enter does not match the format mm/dd/yyyy"
        End If
    Else
        Me.inputDateTxt.BackColor = RGB(255, 255, 0)
        MsgBox "Please enter a date"
    End If
End Sub

Private Sub quitBtn_Click()
    Unload Me
End Sub

Private Sub submitBtn_Click()
    Dim day_rng_str As String
    Dim Size As Integer
    Size = Me.dateLstBox.ListCount - 1
    ReDim ListBoxContents(0 To Size) As String
    
    Dim i As Integer

    For i = 0 To Size
        ListBoxContents(i) = Me.dateLstBox.List(i)
    Next i
    
    day_rng_str = Join(ListBoxContents, ",")
    MsgBox day_rng_str
    
End Sub


Private Sub UserForm_Click()

End Sub
