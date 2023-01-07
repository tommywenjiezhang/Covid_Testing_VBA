VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3072
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   4620
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnCancel_Click()

Unload Me
End Sub

Private Sub btnSubmit_Click()
    Dim myFileName As String
    Dim outputStr As String
    Dim util As New testUtil
    Dim path As String
    
    path = util.getDriveName()
    
    myFileName = path & "\programs\automateTesting" & "\output.txt"

    
    If Not Me.userName.Value = "" And Not Me.password.Value = "" Then
    Open myFileName For Output As #1
        outputStr = "username=" & Me.userName.Value & vbCrLf & _
        "password=" & Me.password.Value
        Write #1, outputStr
    Close #1
        
    Call registerTest.registerTest
    
    Else
     With Me
        .userName.BackColor = RGB(255, 255, 0)
        .password.BackColor = RGB(255, 255, 0)
     End With
     
    End If
    
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    Me.userName.Value = "njvetempmn"
    Me.password.Value = "stars1776"
End Sub
