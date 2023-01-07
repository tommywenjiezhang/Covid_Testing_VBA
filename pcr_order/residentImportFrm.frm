VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} residentImportFrm 
   Caption         =   "UserForm2"
   ClientHeight    =   5118
   ClientLeft      =   48
   ClientTop       =   390
   ClientWidth     =   8070
   OleObjectBlob   =   "residentImportFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "residentImportFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnImport_Click()
    If Me.wingsCbo.Value = "" Then
         Me.wingsCbo.BackColor = RGB(255, 255, 102)
    Else
    Call residentCrtl.importName(Me.wingsCbo.Value)
    Unload Me
    End If
End Sub

Private Sub updateDbBtn_Click()
    If Me.wingsCbo.Value = "" Then
         Me.wingsCbo.BackColor = RGB(255, 255, 102)
    Else
        Call openFile.updateDatabase(Me.wingsCbo.Value)
        Call residentCrtl.importName(Me.wingsCbo.Value)
    Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()

    With Me.wingsCbo
        .AddItem "FREEDOM"
        .AddItem "LIBERTY"
        .AddItem "EAGLE"
        .AddItem "INDEPENDENCE"
        .AddItem "OLD GLORY"
    End With
    
End Sub
