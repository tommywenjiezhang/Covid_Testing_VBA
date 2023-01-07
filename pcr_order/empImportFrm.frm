VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} empImportFrm 
   Caption         =   "UserForm2"
   ClientHeight    =   4770
   ClientLeft      =   48
   ClientTop       =   390
   ClientWidth     =   6570
   OleObjectBlob   =   "empImportFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "empImportFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnImport_Click()
    If Me.cboNameInitial.Value = "" Then
        Me.cboNameInitial.BackColor = RGB(255, 255, 102)
    ElseIf Me.cboNameInitial.Value = "ALL" Then
        
        Call empCrtl.bulkImport
        Unload Me
    Else
        Call empCrtl.importName(Me.cboNameInitial.Value)
    Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.cboNameInitial.AddItem "A-F"
    Me.cboNameInitial.AddItem "G-L"
    Me.cboNameInitial.AddItem "M-R"
    Me.cboNameInitial.AddItem "S-Z"
    Me.cboNameInitial.AddItem "ALL"
End Sub
