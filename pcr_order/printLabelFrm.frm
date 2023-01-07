VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} printLabelFrm 
   Caption         =   "UserForm2"
   ClientHeight    =   3144
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   5988
   OleObjectBlob   =   "printLabelFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "printLabelFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnSumit_Click()
    Dim collection_date_str As String
    If Not Me.collectionDateTxt.Value = "" Then
        collection_date_str = validationHelper.birthdayExtract(Me.collectionDateTxt.Value)
        Call printLabel.printLabel(collection_date_str)
        Unload Me
    Else
         Me.collectionDateTxt.BackColor = RGB(255, 255, 0)
    End If
        
End Sub

Private Sub UserForm_Click()
    
End Sub

Private Sub UserForm_Initialize()
  Me.collectionDateTxt.Value = format(Date, "mm/dd/yyyy")
End Sub


