VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   5526
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   6438
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSubmit_Click()
        Dim sht As Worksheet
        Dim export_rng As Range
        Dim last_row As Long
        Dim temp_wb As Workbook
        Dim workdir As String
        Dim m_folders As New TestExport
        
        
        workdir = m_folders.full_path
        Debug.Print workdir
                
        
        If Not m_folders.FolderExists(m_folders.full_path) Then
                m_folders.FolderCreate (m_folders.full_path)
                workdir = m_folders.full_path
                Debug.Print workdir
                
        End If
        
        
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
     With Me.reasonTestcbo
            .AddItem "Routine"
            .AddItem "New Admit/Readmit"
            .AddItem "Post-Exposure"
            .AddItem "Symptoms"
     End With
     
     With Me.typeOfTestcbo
            .AddItem "BinaxNow"
            .AddItem "QuickVue"
     End With
     
     
    With Me.result_txt
            .AddItem "Positive"
            .AddItem "Negative"
    End With
End Sub
