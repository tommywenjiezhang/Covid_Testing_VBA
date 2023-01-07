VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} residentActionFrm 
   Caption         =   "Make Change"
   ClientHeight    =   7728
   ClientLeft      =   48
   ClientTop       =   390
   ClientWidth     =   7590
   OleObjectBlob   =   "residentActionFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "residentActionFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub actionCbo_Change()
    Me.headerLbl = "Please select action"
    If Not ActiveCell.Value = "" And checkifInrange() Then
        Me.headerLbl.Caption = Me.headerLbl.Caption & " for " & ActiveCell.Value
            Dim choice As String
             If Not Me.actionCbo.Value = "" Then
                 choice = Me.actionCbo.Value
                 Select Case choice
               Case "Delete Resident"
                 Me.deleteFrame.Visible = True
                 Me.editResidentframe.Visible = False
                 Me.addResidentframe.Visible = False
               Case "Edit Resident"
                  Me.editResidentframe.Visible = True
                  Me.deleteFrame.Visible = False
                  Me.addResidentframe.Visible = False
                  Me.editResidentNametxt.FontSize = 16
                  Me.editResidentNametxt.Value = ActiveCell.Value
               Case "Add Resident"
                 Me.headerLbl.Caption = "Add Resident"
                  Me.addResidentframe.Visible = True
                  Me.editResidentframe.Visible = False
                  Me.deleteFrame.Visible = False
               Case Else
                  MsgBox "Unknown Number"
            End Select
            Else
                MsgBox "Please close the from and select a value"
                Me.actionCbo.BackColor = RGB(255, 255, 0)
            End If
        Else
            Unload Me
            MsgBox "Please select a name"
    End If
    
End Sub
Private Function checkifInrange() As Boolean
        Dim util As New testUtil
        
        If util.InRange(ActiveCell, residentList.Range("A:A")) Then
            
            checkifInrange = True
        Else
            checkifInrange = False
        
        End If
End Function

Private Sub btnDeleteResident_Click()
    Dim residentName As String
    
    Dim answer As Integer
 
    answer = MsgBox("Are you sure you want to delete resident", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Resident")
    
    If answer = vbYes Then
        residentName = ActiveCell.Value
        deleteResident residentName
        Unload Me
        ActiveCell.EntireRow.Delete
    Else
        Exit Sub
    End If
    

End Sub

Private Function combineName(firstName As String, lastName As String)
        Dim fullname As String
        fullname = UCase(Trim(lastName)) & "," & UCase(Trim(firstName))
        
        combineName = fullname
        
    
End Function


Private Sub btnAddResident_Click()
    Dim residentName As String
    Dim wingsCbo As String
    Dim db As New residentDb
    Dim last_row As Integer
    Dim idx As Integer
    
    
    With residentList
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    End With
    
    If Not Me.firstNameTxt.Value = "" And Not Me.lastNameTxt.Value = "" And Not Me.wingsCbo.Value = "" Then
        residentName = combineName(Me.firstNameTxt.Value, Me.lastNameTxt.Value)
        
        wingsCbo = Me.wingsCbo.Value
        
        db.insertResidentName residentName, wingsCbo
        residentList.Range("A" & last_row).Value = residentName
        residentList.Range("A" & last_row).Select
        Unload Me
    Else
    MsgBox "Please enter correct first name and last name"
    End If
End Sub


Private Sub editResidentNameBtn_Click()
    Dim newName As String
    Dim residentName As String
    Dim db As New residentDb
    
    
    residentName = Trim(ActiveCell.Value)
    
    If Not Me.editResidentNametxt.Value = "" Then
        newName = Me.editResidentNametxt.Value
        db.updateResidentName residentName, newName
        ActiveCell.Value = newName
    Else
        MsgBox "Please enter the new name you enter"
        Me.editResidentNametxt.BackColor = RGB(255, 255, 0)
    End If
    
End Sub

Private Sub UserForm_Initialize()
    ui_init
    With Me.actionCbo
        .AddItem "Delete Resident"
        .AddItem "Edit Resident"
        .AddItem "Add Resident"
    End With
End Sub

Private Sub ui_init()
    With Me
        .deleteFrame.BackColor = RGB(255, 255, 255)
        .addResidentframe.BackColor = RGB(255, 255, 255)
        .editResidentframe.BackColor = RGB(255, 255, 255)
        .btnAddResident.BackColor = RGB(178, 236, 93)
        .btnAddResident.Font.Size = 14
        .btnAddResident.Font.Bold = True
        .btnAddResident.ForeColor = RGB(255, 255, 255)
        .btnDeleteResident.BackColor = RGB(178, 34, 34)
        .btnDeleteResident.Font.Size = 14
        .btnDeleteResident.Font.Bold = True
        .btnDeleteResident.ForeColor = RGB(255, 255, 255)
        .editResidentNameBtn.BackColor = RGB(255, 225, 53)
        .editResidentNameBtn.Font.Size = 14
        .editResidentNameBtn.Font.Bold = True
        .editResidentNameBtn.ForeColor = RGB(255, 255, 255)
    End With
    
    With Me.wingsCbo
        .AddItem "FREEDOM"
        .AddItem "LIBERTY"
        .AddItem "EAGLE"
        .AddItem "INDEPENDENCE"
        .AddItem "OLD GLORY"
    End With
End Sub



Private Sub deleteResident(ByVal residentName As String)
    Dim db As New residentDb
    
    db.deleteResident (Trim(residentName))
End Sub
