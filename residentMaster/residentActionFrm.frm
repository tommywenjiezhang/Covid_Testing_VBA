VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} residentActionFrm 
   Caption         =   "Make Change"
   ClientHeight    =   11592
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   7536
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
    
        
            Dim choice As String
             If Not Me.actionCbo.value = "" Then
                 choice = Me.actionCbo.value
                 Select Case choice
               Case "Delete Resident"
                Me.headerLbl.Caption = Me.headerLbl.Caption & " for " & ActiveCell.value
                 Me.deleteFrame.Visible = True
                 Me.editResidentframe.Visible = False
                 Me.addResidentframe.Visible = False
               Case "Edit Resident"
                Me.headerLbl.Caption = Me.headerLbl.Caption & " for " & ActiveCell.value
                  Me.editResidentframe.Visible = True
                  Me.deleteFrame.Visible = False
                  Me.addResidentframe.Visible = False
                  Me.editResidentNametxt.FontSize = 16
                  Me.editResidentNametxt.value = ActiveCell.value
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
    Dim residentID As String
    
    Dim answer As Integer
 
    answer = MsgBox("Are you sure you want to delete resident", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Resident")
    
    If answer = vbYes Then
        residentID = residentList.Range("A" & ActiveCell.Row)
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
    Dim wingsStr As String
    Dim db As New residentDb
    Dim last_row As Integer
    Dim idx As Integer
    Dim dobstr As String
    Dim dobDate As Date
    Dim roomStr As String
    
    
    
    
    With residentList
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    End With
    
    If Not Me.firstNameTxt.value = "" And Not Me.lastNameTxt.value = "" And Not Me.wingsCbo.value = "" Then
        residentName = combineName(Me.firstNameTxt.value, Me.lastNameTxt.value)
        dobstr = validationHelper.birthdayExtract(Me.dobTxt.value)
        dobDate = CDate(dobstr)
        roomStr = Me.roomNumberTxt.value
        wingsStr = Me.wingsCbo.value
        
        db.insertResidents residentName, dobDate, wingsStr, roomStr
        importName wingsStr
        Unload Me
    Else
    MsgBox "Please enter correct first name and last name"
    End If
End Sub


Private Sub editResidentNameBtn_Click()
    Dim newName As String
    Dim residentName As String
    Dim db As New residentDb
    
    
    residentName = Trim(ActiveCell.value)
    
    If Not Me.editResidentNametxt.value = "" Then
        newName = Me.editResidentNametxt.value
        db.updateResidentName residentName, newName
        ActiveCell.value = newName
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
