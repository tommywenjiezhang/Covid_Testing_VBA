VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} empActionFrm 
   Caption         =   "Make Change"
   ClientHeight    =   8118
   ClientLeft      =   48
   ClientTop       =   390
   ClientWidth     =   7230
   OleObjectBlob   =   "empActionFRM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "empActionFRM"
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
               Case "Delete Employee"
                 Me.deleteFrame.Visible = True
                 Me.editEmpframe.Visible = False
                 Me.addEmployeeframe.Visible = False
               Case "Edit Employee"
                  Me.editEmpframe.Visible = True
                  Me.deleteFrame.Visible = False
                  Me.addEmployeeframe.Visible = False
                  Me.editEmpNametxt.FontSize = 16
                  Me.editEmpNametxt.Value = ActiveCell.Value
               Case "Add Employee"
                 Me.headerLbl.Caption = "Add Employee"
                  Me.addEmployeeframe.Visible = True
                  Me.editEmpframe.Visible = False
                  Me.deleteFrame.Visible = False
               Case Else
                  MsgBox "Unknown Number"
            End Select
            Else
                MsgBox "Please close the from and select a value"
                Me.actionCbo.BackColor = RGB(255, 255, 0)
            End If
        Else
        
            MsgBox "Please select a name"
    End If
    
End Sub

Private Function checkifInrange() As Boolean
        Dim util As New testUtil
        
        If util.InRange(ActiveCell, empList.Range("A:A")) Then
            
            checkifInrange = True
        Else
            checkifInrange = False
        
        End If
End Function

Private Sub btnAddEmp_Click()
    Dim empName As String
    Dim db As New empDb
    Dim last_row As Integer
    
    
    With empList
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    End With
    
    If Not Me.firstNameTxt.Value = "" And Not Me.lastNameTxt.Value = "" Then
        empName = combineName(Me.firstNameTxt.Value, Me.lastNameTxt.Value)
        
        
        db.insertEmpName empName
        empList.Range("A" & last_row).Value = empName
        empList.Range("A" & last_row).Select
        Unload Me
    Else
    MsgBox "Please enter correct first name and last name"
    End If
    
End Sub

Private Sub btnDeleteEmployee_Click()
    Dim empName As String
    
    Dim answer As Integer
 
    answer = MsgBox("Are you sure you want to delete employee", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Employee")
    
    If answer = vbYes Then
         empName = ActiveCell.Value
        deleteEmp empName
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


Private Sub editEmployeeName_Click()
    Dim newName As String
    Dim empName As String
    Dim db As New empDb
    
    
    empName = Trim(ActiveCell.Value)
    
    If Not Me.editEmpNametxt.Value = "" Then
        newName = Me.editEmpNametxt.Value
        db.updateEmpName empName, newName
        ActiveCell.Value = newName
    Else
        MsgBox "Please enter the new name you enter"
        Me.editEmpNametxt.BackColor = RGB(255, 255, 0)
    End If
    
End Sub

Private Sub headerLbl_Click()

End Sub

Private Sub UserForm_Initialize()
    ui_init
    With Me.actionCbo
        .AddItem "Delete Employee"
        .AddItem "Edit Employee"
        .AddItem "Add Employee"
    End With
End Sub

Private Sub ui_init()
    With Me
        .deleteFrame.BackColor = RGB(255, 255, 255)
        .addEmployeeframe.BackColor = RGB(255, 255, 255)
        .editEmpframe.BackColor = RGB(255, 255, 255)
        .btnAddEmp.BackColor = RGB(178, 236, 93)
        .btnAddEmp.Font.Size = 14
        .btnAddEmp.Font.Bold = True
        .btnAddEmp.ForeColor = RGB(255, 255, 255)
        .btnDeleteEmployee.BackColor = RGB(178, 34, 34)
        .btnDeleteEmployee.Font.Size = 14
        .btnDeleteEmployee.Font.Bold = True
        .btnDeleteEmployee.ForeColor = RGB(255, 255, 255)
        .editEmployeeName.BackColor = RGB(255, 225, 53)
        .editEmployeeName.Font.Size = 14
        .editEmployeeName.Font.Bold = True
        .editEmployeeName.ForeColor = RGB(255, 255, 255)
    End With
End Sub


Private Sub deleteEmp(ByVal empName As String)
    Dim db As New empDb
    
    db.deleteEmployee (Trim(empName))
End Sub
