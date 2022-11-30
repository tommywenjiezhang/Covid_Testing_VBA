VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exportForm 
   Caption         =   "Enter your shift"
   ClientHeight    =   20790
   ClientLeft      =   390
   ClientTop       =   1638
   ClientWidth     =   32658
   OleObjectBlob   =   "exportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btnSumbit_Click()
Dim strShift As String
Dim todayDateStr As String
Dim inputDate As Date


If Trim(Me.shiftCombo.value) = "" Then
  Me.shiftCombo.SetFocus
  MsgBox "Please enter Your shift"
End If

    strShift = Me.shiftCombo.value
    
    If Not Me.startDateTxt.value = "" Then
        todayDateStr = validationHelper.birthdayExtract(Me.startDateTxt.value)
        If Not IsError(CDate(todayDateStr)) Then
            inputDate = CDate(todayDateStr)
        Else
            inputDate = Date
        End If
    Else
        inputDate = Date
    End If
    
    If Me.createCopyChk.value = True Then
        Call exportPDF.exportPDF(inputDate)
    Else
         Call exportPDF.exportPDF(inputDate)
    End If
    
    
    
    Me.shiftCombo.value = ""
    Me.shiftCombo.SetFocus

    Unload Me
done:
    ThisWorkbook.Save
    Exit Sub
End Sub

Private Sub closeBtn_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
    shiftCombo.AddItem "Day Shift"
    shiftCombo.AddItem "Night Shift"
    Me.startDateTxt.value = format(Date, "mm/dd/yyyy")
End Sub
