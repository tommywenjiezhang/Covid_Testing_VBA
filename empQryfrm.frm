VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} empQryfrm 
   Caption         =   "UserForm1"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7020
   OleObjectBlob   =   "empQryfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "empQryfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
 Unload Me
End Sub

Private Sub btnSumit_Click()
    Dim startDateStr  As String
    Dim endDateStr As String
    Dim empID As String
    Dim startDate As Date, endDate As Date
    
    
    If Not Me.endDateTxt.value = "" And Not Me.startDateTxt.value = "" Then
        startDateStr = validationHelper.birthdayExtract(Me.startDateTxt.value)
        endDateStr = validationHelper.birthdayExtract(Me.endDateTxt.value)
        
        
        If Not IsError(CDate(startDateStr)) And Not IsError(CDate(endDateStr)) Then
            startDate = CDate(startDateStr)
            endDate = CDate(endDateStr)
            Shell "taskkill /IM ""AcroRd32.exe"" /F"
            empID = empList.Range("A" & ActiveCell.Row).value
            Call exportPDF.exportPDFEmp(startDate, endDate, empID)
        End If
    Else
        Me.endDateTxt.BackColor = RGB(255, 255, 0)
        Me.startDateTxt.BackColor = RGB(255, 255, 0)
    End If
    
    
End Sub


Private Sub UserForm_Initialize()
    Me.endDateTxt.value = format(Date, "mm/dd/yyyy")
End Sub

