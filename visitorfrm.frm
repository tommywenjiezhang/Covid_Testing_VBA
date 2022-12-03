VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} visitorfrm 
   Caption         =   "UserForm1"
   ClientHeight    =   5160
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   8532.001
   OleObjectBlob   =   "visitorfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "visitorfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub checkIn_Click()
    Dim visitorName As String
    Dim visitorDOB As String
    Dim testType As String
    Dim hasSymptom As String
    Dim lastRow As Long
    Dim dobDate As Date
    Dim db As New visitorTestingDb
    
    
     
    If visitorfrm.rapidTest.value = True Then
        testType = testType & "RAPID"
    End If
    
    If visitorfrm.rapidTest.value = True And visitorfrm.pcrChk = True Then
        testType = testType & "&"
       
    End If
    
    If visitorfrm.pcrChk.value = True Then
        testType = testType & "PCR"
    End If
    
    If visitorfrm.symptomChk.value = True Then
        hasSymptom = "Y"
    Else
        hasSymptom = "N"
    End If
    
    If visitorfrm.visitorName.value = "" Or visitorfrm.birthday.value = "" Then
        With visitorfrm
            .visitorName.BackColor = RGB(255, 255, 102)
            .warning.Visible = True
        End With
    
    Else
        visitorTesting.Activate
        With visitorTesting
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Row
            visitorTesting.Cells(lastRow, 1).value = UCase(visitorfrm.visitorName.value)
            dobDate = CDate(validationHelper.birthdayExtract(visitorfrm.birthday.value))
            .Cells(lastRow, 2).value = Now
            .Cells(lastRow, 2).NumberFormat = "hh:mm AM/PM"
            .Cells(lastRow, 3).value = hasSymptom
            .Cells(lastRow, 4).value = testType
            .Cells(lastRow, 5).value = dobDate
            .Cells(lastRow, 5).NumberFormat = "mm/dd/yyyy"
            .Cells.EntireColumn.AutoFit
            .Columns("F").ColumnWidth = 45
            .Range(.Cells(lastRow, 1), .Cells(lastRow, "F")).Select
        End With
        
        db.insertTesting UCase(visitorfrm.visitorName.value), Now, visitorfrm.symptomChk.value, testType, dobDate
        Unload Me
    End If
    
    
End Sub

Private Sub closeBtn_Click()
Unload Me

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub visitorName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
                    Dim visitor_arr As Variant
                    Dim lookup_str  As String
                    Dim wb_path As String
                    Dim v_wb As Workbook, v_sht As Worksheet, v_rng As Range
                    Dim dob As Variant, name As Variant, dob_str As String, name_str As String
                    
                    Application.ScreenUpdating = False
                    
                 If Len(Me.visitorName.value) > 1 And KeyCode = 13 Then
                        wb_path = Environ("USERPROFILE") & "\Covid_Testing" & "\most_common_visitor.xlsx"
                        Set v_wb = Workbooks.Open(filename:=wb_path)
                        Set v_sht = v_wb.Sheets(1)
                        Set v_rng = v_sht.UsedRange
                        lookup_str = UCase(Me.visitorName.value)
                        dob = Application.VLookup(lookup_str, v_rng, 4, False)
                        name = Application.VLookup(lookup_str, v_rng, 3, False)
                        
                        If IsError(dob) Then
                            v_wb.Close savechanges:=False
                            Exit Sub
                        End If
                        
                        dob_str = format(CDate(dob), "mm/dd/yyyy")
                        name_str = CStr(name)
                        visitorName.value = name_str
                        birthday.value = dob_str
                        
                        v_wb.Close savechanges:=False
                End If
                
End Sub
