VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} testForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4944
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   6900
   OleObjectBlob   =   "testForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "testForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub checkIn_Click()
    Dim lastRow As Long
    Dim testType As String
    testType = ""
    Dim hasSymptom As String
    Dim db As New residentDb
    Dim pcr As Boolean
    Dim rapid As Boolean
    Dim name As String
    Dim dobstr As String
    Dim dobDate As Date
    Dim lotNumber As String
    Dim expirationDateStr As String
    Dim expirationDate As Date
    Dim wingName As String
    Dim residentName As String
    Dim residentID As String
    Dim testKind As String
    Dim testChoice As String
    Dim testReason As String
    
    
    Dim select_rng As Range
        
        Set select_rng = Selection
        Dim idx As Long
        
For Each cell In select_rng
         
        With residentList
        .Unprotect
            lotNumber = .Range("D1").value
            .Range("D2").NumberFormat = "mm/dd/yyyy"
            expirationDateStr = .Range("D2").value
            
            residentName = .Range("B" & cell.Row).value
            residentID = .Range("A" & cell.Row).value
            wingName = .Range("D3").value
            testKind = .Range("D4").value
            testReason = .Range("D5").value
            
            
            
            If Len(lotNumber) = 0 Then
                MsgBox ("Lot Number is not entered")
                lotNumber = InputBox("Please enter the lot number")
            End If
            
            If Len(expirationDate) = 0 Or Not IsDate(expirationDateStr) Then
                MsgBox ("Expiration Date is not entered")
                expirationDateStr = InputBox("Please enter the expiration Date")
            End If
            
            
                expirationDateStr = validationHelper.birthdayExtract(expirationDateStr)
                expirationDate = CDate(expirationDateStr)
                
            If Len(testKind) = 0 Then
                MsgBox ("The test is not entered ")
                testChoice = InputBox("Please enter follow test " & vbNewLine & "1- BinaxNow" & vbNewLine & "2- QuickVue")
                If CLng(testChoice) = 1 Then
                    testKind = "BinaxNow"
                Else
                    testKind = "QuickVue"
                End If
            End If
            
            If Len(testReason) = 0 Then
                testReason = "Routine"
            End If
            .Protect
    End With
    
    
    With testRoster
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
            .Range("A" & lastRow & ":Q" & .Rows.Count).ClearContents
    End With
    
    If select_rng.Count > 1 Then
        populate_birthday (cell.Row)
    End If
  
    
     If Not testForm.dobTxt.value = "" Then
        If IsDate(testForm.dobTxt.value) Then
            dobstr = validationHelper.birthdayExtract(testForm.dobTxt.value)
            dobDate = CDate(dobstr)
        End If
        'db.updateBirthday dobDate, Trim(name)
    End If
    
    
    If testForm.rapidTest.value = True And testForm.pcrChk = True Then
        
        insertTest lastRow, residentID:=residentID, residentName:=residentName, _
        wing:=wingName, dobDate:=dobDate, _
        lotNumber:=lotNumber, expirationDate:=expirationDate, testType:="RAPID", testKind:=testKind, testReason:=testReason
        
        
        insertTest (lastRow + 1), residentID:=residentID, residentName:=residentName, _
        wing:=wingName, dobDate:=dobDate, lotNumber:=lotNumber, expirationDate:=expirationDate, testType:="PCR", testKind:=testKind, testReason:=testReason
       
       
    ElseIf testForm.pcrChk.value = True Then
        
        
    
        insertTest lastRow, residentID:=residentID, residentName:=residentName, _
        wing:=wingName, dobDate:=dobDate, lotNumber:=lotNumber, expirationDate:=expirationDate, testType:="PCR", testKind:=testKind, testReason:=testReason
        
    ElseIf testForm.rapidTest.value = True Then
    
        insertTest lastRow, residentID:=residentID, residentName:=residentName, _
        wing:=wingName, dobDate:=dobDate, lotNumber:=lotNumber, expirationDate:=expirationDate, testType:="RAPID", testKind:=testKind, testReason:=testReason
    End If
    
    testRoster.Cells.EntireColumn.AutoFit
    testRoster.Select
    testRoster.Cells(lastRow, "A").Select
    

Next cell

done:

    Unload Me
    Exit Sub
    
wrong_dob:
    With testForm.dobTxt
        .SetFocus
        .BackColor = RGB(255, 255, 153)
    End With
    testForm.warning.Visible = True
    testForm.warning.Caption = "Please enter correct DOB"
    
End Sub




Private Sub insertTest(ByVal lastRow As Long, _
                                        residentID As String, testType As String, residentName As String, _
                                        wing As String, dobDate As Date, lotNumber As String, expirationDate As Date, testKind As String, testReason As String)
                                        
        Dim db As New residentDb
        Dim hasSymptom As String
        
          If testForm.symptomChk.value = True Then
                hasSymptom = "Y"
            Else
                hasSymptom = "N"
            End If
    
        With testRoster
                 .Cells(lastRow, "A").value = residentID
                .Cells(lastRow, "B").value = residentName
                .Cells(lastRow, "C").value = wing
                .Cells(lastRow, "D").value = Now
                .Cells(lastRow, "D").NumberFormat = "hh:mm:ss AM/PM"
                .Range("E" & lastRow).value = format(dobDate, "mm/dd/yyyy")
                .Range("F" & lastRow).value = hasSymptom
                .Range("G" & lastRow).value = testType
                .Range("H" & lastRow).value = lotNumber
                .Range("I" & lastRow).value = format(expirationDate, "mm/dd/yyyy")
                .Range("I" & lastRow).NumberFormat = "mm/dd/yyyy"
                .Range("J" & lastRow).value = testKind
                .Range("K" & lastRow).value = testReason
        End With
        
        db.insertTesting residentID:=residentID, residentName:=residentName, _
        residentWings:=wing, lotNumber:=lotNumber, expirationDate:=expirationDate, typeOfTest:=testType, timeTested:=Now, symptom:=Me.symptomChk.value, testKind:=testKind, testReason:=testReason
        
       
End Sub

Private Sub closeBtn_Click()
Unload Me

End Sub



Private Sub populate_birthday(ByVal c_row As Long)
    Dim last_row As Long
    Dim value As Variant
    Dim dob As String
    
    
    
    
    Dim residentName As String
    residentName = residentList.Range("A" & c_row).value
    
    With ResidentInfo
        last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        value = Application.VLookup(residentName, .Range("A1:B" & last_row), 2, False)
        If Not IsError(value) Then
            dob = format(CDate(value), "mm/dd/yyyy")
            testForm.dobTxt.value = dob
        End If
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim residentName As String
    Dim wing As String
    residentName = residentList.Range("A" & ActiveCell.Row).value
    wing = residentList.Range("B" & ActiveCell.Row).value
    
    Me.empNameLal.Caption = "Testing for: " & residentName
    
    On Error Resume Next
    populate_birthday (ActiveCell.Row)
End Sub
