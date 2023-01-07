Attribute VB_Name = "Module3"
Sub importMain()
    Dim currWk_sht_name As String
    currWk_sht_name = ActiveSheet.CodeName
    
    If Not currWk_sht_name = "empList" And Not currWk_sht_name = "residentList" Then
        MsgBox "Please select the sheet in PCR Order"
        Exit Sub
    Else
        'active sheet is resident call resident controller
        If currWk_sht_name = "residentList" Then
            residentImportFrm.Show
        Else
        'active sheet is employee call employee controller
            empImportFrm.Show
        End If
    End If
End Sub
