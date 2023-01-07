Attribute VB_Name = "openFile"
Function SelectFolder() As String
    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        SelectFolder = sFolder
    End If
End Function


Function SelectFile() As String
    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        SelectFile = sFolder
    End If
End Function

Sub updateEmpDatabase()
    Dim fileName As String
    fileName = SelectFile()
    

End Sub

Sub updateDatabase(ByVal wingName As String)
Dim folderName As String
Dim fileName As String

Dim name As String

fileName = SelectFolder() & "\" & StrConv(wingName, vbProperCase) & ".xlsx"

If fs.FileExists(fileName) Then
    openWorkbook fileName, wingName
Else
    MsgBox wingName & " can't be found. Please manually select the file"
    fileName = SelectFile()
    openWorkbook fileName, wingName
    
End If

    
    

End Sub


Sub openWorkbook(ByVal fileName As String, wingName As String)
    Dim wb As Workbook
    Dim lastRow As Long
    Dim copy_rng As Range
    Dim db As New residentDb
    Dim residentName As String
    
    
    Set wb = Workbooks.Open(fileName)
    db.deleteResidentByWing (wingName)
    With wb.Sheets(1)
        lastRow = .Cells(.Rows.Count, 2).End(xlUp).Row
        Set copy_rng = .Range("B3:B" & lastRow)
        For Each cl In copy_rng
            If cl.value <> "" And InStr(cl.value, ",") Then
                If InStr(cl.value, "DNR") Then
                    residentName = Left(cl.value, InStr(cl.value, "DNR") - 1)
                Else
                    residentName = cl.value
                End If
                    db.insertResidentName Trim(residentName), wingName
            End If
            
        Next
    End With
End Sub
