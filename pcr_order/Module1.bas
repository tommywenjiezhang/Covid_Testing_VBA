Attribute VB_Name = "Module1"
Sub viewReport()
    Dim fso As Object
    Dim path As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.getDriveName(ThisWorkbook.path)
    Workbooks.Open path & "\programs\automateTesting\RegisterReport.csv", ReadOnly:=True
End Sub

Sub changeLogin()

 UserForm1.Show
End Sub



Sub ListControls()
    Dim lCntr As Long
    Dim aCtrls() As Variant
    Dim ctlLoop As MSForms.control

     'Change UserForm Name In The Next Line
    For Each ctlLoop In residentActionFrm.Controls
        lCntr = lCntr + 1: ReDim Preserve aCtrls(1 To lCntr)
        'Gets Type and name of Control
        aCtrls(lCntr) = TypeName(ctlLoop) & ":" & ctlLoop.name
    Next ctlLoop
     'Change Worksheet Name In The Next Line
    Worksheets(3).Range("A1").Resize(UBound(aCtrls)).Value = Application.Transpose(aCtrls)
End Sub
