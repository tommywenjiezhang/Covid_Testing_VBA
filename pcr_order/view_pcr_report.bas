Attribute VB_Name = "view_pcr_report"
Sub view_pcr_report(ByVal collectionDate As String)
    Dim name As String
    Dim execute_str As String
    
    Dim util As New testUtil
    
    
    Dim path As String
    path = util.getDriveName()
    name = ActiveCell.Value

    If ActiveCell.Value = "" Then
        MsgBox "No Person selected exiting........."
        Exit Sub
    Else
        If Not (util.InRange(ActiveCell, Range("A2:B1000"))) Then
            MsgBox "Selecting Wrong Area please select under empolyee name........."
            Exit Sub
        Else
                execute_str = path & "\programs\python\python -i " & path & "\programs\automateTesting\printLabel.py " & _
                "--name " & name
                obj = Shell(execute_str, vbMinimizedFocus)
        End If
    End If
End Sub


