Attribute VB_Name = "CheckInResident"
Option Explicit

Sub checkIn()
    
'Find the last used row in a Column: column A in this example
    Dim util As New testUtil
    
    If ActiveCell.value = "" Then
        MsgBox "No Person selected exiting........."
        Exit Sub
    Else
        If Not (util.InRange(ActiveCell, residentList.Range("A2:B1000"))) Then
            MsgBox "Selecting Wrong Area please select under Resident name........."
            Exit Sub
        Else
            testForm.Show
        End If
    End If
End Sub

