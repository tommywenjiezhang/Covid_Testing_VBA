VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} printLabelFrm 
   Caption         =   "Print PCR Form"
   ClientHeight    =   1350
   ClientLeft      =   192
   ClientTop       =   810
   ClientWidth     =   5988
   OleObjectBlob   =   "printLabelFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "printLabelFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnSumit_Click()
    Dim collection_date_str As String
    
    If Not Me.collectionDateTxt.value = "" Then
        collection_date_str = validationHelper.birthdayExtract(Me.collectionDateTxt.value)
        printLabel (collection_date_str)
        Unload Me
    Else
         Me.collectionDateTxt.BackColor = RGB(255, 255, 0)
    End If
        
End Sub

Private Sub UserForm_Initialize()
  Me.collectionDateTxt.value = format(Date, "mm/dd/yyyy")
End Sub



Private Sub printLabel(ByVal collectionDate As String)
    Dim name As String
    Dim execute_str As String
    
    Dim util As New testUtil
    
    
    Dim path As String
    name = ActiveCell.value

    If ActiveCell.value = "" Then
        MsgBox "No Person selected exiting........."
        Exit Sub
    Else
        If Not (util.InRange(ActiveCell, Range("B2:B1000"))) Then
            MsgBox "Selecting Wrong Area please select under empolyee name........."
            Exit Sub
        Else
                execute_str = "printLabel " & _
                "--name " & Chr(34) & name & Chr(34) & " --date " & collectionDate
                Debug.Print execute_str
                Call run_exe.run_exe(execute_str)
                
        End If
    End If
End Sub

