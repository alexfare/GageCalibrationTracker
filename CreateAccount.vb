Dim R As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify

Private Sub btnCreate_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Credentials" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    If IsError(Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), Ws.Columns(1), 0)) Then
  
    Dim lLastRow As Long    ' lLastRow = variable to store the result of the row count calculation
    lLastRow = Ws.ListObjects.Item(1).ListRows.Count
    R = lLastRow + 2 ' Add number for every header tab created
    
                Dim gnString As String
                    If IsNumeric(inputUser) Then
                        gnString = Val(inputUser.Value)
                    Else
                        gnString = inputUser
                    End If
    
    Ws.Cells(R, "A") = gnString
    Ws.Cells(R, "B") = inputPass
    
    btnCreate.Caption = "Created!" ' change caption of add button for confirmation
    Application.Wait (Now + TimeValue("0:00:02")) ' Wait to avoid crash
    btnCreate.Caption = "Create"
    Clear_Form
    inputUser.SetFocus
    Unload CreateAccount
    AdminForm.Show
    Else
        ErrMsg_Duplicate
    End If
    
End Sub


Sub ErrMsg()
MsgBox ("Username Not Found"), , "Not Found"
inputUser.SetFocus
End Sub

Sub ErrMsg_Duplicate()
MsgBox ("Username Taken"), , "Duplicate"
inputUser.SetFocus
End Sub

Private Sub Clear_Form()
        inputUser = ""
        inputPass = ""
End Sub

Private Sub btnBack_click()
Unload CreateAccount
AdminForm.Show
End Sub



