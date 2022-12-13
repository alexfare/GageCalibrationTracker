Dim r As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim btnUpdate_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify


Private Sub inputUser_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
        Insp_Date.SetFocus
    End If
End Sub

Public Sub Search_Button_Click()

' clear previous data from form, except "Gage Number"
' --------------------------------------------------------
        inputPass = ""
        Descriptiontxt = ""
        
' ---------------------------------------------------------

Dim Ws As Worksheet

List_Select = "Credentials"
Set Ws = Sheets(List_Select)
Set Worksheet_Set = Ws




    If IsError(Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), Ws.Columns(1), 0)) Then
            btnUpdate_Enable = False
            ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), Ws.Columns(1), 0)
        GN_Verify = inputUser
        inputPass = Ws.Cells(r, "B")
        btnUpdate_Enable = True
            
            
        Dim FS
        Set FS = CreateObject("Scripting.FileSystemObject")

        If FS.FileExists(TextFile_FullPath) Then
            Else
        End If
    End If

inputUser.SetFocus

End Sub



Private Sub btnUpdate_Click()
If btnUpdate_Enable = True Then
    If GN_Verify = inputUser Then
        Update_Worksheet
    Else
        MSG_Verify_Update
    End If
Else
     MsgBox ("Must search for entry before updating"), , "Nothing to Update"
End If
End Sub



Sub ErrMsg()
MsgBox ("Username Not Found"), , "Not Found"
inputUser.SetFocus
End Sub

Sub ErrMsg_Duplicate()
MsgBox ("Username already in use"), , "Duplicate"
inputUser.SetFocus
End Sub



Private Sub Clear_Form()
        inputUser = ""
        inputPass = ""
End Sub

Private Sub Update_Worksheet()
If btnUpdate_Enable = True Then
Dim gnString As String
Set Ws = Worksheet_Set
    If IsNumeric(inputUser) Then
        gnString = Val(inputUser.Value)
    Else
        gnString = inputUser
    End If
Ws.Cells(r, "A") = gnString
Ws.Cells(r, "B") = inputPass


btnUpdate.Caption = "Updated!"
Application.Wait (Now + TimeValue("0:00:02"))
btnUpdate.Caption = "Update"
'Clear_Form 'Clear form after update
inputUser.SetFocus

Else
    MsgBox ("Must search for entry before updating"), , "Nothing to Update"
    
End If

'btnUpdate_Enable = False 'Remove ' if you want to require searching again after an update.

End Sub

Sub MSG_Verify_Update()

MSG1 = MsgBox("Are you sure you want to change the Gage ID?", vbYesNo, "Verify")

If MSG1 = vbYes Then
  Update_Worksheet
Else
  inputUser = GN_Verify
End If

End Sub

Private Sub btnBack_click()
Unload ChangePassword
AdminForm.Show
End Sub




