'Login - Created by Alex Fare
'Version - 1.0.1
'Updated - 12/05/2022
'
'
'
Private Sub btnLogin_Click()

If inputUser.Value = "" Then
MsgBox "User Cannot be Blank.", vbInformation, ""
Exit Sub
End If

If inputPass.Value = "" Then
MsgBox "Password Cannot be Blank!", vbInformation, "Password"
Exit Sub
End If

If inputUser.Value = "Admin" And inputPass.Value = "qwerty" Then
Unload Me
Sheets("CreatedByAlexFare").Activate
End If

UserForm1.Show
End Sub

