'Login - Created by Alex Fare
'Version - 1.3.0
'Updated - 12/13/2022
'
' Updated Default Credentials
' Added multiple logins (Wrong username will cause crash)
' Logging in now redirects to Admin Panel
' Added Failed Login Warning
' Removed unused code
'
'
Private Sub btnLogin_Click()

If inputUser.Value = "" Then
MsgBox "User Cannot be Blank.", vbInformation, ""
Exit Sub
End If

If inputPass.Value = "" Then
MsgBox "Password Cannot be Blank!", vbInformation, ""
Exit Sub
End If

If inputUser.Value = "Admin" And inputPass.Value = "o9!A62sSimZmiHNkQq%3" Then
Unload Me
Sheets("CreatedByAlexFare").Activate
' UserForm1.Show
AdminForm.Show
Else

Dim inputUsername As String
Dim Password As Variant

inputUsername = inputUser.Value
Password = Application.WorksheetFunction.VLookup(inputUsername, Sheets("Credentials").Range("A:B"), 2, 0)

If Password <> inputPass.Value Then
MsgBox "Login Failed, Wrong Password or Username.", vbInformation, "Wrong Password"
Exit Sub
End If

If Password = inputPass.Value Then
Unload Me
Sheets("CreatedByAlexFare").Activate
AdminForm.Show
End If
End If


End Sub


Private Sub btnBack_click()
Unload LoginForm
UserForm1.Show
End Sub


