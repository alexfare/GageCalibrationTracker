'Login - Created by Alex Fare
'Version - 1.2.0
'Updated - 12/12/2022
'
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

If inputUser.Value = "Admin" And inputPass.Value = "Admin" Then
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

' Change Login to be Admin Panel only



