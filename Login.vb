'Login - Created by Alex Fare
'Version - 1.1.0
'Updated - 12/**/2022
'
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
MsgBox "Login Failed, Wrong Password or Username.", vbInformation, ""
End If

End Sub

' Change Login to be Admin Panel only
