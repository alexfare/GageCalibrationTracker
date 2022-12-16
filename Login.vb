Sub HashMD5()
    Dim hashPass As String
        hashPass = inputPass
        Debug.Print StringToMD5Hex(hashPass)

End Sub

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
PassCompare = StringToMD5Hex(outstr)

If Password <> PassCompare Then
MsgBox "Login Failed, Wrong Password or Username.", vbInformation, "Wrong Password"
Exit Sub
End If

If Password = PassCompare Then
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



