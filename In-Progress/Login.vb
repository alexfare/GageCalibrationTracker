Private Sub btnLogin_Click()

If inputUser.Value = "" Then
MsgBox "User Cannot be Blank.", vbInformation, ""
Exit Sub
End If

If inputPass.Value = "" Then
MsgBox "Password Cannot be Blank!", vbInformation, "Password"
Exit Sub
End If

If inputUser.Value = "Admin" And inputPass.Value = "Admin" Then
Unload Me
'Sheets(CreatedByAlexFare).Visible = True
'Sheets(CreatedByAlexFare).Select
Sheets("CreatedByAlexFare").Activate
'ActiveSheet.Range("A1").Select
End If
End If

End Sub

Private Sub inputUsername_Click()

End Sub
