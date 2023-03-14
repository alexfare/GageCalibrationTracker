Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub inputPass_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnLogin_Click
    End If
End Sub

Private Sub btnLogin_Click()

'/ Hash /'
    s = inputPass
    
    Dim sIn As String, sOut As String, b64 As Boolean
    Dim sH As String, sSecret As String
    
    'Password to be converted
    sIn = s
    sSecret = "G4g3Tr4ck3r" 'secret key for StrToSHA512Salt only
    
    'select as required
    'b64 = False   'output hex
    b64 = True   'output base-64
    
    sH = SHA512(sIn, b64)
    'Add salt to the encryption
    'sH = StrToSHA512Salt(sIn, sSecretKey, b64)
    
    'message box and immediate window outputs
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"

  savePass = sH
'/ Hash /'

'User set up
If inputUser.Value = "" Then
MsgBox "User Cannot be Blank.", vbInformation, ""
Exit Sub
End If

'Password set up
If inputPass.Value = "" Then
MsgBox "Password Cannot be Blank!", vbInformation, ""
Exit Sub
End If

Dim inputUsername As Integer
Dim Password As Variant

inputUsername = inputUser.Value
Password = Application.WorksheetFunction.VLookup(inputUsername, Sheets("Codes").Range("A:B"), 2, 0)
PassCompare = savePass

If Password <> PassCompare Then
MsgBox "Login Failed, Wrong Password or Username.", vbInformation, "Wrong Password"
Exit Sub
End If

If Password = PassCompare Then
Unload LoginForm
Sheets("CreatedByAlexFare").Activate
AdminForm.Show
End If
End Sub

Private Sub btnBack_click()
Unload LoginForm
Menu.Show
End Sub