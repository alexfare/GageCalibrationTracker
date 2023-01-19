'/Positioning /'
Private Sub UserForm_Initialize()
    Dim sngLeft     As Single
    Dim sngTop      As Single
    
    Call ReturnPosition_CenterScreen(Me.Height, Me.Width, sngLeft, sngTop)
    Me.Left = sngLeft
    Me.Top = sngTop
End Sub

Private Sub inputPass_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnLogin_Click
    End If
End Sub

Private Sub btnLogin_Click()
    
    '/ Hash /'
    s = inputPass
    
    Dim sIn         As String, sOut As String, b64 As Boolean
    Dim sH          As String, sSecret As String
    
    'Password to be converted
    sIn = s
    sSecret = "G4g3Tr4ck3r"        'secret key for StrToSHA512Salt only
    
    'select         as required
    'b64 = False   'output hex
    b64 = TRUE        'output base-64
    
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
    
    Dim inputUsername As String
    Dim Password    As Variant
    
    inputUsername = inputUser.Value
    Password = Application.WorksheetFunction.VLookup(inputUsername, Sheets("Credentials").Range("A:B"), 2, 0)
    PassCompare = savePass
    
    If Password <> PassCompare Then
        MsgBox "Login Failed, Wrong Password Or Username.", vbInformation, "Wrong Password"
        Exit Sub
    End If
    
    If Password = PassCompare Then
        Unload LoginForm
        Sheets("CreatedByAlexFare").Activate
        AdminForm.Show
        
        '/Add to the login count /'
        Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
        Dim LoginCount As Integer
        
        Dim Ws      As Worksheet
        Dim List_Select
        List_Select = "Admin"        ' Tab name
        Set Ws = Sheets(List_Select)
        Set Worksheet_Set = Ws
        
        LoginCount = Ws.Range("B48")
        LoginCountPlusOne = LoginCount + 1
        Ws.Range("B48") = LoginCountPlusOne
        Ws.Range("B52") = inputUsername
        Ws.Range("B55") = "2"
    End If
End Sub

Private Sub btnBack_click()
    Unload LoginForm
    Menu.Show
End Sub