VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Admin Login"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3870
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/Admin Panel Login /'

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'

inputUser.SetFocus
End Sub

Private Sub inputPass_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnLogin_Click
    End If
End Sub

Private Sub btnLogin_Click()
    ' Check if the user provided input
    If inputUser <> "" And inputPass <> "" Then
        Login_Sub
    Else
        Err_Blank
    End If
End Sub

Private Sub Login_Sub()
    '/ Hash /'
    s = inputPass
    
    Dim sIn As String, sOut As String, b64 As Boolean
    Dim sH As String, sSecret As String
    
    'Password to be converted
    sIn = s
    sSecret = "G4g3Tr4ck3r"        'secret key for StrToSHA512Salt only
    
    'select         as required
    'b64 = False   'output hex
    b64 = True        'output base-64
    
    sH = SHA512(sIn, b64)
    'sH = StrToSHA512Salt(sIn, sSecretKey, b64) 'Add salt to the encryption
    
    'message box and immediate window outputs
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
    
    savePass = sH
    '/ Hash /'
    
    Dim inputUsername As String
    Dim Password As Variant
    
    inputUsername = inputUser.Value
    
    Dim searchValue As String
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim une As Boolean
    
    'Set the worksheet to search in
    Set ws = ThisWorkbook.Worksheets("Credentials")
    une = False
    
    'Set the value to search for
    searchValue = inputUsername
    
    'Get the last row in Column A
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through each cell in Column A and check if the value matches the search value
    For i = 1 To lastRow
        If ws.Cells(i, "A").Value = searchValue Then
            une = True
        End If
    Next i
    If une = False Then
        Failed_Login
    Exit Sub
    End If
        
    Password = Application.WorksheetFunction.VLookup(inputUsername, Sheets("Credentials").Range("A:B"), 2, 0)
    PassCompare = savePass
    
    If Password <> PassCompare Then
        Failed_Login
    Exit Sub
    End If
    
    If Password = PassCompare Then
        '/Add to the login count /'
        Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
        Dim LoginCount As Integer
        
        Dim List_Select
        List_Select = "Admin"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        
        LoginCount = ws.Range("B48")
        LoginCountPlusOne = LoginCount + 1
        ws.Range("B48") = LoginCountPlusOne
        ws.Range("B52") = inputUsername
        ws.Range("B55") = True
        
        Unload LoginForm
        Sheets("CreatedByAlexFare").Activate
        AdminForm.Show
    End If
End Sub

Private Sub btnBack_click()
    Unload LoginForm
    Menu.Show
End Sub

Private Sub SALogin_Click()
    Unload LoginForm
    SuperAdminLogin.Show
End Sub

Private Sub Failed_Login()
    MsgBox "Login Failed, Wrong Username Or Password.", vbInformation, "Failed Login"
End Sub

Private Sub Err_Blank()
    MsgBox "Username or Password cannot be blank.", vbInformation, "Failed Login"
End Sub
