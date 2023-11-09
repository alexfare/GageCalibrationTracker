VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SuperAdminLogin 
   Caption         =   "Super Admin Login"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2850
   OleObjectBlob   =   "SuperAdminLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SuperAdminLogin"
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
End Sub

Private Sub inputPass_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnLogin_Click
    End If
End Sub

Private Sub btnLogin_Click()
    ' Check if the user provided input
    If inputPass <> "" Then
    Login_Sub
    Else
        Failed_Login
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

    b64 = True        'output base-64
    
    sH = SHA512(sIn, b64)
    
    savePass = sH
    '/ End Hash /'
    
    Dim Password As Variant

    Password = "5a6WKkpPucxU75yOvrlND6xY549SrkucxhEg+SukLGzG4pdyY5I1X+51fP5BpkMC1RwXMRw9VZTFXXpXcWeemQ=="
    PassCompare = savePass
    
    If Password <> PassCompare Then
        Failed_Login
    Exit Sub
    End If
    
    If Password = PassCompare Then
        Unload SuperAdminLogin
        SuperAdminMenu.Show
    End If
End Sub

Private Sub btnBack_click()
    Unload SuperAdminLogin
    AdminForm.Show
End Sub

Private Sub Failed_Login()
    MsgBox "Login Failed, Wrong Username Or Password.", vbInformation, "Failed Login"
End Sub

