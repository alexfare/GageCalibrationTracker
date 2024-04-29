VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SuperAdminLogin 
   Caption         =   "Super Admin Login"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
   OleObjectBlob   =   "SuperAdminLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SuperAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/Admin Panel Login /'
Dim SAP As String

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'

'/Retrieve password /'
    Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws

    SAP = ws.Range("B65")
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

    Password = SAP
    PassCompare = savePass
    
    If Password <> PassCompare Then
        Failed_Login
    Exit Sub
    End If
    
    If Password = PassCompare Then
        Unload SuperAdminLogin
        SuperAdminMenu.Show
        
        '/Add to the login count /'
        Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
        Dim LoginCount As Integer
        
        Dim List_Select
        List_Select = "Admin"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        
        ws.Range("B64") = True '/True remains logged in & False means login required
    End If
End Sub

Private Sub btnBack_click()
    Unload SuperAdminLogin
    'AdminForm.Show '/Removed for now - Prevents users from accessing Admin Panel without Admin Access
End Sub

Private Sub Failed_Login()
    MsgBox "Login Failed, Wrong Username Or Password.", vbInformation, "Failed Login"
End Sub

