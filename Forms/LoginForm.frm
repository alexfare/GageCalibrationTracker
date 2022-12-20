VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3825
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLogin_Click()

'/ Hash /'
    s = inputPass
    
    Dim sIn As String, sOut As String, b64 As Boolean
    Dim sH As String, sSecret As String
    
    'Password to be converted
    sIn = s
    sSecret = "" 'secret key for StrToSHA512Salt only
    
    'select as required
    'b64 = False   'output hex
    b64 = True   'output base-64
    
    sH = SHA512(sIn, b64)
    
    'message box and immediate window outputs
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
    ' MsgBox sH & vbNewLine & Len(sH) & " characters in length"
  savePass = sH
'/ Hash /'





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
PassCompare = savePass

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



