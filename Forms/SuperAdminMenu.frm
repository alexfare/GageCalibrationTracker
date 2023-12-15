VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SuperAdminMenu 
   Caption         =   "Super Admin Menu"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   OleObjectBlob   =   "SuperAdminMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SuperAdminMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnAdminWS_Click()
    Unload SuperAdminMenu
    Worksheets("Admin").Activate
End Sub

Private Sub btnCredentialsWS_Click()
    Unload SuperAdminMenu
    Worksheets("Credentials").Activate
End Sub

Private Sub btnListsWS_Click()
    Unload SuperAdminMenu
    Worksheets("Lists").Activate
End Sub

Private Sub btnCustomersWS_Click()
    Unload SuperAdminMenu
    Worksheets("Customers").Activate
End Sub

Private Sub btnGageRRWS_Click()
    Unload SuperAdminMenu
    Worksheets("GageRnR").Activate
End Sub

Private Sub btnGageRRCal_Click()
    Unload SuperAdminMenu
    Worksheets("Calculations").Activate
End Sub

Private Sub btnAudit_Click()
    Unload SuperAdminMenu
    Worksheets("Audit").Activate
End Sub

Private Sub btnBack_click()
    Unload SuperAdminMenu
    AdminForm.Show
End Sub

Private Sub btnSAPass_click()
    Unload SuperAdminMenu
    SuperAdminPassword.Show
End Sub

Private Sub btnPassword_click()
    Dim Worksheet_Set
    Dim ws As Worksheet
    Dim List_Select
    Dim msgBoxPW As String
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    msgBoxPW = Base64DecodeString("UmVwdXJwb3NlNSE=")
    ws.Range("BL1").Value = msgBoxPW
    ws.Range("BL1").Copy
    
    List_Select = "Admin" ' Tab name
    MsgBox "Text copied to clipboard: " & msgBoxPW, vbInformation
End Sub
