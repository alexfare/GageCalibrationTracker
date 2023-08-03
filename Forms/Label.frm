VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Label 
   Caption         =   "Print Labels"
   ClientHeight    =   2370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3930
   OleObjectBlob   =   "Label.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnLargeLabel_Click()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    Dim x As Variant
    Dim Path As String
        Path = ws.Range("B27")
        x = Shell("explorer.exe " + Path, vbNormalFocus) 'explorer.exe is needed due to vba expecting a .exe
        Unload Me
End Sub

Private Sub btnSmallLabel_Click()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    Dim x As Variant
    Dim Path As String
        Path = ws.Range("B26")
        x = Shell("explorer.exe " + Path, vbNormalFocus) 'explorer.exe is needed due to vba expecting a .exe
        Unload Me
End Sub

Private Sub btnSetUp_Click()
    Unload Me
    LabelSetUp.Show
End Sub

Private Sub btnBack_click()
    Unload Me
End Sub

Private Sub btnCert_Click()
    MsgBox ("Coming Soon!")
End Sub
