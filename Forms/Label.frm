VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Label 
   Caption         =   "Print Labels"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   OleObjectBlob   =   "Label.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet

'/Positioning /'
Private Sub UserForm_Initialize()
Dim sngLeft As Single
Dim sngTop As Single

    Call ReturnPosition_CenterScreen(Me.Height, Me.Width, sngLeft, sngTop)
    Me.Left = sngLeft
    Me.Top = sngTop
End Sub

Private Sub btnLargeLabel_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    Dim x As Variant
    Dim Path As String
        Path = Ws.Range("C27")
        'MsgBox (Path) 'Confirms the path works
        x = Shell("explorer.exe " + Path, vbNormalFocus) 'explorer.exe is needed due to vba expecting a .exe
        Unload Me
End Sub

Private Sub btnSmallLabel_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    Dim x As Variant
    Dim Path As String
        Path = Ws.Range("C26")
        'MsgBox (Path) 'Confirms the path works
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



