VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LabelSetUp 
   Caption         =   "Label Printer Set Up"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LabelSetUp.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "LabelSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet

'/Positioning /'
Private Sub UserForm_Initialize()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    inputSmallLabel = Ws.Range("C26")
    inputLargeLabel = Ws.Range("C27")
End Sub

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnBack_click()
    Unload Me
    Label.Show
End Sub

Private Sub btnSubmit_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    Ws.Range("C26") = inputSmallLabel
    Ws.Range("C27") = inputLargeLabel
    
    btnSubmit.Caption = "Updated!" ' change caption of add button for confirmation
    Application.Wait (Now + TimeValue("0:00:01")) ' Wait to avoid crash
    btnSubmit.Caption = "Update"
End Sub


