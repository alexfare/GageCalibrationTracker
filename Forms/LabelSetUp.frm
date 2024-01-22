VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LabelSetUp 
   Caption         =   "Label Printer Set Up"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
   OleObjectBlob   =   "LabelSetUp.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "LabelSetUp"
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

    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    inputSmallLabel = ws.Range("B26")
    inputLargeLabel = ws.Range("B27")
End Sub

Private Sub btnBack_click()
    Unload Me
    Label.Show
End Sub

Private Sub btnSubmit_Click()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    ws.Range("B26") = inputSmallLabel
    ws.Range("B27") = inputLargeLabel
    
    '/Status/'
    statusLabel.Caption = "Status:"
    statusLabelLog.Caption = "Updating..."
    Status
End Sub

'/------- Status -------/'
Private Sub Status()
    Dim startTime As Date
    Dim elapsedTime As Long
    Dim waitTimeInSeconds As Long
    
    waitTimeInSeconds = 2 'change this to the desired wait time in seconds
    
    startTime = Now
    Do While elapsedTime < waitTimeInSeconds
        DoEvents 'allow the program to process any pending events
        elapsedTime = DateDiff("s", startTime, Now)
    Loop
        statusLabel.Caption = ""
        statusLabelLog.Caption = ""
End Sub
