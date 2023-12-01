VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SuperAdminPassword 
   Caption         =   "Change SuperAdmin Password"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2895
   OleObjectBlob   =   "SuperAdminPassword.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SuperAdminPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PassMatch As Boolean

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnBack_click()
    Unload SuperAdminPassword
    SuperAdminMenu.Show
End Sub

Private Sub inputPassx2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnUpdate_Click
    End If
End Sub

Private Sub btnUpdate_Click()
    If inputPass <> "" And inputPassx2 <> "" Then
        If inputPass = inputPassx2 Then
            Update_Worksheet
        Else
            MsgBox "Password fields do not match.", vbInformation, "Error"
        End If
    Else
        MsgBox "Password fields cannot be empty.", vbInformation, "Error"
    End If
End Sub

Private Sub Update_Worksheet()
        '/ Hash /'
        s = inputPass
        
        Dim sIn As String, sOut As String, b64 As Boolean
        Dim sH As String, sSecret As String
        
        'Password to be converted
        sIn = s
        sSecret = ""
        
        b64 = True        'output base-64
        
        sH = SHA512(sIn, b64)
        
        'message box and immediate window outputs
        Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
        savePass = sH
        '/ Hash /'
        
        Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
        Dim ws As Worksheet
        Dim List_Select
        List_Select = "Admin" ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws

        ws.Range("B65") = savePass

        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Password Updated!"
        Status
        Clear_Form
End Sub

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

Private Sub Clear_Form()
    inputPass = ""
    inputPassx2 = ""
End Sub
