VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportIssue 
   Caption         =   "Report Issue"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   OleObjectBlob   =   "ReportIssue.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ReportIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'

    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    vDisplay = ws.Range("Z1")
    inputName.SetFocus
End Sub

Private Sub btnSubmit_Click()
    ' Check if the user provided input
    If inputName <> "" And inputDescription <> "" Then
        Send_Emails
    Else
        MsgBox "Please provide all the required information.", vbExclamation
    End If
End Sub

Sub Send_Emails()
    Dim NewMail     As CDO.Message
    Dim mailConfig  As CDO.Configuration
    Dim fields      As Variant
    Dim msConfigURL As String
    On Error GoTo Err:
    Dim FromEmailToken As String
    Dim TokenString As String
    Dim FromEmailSend As String
    Dim ToEmailSend As String
    Dim EmailString As String
    Dim EmailSetPort As String
    EmailSetPort = Base64DecodeString("NDY1")
    EmailString = Base64DecodeString("c210cC5nbWFpbC5jb20=")
    TokenString = Base64DecodeString("aGN4eGpycHZ0bnR0am5lbQ==")
    FromEmailToken = TokenString
    FromEmailSend = Base64DecodeString("bmluc29zb2Z0QGdtYWlsLmNvbQ==")
    ToEmailSend = Base64DecodeString("YWxleGZhcmU5NEBnbWFpbC5jb20=")
    
    'Version Number
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim ws          As Worksheet
    Dim List_Select
    List_Select = "CreatedByAlexFare" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    GTVersion = ws.Range("Z1")
    
    'early binding
    Set NewMail = New CDO.Message
    Set mailConfig = New CDO.Configuration
    
    'load all default configurations
    mailConfig.Load -1
    
    Set fields = mailConfig.fields
    
    'Set All Email Properties
    With NewMail
        .From = FromEmailSend
        .To = ToEmailSend
        .CC = ""
        .BCC = ""
        .Subject = "GageTracker - Report An Issue"
        .TextBody = "Name: " + inputName + " | Email: " + inputEmail + " | GageTracker version: " + GTVersion + " | Description: " + inputDescription
    End With
    
    msConfigURL = "http://schemas.microsoft.com/cdo/configuration"
    
    With fields
        .Item(msConfigURL & "/smtpusessl") = True
        .Item(msConfigURL & "/smtpauthenticate") = 1
        .Item(msConfigURL & "/smtpserver") = EmailString
        .Item(msConfigURL & "/smtpserverport") = EmailSetPort
        .Item(msConfigURL & "/sendusing") = 2
        .Item(msConfigURL & "/sendusername") = FromEmailSend
        .Item(msConfigURL & "/sendpassword") = FromEmailToken
        .Update        'Update the configuration fields
    End With
    NewMail.Configuration = mailConfig
    NewMail.Send
    
    MsgBox "Your report has been sent. ", vbInformation
    
Exit_Err:
    'Release object memory
    Set NewMail = Nothing
    Set mailConfig = Nothing
    End
    
Err:
    Select Case Err.Number
        Case -2147220973        'Could be because of Internet Connection
            MsgBox "Check your internet connection." & vbNewLine & Err.Number & ": " & Err.Description
        Case -2147220975        'Incorrect credentials User ID or password
            MsgBox "Check your login credentials And try again." & vbNewLine & Err.Number & ": " & Err.Description
        Case Else        'Report other errors
            MsgBox "Error encountered While sending email." & vbNewLine & Err.Number & ": " & Err.Description
    End Select
    
    Resume Exit_Err

End Sub

Private Sub btnBack_click()
    Unload Me
    Menu.Show
End Sub
