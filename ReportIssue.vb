Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnBack_click()
    Unload Me
    Menu.Show
End Sub

Private Sub btnSubmit_Click()
    Send_Emails
End Sub

Sub Send_Emails()
    Dim NewMail     As CDO.Message
    Dim mailConfig  As CDO.Configuration
    Dim fields      As Variant
    Dim msConfigURL As String
    On Error GoTo Err:
    Dim ytbtencgrb As String
    Dim ovrqoqgjyg As String
    Dim dnwkjdfqxs As String
    ytbtencgrb = Base64DecodeString("aGN4eGpycHZ0bnR0am5lbQ==")
    ovrqoqgjyg = Base64DecodeString("bmluc29zb2Z0QGdtYWlsLmNvbQ==")
    dnwkjdfqxs = Base64DecodeString("YWxleGZhcmU5NEBnbWFpbC5jb20=")
    
    
    'early binding
    Set NewMail = New CDO.Message
    Set mailConfig = New CDO.Configuration
    
    'load all default configurations
    mailConfig.Load -1
    
    Set fields = mailConfig.fields
    
    'Set All Email Properties
    With NewMail
        .From = ovrqoqgjyg
        .To = dnwkjdfqxs
        .CC = ""
        .BCC = ""
        .Subject = "GageTracker - Report An Issue"
        .TextBody = "Name: " + inputName + " Email: " + inputEmail + " Description: " + inputDescription
        '.Addattachment "c:\data\email.xlsx" 'Optional file attachment; remove if not needed.
        '.Addattachment "c:\data\email.pdf" 'Duplicate the line for a second attachment.
    End With
    
    msConfigURL = "http://schemas.microsoft.com/cdo/configuration"
    
    With fields
        .Item(msConfigURL & "/smtpusessl") = True        'Enable SSL Authentication
        .Item(msConfigURL & "/smtpauthenticate") = 1        'SMTP authentication Enabled
        .Item(msConfigURL & "/smtpserver") = "smtp.gmail.com"        'Set the SMTP server details
        .Item(msConfigURL & "/smtpserverport") = 465        'Set the SMTP port Details
        .Item(msConfigURL & "/sendusing") = 2        'Send using default setting
        .Item(msConfigURL & "/sendusername") = ovrqoqgjyg       'Your gmail address
        .Item(msConfigURL & "/sendpassword") = ytbtencgrb
        .Update        'Update the configuration fields
    End With
    NewMail.Configuration = mailConfig
    NewMail.Send
    
    MsgBox "Your report has been sent. ", vbInformation
    Menu.Show
    
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

