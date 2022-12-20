VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportIssue 
   Caption         =   "Report Issue"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ReportIssue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_click()
Unload Me
UserForm1.Show
End Sub

Private Sub btnSubmit_click()
Send_Emails
End Sub

Sub Send_Emails()
 Dim NewMail As CDO.Message
 Dim mailConfig As CDO.Configuration
 Dim fields As Variant
 Dim msConfigURL As String
 On Error GoTo Err:
 
Dim l1llll1l1l1l111ll1l1l1l1ll1 As String
Dim l1l1l1l1ll1l1ll1l11l1l1l1l1 As String
Dim l1l1l1l1l1l1ll1l1l11l1l1l11 As String
Dim l1l1l1ll1l1l1l1ll1l1l1ll1l1 As String
Dim l1l1ll1l11l1l1l1ll1lll1ll11 As String
Dim l1l1l1l1l1l1l11lll11ll1l1l1 As String
Dim l1l1l1l1l1ll111l1l1l1l1ll11 As String
l1llll1l1l1l111ll1l1l1l1ll1 = "reuk1ymy1rchc0we"
l1l1l1l1ll1l1ll1l11l1l1l1l1 = "reukiymyirchc0we"
l1l1l1l1l1l1ll1l1l11l1l1l11 = "revkiymyirchcowe"
l1l1l1ll1l1l1l1ll1l1l1ll1l1 = "reukiymyirchcowe"
l1l1ll1l11l1l1l1ll1lll1ll11 = "revk1ymm1rchoowe"
l1l1l1l1l1l1l11lll11ll1l1l1 = "rvvkimmmirchcowe"
l1l1l1l1l1ll111l1l1l1l1ll11 = "reukiyyiirchcome"

 'early binding
 Set NewMail = New CDO.Message
 Set mailConfig = New CDO.Configuration

 'load all default configurations
 mailConfig.Load -1

 Set fields = mailConfig.fields

 'Set All Email Properties
 With NewMail
   .From = "ninsosoft@gmail.com"
   .To = "alexfare94@gmail.com"
   .CC = ""
   .BCC = ""
   .Subject = "Report An Issue"
   .TextBody = "Name: " + inputName + " Email: " + inputEmail + " Description: " + inputDescription
   '.Addattachment "c:\data\email.xlsx" 'Optional file attachment; remove if not needed.
   '.Addattachment "c:\data\email.pdf" 'Duplicate the line for a second attachment.
 End With

 msConfigURL = "http://schemas.microsoft.com/cdo/configuration"

 With fields
   .Item(msConfigURL & "/smtpusessl") = True 'Enable SSL Authentication
   .Item(msConfigURL & "/smtpauthenticate") = 1 'SMTP authentication Enabled
   .Item(msConfigURL & "/smtpserver") = "smtp.gmail.com" 'Set the SMTP server details
   .Item(msConfigURL & "/smtpserverport") = 465 'Set the SMTP port Details
   .Item(msConfigURL & "/sendusing") = 2 'Send using default setting
   .Item(msConfigURL & "/sendusername") = "ninsosoft@gmail.com" 'Your gmail address
   .Item(msConfigURL & "/sendpassword") = l1l1l1ll1l1l1l1ll1l1l1ll1l1
   .Update 'Update the configuration fields
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
 Case -2147220973 'Could be because of Internet Connection
   MsgBox "Check your internet connection." & vbNewLine & Err.Number & ": " & Err.Description
 Case -2147220975 'Incorrect credentials User ID or password
   MsgBox "Check your login credentials and try again." & vbNewLine & Err.Number & ": " & Err.Description
 Case Else 'Report other errors
   MsgBox "Error encountered while sending email." & vbNewLine & Err.Number & ": " & Err.Description
 End Select

 Resume Exit_Err

End Sub



