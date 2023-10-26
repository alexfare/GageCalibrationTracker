Attribute VB_Name = "DueDateEmailer"
Sub DueDateEmailerSub()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filterCriteria As String
    Dim EmailConfig As Object
    Dim EmailMsg As Object
    Dim FromEmailToken As String
    Dim TokenString As String
    Dim FromEmailSend As String
    Dim ToEmailSend As String
    Dim EmailString As String
    Dim EmailSetPort As String
    EmailSetPort = Base64DecodeString("NDY1")
    EmailString = Base64DecodeString("c210cC5nbWFpbC5jb20=")
    TokenString = Base64DecodeString("ZnhjcCBpbWJtIGhjc2YgbWRsbA==")
    FromEmailToken = TokenString
    FromEmailSend = Base64DecodeString("bmluc29zb2Z0QGdtYWlsLmNvbQ==")
    ToEmailSend = Base64DecodeString("Z2FnZXRyYWNrQGZhcmVnYW1pbmcuY29t")
    
    ' Set the worksheet where the data is located
    Set ws = ThisWorkbook.Sheets("CreatedByAlexFare") ' Change "Sheet1" to the name of your worksheet
    
    ' Determine the last row in column G
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Initialize the filter criteria
    filterCriteria = ""
    
    ' Loop through the data to find rows meeting the criteria
    For i = 2 To lastRow ' Assuming the header is in row 1
        If ws.Cells(i, "G").Value < Date And ws.Cells(i, "G").Value >= Date - 30 Then
            ' If the due date is past today and within 30 days
            filterCriteria = filterCriteria & ws.Cells(i, "A").Value & " - " & ws.Cells(i, "G").Value & vbCrLf
        End If
    Next i
    
    ' For Dev Testing Only
    'If filterCriteria = "" Then
        'MsgBox "No rows match the criteria.", vbInformation
        'Exit Sub
    'End If

    ' Create and configure the email
    Set EmailConfig = CreateObject("CDO.Configuration")
    EmailConfig.Load -1 ' CDO Source Defaults
    EmailConfig.fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailString
    EmailConfig.fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = EmailSetPort
    EmailConfig.fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    EmailConfig.fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    EmailConfig.fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = FromEmailSend
    EmailConfig.fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = FromEmailToken
    EmailConfig.fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EmailConfig.fields.Update

    ' Create and send the email
    Set EmailMsg = CreateObject("CDO.Message")
    With EmailMsg
        Set .Configuration = EmailConfig
        .To = ToEmailSend
        .CC = ""
        .BCC = ""
        .Subject = "Overdue Gage Due Dates"
        .From = FromEmailSend
        .TextBody = "The following Gages are overdue:" & vbCrLf & vbCrLf & filterCriteria
        .Send
    End With

    ' Clean up
    Set EmailMsg = Nothing
    Set EmailConfig = Nothing
End Sub

