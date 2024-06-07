Attribute VB_Name = "Updater"
Option Explicit

Sub CheckUpdate()
    Dim url         As String
    url = "https://github.com/alexfare/GageCalibrationTracker/releases/latest"
    ActiveWorkbook.FollowHyperlink url
End Sub

Sub UpdateVersion()
    Dim url         As String
    url = "https://github.com/alexfare/GageTracker/releases/latest"
    ActiveWorkbook.FollowHyperlink url
End Sub

Sub ReportPasswordUpdater()
    On Error GoTo Err
    
    Dim http As Object
    Dim url As String
    Dim data As String
    Dim ws As Worksheet

    ' URL of the text file
    url = "http://faregaming.com/reportpassword.txt"

    ' Create a new HTTP request
    Set http = CreateObject("MSXML2.ServerXMLHTTP")

    ' Send HTTP request
    http.Open "GET", url, False
    http.Send

    ' Check if request was successful
    If http.Status = 200 Then
        ' Get the response text
        data = http.responseText

        ' Assign the data to a cell
        Set ws = ThisWorkbook.Sheets("Admin")
        ws.Range("B69").Value = data
    Else
        'MsgBox "Failed to fetch data. Status code: " & http.Status 'Uncomment for debugging
    End If

    ' Clean up
    Set http = Nothing
    
Err:
    Exit Sub
End Sub
