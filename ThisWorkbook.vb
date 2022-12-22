Private Sub Workbook_Open()

'/ Require Login to open /
    'Worksheets("Login").Activate
    'LoginForm.Show
    
'/ Skip Login /
    Worksheets("CreatedByAlexFare").Activate
    Application.DisplayFullScreen = True
     'MsgBox ""
    Menu.Show
    
End Sub


