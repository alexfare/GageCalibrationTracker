Private Sub Workbook_Open()

'/ Require Login to open /
    'Worksheets("Login").Activate
    'LoginForm.Show
    
'/ Skip Login /
    Worksheets("CreatedByAlexFare").Activate
	'MsgBox _
		'""
    UserForm1.Show
    Application.DisplayFullScreen = True
    
End Sub
