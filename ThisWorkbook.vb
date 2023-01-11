Private Sub Workbook_Open()
    Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
    Dim WorkBookCount As Integer

    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws

     WorkBookCount = Ws.Range("B47")
     WorkBookPlusOne = WorkBookCount + 1
     Ws.Range("B47") = WorkBookPlusOne 

'/ Require Login to open /
    'Worksheets("Login").Activate
    'LoginForm.Show
    
'/ Skip Login /
    Worksheets("CreatedByAlexFare").Activate
    Application.DisplayFullScreen = True
     'MsgBox ""
    Menu.Show
    
End Sub