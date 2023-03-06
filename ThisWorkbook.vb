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
     Ws.Range("B52") = "" 'Clears logged in user.
     Ws.Range("B55") = "1" 'Clears Persistent Login From Last Session
     
     Dim requireLogin As Integer
     requireLogin = Ws.Range("B59")
     If requireLogin = "1" Then
        Worksheets("Login").Activate
        UserLogin.Show
     Else
        Worksheets("CreatedByAlexFare").Activate
        Menu.Show
    End If
    
    Dim skipDevMode As Integer
    skipDevMode = Ws.Range("B61")
    If skipDevMode = "1" Then
        Application.DisplayFullScreen = False
    Else
        Application.DisplayFullScreen = True
    End If

End Sub

