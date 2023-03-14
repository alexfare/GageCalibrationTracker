Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet

'/Positioning /'
Private Sub UserForm_Initialize()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    inputCName = Ws.Range("B2")
    inputCPhone = Ws.Range("B3")
    inputCAddress = Ws.Range("B4")
    inputCWebsite = Ws.Range("B5")
End Sub

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnBack_click()
    Unload Me
End Sub

Private Sub btnSubmit_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    Ws.Range("B2") = inputCName
    Ws.Range("B3") = inputCPhone
    Ws.Range("B4") = inputCAddress
    Ws.Range("B5") = inputCWebsite
    
    btnSubmit.Caption = "Updated!" ' change caption of add button for confirmation
    Application.Wait (Now + TimeValue("0:00:01")) ' Wait to avoid crash
    btnSubmit.Caption = "Update"
End Sub


