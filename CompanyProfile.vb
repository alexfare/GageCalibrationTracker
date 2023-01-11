Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet

'/Positioning /'
Private Sub UserForm_Initialize()
Dim sngLeft As Single
Dim sngTop As Single

    Call ReturnPosition_CenterScreen(Me.Height, Me.Width, sngLeft, sngTop)
    Me.Left = sngLeft
    Me.Top = sngTop
	
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
    Application.Wait (Now + TimeValue("0:00:02")) ' Wait to avoid crash
    btnSubmit.Caption = "Update"
End Sub

