Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet

Private Sub AutoFill_Initialize()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    inputSmallLabel = Ws.Range("C26")
    inputLargeLabel = Ws.Range("C27")
    MsgBox inputSmallLabel
    MsgBox inputLargeLabel
End Sub

'/Positioning /'
Private Sub UserForm_Initialize()
Dim sngLeft As Single
Dim sngTop As Single

    Call ReturnPosition_CenterScreen(Me.Height, Me.Width, sngLeft, sngTop)
    Me.Left = sngLeft
    Me.Top = sngTop
End Sub

Private Sub btnBack_click()
    Unload Me
    Label.Show
End Sub

Private Sub btnSubmit_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    Ws.Range("C26") = inputSmallLabel
    Ws.Range("C27") = inputLargeLabel
End Sub

