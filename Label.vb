Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnLargeLabel_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    Dim x As Variant
    Dim Path As String
        Path = Ws.Range("C27")
        'MsgBox (Path) 'Confirms the path works
        x = Shell("explorer.exe " + Path, vbNormalFocus) 'explorer.exe is needed due to vba expecting a .exe
        Unload Me
End Sub

Private Sub btnSmallLabel_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    Dim x As Variant
    Dim Path As String
        Path = Ws.Range("C26")
        'MsgBox (Path) 'Confirms the path works
        x = Shell("explorer.exe " + Path, vbNormalFocus) 'explorer.exe is needed due to vba expecting a .exe
        Unload Me
End Sub

Private Sub btnSetUp_Click()
    Unload Me
    LabelSetUp.Show
End Sub

Private Sub btnBack_click()
    Unload Me
End Sub



