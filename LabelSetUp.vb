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
