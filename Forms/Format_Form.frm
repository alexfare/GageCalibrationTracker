VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Format_Form 
   Caption         =   "Format"
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "Format_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Format_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnBack_click()
    Unload Format_Form
End Sub
    
Sub FormatConfirmed()
    Clear_All.Clear_Run
    Unload Format_Form
    Unload AdminForm
End Sub

Sub btnFormatConfirm_Click()

    MSG1 = MsgBox("WARNING: THIS WILL AUTOSAVE AT THE END, YOUR DATA WILL BE LOST!", vbYesNo, "WARNING - Formatting")
    
    If MSG1 = vbYes Then
        FormatConfirmed
    Else
        Unload Format_Form
    End If
End Sub
