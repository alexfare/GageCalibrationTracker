VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateAccount 
   Caption         =   "Create Account"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4065
   OleObjectBlob   =   "CreateAccount.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify

'/Positioning /'
Private Sub UserForm_Initialize()
Dim sngLeft As Single
Dim sngTop As Single

    Call ReturnPosition_CenterScreen(Me.Height, Me.Width, sngLeft, sngTop)
    Me.Left = sngLeft
    Me.Top = sngTop
End Sub

Private Sub btnCreate_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Credentials" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws

    
    If IsError(Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), Ws.Columns(1), 0)) Then
  
    Dim lLastRow As Long    ' lLastRow = variable to store the result of the row count calculation
    lLastRow = Ws.ListObjects.Item(1).ListRows.Count
    r = lLastRow + 2 ' Add number for every header tab created
                    Dim gnString As String
                        If IsNumeric(inputUser) Then
                            gnString = Val(inputUser.Value)
                        Else
                            gnString = inputUser
                        End If
                        
                        
'/ Hash /'
    s = inputPass
    
    Dim sIn As String, sOut As String, b64 As Boolean
    Dim sH As String, sSecret As String
    
    'Password to be converted
    sIn = s
    sSecret = "" 'secret key for StrToSHA512Salt only
    
    'select as required
    'b64 = False   'output hex
    b64 = True   'output base-64
    
    sH = SHA512(sIn, b64)
    
    'message box and immediate window outputs
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
    ' MsgBox sH & vbNewLine & Len(sH) & " characters in length"
  savePass = sH
'/ Hash /'
                        
                        
    Ws.Cells(r, "A") = gnString
    Ws.Cells(r, "B") = savePass
    
    btnCreate.Caption = "Created!" ' change caption of add button for confirmation
    Application.Wait (Now + TimeValue("0:00:02")) ' Wait to avoid crash
    btnCreate.Caption = "Create"
    Clear_Form
    inputUser.SetFocus
    Unload CreateAccount
    AdminForm.Show
    Else
        ErrMsg_Duplicate
    End If
    
End Sub


Sub ErrMsg()
MsgBox ("Username Not Found"), , "Not Found"
inputUser.SetFocus
End Sub

Sub ErrMsg_Duplicate()
MsgBox ("Username Taken"), , "Duplicate"
inputUser.SetFocus
End Sub

Private Sub Clear_Form()
        inputUser = ""
        inputPass = ""
End Sub

Private Sub btnBack_click()
Unload CreateAccount
AdminForm.Show
End Sub




