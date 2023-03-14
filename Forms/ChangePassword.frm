VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangePassword 
   Caption         =   "Change Password"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3585
   OleObjectBlob   =   "ChangePassword.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long        ' variable used for storing row number
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim btnUpdate_Enable As Boolean        ' to store update enable flag after search
Dim GN_Verify

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub inputUser_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnUpdate_Click
    End If
End Sub

Public Sub Search_Button_Click()
    
    ' clear previous data from form, except "Gage Number"
    ' --------------------------------------------------------
    inputPass = ""
    Descriptiontxt = ""
    ' ---------------------------------------------------------
    
    Dim Ws As Worksheet
    
    List_Select = "Credentials"
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    If IsError(Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), Ws.Columns(1), 0)) Then
        btnUpdate_Enable = False
        ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), Ws.Columns(1), 0)
        GN_Verify = inputUser
        btnUpdate_Enable = True
        
        Dim FS
        Set FS = CreateObject("Scripting.FileSystemObject")
        
        If FS.FileExists(TextFile_FullPath) Then
        Else
        End If
    End If
    inputUser.SetFocus
End Sub

Private Sub btnUpdate_Click()
    If btnUpdate_Enable = True Then
        If GN_Verify = inputUser Then
            Update_Worksheet
        Else
            MSG_Verify_Update
        End If
    Else
        MsgBox ("Must search For entry before updating"), , "Nothing To Update"
    End If
End Sub

Sub ErrMsg()
    MsgBox ("Username Not Found"), , "Not Found"
    inputUser.SetFocus
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Username already in use"), , "Duplicate"
    inputUser.SetFocus
End Sub

Private Sub Clear_Form()
    inputUser = ""
    inputPass = ""
End Sub

Private Sub Update_Worksheet()
    If btnUpdate_Enable = True Then
        Dim gnString As String
        Set Ws = Worksheet_Set
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
        sSecret = ""        'secret key for StrToSHA512Salt only
        
        'b64 = False   'output hex
        b64 = True        'output base-64
        
        sH = SHA512(sIn, b64)
        
        'message box and immediate window outputs
        Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
        ' MsgBox sH & vbNewLine & Len(sH) & " characters in length"
        savePass = sH
        '/ Hash /'
        
        Ws.Cells(r, "A") = gnString
        Ws.Cells(r, "B") = savePass
        
        btnUpdate.Caption = "Updated!"
        Application.Wait (Now + TimeValue("0:00:01"))
        btnUpdate.Caption = "Update"
        'Clear_Form 'Clear form after update
        inputUser.SetFocus
        
    Else
        MsgBox ("Must search For entry before updating"), , "Nothing To Update"
        
    End If
    
    'btnUpdate_Enable = False 'Remove ' if you want to require searching again after an update.
    
End Sub

Sub MSG_Verify_Update()
    MSG1 = MsgBox("Are you sure you want To change the Username?", vbYesNo, "Verify")
    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        inputUser = GN_Verify
    End If
End Sub

Private Sub btnBack_click()
    Unload ChangePassword
    AdminForm.Show
End Sub

