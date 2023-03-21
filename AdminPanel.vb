Dim r As Long ' variable used for storing row number
Dim Worksheet_Set ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify
Dim currrentUser    As String
Dim rlStatus As Integer

Private Sub UserForm_Initialize()
    '/ Display Admin Audit Log/'
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim Ws          As Worksheet
    Dim List_Select
    List_Select = "Admin"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    txtWorkbookOpened = Ws.Range("B47")
    txtLogins = Ws.Range("B48")
    txtGageCount = Ws.Range("B49")
    txtGageUpdates = Ws.Range("B50")
    txtUserCounts = Ws.Range("B51")
    txtCustomerCount = Ws.Range("B53")
    lblLoggedUser = Ws.Range("B52")
    txtGageRnRCount = Ws.Range("B54")
    
    '/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
End Sub

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

'/ Pressing Enter will instantly search /'
Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
    End If
End Sub

Public Sub Search_Button_Click()
    
    ' clear previous data from form, except "Gage Number"
    ' --------------------------------------------------------
    PartNumbertxt = ""
    lblDateAdded = ""
    lblDateEdit = ""
    lblSearchedDate = ""
    lastUser = ""
    
    ' ---------------------------------------------------------
    
    Dim Ws          As Worksheet
    
    List_Select = "CreatedByAlexFare"
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)) Then
        Update_Button_Enable = False
        ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)
        GN_Verify = Gage_Number
        PartNumbertxt = Ws.Cells(r, "B")
        lblDateAdded = Ws.Cells(r, "AK")
        lblDateEdit = Ws.Cells(r, "AL")
        lblSearchedDate = Ws.Cells(r, "AM")
        lastUser = Ws.Cells(r, "AN")
        
        'Below might be doing nothing anymore? Check into this
        Update_Button_Enable = True
        Option4_Custom = True
        Dim FS
        Set FS = CreateObject("Scripting.FileSystemObject")
        
        If FS.FileExists(TextFile_FullPath) Then
        Else
        End If
    End If
    
    Gage_Number.SetFocus
    
End Sub

Sub ErrMsg()
    MsgBox ("Gage Number Not Found"), , "Not Found"
    Gage_Number.SetFocus
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Gage number already in use"), , "Duplicate"
    Gage_Number.SetFocus
End Sub

Private Sub Update_Button_Click()
    If Update_Button_Enable = True Then
        If GN_Verify = Gage_Number Then
            Update_Worksheet
        Else
            MSG_Verify_Update
        End If
    Else
        MsgBox ("Must search For entry before updating"), , "Nothing To Update"
    End If
End Sub
Private Sub Update_Worksheet()
    If Update_Button_Enable = True Then
        Dim gnString As String
        Set Ws = Worksheet_Set
        If IsNumeric(Gage_Number) Then
            gnString = Val(Gage_Number.Value)
        Else
            gnString = Gage_Number
        End If
        '/ Audit
        Ws.Cells(r, "A") = gnString
        Ws.Cells(r, "B") = PartNumbertxt
        'Ws.Cells(r, "AL") = Now        'Update Last edited
        Ws.Cells(r, "AK") = lblDateAdded        'Date Added
        currrentUser = Application.userName
        lastUser = currrentUser
        Ws.Cells(r, "AN") = lastUser
        
        Update_Button.Caption = "Updated!"
        Application.Wait (Now + TimeValue("0:00:01"))
        Update_Button.Caption = ""
        Gage_Number.SetFocus
        
    Else
        MsgBox ("Must search For entry before updating"), , "Nothing To Update"
        
    End If
    
    'Update_Button_Enable = False 'Remove ' if you want to require searching again after an update.
    
End Sub

Sub MSG_Verify_Update()
    
    MSG1 = MsgBox("Are you sure you want To change the Gage ID?", vbYesNo, "Verify")
    
    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        Gage_Number = GN_Verify
    End If
    
End Sub

Private Sub Clear_Form()
    Gage_Number = ""
    PartNumbertxt = ""
    lblDateAdded = ""
    lblDateEdit = "-"
    lblSearchedDate = ""
    lastUser = ""
End Sub

Private Sub btnClear_Click()
    Update_Button_Enable = False
    Clear_Form
    Gage_Number.SetFocus
End Sub

Sub CheckForUpdate_Click()
    Dim URL         As String
    URL = "https://github.com/alexfare/GageCalibrationTracker"
    ActiveWorkbook.FollowHyperlink URL
End Sub

Private Sub btnClose_Click()
    Unload AdminForm
    
    '/Save Logged In User For The Session /'
    List_Select = "Admin"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    Ws.Range("B55") = "2"       ' 1 = Required | 2 = Not Required
End Sub

Private Sub btnCreateAccount_click()
    Unload AdminForm
    CreateAccount.Show
End Sub

Private Sub btnUpdateUser_click()
    Unload AdminForm
    ChangePassword.Show
End Sub

Private Sub btnDevMode_click()
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
End Sub

Private Sub btnEditLists_Click()
    Unload AdminForm
    Worksheets("Lists").Activate
End Sub

Private Sub btnAbout_Click()
    MsgBox "Code protection password Is GageTracker2022"
End Sub

Private Sub btnCustomers_Click()
    Unload AdminForm
    Worksheets("Customers").Activate
    FormCustomer.Show
End Sub

Private Sub btnCompanyProfile_Click()
    CompanyProfile.Show
End Sub

'/ Settings Tab /'
Private Sub btnRequireLogin_click()
    List_Select = "Admin"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    rlStatus = Ws.Range("B59")
    
    If rlStatus = "1" Then
        Ws.Range("B59") = "2"
        btnRequireLogin.Caption = "Off"
    End If
    If rlStatus = "2" Then
        Ws.Range("B59") = "1"
        btnRequireLogin.Caption = "On"
    End If
End Sub

