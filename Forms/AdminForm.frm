VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminForm 
   Caption         =   "Admin Panel  - Created By Alex Fare"
   ClientHeight    =   9360.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "AdminForm.frx":0000
End
Attribute VB_Name = "AdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long ' variable used for storing row number
Dim Worksheet_Set ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify
Dim currentUser    As String
Dim rlStatus As Integer

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'

'/ Display Admin Audit Log/'
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim ws          As Worksheet
    Dim List_Select
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    txtWorkbookOpened = ws.Range("B47")
    txtLogins = ws.Range("B48")
    txtGageCount = ws.Range("B49")
    txtGageUpdates = ws.Range("B50")
    txtUserCounts = ws.Range("B51")
    txtCustomerCount = ws.Range("B53")
    lblLoggedUser = ws.Range("B52")
    txtGageRnRCount = ws.Range("B54")
    
    '/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    vDisplay = ws.Range("Z1")
    
    Gage_Number.SetFocus
End Sub

'/ Pressing Enter will instantly search /'
Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Confirm_Click
    End If
End Sub

'/------- Search Button -------/'
Private Sub Search_Confirm_Click()
    If Gage_Number <> "" Then
    Search_Button_Click
    Else
        ErrMsg_Blank
    End If
End Sub

Public Sub Search_Button_Click()
    
    Clear_Form ' clear previous data from form, except "Gage Number"
    
    Dim ws As Worksheet
    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)) Then
        Update_Button_Enable = False
        ErrMsg_NotFound
    Else
        r = Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)
        GN_Verify = Gage_Number
        PartNumbertxt = ws.Cells(r, "B")
        Descriptiontxt = ws.Cells(r, "C")
        comboGageType = ws.Cells(r, "D")
        Customer = ws.Cells(r, "E")
        Insp_Date = ws.Cells(r, "F")
        Due_Date = ws.Cells(r, "G")
        Department = ws.Cells(r, "I")
        Comments = ws.Cells(r, "J")
        Revtxt = ws.Cells(r, "K")
        serialInput = ws.Cells(r, "L")
        lblDateAdded = ws.Cells(r, "AK")
        lblDateEdit = ws.Cells(r, "AL")
        lblSearchedDate = ws.Cells(r, "AM")
        lastUser = ws.Cells(r, "AN")
        Ownertxt = ws.Cells(r, "M")
        comboStatus = ws.Cells(r, "Z")
        Update_Button_Enable = True
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Searching..."
        Status
    End If
End Sub

'/------- Update Button -------/'
Private Sub Update_Button_Click()
    If Update_Button_Enable = True Then
        If GN_Verify = Gage_Number Then
            Update_Worksheet
        Else
            MSG_Verify_Update
        End If
    Else
        ErrMsg_Search
    End If
End Sub

Private Sub Update_Worksheet()
    If Update_Button_Enable = True Then
        Dim gnString As String
        Set ws = Worksheet_Set
        If IsNumeric(Gage_Number) Then
            gnString = Val(Gage_Number.Value)
        Else
            gnString = Gage_Number
        End If

        '/------- Gage Info -------/'
        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = PartNumbertxt
        ws.Cells(r, "C") = Descriptiontxt
        ws.Cells(r, "D") = comboGageType
        ws.Cells(r, "E") = Customer
        ws.Cells(r, "I") = Department
        ws.Cells(r, "K") = Revtxt
        ws.Cells(r, "L") = serialInput
        ws.Cells(r, "AK") = lblDateAdded        'Date Added
        ws.Cells(r, "AL") = lblDateEdit
        ws.Cells(r, "AM") = lblSearchedDate
        ws.Cells(r, "AN") = lastUser
        ws.Cells(r, "M") = Ownertxt
        ws.Cells(r, "Z") = comboStatus
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Updating..."
        Status
        
    Else
        ErrMsg_Search
    End If
End Sub

Sub MSG_Verify_Update()
    MSG1 = MsgBox("Are you sure you want To change the Gage ID?", vbYesNo, "Verify")
    
    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        Gage_Number = GN_Verify
    End If
End Sub

'/------- Clear -------/'
Private Sub Clear_Form()
    PartNumbertxt = ""
    serialNumberTxt = ""
    lblDateAdded = ""
    lblDateEdit = ""
    lblSearchedDate = ""
    lastUser = ""
    Ownertxt = ""
    Revtxt = ""
    serialInput = ""
    Descriptiontxt = ""
    comboGageType = ""
    Customer = ""
    Department = ""
    Comments = ""
    comboStatus = ""
End Sub

'/------- Clear Button -------/'
Private Sub btnClear_Click()
    Update_Button_Enable = False
    Gage_Number = ""
    Clear_Form
End Sub

Sub CheckForUpdate_Click()
    Dim url         As String
    url = "https://github.com/alexfare/GageCalibrationTracker"
    ActiveWorkbook.FollowHyperlink url
End Sub

Private Sub btnClose_Click()
    Unload AdminForm
    Menu.Show
End Sub

Private Sub btnCreateAccount_click()
    Unload AdminForm
    CreateAccount.Show
End Sub

Private Sub btnUpdateUser_click()
    Unload AdminForm
    ChangePassword.Show
End Sub

Private Sub btnEditLists_Click()
    Unload AdminForm
    Worksheets("Lists").Activate
End Sub

Private Sub btnCustomers_Click()
    Unload AdminForm
    FormCustomer.Show
End Sub

Private Sub btnCompanyProfile_Click()
    CompanyProfile.Show
End Sub

'/------- Status -------/'
Private Sub Status()
    Dim startTime As Date
    Dim elapsedTime As Long
    Dim waitTimeInSeconds As Long
    
    waitTimeInSeconds = 2
    startTime = Now
    Do While elapsedTime < waitTimeInSeconds
        DoEvents 'allow the program to process any pending events
        elapsedTime = DateDiff("s", startTime, Now)
    Loop
        statusLabel.Caption = ""
        statusLabelLog.Caption = ""
End Sub

Private Sub btnLogout_Click()
        List_Select = "Admin"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        ws.Range("B55") = "1"
        Unload AdminForm
End Sub

Private Sub btnSave_click()
    ThisWorkbook.Save
    
    '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Saving..."
        Status
End Sub

Private Sub btnReleaseNotes_click()
    Dim url         As String
    url = "https://github.com/alexfare/GageCalibrationTracker/releases/latest"
    ActiveWorkbook.FollowHyperlink url
End Sub

Private Sub SuperAdminBTN_click()
    Dim isSuperAdmin As Boolean
    Dim loggedUser As String
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    loggedUser = ws.Range("B52")

    List_Select = "Credentials"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(loggedUser), Val(loggedUser), loggedUser), ws.Columns(1), 0)) Then
        btnUpdate_Enable = False
        ErrMsg_UserError
        Exit Sub
    Else
        r = Application.Match(IIf(IsNumeric(loggedUser), Val(loggedUser), loggedUser), ws.Columns(1), 0)
        GN_Verify = loggedUser
        btnUpdate_Enable = True
        isSuperAdmin = ws.Cells(r, "H")
    End If
    
    If isSuperAdmin = True Then
        Unload AdminForm
        SuperAdminMenu.Show
    Else
        ErrMsg_NotSuperAdmin
    End If
End Sub

Private Sub btnExport_click()
    ExportGCTData
End Sub

Sub ExportGCTData()
    Dim FilePath As Variant
    Dim ws As Worksheet
    Dim defaultFileName As String
    
    Set ws = ThisWorkbook.Worksheets("CreatedByAlexFare")
    
    defaultFileName = "GageTracker_" & Format(Date, "yyyy-mm-dd") & ".csv"
    
    FilePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, FileFilter:="CSV Files (*.csv), *.csv")
    
    If FilePath <> False Then
        ws.Copy
        ActiveWorkbook.SaveAs FilePath, xlCSV
        ActiveWorkbook.Close SaveChanges:=False
    End If
End Sub


Sub btnImport_click()

    MSG1 = MsgBox("Importing is a WIP, Current state will not import certain formatting conditions.", vbYesNo, "WARNING")
    
    If MSG1 = vbYes Then
        ImportGCTData
    Else
    End If
End Sub

Sub ImportGCTData()
    Dim FilePath As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set ws = ThisWorkbook.Worksheets("CreatedByAlexFare")
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select CSV File to Import"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
        End If
    End With
    
    If FilePath <> "" Then
        ws.Cells.ClearContents
        ws.Cells.FormatConditions.Delete
        
        With ws.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=ws.Cells(1, 1))
            .TextFileParseType = xlDelimited
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        
        ws.Cells.EntireColumn.AutoFit
    End If
End Sub

'/------- Error Handling -------/'
Sub ErrMsg_NotFound()
    MsgBox ("Gage Number Not Found"), , "Not Found"
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Gage number already in use"), , "Duplicate"
End Sub

Sub ErrMsg_Search()
    MsgBox ("Must search For entry before updating"), , "Nothing To Update"
End Sub

Sub ErrMsg_Blank()
    MsgBox ("Gage ID cannot be blank."), , "Nothing To Update"
End Sub

Sub ErrMsg_NotSuperAdmin()
    MSG1 = MsgBox("User not Super Admin, Sign in with a password?", vbYesNo, "Admin Error")
    
    If MSG1 = vbYes Then
        Unload AdminForm
        SuperAdminLogin.Show
    Else
    End If
End Sub

Sub ErrMsg_UserError()
    MsgBox ("Having Issues Logging In, Please Try Again."), , "Admin Error"
End Sub


