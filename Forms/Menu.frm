VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Gage Calibration Tracker - Menu"
   ClientHeight    =   8295.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10695
   OleObjectBlob   =   "Menu.frx":0000
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long        ' variable used for storing row number
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean        ' to store update enable flag after search
Dim GN_Verify
Dim Due_Date_Original
Dim Date_Due
Dim ActionLog As String 'Audit Log
Dim AuditTime As String 'Audit Log
Dim AuditUser As String 'Audit Log
Dim auditDate As String 'Audit Log
Dim GageList As String
Dim AdminList As String
Dim List_Select
Dim ws As Worksheet

'/Start up script /'
Private Sub UserForm_Activate()
    '/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    '/End Positioning /'

    '/ Setup /'
    GageList = "CreatedByAlexFare"
    AdminList = "Admin"

    List_Select = GageList
    Set ws = Sheets(List_Select)
    vDisplay = ws.Range("Z1")
    Gage_Number.SetFocus
    SettingsModule.DueDateColor
    lblPCUser = Application.userName
    SettingsModule.GetCurrentVersion
End Sub

Private Sub UserForm_Initialize()
    With cboInterval
        .AddItem "6 Months"  ' Corresponds to Interval_6
        .AddItem "1 Year"    ' Corresponds to Interval_1
        .AddItem "2 Years"   ' Corresponds to Interval_2
        .AddItem "Custom"    ' Corresponds to Interval_Custom
    End With
End Sub

Private Sub Add_Button_Click()
    ' Check if the user provided input
    If Gage_Number <> "" Then
        AddNewGage
    Else
        ErrMsg_NoGageID
    End If
End Sub

'/------- Add Gage -------/'
Private Sub AddNewGage()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = GageList
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws

    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)) Then
        Dim lLastRow As Long        ' lLastRow = variable to store the result of the row count calculation
        lLastRow = ws.ListObjects.Item(1).ListRows.Count
        r = lLastRow + 3        ' Add number for every header tab created
        Dim gnString As String
        If IsNumeric(Gage_Number) Then
            gnString = Val(Gage_Number.Value)
        Else
            gnString = Gage_Number
        End If
        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = PartNumbertxt
        ws.Cells(r, "C") = Descriptiontxt
        ws.Cells(r, "D") = comboGageType
        ws.Cells(r, "E") = Customer
        ws.Cells(r, "F") = Insp_Date

        ' Setting Due Date based on cboInterval selection
        Select Case cboInterval.Text
            Case "6 Months"
                Date_Due = Format(DateAdd("m", 6, Insp_Date), "m/d/yyyy")
            Case "1 Year"
                Date_Due = Format(DateAdd("yyyy", 1, Insp_Date), "m/d/yyyy")
            Case "2 Years"
                Date_Due = Format(DateAdd("yyyy", 2, Insp_Date), "m/d/yyyy")
            Case "Custom"
                Date_Due = Format(InputBox("Enter custom due date:", "Custom Date", Format(Date, "m/d/yyyy")), "m/d/yyyy")
            Case Else
                MsgBox "Invalid Interval. Please select a valid interval.", vbCritical
                Exit Sub
        End Select

        ws.Cells(r, "G") = Date_Due
        ws.Cells(r, "O") = cboInterval.Text

        ws.Cells(r, "H") = Initials
        ws.Cells(r, "I") = Department
        ws.Cells(r, "J") = Comments
        ws.Cells(r, "K") = Revtxt
        ws.Cells(r, "L") = serialInput
        ws.Cells(r, "N") = nistinput
        ws.Cells(r, "Z") = comboStatus
        ws.Cells(r, "AA") = aN1
        ws.Cells(r, "AB") = aA1
        ws.Cells(r, "AC") = aN2
        ws.Cells(r, "AD") = aA2
        ws.Cells(r, "AE") = aN3
        ws.Cells(r, "AF") = aA3
        ws.Cells(r, "AG") = aN4
        ws.Cells(r, "AH") = aA4
        ws.Cells(r, "AI") = aN5
        ws.Cells(r, "AJ") = aA5
        ws.Cells(r, "AK") = Now

        '/ Audit Log
        lastUser = Application.userName
        ws.Cells(r, "AN") = lastUser
        ActionLog = "Added Gage"
        auditLog

        btnClear_Click
        AddGageCount

        '/Status /'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Adding..."
        Status
    Else
        ErrMsg_Duplicate
    End If
End Sub

'/------- Press Enter -------/'
Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Confirm_Click
    End If
End Sub

'/------- Search Button -------/'
Public Sub Search_Confirm_Click()
    If Gage_Number <> "" Then
        Search_Button
    Else
        ErrMsg_Blank
    End If
End Sub

Public Sub Search_Button()
    Dim ws As Worksheet
    Dim DateEdit 'Update Last searched
    Clear_Form ' clear previous data from form, except "Gage Number"
    List_Select = GageList
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
        Due_Date_Original = ws.Cells(r, "G")
        Due_Date = Format(Due_Date_Original, "m/d/yyyy")
        Initials = ws.Cells(r, "H")
        Department = ws.Cells(r, "I")
        Comments = ws.Cells(r, "J")
        Revtxt = ws.Cells(r, "K")
        serialInput = ws.Cells(r, "L")
        nistinput = ws.Cells(r, "N")
        comboStatus = ws.Cells(r, "Z")
        aN1 = ws.Cells(r, "AA")
        aA1 = ws.Cells(r, "AB")
        aN2 = ws.Cells(r, "AC")
        aA2 = ws.Cells(r, "AD")
        aN3 = ws.Cells(r, "AE")
        aA3 = ws.Cells(r, "AF")
        aN4 = ws.Cells(r, "AG")
        aA4 = ws.Cells(r, "AH")
        aN5 = ws.Cells(r, "AI")
        aA5 = ws.Cells(r, "AJ")
        cboInterval = ws.Cells(r, "O")
        DateEdit = ws.Cells(r, "AM") 'Update Last searched
        ws.Cells(r, "AM") = Now        'Update Last searched
        Update_Button_Enable = True

        '/ Audit Log
        lblDateAdded = ws.Cells(r, "AK")
        lblDateEdit = ws.Cells(r, "AL")
        lblSearchedDate = DateEdit 'Update Last searched
        lastUser = ws.Cells(r, "AN")
        ActionLog = "Searched"
        auditLog

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
        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = PartNumbertxt
        ws.Cells(r, "C") = Descriptiontxt
        ws.Cells(r, "D") = comboGageType
        ws.Cells(r, "E") = Customer
        ws.Cells(r, "F") = Insp_Date
        ws.Cells(r, "H") = Initials
        ws.Cells(r, "I") = Department
        ws.Cells(r, "J") = Comments
        ws.Cells(r, "K") = Revtxt
        ws.Cells(r, "L") = serialInput
        ws.Cells(r, "N") = nistinput
        ws.Cells(r, "Z") = comboStatus
        ws.Cells(r, "AA") = aN1
        ws.Cells(r, "AB") = aA1
        ws.Cells(r, "AC") = aN2
        ws.Cells(r, "AD") = aA2
        ws.Cells(r, "AE") = aN3
        ws.Cells(r, "AF") = aA3
        ws.Cells(r, "AG") = aN4
        ws.Cells(r, "AH") = aA4
        ws.Cells(r, "AI") = aN5
        ws.Cells(r, "AJ") = aA5
        ws.Cells(r, "AL") = Now        'Update Last edited

        ' Setting Due Date based on cboInterval selection
        Select Case cboInterval.Text
            Case "6 Months"
                Date_Due = Format(DateAdd("m", 6, Insp_Date), "m/d/yyyy")
            Case "1 Year"
                Date_Due = Format(DateAdd("yyyy", 1, Insp_Date), "m/d/yyyy")
            Case "2 Years"
                Date_Due = Format(DateAdd("yyyy", 2, Insp_Date), "m/d/yyyy")
            Case "Custom"
                Date_Due = Format(InputBox("Enter custom due date:", "Custom Date", Format(Date, "m/d/yyyy")), "m/d/yyyy")
            Case Else
                MsgBox "Invalid Interval. Please select a valid interval.", vbCritical
                Exit Sub
        End Select

        ws.Cells(r, "G") = Date_Due
        ws.Cells(r, "O") = cboInterval.Text

        '/ Audit Log
        lastUser = Application.userName
        ws.Cells(r, "AN") = lastUser

        ActionLog = "Updated Gage"
        auditLog

        '/Prevent Issues in the future, Call back the main page/'
        List_Select = GageList
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        '/ End Audit Log /'

        '/Status /'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Updating..."
        Status

        Search_Button
    Else
        ErrMsg_Search
    End If
End Sub

'/------- Clear Form -------/'
Private Sub btnClear_Click()
    Update_Button_Enable = False
    Gage_Number = ""
    Clear_Form
End Sub

Private Sub Clear_Form()
    PartNumbertxt = ""
    Descriptiontxt = ""
    comboGageType = ""
    Customer = ""
    Insp_Date = ""
    Due_Date = ""
    Initials = ""
    Department = ""
    Comments = ""
    Revtxt = ""
    serialInput = ""
    nistinput = ""
    comboStatus = ""
    aN1 = ""
    aA1 = ""
    aN2 = ""
    aA2 = ""
    aN3 = ""
    aA3 = ""
    aN4 = ""
    aA4 = ""
    aN5 = ""
    aA5 = ""
    lblDateAdded = ""
    lblDateEdit = ""
    lblSearchedDate = ""
    lastUser = ""
    cboInterval.ListIndex = -1
End Sub

Sub MSG_Verify_Update()

    MSG1 = MsgBox("Are you sure you want to change the Gage ID?", vbYesNo, "Verify")

    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        Gage_Number = GN_Verify
    End If

End Sub

Private Sub btnSave_click()
    ThisWorkbook.Save

    '/Status/'
    statusLabel.Caption = "Status:"
    statusLabelLog.Caption = "Saving..."
    Status
End Sub

Private Sub bgSave()
    ThisWorkbook.Save
End Sub

'/------- Admin Panel -------/'
Private Sub btnAdmin_click()
    '/Add to the login count /'
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim LoginCount  As Integer
    Dim ws As Worksheet
    Dim List_Select
    Dim TempLogin   As Integer
    List_Select = AdminList
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Persistent_Login = ws.Range("B55")

    If Persistent_Login = True Then
        Sheets("CreatedByAlexFare").Activate
        Unload Menu
        AdminForm.Show
    Else
        Unload Menu
        LoginForm.Show
    End If
End Sub

'/Report Issue Panel /'
Private Sub btnReportIssue_click()
    Unload Menu
    ReportIssue.Show
End Sub

'/Label Printing /'
Private Sub btnLabel_Click()
    Label.Show
End Sub

'/------- Gage R&R -------/'
Private Sub btnGageRR_Click()
    Unload Menu
    GageRnR.Show
End Sub

'/------- Status -------/'
Private Sub Status()
    Dim startTime As Date
    Dim elapsedTime As Long
    Dim waitTimeInSeconds As Long
    SettingsModule.DueDateColor
    bgSave

    waitTimeInSeconds = 1
    startTime = Now
    Do While elapsedTime < waitTimeInSeconds
        DoEvents 'allow the program to process any pending events
        elapsedTime = DateDiff("s", startTime, Now)
    Loop
    statusLabel.Caption = ""
    statusLabelLog.Caption = ""
End Sub

'/ ------- Audit Log ------- /'
Private Sub auditLog()
    Dim ws As Worksheet
    Dim auditLog As String
    Dim auditAdd As String
    Dim auditDate As String

    Set ws = ThisWorkbook.Sheets("Audit")

    AuditTime = Now
    AuditUser = Application.userName

    auditLog = ws.Range("A2").Value
    auditDate = Now
    auditAdd = " | Date: " & auditDate & vbCrLf & " User: " & AuditUser & vbCrLf & " Action: " & ActionLog & " | " & vbCrLf & " "
    auditLog = auditLog & vbCrLf & auditAdd

    ws.Range("A2").Value = auditLog
End Sub

'/ -------  Auto Due Date ------- /'
Private Sub cboInterval_Change()
    On Error GoTo Err

    If Not IsDate(Insp_Date) Then
        Exit Sub
    End If

    Select Case cboInterval.Text
        Case "6 Months"
            Due_Date = Format(DateAdd("m", 6, Insp_Date), "m/d/yyyy")

        Case "1 Year"
            Due_Date = Format(DateAdd("yyyy", 1, Insp_Date), "m/d/yyyy")

        Case "2 Years"
            Due_Date = Format(DateAdd("yyyy", 2, Insp_Date), "m/d/yyyy")

        Case "Custom"
            Due_Date = Format(Due_Date, "m/d/yyyy")
        
        Case ""
            Due_Date = ""

        Case Else
            ErrMsg_InvalidDate
    End Select

    Exit Sub

Err:
    ErrMsg_InvalidDate
End Sub

Private Sub AddGageCount()
'/Add to Gage Number count/'
        Dim AddCount As Integer

        List_Select = AdminList
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws

        AddCount = ws.Range("B49")
        AddCountPlusOne = AddCount + 1
        ws.Range("B49") = AddCountPlusOne

        '/Prevent Issues in the future, Call back the main page/'
        List_Select = GageList
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
End Sub

'/------- Error Handling -------/'
Sub ErrMsg_NotFound()
    MsgBox ("Gage number not found."), vbInformation, "Error - Not Found"
End Sub

Sub ErrMsg_Duplicate()
    MsgBox "Gage number already exists. Please use a different gage number.", vbInformation, "Error - Duplicate"
End Sub

Sub ErrMsg_InvalidDate()
    MsgBox "Invalid date format. Please enter the date in mm/dd/yyyy or m/d/yyyy format.", vbInformation, "Error - Date"
End Sub

Sub ErrMsg_NoGageID()
    MsgBox "Please provide a gage name.", vbInformation, "Error - Blank"
End Sub

Sub ErrMsg_Search()
    MsgBox ("Must search for entry before updating."), vbInformation, "Error - Search"
End Sub

Sub ErrMsg_Blank()
    MsgBox ("Gage number cannot be blank."), vbInformation, "Error - Blank"
End Sub

Private Sub UserForm_Terminate()
    SettingsModule.DueDateColor
End Sub

