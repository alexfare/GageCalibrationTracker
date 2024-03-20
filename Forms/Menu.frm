VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "GageTracker - Created By Alex Fare"
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
' Gage Tracker
' Created By: Alex Fare

Dim r As Long        ' variable used for storing row number
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean        ' to store update enable flag after search
Dim GN_Verify
Dim Due_Date_Original
Dim Date_Due_6mos
Dim Date_Due_1yr
Dim Date_Due_2yr
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
    '/ Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    '/ End Positioning /'
    
    '/ Setup /'
    GageList = "CreatedByAlexFare"
    AdminList = "Admin"
    
    List_Select = GageList
    Set ws = Sheets(List_Select)
    vDisplay = ws.Range("Z1")
    Gage_Number.SetFocus
    DueDateColorRange
    lblPCUser = Application.userName
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
        ws.Cells(r, "G") = Due_Date
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
        DateEdit = ws.Cells(r, "AM") 'Update Last searched
        ws.Cells(r, "AM") = Now        'Update Last searched
        Update_Button_Enable = True
        Interval_Custom = True
        
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
        
        '/ Audit Log
        lastUser = Application.userName
        ws.Cells(r, "AN") = lastUser
        
        If Interval_6 = True Then        ' option1 = 1month, option2 = 6months, option3 = 1year, option4 = custom or original
        Due_Date = Date_Due_6mos
    End If
    If Interval_1 = True Then
        Due_Date = Date_Due_1yr
    End If
    If Interval_2 = True Then
        Due_Date = Date_Due_2yr
    End If
    If Interval_Custom = True Then
        Interval_Custom_Click
        Due_Date = Date_Due
    End If
    
    ws.Cells(r, "G") = Due_Date
    
    '/Audit Log/'
    Dim UpdateCount As Integer
    
    List_Select = AdminList
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    UpdateCount = ws.Range("B50")
    UpdateCountPlusOne = UpdateCount + 1
    ws.Range("B50") = UpdateCountPlusOne
    
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
    DueDateColorRange
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

'/------- Update Due Date Color -------/'
Sub DueDateColorRange()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim targetDate As Date
    Dim currentDate As Date
    Dim Worksheet_Set
    Dim List_Select
    Dim ColorRangeLeadTime As Integer
    
    List_Select = AdminList
    Set Worksheet_Set = ws
    Set ws = Sheets(List_Select)
    ColorRangeLeadTime = ws.Range("B63")
    
    List_Select = GageList
    Set Worksheet_Set = ws
    Set ws = Sheets(List_Select)

    targetDate = Range("I1").Value
    Set rng = ws.Range("G3:G2000")
    
    For Each cell In rng
        If IsDate(cell.Value) Then
            currentDate = cell.Value
            
            monthsDiff = DateDiff("m", targetDate, currentDate)
            
            If currentDate < targetDate Then
                cell.Interior.Color = RGB(255, 0, 0) 'Red
            ElseIf monthsDiff <= ColorRangeLeadTime Then
                cell.Interior.Color = RGB(255, 255, 0) 'Yellow
            Else
                cell.Interior.Color = RGB(0, 255, 0) 'Green
            End If
        End If
    Next cell
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
Private Sub Interval_6_Click() ' auto format for 1 year interval
On Error GoTo Err
If IsDate(Insp_Date) Then 'check if Insp_Date is a valid date
    Date_Due_6mos = DateAdd("m", 6, Insp_Date)
    Date_Due_6mos = Format(Date_Due_6mos, "m/d/yyyy")
    Due_Date = Date_Due_6mos
    Exit Sub
End If

Err:
    ErrMsg_InvalidDate
End Sub

Private Sub Interval_1_Click() ' auto format for 1 year interval
On Error GoTo Err
If IsDate(Insp_Date) Then 'check if Insp_Date is a valid date
    Date_Due_1yr = DateAdd("yyyy", 1, Insp_Date)
    Date_Due_1yr = Format(Date_Due_1yr, "m/d/yyyy")
    Due_Date = Date_Due_1yr
    Exit Sub
End If

Err:
    ErrMsg_InvalidDate
End Sub

Private Sub Interval_2_Click() ' auto format for 2 year interval
On Error GoTo Err
If IsDate(Insp_Date) Then 'check if Insp_Date is a valid date
    Date_Due_2yr = DateAdd("yyyy", 2, Insp_Date)
    Date_Due_2yr = Format(Date_Due_2yr, "m/d/yyyy")
    Due_Date = Date_Due_2yr
    Exit Sub
End If

Err:
    ErrMsg_InvalidDate
End Sub

Private Sub Interval_Custom_Click() ' formatting for either original record, or new custom date
On Error GoTo Err
If IsDate(Insp_Date) Then 'check if Insp_Date is a valid date
    Date_Due = Format(Due_Date, "m/d/yyyy")
    Due_Date = Date_Due
    Exit Sub
End If

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
    MsgBox ("Gage Number Not Found."), vbInformation, "Not Found"
End Sub

Sub ErrMsg_Duplicate()
    MsgBox "Gage number already exists. Please use a different Gage number.", vbInformation, "Duplicate Gage Number"
End Sub

Sub ErrMsg_InvalidDate()
    MsgBox "Invalid date format. Please enter the date in mm/dd/yyyy or m/d/yyyy format.", vbInformation, "Error"
End Sub

Sub ErrMsg_NoGageID()
    MsgBox "Please Provide A Gage Name.", vbInformation, "Error"
End Sub

Sub ErrMsg_Search()
    MsgBox ("Must search for entry before updating."), vbInformation, "Error"
End Sub

Sub ErrMsg_Blank()
    MsgBox ("Gage number cannot be blank."), vbInformation, "Error"
End Sub

Private Sub UserForm_Terminate()
    DueDateColorRange
End Sub

