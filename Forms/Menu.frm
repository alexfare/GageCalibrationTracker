VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "GageTracker - Created By Alex Fare"
   ClientHeight    =   8895.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9960.001
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   2  'CenterScreen
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
Dim currrentUser As String

'/Start up script /'
Private Sub UserForm_Initialize()
'/Code Confirm for production use only/'
    Dim CodeCompare As Integer
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim LoginCount  As Integer
    Dim ws          As Worksheet
    Dim List_Select
    
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    CodeCompare = ws.Range("B56")
    If CodeCompare = "1" Then
        Unload Menu
        CodeConfirm.Show
    End If
'/ End code confirm /'

'/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
End Sub
Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

'/Auto Due Date
Private Sub Option1_6_Click()        ' auto format for 6 month interval
If IsDate(Insp_Date) Then 'check if Insp_Date is a valid date
    Date_Due_6mos = DateAdd("m", 6, Insp_Date)
    Date_Due_6mos = Format(Date_Due_6mos, "m/d/yyyy")
    Due_Date = Date_Due_6mos
Else
    MsgBox "Invalid date format. Please enter the date in mm/dd/yyyy or m/d/yyyy format."
End If
End Sub
Private Sub Option2_12_Click()        ' auto format for 1 year interval
If IsDate(Insp_Date) Then 'check if Insp_Date is a valid date
    Date_Due_1yr = DateAdd("yyyy", 1, Insp_Date)
    Date_Due_1yr = Format(Date_Due_1yr, "m/d/yyyy")
    Due_Date = Date_Due_1yr
Else
    MsgBox "Invalid date format. Please enter the date in mm/dd/yyyy or m/d/yyyy format."
End If
End Sub
Private Sub Option3_24_Click()
If IsDate(Insp_Date) Then 'check if Insp_Date is a valid date
    Date_Due_2yr = DateAdd("yyyy", 2, Insp_Date)
    Date_Due_2yr = Format(Date_Due_2yr, "m/d/yyyy")
    Due_Date = Date_Due_2yr
Else
    MsgBox "Invalid date format. Please enter the date in mm/dd/yyyy or m/d/yyyy format."
End If
End Sub
Private Sub Option4_Custom_Click()        ' formatting for either original record, or new custom date
If IsDate(Insp_Date) Then 'check if Insp_Date is a valid date
    Date_Due = Format(Due_Date, "m/d/yyyy")
    Due_Date = Date_Due
Else
    MsgBox "Invalid date format. Please enter the date in mm/dd/yyyy or m/d/yyyy format."
End If
End Sub

'/ Add Gage
Private Sub Add_Button_Click()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "CreatedByAlexFare"        ' Tab name
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
        currrentUser = Application.userName
        lastUser = currrentUser
        ws.Cells(r, "AN") = lastUser
        
        Clear_Form
        Gage_Number.SetFocus
        
        '/Add to Gage Number count/'
        Dim AddCount As Integer
        
        List_Select = "Admin"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        
        AddCount = ws.Range("B49")
        AddCountPlusOne = AddCount + 1
        ws.Range("B49") = AddCountPlusOne
        
        '/Prevent Issues in the future, Call back the main page/'
        List_Select = "CreatedByAlexFare"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Adding..."
        Status
    Else
        ErrMsg_Duplicate
    End If
End Sub

'/ Clear Button
Private Sub btnClear_Click()
    Update_Button_Enable = False
    Clear_Form
    Gage_Number.SetFocus
End Sub

'/ Done Button
Private Sub Done_Button_Click()
    Unload Menu
End Sub

Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
        Gage_Number.SetFocus
    End If
End Sub

'/ Search Button
Public Sub Search_Button_Click()
    Dim ws          As Worksheet
    Dim DateEdit 'Update Last searched
    Dim Gage_Number_Save
    
    ' clear previous data from form, except "Gage Number"
    ' --------------------------------------------------------
    Gage_Number_Save = Gage_Number
    Clear_Form
    Gage_Number = Gage_Number_Save
    ' ---------------------------------------------------------
    
    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)) Then
        Update_Button_Enable = False
        ErrMsg
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
        Option4_Custom = True
        
        '/ Audit Log
        lblDateAdded = ws.Cells(r, "AK")
        lblDateEdit = ws.Cells(r, "AL")
        lblSearchedDate = DateEdit 'Update Last searched
        lastUser = ws.Cells(r, "AN")
                
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Searching..."
        Status
        
    End If
    Gage_Number.SetFocus
End Sub

'/ Update Button
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

Sub ErrMsg()
    MsgBox ("Gage Number Not Found"), , "Not Found"
    Gage_Number.SetFocus
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Gage number already in use"), , "Duplicate"
    Gage_Number.SetFocus
End Sub

Private Sub Clear_Form()
    Gage_Number = ""
    PartNumbertxt = ""
    Descriptiontxt = ""
    comboGageType = ""
    Customer = ""
    Insp_Date = ""
    Due_Date = ""
    Initials = ""
    Department = ""
    Comments = ""
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
        currrentUser = Application.userName
        lastUser = currrentUser
        ws.Cells(r, "AN") = lastUser
        
        If Option1_6 = True Then        ' option1 = 1month, option2 = 6months, option3 = 1year, option4 = custom or original
        Due_Date = Date_Due_6mos
    End If
    If Option2_12 = True Then
        Due_Date = Date_Due_1yr
    End If
    If Option3_24 = True Then
        Due_Date = Date_Due_2yr
    End If
    If Option4_Custom = True Then
        Option4_Custom_Click
        Due_Date = Date_Due
    End If
    
    ws.Cells(r, "G") = Due_Date
    
    Gage_Number.SetFocus 'Clear_Form 'Clear form after update
    
    '/Audit Log/'
    Dim UpdateCount As Integer
    
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    UpdateCount = ws.Range("B50")
    UpdateCountPlusOne = UpdateCount + 1
    ws.Range("B50") = UpdateCountPlusOne
    
    '/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    '/ End Audit Log /'
    
    '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Updating..."
        Status
Else
    MsgBox ("Must search For entry before updating"), , "Nothing To Update"
End If

'Update_Button_Enable = False 'Remove comment if you want to require searching again after an update.

End Sub

Sub MSG_Verify_Update()
    
    MSG1 = MsgBox("Are you sure you want To change the Gage ID?", vbYesNo, "Verify")
    
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

Private Sub btnLogout_Click()
    Unload Menu
    Worksheets("Login").Activate
    LoginForm.Show
    ThisWorkbook.Save
End Sub

'/Admin Panel - Bring up admin menu to edit audit dates/'
Private Sub btnAdmin_click()
    '/Add to the login count /'
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim LoginCount  As Integer
    
    Dim ws          As Worksheet
    Dim List_Select
    Dim TempLogin   As Integer
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Persistent_Login = ws.Range("B55")
    
    If Persistent_Login = "1" Then
        Unload Menu
        LoginForm.Show
    End If
    
    If Persistent_Login = "2" Then
        Sheets("CreatedByAlexFare").Activate
        Unload Menu
        AdminForm.Show
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

'/Gage R&R /'
Private Sub btnGageRR_Click()
    MsgBox "NOTE: This is a WIP preview. Calculation formula is not displaying correctly!"
    GageRnR.Show
End Sub

Private Sub Status()
    Dim startTime As Date
    Dim elapsedTime As Long
    Dim waitTimeInSeconds As Long
    
    waitTimeInSeconds = 2 'change this to the desired wait time in seconds
    
    startTime = Now
    Do While elapsedTime < waitTimeInSeconds
        DoEvents 'allow the program to process any pending events
        elapsedTime = DateDiff("s", startTime, Now)
    Loop
        statusLabel.Caption = ""
        statusLabelLog.Caption = ""
End Sub
