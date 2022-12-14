VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "GageTracker - Created By Alex Fare"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12105
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Gage Tracker
' Managed By: Alex Fare
' Rev: 3.11.0
' Updated: 01/12/2022

Dim r As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify
Dim Due_Date_Original
Dim Date_Due_6mos
Dim Date_Due_1yr
Dim Date_Due_2yr
Dim Date_Due
Dim currrentUser As String

'/Positioning /'
Private Sub UserForm_Initialize()
Dim sngLeft As Single
Dim sngTop As Single

    Call ReturnPosition_CenterScreen(Me.Height, Me.Width, sngLeft, sngTop)
    Me.Left = sngLeft
    Me.Top = sngTop
End Sub

'/Auto Due Date
Private Sub Option1_6_Click() ' auto format for 6 month interval
    Date_Due_6mos = DateAdd("m", 6, Insp_Date)
    Date_Due_6mos = Format(Date_Due_6mos, "mm/dd/yyyy")
    Due_Date = Date_Due_6mos
End Sub
Private Sub Option2_12_Click() ' auto format for 1 year interval
    Date_Due_1yr = DateAdd("yyyy", 1, Insp_Date)
    Date_Due_1yr = Format(Date_Due_1yr, "mm/dd/yyyy")
    Due_Date = Date_Due_1yr
End Sub
Private Sub Option3_24_Click() ' auto format for 2 year interval
    Date_Due_2yr = DateAdd("yyyy", 2, Insp_Date)
    Date_Due_2yr = Format(Date_Due_2yr, "mm/dd/yyyy")
    Due_Date = Date_Due_2yr
End Sub
Private Sub Option4_Custom_Click() ' formatting for either original record, or new custom date
Date_Due = Format(Due_Date, "mm/dd/yyyy")
Due_Date = Date_Due
End Sub

'/ Add Gage
Private Sub Add_Button_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "CreatedByAlexFare" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)) Then
  
    Dim lLastRow As Long    ' lLastRow = variable to store the result of the row count calculation
    lLastRow = Ws.ListObjects.Item(1).ListRows.Count
    r = lLastRow + 3 ' Add number for every header tab created
                Dim gnString As String
                    If IsNumeric(Gage_Number) Then
                        gnString = Val(Gage_Number.Value)
                    Else
                        gnString = Gage_Number
                    End If
    
    Ws.Cells(r, "A") = gnString
    Ws.Cells(r, "B") = PartNumbertxt
    Ws.Cells(r, "C") = Descriptiontxt
    Ws.Cells(r, "D") = comboGageType
    Ws.Cells(r, "E") = Customer
    Ws.Cells(r, "F") = Insp_Date
    Ws.Cells(r, "G") = Due_Date
    Ws.Cells(r, "H") = Initials
    Ws.Cells(r, "I") = Department
    Ws.Cells(r, "J") = Comments
    Ws.Cells(r, "Z") = comboStatus
    Ws.Cells(r, "AA") = aN1
    Ws.Cells(r, "AB") = aA1
    Ws.Cells(r, "AC") = aN2
    Ws.Cells(r, "AD") = aA2
    Ws.Cells(r, "AE") = aN3
    Ws.Cells(r, "AF") = aA3
    Ws.Cells(r, "AG") = aN4
    Ws.Cells(r, "AH") = aA4
    Ws.Cells(r, "AI") = aN5
    Ws.Cells(r, "AJ") = aA5
    Ws.Cells(r, "AK") = Now
    
    '/ Audit Log
    currrentUser = Application.userName
    lastUser = currrentUser
    Ws.Cells(r, "AN") = lastUser
    
    Add_Button.Caption = "Added!" ' change caption of add button for confirmation
    Application.Wait (Now + TimeValue("0:00:02")) ' Wait to avoid crash
    Add_Button.Caption = "Add"
    Clear_Form
    Gage_Number.SetFocus
    
'/Add to Gage Number count/'
    Dim AddCount As Integer

    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws

     AddCount = Ws.Range("B49")
     AddCountPlusOne = AddCount + 1
     Ws.Range("B49") = AddCountPlusOne
     
     '/Prevent Issues in the future, Call back the main page/'
     List_Select = "CreatedByAlexFare" ' Tab name
     Set Ws = Sheets(List_Select)
     Set Worksheet_Set = Ws
     
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
' clear previous data from form, except "Gage Number"
' --------------------------------------------------------
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
' ---------------------------------------------------------

Dim Ws As Worksheet

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
        Descriptiontxt = Ws.Cells(r, "C")
        comboGageType = Ws.Cells(r, "D")
        Customer = Ws.Cells(r, "E")
        Insp_Date = Ws.Cells(r, "F")
        Due_Date_Original = Ws.Cells(r, "G")
        Due_Date = Format(Due_Date_Original, "mm/dd/yyyy")
        Initials = Ws.Cells(r, "H")
        Department = Ws.Cells(r, "I")
        Comments = Ws.Cells(r, "J")
        comboStatus = Ws.Cells(r, "Z")
        aN1 = Ws.Cells(r, "AA")
        aA1 = Ws.Cells(r, "AB")
        aN2 = Ws.Cells(r, "AC")
        aA2 = Ws.Cells(r, "AD")
        aN3 = Ws.Cells(r, "AE")
        aA3 = Ws.Cells(r, "AF")
        aN4 = Ws.Cells(r, "AG")
        aA4 = Ws.Cells(r, "AH")
        aN5 = Ws.Cells(r, "AI")
        aA5 = Ws.Cells(r, "AJ")
        Ws.Cells(r, "AM") = Now 'Update Last searched
        Update_Button_Enable = True
        Option4_Custom = True
        
        '/ Audit Log
        lblDateAdded = Ws.Cells(r, "AK")
        lblDateEdit = Ws.Cells(r, "AL")
        lblSearchedDate = Ws.Cells(r, "AM")
        lastUser = Ws.Cells(r, "AN")
            
        Dim FS
        Set FS = CreateObject("Scripting.FileSystemObject")

        If FS.FileExists(TextFile_FullPath) Then
            Else
        End If
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
     MsgBox ("Must search for entry before updating"), , "Nothing to Update"
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
        lblDateAdded = "-"
        lblDateEdit = "-"
        lblSearchedDate = "-"
        lastUser = "-"
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
Ws.Cells(r, "A") = gnString
Ws.Cells(r, "B") = PartNumbertxt
Ws.Cells(r, "C") = Descriptiontxt
Ws.Cells(r, "D") = comboGageType
Ws.Cells(r, "E") = Customer
Ws.Cells(r, "F") = Insp_Date
Ws.Cells(r, "H") = Initials
Ws.Cells(r, "I") = Department
Ws.Cells(r, "J") = Comments
Ws.Cells(r, "Z") = comboStatus
Ws.Cells(r, "AA") = aN1
Ws.Cells(r, "AB") = aA1
Ws.Cells(r, "AC") = aN2
Ws.Cells(r, "AD") = aA2
Ws.Cells(r, "AE") = aN3
Ws.Cells(r, "AF") = aA3
Ws.Cells(r, "AG") = aN4
Ws.Cells(r, "AH") = aA4
Ws.Cells(r, "AI") = aN5
Ws.Cells(r, "AJ") = aA5
Ws.Cells(r, "AL") = Now 'Update Last edited

'/ Audit Log
    currrentUser = Application.userName
    lastUser = currrentUser
    Ws.Cells(r, "AN") = lastUser

If Option1_6 = True Then                ' option1 = 1month, option2 = 6months, option3 = 1year, option4 = custom or original
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
    
Ws.Cells(r, "G") = Due_Date

Update_Button.Caption = "Updated!"
Application.Wait (Now + TimeValue("0:00:02"))
Update_Button.Caption = "Update"
'Clear_Form 'Clear form after update
Gage_Number.SetFocus

'/Add to update count/'
    Dim UpdateCount As Integer

    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws

     UpdateCount = Ws.Range("B50")
     UpdateCountPlusOne = UpdateCount + 1
     Ws.Range("B50") = UpdateCountPlusOne
     
     '/Prevent Issues in the future, Call back the main page/'
     List_Select = "CreatedByAlexFare" ' Tab name
     Set Ws = Sheets(List_Select)
     Set Worksheet_Set = Ws

Else
    MsgBox ("Must search for entry before updating"), , "Nothing to Update"
End If

'Update_Button_Enable = False 'Remove comment if you want to require searching again after an update.

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
End Sub

Private Sub btnLogOut_click()
Unload Menu
Worksheets("Login").Activate
LoginForm.Show
ThisWorkbook.Save
End Sub

'/Admin Panel - Bring up admin menu to edit audit dates/'
Private Sub btnAdmin_click()
Unload Menu
LoginForm.Show
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

