' Gage Tracker
' Managed By: Alex Fare
' Rev: 3.9.1
' Updated: 12/20/2022

Dim R As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify
Dim Due_Date_Original
Dim Date_Due_6mos
Dim Date_Due_1yr
Dim Date_Due_2yr
Dim Date_Due

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

Private Sub Add_Button_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "CreatedByAlexFare" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)) Then
  
    Dim lLastRow As Long    ' lLastRow = variable to store the result of the row count calculation
    lLastRow = Ws.ListObjects.Item(1).ListRows.Count
    R = lLastRow + 3 ' Add number for every header tab created
    
                Dim gnString As String
                    If IsNumeric(Gage_Number) Then
                        gnString = Val(Gage_Number.Value)
                    Else
                        gnString = Gage_Number
                    End If
    
    Ws.Cells(R, "A") = gnString
    Ws.Cells(R, "B") = PartNumbertxt
    Ws.Cells(R, "C") = Descriptiontxt
    Ws.Cells(R, "D") = GageType
    Ws.Cells(R, "E") = Customer
    Ws.Cells(R, "F") = Insp_Date
    Ws.Cells(R, "G") = Due_Date
    Ws.Cells(R, "H") = Initials
    Ws.Cells(R, "I") = Department
    Ws.Cells(R, "J") = Comments
    Ws.Cells(R, "Z") = Statustxt
    Ws.Cells(R, "AA") = aN1
    Ws.Cells(R, "AB") = aA1
    Ws.Cells(R, "AC") = aN2
    Ws.Cells(R, "AD") = aA2
    Ws.Cells(R, "AE") = aN3
    Ws.Cells(R, "AF") = aA3
    Ws.Cells(R, "AG") = aN4
    Ws.Cells(R, "AH") = aA4
    Ws.Cells(R, "AI") = aN5
    Ws.Cells(R, "AJ") = aA5
    Ws.Cells(R, "AK") = Now
    
    Add_Button.Caption = "Added!" ' change caption of add button for confirmation
    Application.Wait (Now + TimeValue("0:00:02")) ' Wait to avoid crash
    Add_Button.Caption = "Add"
    Clear_Form
    Gage_Number.SetFocus
    Else
        ErrMsg_Duplicate
    End If
    
End Sub

Private Sub btnClear_Click()
Update_Button_Enable = False
Clear_Form
Gage_Number.SetFocus
End Sub

Private Sub Done_Button_Click()
Unload UserForm1
End Sub

Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
        Insp_Date.SetFocus
    End If
End Sub

Public Sub Search_Button_Click()

' clear previous data from form, except "Gage Number"
' --------------------------------------------------------
        PartNumbertxt = ""
        Descriptiontxt = ""
        GageType = ""
        Customer = ""
        Insp_Date = ""
        Due_Date = ""
        Initials = ""
        Department = ""
        Comments = ""
        Statustxt = ""
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
        lbSearchedDate = ""
        
' ---------------------------------------------------------

Dim Ws As Worksheet

List_Select = "CreatedByAlexFare"
Set Ws = Sheets(List_Select)
Set Worksheet_Set = Ws




    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)) Then
            Update_Button_Enable = False
            ErrMsg
    Else
        R = Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)
        GN_Verify = Gage_Number
        PartNumbertxt = Ws.Cells(R, "B")
        Descriptiontxt = Ws.Cells(R, "C")
        GageType = Ws.Cells(R, "D")
        Customer = Ws.Cells(R, "E")
        Insp_Date = Ws.Cells(R, "F")
        Due_Date_Original = Ws.Cells(R, "G")
        Due_Date = Format(Due_Date_Original, "mm/dd/yyyy")
        Initials = Ws.Cells(R, "H")
        Department = Ws.Cells(R, "I")
        Comments = Ws.Cells(R, "J")
        Statustxt = Ws.Cells(R, "Z")
        aN1 = Ws.Cells(R, "AA")
        aA1 = Ws.Cells(R, "AB")
        aN2 = Ws.Cells(R, "AC")
        aA2 = Ws.Cells(R, "AD")
        aN3 = Ws.Cells(R, "AE")
        aA3 = Ws.Cells(R, "AF")
        aN4 = Ws.Cells(R, "AG")
        aA4 = Ws.Cells(R, "AH")
        aN5 = Ws.Cells(R, "AI")
        aA5 = Ws.Cells(R, "AJ")
        Ws.Cells(R, "AM") = Now 'Update Last searched
        Update_Button_Enable = True
        Option4_Custom = True
        
        lblDateAdded = Ws.Cells(R, "AK")
        lblDateEdit = Ws.Cells(R, "AL")
        lbSearchedDate = Ws.Cells(R, "AM")
            
            
        Dim FS
        Set FS = CreateObject("Scripting.FileSystemObject")

        If FS.FileExists(TextFile_FullPath) Then
            Else
        End If
    End If

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
        GageType = ""
        Customer = ""
        Insp_Date = ""
        Due_Date = ""
        Initials = ""
        Department = ""
        Comments = ""
        Statustxt = ""
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
        lbSearchedDate = "-"
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
Ws.Cells(R, "A") = gnString
Ws.Cells(R, "B") = PartNumbertxt
Ws.Cells(R, "C") = Descriptiontxt
Ws.Cells(R, "D") = GageType
Ws.Cells(R, "E") = Customer
Ws.Cells(R, "F") = Insp_Date
Ws.Cells(R, "H") = Initials
Ws.Cells(R, "I") = Department
Ws.Cells(R, "J") = Comments
Ws.Cells(R, "Z") = Statustxt
Ws.Cells(R, "AA") = aN1
Ws.Cells(R, "AB") = aA1
Ws.Cells(R, "AC") = aN2
Ws.Cells(R, "AD") = aA2
Ws.Cells(R, "AE") = aN3
Ws.Cells(R, "AF") = aA3
Ws.Cells(R, "AG") = aN4
Ws.Cells(R, "AH") = aA4
Ws.Cells(R, "AI") = aN5
Ws.Cells(R, "AJ") = aA5
Ws.Cells(R, "AL") = Now 'Update Last edited

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
    
Ws.Cells(R, "G") = Due_Date

Update_Button.Caption = "Updated!"
Application.Wait (Now + TimeValue("0:00:02"))
Update_Button.Caption = "Update"
'Clear_Form 'Clear form after update
Gage_Number.SetFocus

Else
    MsgBox ("Must search for entry before updating"), , "Nothing to Update"
    
End If

'Update_Button_Enable = False 'Remove ' if you want to require searching again after an update.

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
Unload UserForm1
Worksheets("Login").Activate
LoginForm.Show
ThisWorkbook.Save
End Sub

'/Admin Panel - Bring up admin menu to edit audit dates/'
Private Sub btnAdmin_click()
Unload UserForm1
LoginForm.Show
End Sub

'/Report Issue Panel /'
Private Sub btnReportIssue_click()
Unload UserForm1
ReportIssue.Show
End Sub

