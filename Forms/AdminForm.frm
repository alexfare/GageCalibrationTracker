VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminForm 
   Caption         =   "Admin Panel  - Created By Alex Fare"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120.001
   OleObjectBlob   =   "AdminForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify
Dim currrentUser As String


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
        Ws.Cells(r, "AM") = Now 'Update Last searched
        Update_Button_Enable = True
        Option4_Custom = True
        
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
     MsgBox ("Must search for entry before updating"), , "Nothing to Update"
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
Ws.Cells(r, "AL") = Now 'Update Last edited
Ws.Cells(r, "AK") = lblDateAdded 'Date Added
currrentUser = Application.userName
lastUser = currrentUser
Ws.Cells(r, "AN") = lastUser

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

Private Sub Clear_Form()
        Gage_Number = ""
        PartNumbertxt = ""
        lblDateAdded = "-"
        lblDateEdit = "-"
        lblSearchedDate = "-"
        lastUser = "-"
End Sub

Private Sub btnClear_Click()
Update_Button_Enable = False
Clear_Form
Gage_Number.SetFocus
End Sub

Sub CheckForUpdate_Click()
    Dim URL As String
    URL = "https://github.com/alexfare/GageCalibrationTracker"
    ActiveWorkbook.FollowHyperlink URL
End Sub


Private Sub btnLogOut_click()
Unload AdminForm
Menu.Show
ThisWorkbook.Save
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
    MsgBox "Code protection password is GageTracker2022"
End Sub
