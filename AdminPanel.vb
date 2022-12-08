' v1.0.0


Dim r As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify


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
        r = Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)
        GN_Verify = Gage_Number
        PartNumbertxt = Ws.Cells(r, "B")
        Descriptiontxt = Ws.Cells(r, "C")
        GageType = Ws.Cells(r, "D")
        Customer = Ws.Cells(r, "E")
        Insp_Date = Ws.Cells(r, "F")
        Due_Date_Original = Ws.Cells(r, "G")
        Due_Date = Format(Due_Date_Original, "mm/dd/yyyy")
        Initials = Ws.Cells(r, "H")
        Department = Ws.Cells(r, "I")
        Comments = Ws.Cells(r, "J")
        Statustxt = Ws.Cells(r, "Z")
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
        
        lblDateAdded = Ws.Cells(r, "AK")
        lblDateEdit = Ws.Cells(r, "AL")
        lbSearchedDate = Ws.Cells(r, "AM")
            
            
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



