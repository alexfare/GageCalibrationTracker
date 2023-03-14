Dim r               As Long        ' variable used for storing row number
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean        ' to store update enable flag after search
Dim GN_Verify
Dim currrentUser    As String

'/Positioning /'
Private Sub UserForm_Initialize()
    Dim Ws          As Worksheet
    Dim List_Select
    List_Select = "GageRnR"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    Dim rng         As Range
    For Each rng In Ws.Range("A3:A50")
    Me.GageRnR_List.AddItem rng.Value
    Next rng
    
End Sub

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

'/ Add Gage
Private Sub Add_Button_Click()
    Dim Ws          As Worksheet
    Dim List_Select
    List_Select = "GageRnR"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)) Then
        
        Dim lLastRow As Long        ' lLastRow = variable to store the result of the row count calculation
        lLastRow = Ws.ListObjects.Item(1).ListRows.Count
        r = lLastRow + 3        ' Add number for every header tab created
        Dim gnString As String
        If IsNumeric(Gage_Number) Then
            gnString = Val(Gage_Number.Value)
        Else
            gnString = Gage_Number
        End If
        
        Ws.Cells(r, "A") = gnString
        Ws.Cells(r, "B") = PartNumbertxt
        
        Add_Button.Caption = "Added!"        ' change caption of add button for confirmation
        Application.Wait (Now + TimeValue("0:00:01"))        ' Wait to avoid crash
        Add_Button.Caption = ""
        Clear_Form
        Gage_Number.SetFocus
        
        '/Add to Gage Number count/'
        Dim AddGageRnR As Integer
        
        List_Select = "Admin"        ' Tab name
        Set Ws = Sheets(List_Select)
        Set Worksheet_Set = Ws
        
        AddGageRnR = Ws.Range("B54")
        AddGageRnRPlusOne = AddGageRnR + 1
        Ws.Range("B54") = AddGageRnRPlusOne
        
        '/Prevent Issues in the future, Call back the main page/'
        List_Select = "GageRnR"        ' Tab name
        Set Ws = Sheets(List_Select)
        Set Worksheet_Set = Ws
    Else
        ErrMsg_Duplicate
    End If
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
    Ap2Name = ""
    Ap3Name = ""
    
    '/ Gage R&R Appraiser 1 /*
    Ap1Name = ""
    
    'Trial 1
    A1T1P1 = ""
    A1T1P2 = ""
    A1T1P3 = ""
    A1T1P4 = ""
    A1T1P5 = ""
    A1T1P6 = ""
    A1T1P7 = ""
    A1T1P8 = ""
    A1T1P9 = ""
    A1T1P10 = ""
    
    'Trial 2
    A1T2P1 = ""
    A1T2P2 = ""
    A1T2P3 = ""
    A1T2P4 = ""
    A1T2P5 = ""
    A1T2P6 = ""
    A1T2P7 = ""
    A1T2P8 = ""
    A1T2P9 = ""
    A1T2P10 = ""
    
    'Trial 3
    A1T2P1 = ""
    A1T2P2 = ""
    A1T2P3 = ""
    A1T2P4 = ""
    A1T2P5 = ""
    A1T2P6 = ""
    A1T2P7 = ""
    A1T2P8 = ""
    A1T2P9 = ""
    A1T2P10 = ""
    
    '/ Gage R&R Appraiser 2 /*
    Ap2Name = ""
    
    'Trial 1
    A2T1P1 = ""
    A2T1P2 = ""
    A2T1P3 = ""
    A2T1P4 = ""
    A2T1P5 = ""
    A2T1P6 = ""
    A2T1P7 = ""
    A2T1P8 = ""
    A2T1P9 = ""
    A2T1P10 = ""
    
    'Trial 2
    A2T2P1 = ""
    A2T2P2 = ""
    A2T2P3 = ""
    A2T2P4 = ""
    A2T2P5 = ""
    A2T2P6 = ""
    A2T2P7 = ""
    A2T2P8 = ""
    A2T2P9 = ""
    A2T2P10 = ""
    
    'Trial 3
    A2T3P1 = ""
    A2T3P2 = ""
    A2T3P3 = ""
    A2T3P4 = ""
    A2T3P5 = ""
    A2T3P6 = ""
    A2T3P7 = ""
    A2T3P8 = ""
    A2T3P9 = ""
    A2T3P10 = ""
    
    '/ Gage R&R Appraiser 3 /*
    Ap3Name = ""
    
    'Trial 1
    A3T1P1 = ""
    A3T1P2 = ""
    A3T1P3 = ""
    A3T1P4 = ""
    A3T1P5 = ""
    A3T1P6 = ""
    A3T1P7 = ""
    A3T1P8 = ""
    A3T1P9 = ""
    A3T1P10 = ""
    
    'Trial 2
    A3T2P1 = ""
    A3T2P2 = ""
    A3T2P3 = ""
    A3T2P4 = ""
    A3T2P5 = ""
    A3T2P6 = ""
    A3T2P7 = ""
    A3T2P8 = ""
    A3T2P9 = ""
    A3T2P10 = ""
    
    'Trial 3
    A3T3P1 = ""
    A3T3P2 = ""
    A3T3P3 = ""
    A3T3P4 = ""
    A3T3P5 = ""
    A3T3P6 = ""
    A3T3P7 = ""
    A3T3P8 = ""
    A3T3P9 = ""
    A3T3P10 = ""
    
    '/ Calculation --------------------------------------------
    

    ' ---------------------------------------------------------
    
    Dim Ws          As Worksheet
    
    List_Select = "GageRnR"
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)) Then
        Update_Button_Enable = False
        ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), Ws.Columns(1), 0)
        GN_Verify = Gage_Number
        PartNumbertxt = Ws.Cells(r, "B")
        PartNametxt = Ws.Cells(r, "C")
        Update_Button_Enable = True
        Option4_Custom = True
        
        '/ Gage R&R Appraiser 1 /*
        Ap1Name = Ws.Cells(r, "D")
        'Trial 1
        A1T1P1 = Ws.Cells(r, "E")
        A1T1P2 = Ws.Cells(r, "F")
        A1T1P3 = Ws.Cells(r, "G")
        A1T1P4 = Ws.Cells(r, "H")
        A1T1P5 = Ws.Cells(r, "I")
        A1T1P6 = Ws.Cells(r, "J")
        A1T1P7 = Ws.Cells(r, "K")
        A1T1P8 = Ws.Cells(r, "L")
        A1T1P9 = Ws.Cells(r, "M")
        A1T1P10 = Ws.Cells(r, "N")
        
        'Trial 2
        A1T2P1 = Ws.Cells(r, "O")
        A1T2P2 = Ws.Cells(r, "P")
        A1T2P3 = Ws.Cells(r, "Q")
        A1T2P4 = Ws.Cells(r, "R")
        A1T2P5 = Ws.Cells(r, "S")
        A1T2P6 = Ws.Cells(r, "T")
        A1T2P7 = Ws.Cells(r, "U")
        A1T2P8 = Ws.Cells(r, "V")
        A1T2P9 = Ws.Cells(r, "W")
        A1T2P10 = Ws.Cells(r, "X")
        
        'Trial 3
        A1T3P1 = Ws.Cells(r, "Y")
        A1T3P2 = Ws.Cells(r, "Z")
        A1T3P3 = Ws.Cells(r, "AA")
        A1T3P4 = Ws.Cells(r, "AB")
        A1T3P5 = Ws.Cells(r, "AC")
        A1T3P6 = Ws.Cells(r, "AD")
        A1T3P7 = Ws.Cells(r, "AE")
        A1T3P8 = Ws.Cells(r, "AF")
        A1T3P9 = Ws.Cells(r, "AG")
        A1T3P10 = Ws.Cells(r, "AH")
        
        '/ Gage R&R Appraiser 2 /*
        Ap2Name = Ws.Cells(r, "AI")
        
        'Trial 1
        A2T1P1 = Ws.Cells(r, "AJ")
        A2T1P2 = Ws.Cells(r, "AK")
        A2T1P3 = Ws.Cells(r, "AL")
        A2T1P4 = Ws.Cells(r, "AM")
        A2T1P5 = Ws.Cells(r, "AN")
        A2T1P6 = Ws.Cells(r, "AO")
        A2T1P7 = Ws.Cells(r, "AP")
        A2T1P8 = Ws.Cells(r, "AQ")
        A2T1P9 = Ws.Cells(r, "AR")
        A2T1P10 = Ws.Cells(r, "AS")
        
        'Trial 2
        A2T2P1 = Ws.Cells(r, "AT")
        A2T2P2 = Ws.Cells(r, "AU")
        A2T2P3 = Ws.Cells(r, "AV")
        A2T2P4 = Ws.Cells(r, "AW")
        A2T2P5 = Ws.Cells(r, "AX")
        A2T2P6 = Ws.Cells(r, "AY")
        A2T2P7 = Ws.Cells(r, "AZ")
        A2T2P8 = Ws.Cells(r, "BA")
        A2T2P9 = Ws.Cells(r, "BB")
        A2T2P10 = Ws.Cells(r, "BC")
        
        'Trial 3
        A2T3P1 = Ws.Cells(r, "BD")
        A2T3P2 = Ws.Cells(r, "BE")
        A2T3P3 = Ws.Cells(r, "BF")
        A2T3P4 = Ws.Cells(r, "BG")
        A2T3P5 = Ws.Cells(r, "BH")
        A2T3P6 = Ws.Cells(r, "BI")
        A2T3P7 = Ws.Cells(r, "BJ")
        A2T3P8 = Ws.Cells(r, "BK")
        A2T3P9 = Ws.Cells(r, "BL")
        A2T3P10 = Ws.Cells(r, "BM")
        
        '/ Gage R&R Appraiser 3 /*
        Ap3Name = Ws.Cells(r, "BN")
        
        'Trial 1
        A3T1P1 = Ws.Cells(r, "BO")
        A3T1P2 = Ws.Cells(r, "BP")
        A3T1P3 = Ws.Cells(r, "BQ")
        A3T1P4 = Ws.Cells(r, "BR")
        A3T1P5 = Ws.Cells(r, "BS")
        A3T1P6 = Ws.Cells(r, "BT")
        A3T1P7 = Ws.Cells(r, "BU")
        A3T1P8 = Ws.Cells(r, "BV")
        A3T1P9 = Ws.Cells(r, "BW")
        A3T1P10 = Ws.Cells(r, "BX")
        
        'Trial 2
        A3T2P1 = Ws.Cells(r, "BY")
        A3T2P2 = Ws.Cells(r, "BZ")
        A3T2P3 = Ws.Cells(r, "CA")
        A3T2P4 = Ws.Cells(r, "CB")
        A3T2P5 = Ws.Cells(r, "CC")
        A3T2P6 = Ws.Cells(r, "CD")
        A3T2P7 = Ws.Cells(r, "CE")
        A3T2P8 = Ws.Cells(r, "CF")
        A3T2P9 = Ws.Cells(r, "CG")
        A3T2P10 = Ws.Cells(r, "CH")
        
        'Trial 3
        A3T3P1 = Ws.Cells(r, "CI")
        A3T3P2 = Ws.Cells(r, "CJ")
        A3T3P3 = Ws.Cells(r, "CK")
        A3T3P4 = Ws.Cells(r, "CL")
        A3T3P5 = Ws.Cells(r, "CM")
        A3T3P6 = Ws.Cells(r, "CN")
        A3T3P7 = Ws.Cells(r, "CO")
        A3T3P8 = Ws.Cells(r, "CP")
        A3T3P9 = Ws.Cells(r, "CQ")
        A3T3P10 = Ws.Cells(r, "CR")
        
        
        '/ Calculation
        Dim A1P1R As Integer
        Dim A1T1P1i As Integer
        Dim A1T2P1i As Integer
        A1T1P1i = A1T1P1
        A1T2P1i = A1T2P1
        A1P1R = A1T1P1i + A1T2P1i
        Range12 = A1P1R
        
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
        
        '/ Gage R&R Appraiser 1 /*
        Ws.Cells(r, "D") = Ap1Name
        'Trial 1
        Ws.Cells(r, "E") = A1T1P1
        Ws.Cells(r, "F") = A1T1P2
        Ws.Cells(r, "G") = A1T1P3
        Ws.Cells(r, "H") = A1T1P4
        Ws.Cells(r, "I") = A1T1P5
        Ws.Cells(r, "J") = A1T1P6
        Ws.Cells(r, "K") = A1T1P7
        Ws.Cells(r, "L") = A1T1P8
        Ws.Cells(r, "M") = A1T1P9
        Ws.Cells(r, "N") = A1T1P10
        
        'Trial 2
        Ws.Cells(r, "O") = A1T2P1
        Ws.Cells(r, "P") = A1T2P2
        Ws.Cells(r, "Q") = A1T2P3
        Ws.Cells(r, "R") = A1T2P4
        Ws.Cells(r, "S") = A1T2P5
        Ws.Cells(r, "T") = A1T2P6
        Ws.Cells(r, "U") = A1T2P7
        Ws.Cells(r, "V") = A1T2P8
        Ws.Cells(r, "W") = A1T2P9
        Ws.Cells(r, "X") = A1T2P10
        
        'Trial 3
        Ws.Cells(r, "Y") = A1T3P1
        Ws.Cells(r, "Z") = A1T3P2
        Ws.Cells(r, "AA") = A1T3P3
        Ws.Cells(r, "AB") = A1T3P4
        Ws.Cells(r, "AC") = A1T3P5
        Ws.Cells(r, "AD") = A1T3P6
        Ws.Cells(r, "AE") = A1T3P7
        Ws.Cells(r, "AF") = A1T3P8
        Ws.Cells(r, "AG") = A1T3P9
        Ws.Cells(r, "AH") = A1T3P10
        
        '/ Gage R&R Appraiser 2 /*
        Ws.Cells(r, "AI") = Ap2Name
        
        'Trial 1
        Ws.Cells(r, "AJ") = A2T1P1
        Ws.Cells(r, "AK") = A2T1P2
        Ws.Cells(r, "AL") = A2T1P3
        Ws.Cells(r, "AM") = A2T1P4
        Ws.Cells(r, "AN") = A2T1P5
        Ws.Cells(r, "AO") = A2T1P6
        Ws.Cells(r, "AP") = A2T1P7
        Ws.Cells(r, "AQ") = A2T1P8
        Ws.Cells(r, "AR") = A2T1P9
        Ws.Cells(r, "AS") = A2T1P10
        
        'Trial 2
        Ws.Cells(r, "AT") = A2T2P1
        Ws.Cells(r, "AU") = A2T2P2
        Ws.Cells(r, "AV") = A2T2P3
        Ws.Cells(r, "AW") = A2T2P4
        Ws.Cells(r, "AX") = A2T2P5
        Ws.Cells(r, "AY") = A2T2P6
        Ws.Cells(r, "AZ") = A2T2P7
        Ws.Cells(r, "BA") = A2T2P8
        Ws.Cells(r, "BB") = A2T2P9
        Ws.Cells(r, "BC") = A2T2P10
        
        'Trial 3
        Ws.Cells(r, "BD") = A2T3P1
        Ws.Cells(r, "BE") = A2T3P2
        Ws.Cells(r, "BF") = A2T3P3
        Ws.Cells(r, "BG") = A2T3P4
        Ws.Cells(r, "BH") = A2T3P5
        Ws.Cells(r, "BI") = A2T3P6
        Ws.Cells(r, "BJ") = A2T3P7
        Ws.Cells(r, "BK") = A2T3P8
        Ws.Cells(r, "BL") = A2T3P9
        Ws.Cells(r, "BM") = A2T3P10
        
        '/ Gage R&R Appraiser 3 /*
        Ws.Cells(r, "BN") = Ap3Name
        
        'Trial 1
        Ws.Cells(r, "BO") = A3T1P1
        Ws.Cells(r, "BP") = A3T1P2
        Ws.Cells(r, "BQ") = A3T1P3
        Ws.Cells(r, "BR") = A3T1P4
        Ws.Cells(r, "BS") = A3T1P5
        Ws.Cells(r, "BT") = A3T1P6
        Ws.Cells(r, "BU") = A3T1P7
        Ws.Cells(r, "BV") = A3T1P8
        Ws.Cells(r, "BW") = A3T1P9
        Ws.Cells(r, "BX") = A3T1P10
        
        'Trial 2
        Ws.Cells(r, "BY") = A3T2P1
        Ws.Cells(r, "BZ") = A3T2P2
        Ws.Cells(r, "CA") = A3T2P3
        Ws.Cells(r, "CB") = A3T2P4
        Ws.Cells(r, "CC") = A3T2P5
        Ws.Cells(r, "CD") = A3T2P6
        Ws.Cells(r, "CE") = A3T2P7
        Ws.Cells(r, "CF") = A3T2P8
        Ws.Cells(r, "CG") = A3T2P9
        Ws.Cells(r, "CH") = A3T2P10
        
        'Trial 3
        Ws.Cells(r, "CI") = A3T3P1
        Ws.Cells(r, "CJ") = A3T3P2
        Ws.Cells(r, "CK") = A3T3P3
        Ws.Cells(r, "CL") = A3T3P4
        Ws.Cells(r, "CM") = A3T3P5
        Ws.Cells(r, "CN") = A3T3P6
        Ws.Cells(r, "CO") = A3T3P7
        Ws.Cells(r, "CP") = A3T3P8
        Ws.Cells(r, "CQ") = A3T3P9
        Ws.Cells(r, "CR") = A3T3P10
        
        Update_Button.Caption = "Updated!"
        Application.Wait (Now + TimeValue("0:00:01"))
        Update_Button.Caption = ""
        'Clear_Form 'Clear form after update
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
    '/ Gage R&R Appraiser 1 /*
    Ap1Name = ""
    
    'Trial 1
    A1T1P1 = ""
    A1T1P2 = ""
    A1T1P3 = ""
    A1T1P4 = ""
    A1T1P5 = ""
    A1T1P6 = ""
    A1T1P7 = ""
    A1T1P8 = ""
    A1T1P9 = ""
    A1T1P10 = ""
    
    'Trial 2
    A1T2P1 = ""
    A1T2P2 = ""
    A1T2P3 = ""
    A1T2P4 = ""
    A1T2P5 = ""
    A1T2P6 = ""
    A1T2P7 = ""
    A1T2P8 = ""
    A1T2P9 = ""
    A1T2P10 = ""
    
    'Trial 3
    A1T3P1 = ""
    A1T3P2 = ""
    A1T3P3 = ""
    A1T3P4 = ""
    A1T3P5 = ""
    A1T3P6 = ""
    A1T3P7 = ""
    A1T3P8 = ""
    A1T3P9 = ""
    A1T3P10 = ""
    
    '/ Gage R&R Appraiser 2 /*
    Ap2Name = ""
    
    'Trial 1
    A2T1P1 = ""
    A2T1P2 = ""
    A2T1P3 = ""
    A2T1P4 = ""
    A2T1P5 = ""
    A2T1P6 = ""
    A2T1P7 = ""
    A2T1P8 = ""
    A2T1P9 = ""
    A2T1P10 = ""
    
    'Trial 2
    A2T2P1 = ""
    A2T2P2 = ""
    A2T2P3 = ""
    A2T2P4 = ""
    A2T2P5 = ""
    A2T2P6 = ""
    A2T2P7 = ""
    A2T2P8 = ""
    A2T2P9 = ""
    A2T2P10 = ""
    
    'Trial 3
    A2T3P1 = ""
    A2T3P2 = ""
    A2T3P3 = ""
    A2T3P4 = ""
    A2T3P5 = ""
    A2T3P6 = ""
    A2T3P7 = ""
    A2T3P8 = ""
    A2T3P9 = ""
    A2T3P10 = ""
    
    '/ Gage R&R Appraiser 3 /*
    Ap3Name = ""
    
    'Trial 1
    A3T1P1 = ""
    A3T1P2 = ""
    A3T1P3 = ""
    A3T1P4 = ""
    A3T1P5 = ""
    A3T1P6 = ""
    A3T1P7 = ""
    A3T1P8 = ""
    A3T1P9 = ""
    A3T1P10 = ""
    
    'Trial 2
    A3T2P1 = ""
    A3T2P2 = ""
    A3T2P3 = ""
    A3T2P4 = ""
    A3T2P5 = ""
    A3T2P6 = ""
    A3T2P7 = ""
    A3T2P8 = ""
    A3T2P9 = ""
    A3T2P10 = ""
    
    'Trial 3
    A3T3P1 = ""
    A3T3P2 = ""
    A3T3P3 = ""
    A3T3P4 = ""
    A3T3P5 = ""
    A3T3P6 = ""
    A3T3P7 = ""
    A3T3P8 = ""
    A3T3P9 = ""
    A3T3P10 = ""
End Sub

Private Sub btnClear_Click()
    Update_Button_Enable = False
    Clear_Form
    Gage_Number.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload GageRnR
End Sub
