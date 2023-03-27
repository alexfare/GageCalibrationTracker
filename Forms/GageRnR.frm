VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GageRnR 
   Caption         =   "Gage R&R"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "GageRnR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GageRnR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r               As Long        ' variable used for storing row number
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean        ' to store update enable flag after search
Dim GN_Verify
Dim currrentUser    As String

'/Positioning /'
Private Sub UserForm_Initialize()
    Dim ws          As Worksheet
    Dim List_Select
    List_Select = "GageRnR"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    Dim rng         As Range
    For Each rng In ws.Range("A3:A50")
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
    Dim ws          As Worksheet
    Dim List_Select
    List_Select = "GageRnR"        ' Tab name
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
        
        Add_Button.Caption = "Added!"        ' change caption of add button for confirmation
        Application.Wait (Now + TimeValue("0:00:01"))        ' Wait to avoid crash
        Add_Button.Caption = ""
        Clear_Form
        Gage_Number.SetFocus
        
        '/Add to Gage Number count/'
        Dim AddGageRnR As Integer
        
        List_Select = "Admin"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        
        AddGageRnR = ws.Range("B54")
        AddGageRnRPlusOne = AddGageRnR + 1
        ws.Range("B54") = AddGageRnRPlusOne
        
        '/Prevent Issues in the future, Call back the main page/'
        List_Select = "GageRnR"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
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
    Gage_Number_Save = Gage_Number
    Clear_Form
    Gage_Number = Gage_Number_Save
    '/ Calculation --------------------------------------------
    

    ' ---------------------------------------------------------
    
    Dim ws          As Worksheet
    
    List_Select = "GageRnR"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)) Then
        Update_Button_Enable = False
        ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)
        GN_Verify = Gage_Number
        PartNumbertxt = ws.Cells(r, "B")
        PartNametxt = ws.Cells(r, "C")
        Update_Button_Enable = True
        Option4_Custom = True
        
        '/ Gage R&R Appraiser 1 /*
        Ap1Name = ws.Cells(r, "D")
        'Trial 1
        A1T1P1 = ws.Cells(r, "E")
        A1T1P2 = ws.Cells(r, "F")
        A1T1P3 = ws.Cells(r, "G")
        A1T1P4 = ws.Cells(r, "H")
        A1T1P5 = ws.Cells(r, "I")
        A1T1P6 = ws.Cells(r, "J")
        A1T1P7 = ws.Cells(r, "K")
        A1T1P8 = ws.Cells(r, "L")
        A1T1P9 = ws.Cells(r, "M")
        A1T1P10 = ws.Cells(r, "N")
        
        'Trial 2
        A1T2P1 = ws.Cells(r, "O")
        A1T2P2 = ws.Cells(r, "P")
        A1T2P3 = ws.Cells(r, "Q")
        A1T2P4 = ws.Cells(r, "R")
        A1T2P5 = ws.Cells(r, "S")
        A1T2P6 = ws.Cells(r, "T")
        A1T2P7 = ws.Cells(r, "U")
        A1T2P8 = ws.Cells(r, "V")
        A1T2P9 = ws.Cells(r, "W")
        A1T2P10 = ws.Cells(r, "X")
        
        'Trial 3
        A1T3P1 = ws.Cells(r, "Y")
        A1T3P2 = ws.Cells(r, "Z")
        A1T3P3 = ws.Cells(r, "AA")
        A1T3P4 = ws.Cells(r, "AB")
        A1T3P5 = ws.Cells(r, "AC")
        A1T3P6 = ws.Cells(r, "AD")
        A1T3P7 = ws.Cells(r, "AE")
        A1T3P8 = ws.Cells(r, "AF")
        A1T3P9 = ws.Cells(r, "AG")
        A1T3P10 = ws.Cells(r, "AH")
        
        '/ Gage R&R Appraiser 2 /*
        Ap2Name = ws.Cells(r, "AI")
        
        'Trial 1
        A2T1P1 = ws.Cells(r, "AJ")
        A2T1P2 = ws.Cells(r, "AK")
        A2T1P3 = ws.Cells(r, "AL")
        A2T1P4 = ws.Cells(r, "AM")
        A2T1P5 = ws.Cells(r, "AN")
        A2T1P6 = ws.Cells(r, "AO")
        A2T1P7 = ws.Cells(r, "AP")
        A2T1P8 = ws.Cells(r, "AQ")
        A2T1P9 = ws.Cells(r, "AR")
        A2T1P10 = ws.Cells(r, "AS")
        
        'Trial 2
        A2T2P1 = ws.Cells(r, "AT")
        A2T2P2 = ws.Cells(r, "AU")
        A2T2P3 = ws.Cells(r, "AV")
        A2T2P4 = ws.Cells(r, "AW")
        A2T2P5 = ws.Cells(r, "AX")
        A2T2P6 = ws.Cells(r, "AY")
        A2T2P7 = ws.Cells(r, "AZ")
        A2T2P8 = ws.Cells(r, "BA")
        A2T2P9 = ws.Cells(r, "BB")
        A2T2P10 = ws.Cells(r, "BC")
        
        'Trial 3
        A2T3P1 = ws.Cells(r, "BD")
        A2T3P2 = ws.Cells(r, "BE")
        A2T3P3 = ws.Cells(r, "BF")
        A2T3P4 = ws.Cells(r, "BG")
        A2T3P5 = ws.Cells(r, "BH")
        A2T3P6 = ws.Cells(r, "BI")
        A2T3P7 = ws.Cells(r, "BJ")
        A2T3P8 = ws.Cells(r, "BK")
        A2T3P9 = ws.Cells(r, "BL")
        A2T3P10 = ws.Cells(r, "BM")
        
        '/ Gage R&R Appraiser 3 /*
        Ap3Name = ws.Cells(r, "BN")
        
        'Trial 1
        A3T1P1 = ws.Cells(r, "BO")
        A3T1P2 = ws.Cells(r, "BP")
        A3T1P3 = ws.Cells(r, "BQ")
        A3T1P4 = ws.Cells(r, "BR")
        A3T1P5 = ws.Cells(r, "BS")
        A3T1P6 = ws.Cells(r, "BT")
        A3T1P7 = ws.Cells(r, "BU")
        A3T1P8 = ws.Cells(r, "BV")
        A3T1P9 = ws.Cells(r, "BW")
        A3T1P10 = ws.Cells(r, "BX")
        
        'Trial 2
        A3T2P1 = ws.Cells(r, "BY")
        A3T2P2 = ws.Cells(r, "BZ")
        A3T2P3 = ws.Cells(r, "CA")
        A3T2P4 = ws.Cells(r, "CB")
        A3T2P5 = ws.Cells(r, "CC")
        A3T2P6 = ws.Cells(r, "CD")
        A3T2P7 = ws.Cells(r, "CE")
        A3T2P8 = ws.Cells(r, "CF")
        A3T2P9 = ws.Cells(r, "CG")
        A3T2P10 = ws.Cells(r, "CH")
        
        'Trial 3
        A3T3P1 = ws.Cells(r, "CI")
        A3T3P2 = ws.Cells(r, "CJ")
        A3T3P3 = ws.Cells(r, "CK")
        A3T3P4 = ws.Cells(r, "CL")
        A3T3P5 = ws.Cells(r, "CM")
        A3T3P6 = ws.Cells(r, "CN")
        A3T3P7 = ws.Cells(r, "CO")
        A3T3P8 = ws.Cells(r, "CP")
        A3T3P9 = ws.Cells(r, "CQ")
        A3T3P10 = ws.Cells(r, "CR")
        
        
        '/ Calculation
        Dim A1P1R As Integer
        Dim A1T1P1i As Integer
        Dim A1T2P1i As Integer
        A1T1P1i = A1T1P1
        A1T2P1i = A1T2P1
        A1P1R = A1T1P1i + A1T2P1i
        Range12 = A1P1R
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
        Set ws = Worksheet_Set
        If IsNumeric(Gage_Number) Then
            gnString = Val(Gage_Number.Value)
        Else
            gnString = Gage_Number
        End If
        '/ Audit
        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = PartNumbertxt
        
        '/ Gage R&R Appraiser 1 /*
        ws.Cells(r, "D") = Ap1Name
        'Trial 1
        ws.Cells(r, "E") = A1T1P1
        ws.Cells(r, "F") = A1T1P2
        ws.Cells(r, "G") = A1T1P3
        ws.Cells(r, "H") = A1T1P4
        ws.Cells(r, "I") = A1T1P5
        ws.Cells(r, "J") = A1T1P6
        ws.Cells(r, "K") = A1T1P7
        ws.Cells(r, "L") = A1T1P8
        ws.Cells(r, "M") = A1T1P9
        ws.Cells(r, "N") = A1T1P10
        
        'Trial 2
        ws.Cells(r, "O") = A1T2P1
        ws.Cells(r, "P") = A1T2P2
        ws.Cells(r, "Q") = A1T2P3
        ws.Cells(r, "R") = A1T2P4
        ws.Cells(r, "S") = A1T2P5
        ws.Cells(r, "T") = A1T2P6
        ws.Cells(r, "U") = A1T2P7
        ws.Cells(r, "V") = A1T2P8
        ws.Cells(r, "W") = A1T2P9
        ws.Cells(r, "X") = A1T2P10
        
        'Trial 3
        ws.Cells(r, "Y") = A1T3P1
        ws.Cells(r, "Z") = A1T3P2
        ws.Cells(r, "AA") = A1T3P3
        ws.Cells(r, "AB") = A1T3P4
        ws.Cells(r, "AC") = A1T3P5
        ws.Cells(r, "AD") = A1T3P6
        ws.Cells(r, "AE") = A1T3P7
        ws.Cells(r, "AF") = A1T3P8
        ws.Cells(r, "AG") = A1T3P9
        ws.Cells(r, "AH") = A1T3P10
        
        '/ Gage R&R Appraiser 2 /*
        ws.Cells(r, "AI") = Ap2Name
        
        'Trial 1
        ws.Cells(r, "AJ") = A2T1P1
        ws.Cells(r, "AK") = A2T1P2
        ws.Cells(r, "AL") = A2T1P3
        ws.Cells(r, "AM") = A2T1P4
        ws.Cells(r, "AN") = A2T1P5
        ws.Cells(r, "AO") = A2T1P6
        ws.Cells(r, "AP") = A2T1P7
        ws.Cells(r, "AQ") = A2T1P8
        ws.Cells(r, "AR") = A2T1P9
        ws.Cells(r, "AS") = A2T1P10
        
        'Trial 2
        ws.Cells(r, "AT") = A2T2P1
        ws.Cells(r, "AU") = A2T2P2
        ws.Cells(r, "AV") = A2T2P3
        ws.Cells(r, "AW") = A2T2P4
        ws.Cells(r, "AX") = A2T2P5
        ws.Cells(r, "AY") = A2T2P6
        ws.Cells(r, "AZ") = A2T2P7
        ws.Cells(r, "BA") = A2T2P8
        ws.Cells(r, "BB") = A2T2P9
        ws.Cells(r, "BC") = A2T2P10
        
        'Trial 3
        ws.Cells(r, "BD") = A2T3P1
        ws.Cells(r, "BE") = A2T3P2
        ws.Cells(r, "BF") = A2T3P3
        ws.Cells(r, "BG") = A2T3P4
        ws.Cells(r, "BH") = A2T3P5
        ws.Cells(r, "BI") = A2T3P6
        ws.Cells(r, "BJ") = A2T3P7
        ws.Cells(r, "BK") = A2T3P8
        ws.Cells(r, "BL") = A2T3P9
        ws.Cells(r, "BM") = A2T3P10
        
        '/ Gage R&R Appraiser 3 /*
        ws.Cells(r, "BN") = Ap3Name
        
        'Trial 1
        ws.Cells(r, "BO") = A3T1P1
        ws.Cells(r, "BP") = A3T1P2
        ws.Cells(r, "BQ") = A3T1P3
        ws.Cells(r, "BR") = A3T1P4
        ws.Cells(r, "BS") = A3T1P5
        ws.Cells(r, "BT") = A3T1P6
        ws.Cells(r, "BU") = A3T1P7
        ws.Cells(r, "BV") = A3T1P8
        ws.Cells(r, "BW") = A3T1P9
        ws.Cells(r, "BX") = A3T1P10
        
        'Trial 2
        ws.Cells(r, "BY") = A3T2P1
        ws.Cells(r, "BZ") = A3T2P2
        ws.Cells(r, "CA") = A3T2P3
        ws.Cells(r, "CB") = A3T2P4
        ws.Cells(r, "CC") = A3T2P5
        ws.Cells(r, "CD") = A3T2P6
        ws.Cells(r, "CE") = A3T2P7
        ws.Cells(r, "CF") = A3T2P8
        ws.Cells(r, "CG") = A3T2P9
        ws.Cells(r, "CH") = A3T2P10
        
        'Trial 3
        ws.Cells(r, "CI") = A3T3P1
        ws.Cells(r, "CJ") = A3T3P2
        ws.Cells(r, "CK") = A3T3P3
        ws.Cells(r, "CL") = A3T3P4
        ws.Cells(r, "CM") = A3T3P5
        ws.Cells(r, "CN") = A3T3P6
        ws.Cells(r, "CO") = A3T3P7
        ws.Cells(r, "CP") = A3T3P8
        ws.Cells(r, "CQ") = A3T3P9
        ws.Cells(r, "CR") = A3T3P10
        
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
