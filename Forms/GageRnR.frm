VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GageRnR 
   Caption         =   "Gage R&R"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7245
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
    
    'Dim rng         As Range
    'For Each rng In ws.Range("A3:A50")
        'Me.GageRnR_List.AddItem rng.Value
    'Next rng
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
        
        '/ Gage R&R Appraiser 1 /*
        ws.Cells(r, "D") = Ap1Name
        'Trial 1
        ws.Cells(r, "E") = A1T1P1
        ws.Cells(r, "F") = A1T1P2
        ws.Cells(r, "G") = A1T1P3
        ws.Cells(r, "H") = A1T1P4
        ws.Cells(r, "I") = A1T1P5
        
        'Trial 2
        ws.Cells(r, "J") = A1T2P1
        ws.Cells(r, "K") = A1T2P2
        ws.Cells(r, "L") = A1T2P3
        ws.Cells(r, "M") = A1T2P4
        ws.Cells(r, "N") = A1T2P5
        
        'Trial 3
        ws.Cells(r, "O") = A1T3P1
        ws.Cells(r, "P") = A1T3P2
        ws.Cells(r, "Q") = A1T3P3
        ws.Cells(r, "R") = A1T3P4
        ws.Cells(r, "S") = A1T3P5
        
        '/ Gage R&R Appraiser 2 /*
        ws.Cells(r, "T") = Ap2Name
        
        'Trial 1
        ws.Cells(r, "U") = A2T1P1
        ws.Cells(r, "V") = A2T1P2
        ws.Cells(r, "W") = A2T1P3
        ws.Cells(r, "X") = A2T1P4
        ws.Cells(r, "Y") = A2T1P5
        
        'Trial 2
        ws.Cells(r, "Z") = A2T2P1
        ws.Cells(r, "AA") = A2T2P2
        ws.Cells(r, "AB") = A2T2P3
        ws.Cells(r, "AC") = A2T2P4
        ws.Cells(r, "AD") = A2T2P5
        
        'Trial 3
        ws.Cells(r, "AE") = A2T3P1
        ws.Cells(r, "AF") = A2T3P2
        ws.Cells(r, "AG") = A2T3P3
        ws.Cells(r, "AH") = A2T3P4
        ws.Cells(r, "AI") = A2T3P5
        
        '/ Gage R&R Appraiser 3 /*
        ws.Cells(r, "AJ") = Ap3Name
        
        'Trial 1
        ws.Cells(r, "AK") = A3T1P1
        ws.Cells(r, "AL") = A3T1P2
        ws.Cells(r, "AM") = A3T1P3
        ws.Cells(r, "AN") = A3T1P4
        ws.Cells(r, "AO") = A3T1P5
        
        'Trial 2
        ws.Cells(r, "AP") = A3T2P1
        ws.Cells(r, "AQ") = A3T2P2
        ws.Cells(r, "AR") = A3T2P3
        ws.Cells(r, "AS") = A3T2P4
        ws.Cells(r, "AT") = A3T2P5
        
        'Trial 3
        ws.Cells(r, "AU") = A3T3P1
        ws.Cells(r, "AV") = A3T3P2
        ws.Cells(r, "AW") = A3T3P3
        ws.Cells(r, "AX") = A3T3P4
        ws.Cells(r, "AY") = A3T3P5
        
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

        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Adding..."
        Status
        
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
        
        'Trial 2
        A1T2P1 = ws.Cells(r, "J")
        A1T2P2 = ws.Cells(r, "K")
        A1T2P3 = ws.Cells(r, "L")
        A1T2P4 = ws.Cells(r, "M")
        A1T2P5 = ws.Cells(r, "N")
        
        'Trial 3
        A1T3P1 = ws.Cells(r, "O")
        A1T3P2 = ws.Cells(r, "P")
        A1T3P3 = ws.Cells(r, "Q")
        A1T3P4 = ws.Cells(r, "R")
        A1T3P5 = ws.Cells(r, "S")
        
        '/ Gage R&R Appraiser 2 /*
        Ap2Name = ws.Cells(r, "T")
        
        'Trial 1
        A2T1P1 = ws.Cells(r, "U")
        A2T1P2 = ws.Cells(r, "V")
        A2T1P3 = ws.Cells(r, "W")
        A2T1P4 = ws.Cells(r, "X")
        A2T1P5 = ws.Cells(r, "Y")
        
        'Trial 2
        A2T2P1 = ws.Cells(r, "Z")
        A2T2P2 = ws.Cells(r, "AA")
        A2T2P3 = ws.Cells(r, "AB")
        A2T2P4 = ws.Cells(r, "AC")
        A2T2P5 = ws.Cells(r, "AD")
        
        'Trial 3
        A2T3P1 = ws.Cells(r, "AE")
        A2T3P2 = ws.Cells(r, "AF")
        A2T3P3 = ws.Cells(r, "AG")
        A2T3P4 = ws.Cells(r, "AH")
        A2T3P5 = ws.Cells(r, "AI")
        
        '/ Gage R&R Appraiser 3 /*
        Ap3Name = ws.Cells(r, "AJ")
        
        'Trial 1
        A3T1P1 = ws.Cells(r, "AK")
        A3T1P2 = ws.Cells(r, "AL")
        A3T1P3 = ws.Cells(r, "AM")
        A3T1P4 = ws.Cells(r, "AN")
        A3T1P5 = ws.Cells(r, "AO")
        
        'Trial 2
        A3T2P1 = ws.Cells(r, "AP")
        A3T2P2 = ws.Cells(r, "AQ")
        A3T2P3 = ws.Cells(r, "AR")
        A3T2P4 = ws.Cells(r, "AS")
        A3T2P5 = ws.Cells(r, "AT")
        
        'Trial 3
        A3T3P1 = ws.Cells(r, "AU")
        A3T3P2 = ws.Cells(r, "AV")
        A3T3P3 = ws.Cells(r, "AW")
        A3T3P4 = ws.Cells(r, "AX")
        A3T3P5 = ws.Cells(r, "AY")
        
        '/ Calculation
        'GageRnR
        List_Select = "Calculations"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        'A1 Trial 1
        ws.Range("C3") = A1T1P1
        ws.Range("C4") = A1T1P2
        ws.Range("C5") = A1T1P3
        ws.Range("C6") = A1T1P4
        ws.Range("C7") = A1T1P5
        
        'A1 Trial 2
        ws.Range("D3") = A1T2P1
        ws.Range("D4") = A1T2P2
        ws.Range("D5") = A1T2P3
        ws.Range("D6") = A1T2P4
        ws.Range("D7") = A1T2P5
        
        'A1 Trial 3
        ws.Range("E3") = A1T3P1
        ws.Range("E4") = A1T3P2
        ws.Range("E5") = A1T3P3
        ws.Range("E6") = A1T3P4
        ws.Range("E7") = A1T3P5
        
        'A2 Trial 1
        ws.Range("C8") = A2T1P1
        ws.Range("C9") = A2T1P2
        ws.Range("C10") = A2T1P3
        ws.Range("C11") = A2T1P4
        ws.Range("C12") = A2T1P5
        
        'A2 Trial 2
        ws.Range("D8") = A2T2P1
        ws.Range("D9") = A2T2P2
        ws.Range("D10") = A2T2P3
        ws.Range("D11") = A2T2P4
        ws.Range("D12") = A2T2P5
        
        'A2 Trial 3
        ws.Range("E8") = A2T3P1
        ws.Range("E9") = A2T3P2
        ws.Range("E10") = A2T3P3
        ws.Range("E11") = A2T3P4
        ws.Range("E12") = A2T3P5
        
        'A3 Trial 1
        ws.Range("C13") = A3T1P1
        ws.Range("C14") = A3T1P2
        ws.Range("C15") = A3T1P3
        ws.Range("C16") = A3T1P4
        ws.Range("C17") = A3T1P5
        
        'A3 Trial 2
        ws.Range("D13") = A3T2P1
        ws.Range("D14") = A3T2P2
        ws.Range("D15") = A3T2P3
        ws.Range("D16") = A3T2P4
        ws.Range("D17") = A3T2P5
        
        'A3 Trial 3
        ws.Range("E13") = A3T3P1
        ws.Range("E14") = A3T3P2
        ws.Range("E15") = A3T3P3
        ws.Range("E16") = A3T3P4
        ws.Range("E17") = A3T3P5
        
        '/ Convert Range to numbers
Dim rng As Range
Dim cell As Range

Set rng = ws.Range("C3:E17")

On Error Resume Next ' Ignore errors and continue execution
For Each cell In rng
    If IsNumeric(cell.Value) Then
        cell.Value = Val(cell.Value)
    End If
Next cell
On Error GoTo 0 ' Disable error handling

' Calculations
On Error Resume Next ' Ignore errors and continue execution
calR = ws.Range("B25")
cald2 = ws.Range("B26")
calk1 = ws.Range("B27")
calEV = ws.Range("B28")
calxdiff = ws.Range("B30")
caln = ws.Range("B31")
calrValue = ws.Range("B32")
cald2Value = ws.Range("B33")
calk2 = ws.Range("B34")
calAV = ws.Range("B37")
calRR = ws.Range("B38")
On Error GoTo 0 ' Disable error handling

'/ Convert score to percentage
Dim pScore As Double
On Error Resume Next ' Ignore errors and continue execution
pScore = ws.Range("B39")
If Not IsError(pScore) Then
    calScore = FormatPercent(pScore, 2)
End If
On Error GoTo 0 ' Disable error handling

'/ Status
On Error Resume Next ' Ignore errors and continue execution
statusLabel.Caption = "Status:"
statusLabelLog.Caption = "Searching..."
Status
On Error GoTo 0 ' Disable error handling

'/ Change back to GageRnR Worksheet
On Error Resume Next ' Ignore errors and continue execution
List_Select = "GageRnR" ' Tab name
Set ws = Sheets(List_Select)
If Not ws Is Nothing Then
    Set Worksheet_Set = ws
Else
    ' Handle the case when the worksheet is not found
    MsgBox "Worksheet 'GageRnR' not found!"
End If
On Error GoTo 0 ' Disable error handling
    Gage_Number.SetFocus
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
        
        'Trial 2
        ws.Cells(r, "J") = A1T2P1
        ws.Cells(r, "K") = A1T2P2
        ws.Cells(r, "L") = A1T2P3
        ws.Cells(r, "M") = A1T2P4
        ws.Cells(r, "N") = A1T2P5
        
        'Trial 3
        ws.Cells(r, "O") = A1T3P1
        ws.Cells(r, "P") = A1T3P2
        ws.Cells(r, "Q") = A1T3P3
        ws.Cells(r, "R") = A1T3P4
        ws.Cells(r, "S") = A1T3P5
        
        '/ Gage R&R Appraiser 2 /*
        ws.Cells(r, "T") = Ap2Name
        
        'Trial 1
        ws.Cells(r, "U") = A2T1P1
        ws.Cells(r, "V") = A2T1P2
        ws.Cells(r, "W") = A2T1P3
        ws.Cells(r, "X") = A2T1P4
        ws.Cells(r, "Y") = A2T1P5
        
        'Trial 2
        ws.Cells(r, "Z") = A2T2P1
        ws.Cells(r, "AA") = A2T2P2
        ws.Cells(r, "AB") = A2T2P3
        ws.Cells(r, "AC") = A2T2P4
        ws.Cells(r, "AD") = A2T2P5
        
        'Trial 3
        ws.Cells(r, "AE") = A2T3P1
        ws.Cells(r, "AF") = A2T3P2
        ws.Cells(r, "AG") = A2T3P3
        ws.Cells(r, "AH") = A2T3P4
        ws.Cells(r, "AI") = A2T3P5
        
        '/ Gage R&R Appraiser 3 /*
        ws.Cells(r, "AJ") = Ap3Name
        
        'Trial 1
        ws.Cells(r, "AK") = A3T1P1
        ws.Cells(r, "AL") = A3T1P2
        ws.Cells(r, "AM") = A3T1P3
        ws.Cells(r, "AN") = A3T1P4
        ws.Cells(r, "AO") = A3T1P5
        
        'Trial 2
        ws.Cells(r, "AP") = A3T2P1
        ws.Cells(r, "AQ") = A3T2P2
        ws.Cells(r, "AR") = A3T2P3
        ws.Cells(r, "AS") = A3T2P4
        ws.Cells(r, "AT") = A3T2P5
        
        'Trial 3
        ws.Cells(r, "AU") = A3T3P1
        ws.Cells(r, "AV") = A3T3P2
        ws.Cells(r, "AW") = A3T3P3
        ws.Cells(r, "AX") = A3T3P4
        ws.Cells(r, "AY") = A3T3P5
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Updating..."
        Status
        Search_Button_Click
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
    
    'Trial 2
    A1T2P1 = ""
    A1T2P2 = ""
    A1T2P3 = ""
    A1T2P4 = ""
    A1T2P5 = ""
    
    'Trial 3
    A1T3P1 = ""
    A1T3P2 = ""
    A1T3P3 = ""
    A1T3P4 = ""
    A1T3P5 = ""
    
    '/ Gage R&R Appraiser 2 /*
    Ap2Name = ""
    
    'Trial 1
    A2T1P1 = ""
    A2T1P2 = ""
    A2T1P3 = ""
    A2T1P4 = ""
    A2T1P5 = ""
    
    'Trial 2
    A2T2P1 = ""
    A2T2P2 = ""
    A2T2P3 = ""
    A2T2P4 = ""
    A2T2P5 = ""
    
    'Trial 3
    A2T3P1 = ""
    A2T3P2 = ""
    A2T3P3 = ""
    A2T3P4 = ""
    A2T3P5 = ""
    
    '/ Gage R&R Appraiser 3 /*
    Ap3Name = ""
    
    'Trial 1
    A3T1P1 = ""
    A3T1P2 = ""
    A3T1P3 = ""
    A3T1P4 = ""
    A3T1P5 = ""
    
    'Trial 2
    A3T2P1 = ""
    A3T2P2 = ""
    A3T2P3 = ""
    A3T2P4 = ""
    A3T2P5 = ""
    
    'Trial 3
    A3T3P1 = ""
    A3T3P2 = ""
    A3T3P3 = ""
    A3T3P4 = ""
    A3T3P5 = ""
    
    'Cal
    calR = ""
    cald2 = ""
    calk1 = ""
    calEV = ""
    calxdiff = ""
    caln = ""
    calrValue = ""
    cald2Value = ""
    calk2 = ""
    calAV = ""
    calRR = ""
    calScore = ""
    
End Sub

Private Sub btnClear_Click()
    Update_Button_Enable = False
    Clear_Form
    Gage_Number.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload GageRnR
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


