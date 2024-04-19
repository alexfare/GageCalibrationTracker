VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    On Error GoTo Err_Save
    Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
    Dim WorkBookCount As Integer
    Dim WorkBookPlusOne As Integer
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    WorkBookCount = ws.Range("B47")
    WorkBookPlusOne = WorkBookCount + 1
    ws.Range("B47") = WorkBookPlusOne
    ws.Range("B52") = "" 'Clears logged in user.
    ws.Range("B55") = "False" 'Clears Persistent Login From Last Session
    ws.Range("B64") = "1" 'Clears Persistent Login From Last Session
    ws.Range("B66") = ""
    Worksheets("CreatedByAlexFare").Activate
    SettingsModule.DueDateColor
    ThisWorkbook.Save
    
ExitSub:
    Exit Sub

Err_Save:
    Resume ExitSub
End Sub
