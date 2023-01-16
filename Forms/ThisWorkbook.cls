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
    Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
    Dim WorkBookCount As Integer

    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws

     WorkBookCount = Ws.Range("B47")
     WorkBookPlusOne = WorkBookCount + 1
     Ws.Range("B47") = WorkBookPlusOne
     Ws.Range("B52") = "" 'Clears logged in user.
     Ws.Range("B55") = "1"

    Worksheets("CreatedByAlexFare").Activate
    Application.DisplayFullScreen = True
    Menu.Show

End Sub
