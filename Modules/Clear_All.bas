Attribute VB_Name = "Clear_All"
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select

Sub Clear_Run()
    Clear_Admin
    Clear_Customers
    Clear_Credentials
    Clear_GageRR
    Delete_Rows
    Clear_Completed
End Sub

Sub Delete_Rows()
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Long
    ws.AutoFilterMode = False
    
    For i = 999 To 3 Step -1
        ws.Rows(i).EntireRow.Delete
    Next i
MsgBox "Rows Deleted!", vbInformation + vbApplicationModal, "Format Status"
End Sub

Sub Clear_Admin()
    Dim SuperAdmin As String
    Dim i As Integer
    Dim SkipVersion As String
    
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    SuperAdmin = "5a6WKkpPucxU75yOvrlND6xY549SrkucxhEg+SukLGzG4pdyY5I1X+51fP5BpkMC1RwXMRw9VZTFXXpXcWeemQ=="
    SkipVersion = ws.Range("B68")
    
    For i = 2 To 999
        ws.Range("B" & i).ClearContents
    Next i
ws.Range("B65") = SuperAdmin
ws.Range("B68") = SkipVersion
MsgBox "Admin Settings Cleared & Super Admin Password Set To Default.", vbInformation + vbApplicationModal, "Format Status"
End Sub

Sub Clear_Customers()
    List_Select = "Customers"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 999 To 2 Step -1
        ws.Rows(i).EntireRow.Delete
    Next i
MsgBox "Customers Cleared.", vbInformation + vbApplicationModal, "Format Status"
End Sub

Sub Clear_Credentials()
    List_Select = "Credentials"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 999 To 3 Step -1
        ws.Rows(i).EntireRow.Delete
    Next i
MsgBox "Credentials Cleared.", vbInformation + vbApplicationModal, "Format Status"
End Sub

Sub Clear_GageRR()
On Error GoTo Err
    List_Select = "GageRnR"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 999 To 3 Step -1
        ws.Rows(i).EntireRow.Delete
    Next i
MsgBox "Gage R&R Cleared.", vbInformation + vbApplicationModal, "Format Status"

ExitSub:
    Exit Sub

Err:
    Resume ExitSub
End Sub

Sub Clear_Completed()
    MsgBox "Formatting has completed."
    ThisWorkbook.Save
End Sub

