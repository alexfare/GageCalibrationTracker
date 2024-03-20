Attribute VB_Name = "Clear_All"
Sub Clear_Run()
    Clear_Admin
    Clear_Customers
    Clear_Credentials
    Clear_GageRR
    Delete_Rows
    Clear_completed
End Sub

Sub Delete_Rows()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select
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
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim SuperAdmin As String
SuperAdmin = "5a6WKkpPucxU75yOvrlND6xY549SrkucxhEg+SukLGzG4pdyY5I1X+51fP5BpkMC1RwXMRw9VZTFXXpXcWeemQ=="
Dim List_Select
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 2 To 999
        ws.Range("B" & i).ClearContents
    Next i
ws.Range("B65") = SuperAdmin
MsgBox "Admin Settings Cleared & Super Admin Password Set To Default.", vbInformation + vbApplicationModal, "Format Status"
End Sub

Sub Clear_Customers()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select
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
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select
    List_Select = "Credentials"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 999 To 3 Step -1
        ws.Rows(i).EntireRow.Delete
    Next i
MsgBox "Credentials Deleted.", vbInformation + vbApplicationModal, "Format Status"
End Sub

Sub Clear_GageRR()
On Error GoTo Err
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select
    List_Select = "GageRnR"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 999 To 3 Step -1
        ws.Rows(i).EntireRow.Delete
    Next i
MsgBox "Gage R&R Deleted.", vbInformation + vbApplicationModal, "Format Status"

ExitSub:
    Exit Sub

Err:
    Resume ExitSub
End Sub

Sub Clear_completed()
    MsgBox "Formatting has completed."
    ThisWorkbook.Save
    ThisWorkbook.Close
End Sub
