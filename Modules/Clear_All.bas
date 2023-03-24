Attribute VB_Name = "Clear_All"
Sub Clear_Run()
    Delete_Rows
    Clear_Admin
    Clear_Customers
    Clear_Credentials
End Sub

Sub Delete_Rows()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Ws As Worksheet
Dim List_Select
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    Dim i As Long
    Ws.AutoFilterMode = False
    
    For i = 999 To 3 Step -1
        Ws.Rows(i).EntireRow.Delete
    Next i
End Sub

Sub Clear_Admin()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Ws As Worksheet
Dim List_Select
    List_Select = "Admin"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    Dim i As Integer
    
    For i = 2 To 500
        Ws.Range("B" & i).ClearContents
    Next i
End Sub

Sub Clear_Customers()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Ws As Worksheet
Dim List_Select
    List_Select = "Customers"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    Dim i As Integer
    
    For i = 999 To 2 Step -1
        Ws.Rows(i).EntireRow.Delete
    Next i
End Sub

Sub Clear_Credentials()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Ws As Worksheet
Dim List_Select
    List_Select = "Credentials"        ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    Dim i As Integer
    
    For i = 999 To 3 Step -1
        Ws.Rows(i).EntireRow.Delete
    Next i
End Sub
