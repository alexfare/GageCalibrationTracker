Dim r As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnBack_click()
    Unload Me
End Sub
Private Sub Add_Button_Click()
    Dim Ws As Worksheet
    Dim List_Select
    List_Select = "Customers" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws
    
    If IsError(Application.Match(IIf(IsNumeric(Customer_Name), Val(Customer_Name), Customer_Name), Ws.Columns(1), 0)) Then
  
    Dim lLastRow As Long    ' lLastRow = variable to store the result of the row count calculation
    lLastRow = Ws.ListObjects.Item(1).ListRows.Count
    r = lLastRow + 2 ' Add number for every header tab created
                Dim cnString As String
                    If IsNumeric(Customer_Name) Then
                        cnString = Val(Customer_Name.Value)
                    Else
                        cnString = Customer_Name
                    End If
    
    Ws.Cells(r, "A") = cnString
    Ws.Cells(r, "B") = inCAddress
    Ws.Cells(r, "C") = inCPhoneNumber
    Ws.Cells(r, "D") = inCWebsite
    
    Add_Button.Caption = "Added!" ' change caption of add button for confirmation
    Application.Wait (Now + TimeValue("0:00:02")) ' Wait to avoid crash
    Add_Button.Caption = "Add"
    'Clear_Form
    Customer_Name.SetFocus
    
    '/Add to Gage Number count/'
    Dim AddCustomer As Integer

    List_Select = "Admin" ' Tab name
    Set Ws = Sheets(List_Select)
    Set Worksheet_Set = Ws

     AddCustomer = Ws.Range("B53")
     AddCustomerPlusOne = AddCustomer + 1
     Ws.Range("B53") = AddCustomerPlusOne
     
     '/Prevent Issues in the future, Call back the main page/'
     List_Select = "Customers" ' Tab name
     Set Ws = Sheets(List_Select)
     Set Worksheet_Set = Ws
    Else
        'ErrMsg_Duplicate
    End If
End Sub

'/ Clear Button
Private Sub btnClear_Click()
Update_Button_Enable = False
Clear_Form
Customer_Number.SetFocus
End Sub

Sub ErrMsg()
MsgBox ("Customer Not Found"), , "Not Found"
Customer_Number.SetFocus
End Sub

Sub ErrMsg_Duplicate()
MsgBox ("Customer already added"), , "Duplicate"
Customer_Number.SetFocus
End Sub

Private Sub Clear_Form()
        Customer_Name = ""
        inCAddress = ""
        inCPhoneNumber = ""
        inCWebsite = ""
End Sub

