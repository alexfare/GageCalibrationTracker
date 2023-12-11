VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCustomer 
   Caption         =   "Customer Manager"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
   OleObjectBlob   =   "FormCustomer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long           ' variable used for storing row number
Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify
Dim cnString As String

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'

    Customer_Name.SetFocus

End Sub

'/------- Search Button -------/'
Private Sub Search_Customer_Click()
    If Customer_Name <> "" Then
            Customer_Profile
        Else
            Err_Blank
        End If
End Sub

Private Sub Customer_Profile()
    Dim ws As Worksheet
    Clear_Form
    List_Select = "Customers"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(Customer_Name), Val(Customer_Name), Customer_Name), ws.Columns(1), 0)) Then
        Update_Button_Enable = False
        ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(Customer_Name), Val(Customer_Name), Customer_Name), ws.Columns(1), 0)
        cnString = Customer_Name
        inCAddress = ws.Cells(r, "B")
        inCPhoneNumber = ws.Cells(r, "C")
        inCWebsite = ws.Cells(r, "D")
        Update_Button_Enable = True
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Searching..."
        Status
    End If
End Sub

'/------- Add Customer -------/'
Private Sub Add_Button_Confirm_Click()
    If Customer_Name <> "" Then
            Add_Button_Click
        Else
            Err_Blank
        End If
End Sub

Private Sub Add_Button_Click()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Customers" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(Customer_Name), Val(Customer_Name), Customer_Name), ws.Columns(1), 0)) Then
  
    Dim lLastRow As Long    ' lLastRow = variable to store the result of the row count calculation
    lLastRow = ws.ListObjects.Item(1).ListRows.Count
    r = lLastRow + 2 ' Add number for every header tab created
    'Dim cnString As String
        If IsNumeric(Customer_Name) Then
            cnString = Val(Customer_Name.Value)
        Else
            cnString = Customer_Name
        End If
    
    ws.Cells(r, "A") = cnString
    ws.Cells(r, "B") = inCAddress
    ws.Cells(r, "C") = inCPhoneNumber
    ws.Cells(r, "D") = inCWebsite
    
    '/Status/'
    statusLabel.Caption = "Status:"
    statusLabelLog.Caption = "Adding..."
    Status
    
    '/Add to Gage Number count/'
    Dim AddCustomer As Integer

    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws

     AddCustomer = ws.Range("B53")
     AddCustomerPlusOne = AddCustomer + 1
     ws.Range("B53") = AddCustomerPlusOne
     
     '/Prevent Issues in the future, Call back the main page/'
     List_Select = "Customers" ' Tab name
     Set ws = Sheets(List_Select)
     Set Worksheet_Set = ws
    Else
        ErrMsg_Duplicate
    End If
End Sub

'/------- Update Button -------/'
Private Sub Update_Button_Click()
    If Update_Button_Enable = True Then
        If cnString = Customer_Name Then
            Update_Worksheet
        Else
            MSG_Verify_Update
        End If
    Else
        ErrMsg_Search
    End If
End Sub

Private Sub Update_Worksheet()
    If Update_Button_Enable = True Then
        Dim cnString As String
        Set ws = Worksheet_Set
        If IsNumeric(Customer_Name) Then
            cnString = Val(Customer_Name.Value)
        Else
            cnString = Customer_Name
        End If
        ws.Cells(r, "A") = cnString
        ws.Cells(r, "B") = inCAddress
        ws.Cells(r, "C") = inCPhoneNumber
        ws.Cells(r, "D") = inCWebsite
    
    '/Status /'
    statusLabel.Caption = "Status:"
    statusLabelLog.Caption = "Updating..."
    Status
        
    Customer_Profile
Else
    ErrMsg_Search
End If
End Sub

Sub MSG_Verify_Update()
    MSG1 = MsgBox("Are you sure you want to change the Customer ID?", vbYesNo, "Verify")
    
    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        Customer_Name = cnString
    End If
End Sub

Private Sub btnBack_click()
    Unload Me
    AdminForm.Show
End Sub

'/ Clear Button
Private Sub btnClear_Click()
    Update_Button_Enable = False
    Customer_Name = ""
    Clear_Form
End Sub

Private Sub Clear_Form()
        inCAddress = ""
        inCPhoneNumber = ""
        inCWebsite = ""
End Sub

'/------- Status -------/'
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

'/------- Error Codes -------/'
Sub Err_Blank()
    MsgBox ("Customer name cannot be blank."), , "Error"
End Sub

Sub ErrMsg()
    MsgBox ("Customer Not Found."), , "Not Found"
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Customer already added."), , "Duplicate"
End Sub

Sub ErrMsg_Search()
    MsgBox ("Must search for entry before updating."), , "Search Error"
End Sub
