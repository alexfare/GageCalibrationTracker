VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCustomer 
   Caption         =   "Customer Manager"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
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
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Customers" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(Customer_Name), Val(Customer_Name), Customer_Name), ws.Columns(1), 0)) Then
  
    Dim lLastRow As Long    ' lLastRow = variable to store the result of the row count calculation
    lLastRow = ws.ListObjects.Item(1).ListRows.Count
    r = lLastRow + 2 ' Add number for every header tab created
                Dim cnString As String
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
        'ErrMsg_Duplicate
    End If
End Sub

'/ Clear Button
Private Sub btnClear_Click()
Update_Button_Enable = False
Clear_Form
End Sub

Sub ErrMsg()
MsgBox ("Customer Not Found"), , "Not Found"
End Sub

Sub ErrMsg_Duplicate()
MsgBox ("Customer already added"), , "Duplicate"
End Sub

Private Sub Clear_Form()
        Customer_Name = ""
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
