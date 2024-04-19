Attribute VB_Name = "SettingsModule"
'/------- Update Due Date Color -------/'
Public Sub DueDateColor()
    Dim ws As Worksheet
    Dim List_Select
    Dim Worksheet_Set
    Dim rng As Range
    Dim cell As Range
    Dim targetDate As Date
    Dim currentDate As Date
    Dim ColorRangeLeadTime As Integer
    
    List_Select = Admin
    Set Worksheet_Set = ws
    Set ws = Sheets("Admin")
    ColorRangeLeadTime = ws.Range("B63")
    
    Set ws = Sheets("CreatedByAlexFare")
    targetDate = Range("I1").Value
    Set rng = ws.Range("G3:G2000")
    
    For Each cell In rng
        If IsDate(cell.Value) Then
            currentDate = cell.Value
            
            monthsDiff = DateDiff("m", targetDate, currentDate)
            
            If currentDate < targetDate Then
                cell.Interior.Color = RGB(255, 0, 0) 'Red
            ElseIf monthsDiff <= ColorRangeLeadTime Then
                cell.Interior.Color = RGB(255, 255, 0) 'Yellow
            Else
                cell.Interior.Color = RGB(0, 255, 0) 'Green
            End If
        End If
    Next cell
End Sub
