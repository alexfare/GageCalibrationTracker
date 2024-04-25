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

'/------- Import Data -------/'
Sub ImportGCTData()
    Dim FilePath As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set ws = ThisWorkbook.Worksheets("CreatedByAlexFare")
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select CSV File to Import"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
        End If
    End With
    
    If FilePath <> "" Then
        ws.Cells.ClearContents
        ws.Cells.FormatConditions.Delete
        
        With ws.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=ws.Cells(1, 1))
            .TextFileParseType = xlDelimited
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        
        ws.Cells.EntireColumn.AutoFit
    End If
End Sub

'/------- Export Data -------/'
Sub ExportGCTData()
    Dim FilePath As Variant
    Dim ws As Worksheet
    Dim defaultFileName As String
    
    Set ws = ThisWorkbook.Worksheets("CreatedByAlexFare")
    
    defaultFileName = "GageTracker_" & Format(Date, "yyyy-mm-dd") & ".csv"
    
    FilePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, FileFilter:="CSV Files (*.csv), *.csv")
    
    If FilePath <> False Then
        ws.Copy
        ActiveWorkbook.SaveAs FilePath, xlCSV
        ActiveWorkbook.Close SaveChanges:=False
    End If
End Sub

Public Sub GetCurrentVersion()
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim CurrentVersion As String
    Dim CurrentVersionV As String

    CurrentVersion = ws.Range("B68")
    
    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    CurrentVersionV = "v" + CurrentVersion
    ws.Range("Z1") = CurrentVersionV
End Sub
