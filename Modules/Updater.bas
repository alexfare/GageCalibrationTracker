Attribute VB_Name = "Updater"
Option Explicit

Sub CheckUpdate()
    Dim url         As String
    url = "https://github.com/alexfare/GageCalibrationTracker"
    ActiveWorkbook.FollowHyperlink url
End Sub
