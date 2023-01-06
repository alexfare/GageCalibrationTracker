VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LabelSetUp 
   Caption         =   "Label Printer Set Up"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LabelSetUp.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "LabelSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_click()
    Unload Me
    Label.Show
End Sub
