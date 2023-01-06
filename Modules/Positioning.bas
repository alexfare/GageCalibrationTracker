Attribute VB_Name = "Positioning"
Option Explicit

#If VBA7 Then
    Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long
    Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As udtRECT) As Long
#Else
    Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
    Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal Index As Long) As Long
    Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As udtRECT) As Long
#End If

Type udtRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Sub ReturnPosition_CenterScreen(ByVal sngHeight As Single, _
                                       ByVal sngWidth As Single, _
                                       ByRef sngLeft As Single, _
                                       ByRef sngTop As Single)
Dim sngAppWidth As Single
Dim sngAppHeight As Single
Dim hWnd As Long
Dim lreturn As Long
Dim lpRect As udtRECT

    hWnd = Application.hWnd   'Used in Excel and Word
    'hWnd = Application.hWndAccessApp  'Used in Access
    
    lreturn = GetWindowRect(hWnd, lpRect)
    sngAppWidth = ConvertPixelsToPoints(lpRect.Right - lpRect.Left, "X")
    sngAppHeight = ConvertPixelsToPoints(lpRect.Bottom - lpRect.Top, "Y")
    sngLeft = ConvertPixelsToPoints(lpRect.Left, "X") + ((sngAppWidth - sngWidth) / 2)
    sngTop = ConvertPixelsToPoints(lpRect.Top, "Y") + ((sngAppHeight - sngHeight) / 2)
End Sub

Public Function ConvertPixelsToPoints(ByVal sngPixels As Single, _
                                      ByVal sXorY As String) As Single
Dim hDC As Long

   hDC = GetDC(0)
   If sXorY = "X" Then
      ConvertPixelsToPoints = sngPixels * (72 / GetDeviceCaps(hDC, 88))
   End If
   If sXorY = "Y" Then
      ConvertPixelsToPoints = sngPixels * (72 / GetDeviceCaps(hDC, 90))
   End If
   Call ReleaseDC(0, hDC)
End Function

