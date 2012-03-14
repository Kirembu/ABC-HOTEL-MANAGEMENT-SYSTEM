Attribute VB_Name = "Curve"
Public gblnRec As Boolean
Public gblnPriv As Boolean
Public gblnCK As Boolean
Public gsngPer As Single
Public gstrAllProduct As String
Public gstrAllRec As String
Public gstrAllMem As String
Public gstrProName As String
Global Const winding = 2
Global Const alternate = 1
Global Const rgn_or = 2

Type POINTAPI
   x As Long
   y As Long
End Type

Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyfillMode As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

