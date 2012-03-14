Attribute VB_Name = "mdlFX"


Public Function MakeGradient(ByRef Frm As Object, Scheme As Integer)
    Dim cR(255) As Integer
    Dim cG(255) As Integer
    Dim cB(255) As Integer
    Dim d As Double
    Dim i As Integer
    
    
    Select Case Scheme
        Case 1
            For i = 0 To 255
                cR(i) = 255 - (i * 0.2)
                cG(i) = 255 - (i * 0.2)
                cB(i) = 255 - (i * 0.2)
            Next

    End Select
    

    Frm.ScaleMode = vbPixels
    d = Frm.ScaleHeight / 255
    Frm.DrawWidth = d + 1
    For i = 0 To 255
        Frm.ForeColor = RGB(cR(i), cG(i), cB(i))
        Frm.Line (0, i * d)-(Frm.ScaleWidth, i * d)
    Next
    'Frm.AutoRedraw = True
End Function

Public Function CenterForm(ByRef Frm As Form)
    Frm.Move (Screen.Width - Frm.Width) / 2, (Screen.Height - Frm.Height) / 2
End Function
Public Function cSentenceCase(sText As String) As String
    
    Dim splitText() As String
    Dim newWord As String
    Dim i As Integer
    
    'check if null---------------
    If Len(sText) < 1 Then
        cSentenceCase = ""
        Exit Function
    End If
    'end Null --------------------
    
    'convert
    sText = Trim(sText)
    
    splitText = Split(sText, " ")
    
    For i = 0 To UBound(splitText)
        If Len(Trim(splitText(i))) > 0 Then
            newWord = UCase(Left(Trim(splitText(i)), 1)) & LCase(Right(Trim(splitText(i)), Len(Trim(splitText(i))) - 1))
            cSentenceCase = cSentenceCase & " " & newWord
        End If
    Next
    
    cSentenceCase = Trim(cSentenceCase)
End Function
Public Function imgBox(path As String, Box As Image)
Dim fso As New FileSystemObject
Set fso = New FileSystemObject
If fso.FileExists(path) Then
Box.Picture = LoadPicture(path)
Else
Box.Picture = LoadPicture(App.path & "\Images\noimg.gif")
End If
End Function
