Attribute VB_Name = "SkinForm"





'######################################################################################
'                          ----------------------------------
'                           CREATE SKINNED TRANSPARENT FORMS
'                          ----------------------------------
'                         Copyright (C) 2002, Kundan, IIT Delhi
'                               Email : imkundan@yahoo.com
'                               http://imkundan.tripod.com
'######################################################################################
'
'   How to use this module?
'
'   This module can create transparent area on forms by making a particular color
'   of the image of the form transparent
'
'
'   You can directly use the function createSkinnedForm
'   various parameters are :
'
'                        targetForm      :   The form which need to be skinned
'                        skinSrc         :   A picture box control which contains
'                                            the picture which you want to be the
'                                            picture of the form
'                        transparentColor:   This is optional parameter
'                                            uses the specified color as transparent
'                                            otherwise uses the color at the color
'                                            of pixel at the (0,0) of the image
'
'######################################################################################




























Option Explicit
Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Function createSkinnedForm(SkinnedForm As Form, skinSrc As PictureBox, Optional transparentColor As Long) As Long
Const RGN_OR = 2
Dim glSkinImage As Long
Dim glHeight    As Long
Dim glwidth     As Long
Dim lReturn     As Long
Dim lRgnTmp     As Long
Dim lSkinRgn    As Long
Dim lStart      As Long
Dim lRow        As Long
Dim lCol        As Long
skinSrc.AutoSize = True
With SkinnedForm
    .Hide
    .Picture = skinSrc.Picture
    .Width = skinSrc.Width
    .Height = skinSrc.Height
End With
lSkinRgn = CreateRectRgn(0, 0, 0, 0)
With skinSrc
    .AutoRedraw = True
    glHeight = .Height / Screen.TwipsPerPixelY
    glwidth = .Width / Screen.TwipsPerPixelX
    If transparentColor < 1 Then transparentColor = GetPixel(.hDC, 0, 0)
    For lRow = 0 To glHeight - 1
        lCol = 0
        Do While lCol < glwidth
            Do While lCol < glwidth And GetPixel(.hDC, lCol, lRow) = transparentColor
                lCol = lCol + 1
            Loop
            If lCol < glwidth Then
                lStart = lCol
                Do While lCol < glwidth And GetPixel(.hDC, lCol, lRow) <> transparentColor
                    lCol = lCol + 1
                Loop
                If lCol > glwidth Then lCol = glwidth
                lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                Call DeleteObject(lRgnTmp)
            End If
        Loop
    Next
End With
Call SetWindowRgn(SkinnedForm.hWnd, lSkinRgn, True)
skinSrc.Picture = LoadPicture("")
SkinnedForm.Show
End Function
