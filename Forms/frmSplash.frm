VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox piclogo 
      Height          =   2055
      Left            =   3960
      Picture         =   "frmSplash.frx":0442
      ScaleHeight     =   1995
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   5520
      Width           =   9855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   2760
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2



Dim isAbout As Boolean



Dim tl As Integer

Public Sub ShowSplash()
    isAbout = False
    Me.Show
    DoEvents
End Sub


Public Sub UnloadSplash()
    Me.Enabled = False
    Timer1.Enabled = True
End Sub



Public Function ShowAbout()
   isAbout = True
'    lblSchoolName.Caption = CurrentSchool.SchoolName
    Me.Show

End Function

 
Private Sub Form_Activate()
    SetWindowPos Me.hWnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS
End Sub


Private Sub Form_DblClick()
        Unload Me
frmLogin.Show
End Sub

Private Sub Form_Deactivate()
    If isAbout = True Then
        UnloadSplash
    End If
End Sub

Private Sub Trans(Level As Integer)
        Dim Msg As Long

        Msg = GetWindowLong(Me.hWnd, G)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong Me.hWnd, G, Msg
        SetLayeredWindowAttributes Me.hWnd, 0, Level, LWA_ALPHA
        MakeSemiTransparent = 1
End Sub

Private Sub Form_Load()
Call createSkinnedForm(Me, piclogo)
tl = 100
End Sub



Private Sub Timer1_Timer()
    Trans tl
    
    tl = tl - 1
    
    If tl < 10 Then
        Timer1.Enabled = False
        tl = 100
        Unload Me
frmLogin.Show
    End If
End Sub



