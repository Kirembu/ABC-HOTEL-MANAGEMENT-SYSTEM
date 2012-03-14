VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8E9EC&
   Caption         =   "HSES"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   StartUpPosition =   2  'CenterScreen
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   -3510
      TabIndex        =   0
      Top             =   555
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   106
   End
   Begin MSComctlLib.ImageList imglistUser 
      Left            =   5760
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":058A
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1464
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox bgHWND 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   0
      Width           =   15
   End
   Begin VB.PictureBox bgUserList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3990
      Left            =   0
      ScaleHeight     =   266
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   615
      Width           =   6525
      Begin MSComctlLib.ListView listUser 
         Height          =   3210
         Left            =   3750
         TabIndex        =   12
         Top             =   60
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   5662
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imglistUser"
         SmallIcons      =   "imglistUser"
         ForeColor       =   8421504
         BackColor       =   16777215
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "frmLogin.frx":233E
         NumItems        =   0
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   3330
         TabIndex        =   13
         Top             =   3375
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   714
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16185592
         Focus           =   0   'False
         cGradient       =   16185592
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14215660
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   405
         Left            =   4890
         TabIndex        =   14
         Top             =   3375
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   714
         Caption         =   "&Next"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16185592
         Focus           =   0   'False
         cGradient       =   16185592
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14215660
      End
      Begin HSES.b8Line b8Line3 
         Height          =   60
         Left            =   0
         TabIndex        =   17
         Top             =   3270
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   106
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[1]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   585
         TabIndex        =   24
         Top             =   75
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1470
         TabIndex        =   23
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   22
         Top             =   150
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[2]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1020
         TabIndex        =   21
         Top             =   75
         Width           =   405
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   -150
         Picture         =   "frmLogin.frx":2C18
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   6705
      End
      Begin VB.Image imgL1Bg 
         Height          =   3600
         Left            =   -30
         Picture         =   "frmLogin.frx":2CB5
         Top             =   -75
         Width           =   4530
      End
   End
   Begin VB.PictureBox bgUP 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   0
      ScaleHeight     =   268
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   615
      Visible         =   0   'False
      Width           =   6525
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3510
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "admin"
         Top             =   1800
         Width           =   2865
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3510
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1020
         Width           =   2865
      End
      Begin VB.CheckBox chkRemeberUserName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remember My Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   3510
         TabIndex        =   3
         Top             =   2145
         Value           =   1  'Checked
         Width           =   2190
      End
      Begin lvButton.lvButtons_H cmdLogIn 
         Default         =   -1  'True
         Height          =   405
         Left            =   4890
         TabIndex        =   6
         Top             =   3360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   714
         Caption         =   "Log-In"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16185592
         Focus           =   0   'False
         cGradient       =   16185592
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14215660
      End
      Begin HSES.b8Line b8Line2 
         Height          =   60
         Left            =   -210
         TabIndex        =   7
         Top             =   3270
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   106
      End
      Begin lvButton.lvButtons_H cmdBack 
         Height          =   405
         Left            =   3330
         TabIndex        =   10
         Top             =   3360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   714
         Caption         =   "Back"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16185592
         Focus           =   0   'False
         cGradient       =   16185592
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14215660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[1]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   585
         TabIndex        =   20
         Top             =   75
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password.."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1470
         TabIndex        =   19
         Top             =   120
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   18
         Top             =   150
         Width           =   390
      End
      Begin VB.Image Image5 
         Height          =   555
         Left            =   -60
         Picture         =   "frmLogin.frx":A855
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   6705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[2]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1020
         TabIndex        =   15
         Top             =   75
         Width           =   405
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   3510
         TabIndex        =   9
         Top             =   1545
         Width           =   825
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   3495
         TabIndex        =   8
         Top             =   750
         Width           =   945
      End
      Begin VB.Image imgL2BG 
         Height          =   3600
         Left            =   -30
         Top             =   -75
         Width           =   4530
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User's Log-in"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   150
      TabIndex        =   16
      Top             =   90
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   0
      Picture         =   "frmLogin.frx":A8F2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6405
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Dim FixWidth As Integer
Dim FixHeight As Integer

Public Sub ShowLogin()
   
    
    Me.Show vbModal

End Sub

Private Sub cmdBack_Click()
    HidePassword
End Sub

Private Sub cmdCancel_Click()
    'temp
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    
    Static AccessCounter As Integer
    'check texboxes
    If Not CheckTextBox(txtUserName, "Invalid User Name") Then
        Exit Sub
    End If
    
    If Not CheckTextBox(txtPassword, "Invalid Password") Then
        Exit Sub
    End If
    
    Dim vUser As User
    
    'get user info
    If GetUserByName(vUser, txtUserName.Text) = False Then
        MsgBox "User does not exist!", vbExclamation
        Call HLTxt(txtUserName)
        Exit Sub
    End If
    
    'check if user allready online
    'If UserOnline(txtUserName.Text) = True Then
    '    MsgBox "The User account that you trying to Log-In was already used." & vbNewLine & _
    '            "Please try another User account." & vbNewLine & vbNewLine & _
    '            "Please contact your Administrator for more information about this.", vbExclamation
    '
    '    'temp
    '    'must enabled EXIT SUB at the next line
    '    'Exit Sub
    'End If
    
    
    'check password
    If vUser.Password <> txtPassword.Text Then
        AccessCounter = AccessCounter + 1
        If AccessCounter > 5 Then
            'temp
            MsgBox "Unxepected error 0000FFF.", vbCritical
            'exit application
            End
        End If
        
        MsgBox "Invalid Password", vbExclamation
        Call HLTxt(txtPassword)
        Exit Sub
    End If
    
    'save to log
    If UserLogin(txtUserName.Text, Now) <> Success Then
        CatchError "frmLogin", "Login_click", "Unable to save in logrecord"
        'critical
        End
    End If
    
    'success
    'set cureent users
    vUser.OnLine = True
    CurrentUser = vUser
    
    
    Unload Me
    Call AfterLogin
    'close this form
    
End Sub



Private Sub cmdNext_Click()
    txtUserName.Text = listUser.SelectedItem.Text
    
    ShowPassword
End Sub

Private Sub Form_Activate()

    Dim SR As RECT
    Dim LR As RECT
    
    GetWindowRect frmSplash.hwnd, SR
    GetWindowRect bgHWND.hwnd, LR

        
    If SR.Top < 1 Then
        CenterForm Me
    Else
        Me.Top = Me.Top - (LR.Top - SR.Top) * 15
    End If

    'refresh user list
    RefreshUserList

    txtUserName.Text = AppGet_LoginUserName
    If Len(AppGet_LoginUserName) > 0 Then
        chkRemeberUserName.Value = vbChecked
        ShowPassword
    Else
        chkRemeberUserName.Value = vbUnchecked
    End If
    
    
    
End Sub

Private Sub ShowPassword()
    bgUserList.Visible = False
    bgUP.Visible = True
    txtPassword.SetFocus
End Sub

Private Sub HidePassword()

    bgUP.Visible = False
    bgUserList.Visible = True
    listUser.SetFocus
End Sub

Private Sub Form_Load()
    FixWidth = Me.Width
    FixHeight = Me.Height
    
    'set images
    'On Error Resume Next
    'Set imgBottomLogo.Picture = LoadPicture(App.Path & "/Resources/Images/BottomLogo.gif")
    
    
    Set imgL2BG.Picture = imgL1Bg.Picture
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    If chkRemeberUserName.Value = vbChecked Then
    
        AppSet_LoginUserName txtUserName.Text
    Else
    
        AppSet_LoginUserName ""
    End If
    
End Sub



Private Sub Form_Resize()
On Error Resume Next
    Me.Width = FixWidth
    Me.Height = FixHeight
End Sub

Private Sub RefreshUserList()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblUser.UserName, tblUser.UserType, First(tblLogRecord.Logout) AS FirstOfLogout, First(tblLogRecord.Logout) AS FirstOfLogout1, First(tblLogRecord.SuccessfullyOut) AS FirstOfSuccessfullyOut" & _
            " FROM tblUser LEFT JOIN tblLogRecord ON tblUser.UserName = tblLogRecord.UserName" & _
            " GROUP BY tblUser.UserName, tblUser.UserType" & _
            " ORDER BY First(tblLogRecord.Logout) DESC;"
    
    Clipboard.SetText sSQL
    listUser.ListItems.Clear
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        'fatal
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        
        GoTo ReleaseAndExit
    End If
    
    
    vRS.MoveFirst
    While vRS.EOF = False
    
        If LCase(ReadField(vRS.Fields("UserType"))) = "administrator" Then
            listUser.ListItems.Add , , ReadField(vRS.Fields("UserName")), "admin"
        Else
            listUser.ListItems.Add , , ReadField(vRS.Fields("UserName")), "user"
        End If
    
        vRS.MoveNext
    Wend


ReleaseAndExit:
    Set vRS = Nothing
End Sub



Private Sub listUser_DblClick()
    Call cmdNext_Click
End Sub

Private Sub listUser_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdNext_Click
    End If
End Sub
