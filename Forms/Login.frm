VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   6990
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5760
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4129.919
   ScaleMode       =   0  'User
   ScaleWidth      =   5408.327
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "="
      TabIndex        =   5
      Top             =   5160
      Width           =   4335
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "CA&NCEL"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "O&K"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   14737632
   End
   Begin MSComctlLib.ListView lvwUser 
      Height          =   4215
      Left            =   360
      TabIndex        =   2
      Top             =   495
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7435
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imglistUser"
      SmallIcons      =   "imglistUser"
      ForeColor       =   11053224
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imglistUser 
      Left            =   1080
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   240
      ImageHeight     =   83
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Login.frx":F12C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "..Enter Password"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   5055
   End
   Begin VB.Label lblUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select a user.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
    End
End Sub

Private Sub cmdOK_Click()

    pubsql = "select * from tbl_Users where UserName='" & Me.lvwUser.SelectedItem.Text & "'"
    Call mdlData.DataBaseToForm(pubsql)
    If pubRst.EOF = True Then
     MsgBox "Password  not correct, try again!", , "Login"
        Exit Sub
    End If

    If (txtPassword.Text) = pubRst("Password") Then
        LoginSucceeded = True
        pubUserID = pubRst("User_ID")
        pubLoginName = pubRst("LoginName")
        mdlVariables.pubUserName = pubRst("LoginName")
        pubUserLevel = pubRst("UserType")
'        frmMain.lblMain.Caption = "Current User: " + pubLoginName
        frmMain.Show
        Unload Me
    Else
        MsgBox "Invalid Password, try again!", vbExclamation, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If

Unload Me
frmMain.Show
End Sub

Private Sub Form_Load()
OpenDataBase
Fill_Area2 "UserName", "tbl_Users", lvwUser
mdlFX.MakeGradient Me, 1
End Sub

Private Sub lvwUser_Click()
Dim i As Integer
If lvwUser.ListItems.Count > 0 Then
For i = 1 To lvwUser.ListItems.Count
If lvwUser.ListItems(i).Index = lvwUser.SelectedItem.Index Then
lvwUser.ListItems.Item(lvwUser.SelectedItem.Index).Ghosted = False
Else
lvwUser.ListItems.Item(i).Ghosted = True
End If
Next i
End If
End Sub

Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK_Click
End Sub
