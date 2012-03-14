VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmNewUser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New User"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   15
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   3000
   End
   Begin VB.TextBox txtPassword2 
      DataSource      =   "adoAddUser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
   End
   Begin VB.ComboBox cboUserType 
      DataSource      =   "adoAddUser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmNewUser.frx":076A
      Left            =   2160
      List            =   "frmNewUser.frx":0774
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox txtLoginName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1590
      Width           =   3135
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtFirstName 
      DataSource      =   "adoAddUser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   3
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox txtContactNumber 
      DataSource      =   "adoAddUser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   5
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox txtAddress 
      DataSource      =   "adoAddUser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2160
      MaxLength       =   70
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5160
      Width           =   3135
   End
   Begin MSComctlLib.ImageList imlUsers 
      Left            =   5400
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewUser.frx":078D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewUser.frx":0F07
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewUser.frx":1681
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewUser.frx":1DFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewUser.frx":2575
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   405
      Left            =   5280
      TabIndex        =   16
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      Caption         =   "&Save"
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
      cBhover         =   14215660
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   3720
      TabIndex        =   17
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      Caption         =   "&Cancel"
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
      cBhover         =   14215660
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New User."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   60
      TabIndex        =   18
      Top             =   60
      Width           =   1470
   End
   Begin VB.Label lblConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   14
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label lblUserLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "User Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   195
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log-In Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   1695
      Width           =   1155
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   9
      Top             =   2175
      Width           =   900
   End
   Begin VB.Label lblFirstname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   8
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label lblContacts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Width           =   1035
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   5160
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmNewUser.frx":2CEF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As VbMsgBoxResult
Dim rsusers As New ADODB.Recordset
Dim Valid As Boolean
Dim strItmX(10) As String
Dim x, i As Integer
Dim strMissing As String
Public Sub Form_Cancel()
Unload Me
End Sub

Public Sub Form_Save()
End Sub


Private Function Check()
strMissing = ""
If txtPassword1 <> txtPassword2 Then
MsgBox "Please verify the password correctly", vbExclamation, "Password verification"
Exit Function
End If
If Me.txtFirstName.Text = "" Then
strMissing = strMissing + vbCr + "-First Name"
txtFirstName.BackColor = vbYellow
End If
If Me.txtLoginName.Text = "" Then
strMissing = strMissing + vbCr + "-Second Name"
txtLoginName.BackColor = vbYellow
End If
If Me.txtPassword1.Text = "" Then
strMissing = strMissing + vbCr + "-Password"
Me.txtPassword1.BackColor = vbYellow
End If

If Me.txtPassword2.Text = "" Then
strMissing = strMissing + vbCr + "-Password"
Me.txtPassword2.BackColor = vbYellow
End If

If Me.txtContactNumber.Text = "" Then
strMissing = strMissing + vbCr + "-Phone Number"
Me.txtContactNumber.BackColor = vbYellow
End If

If txtAddress.Text = "" Then
strMissing = strMissing + vbCr + "-Address"
txtAddress.BackColor = vbYellow
End If

If Me.cboUserType.ListIndex < 0 Then
strMissing = strMissing + vbCr + "-User type"
Me.cboUserType.BackColor = vbYellow
End If
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
strMissing = ""
Check
    If strMissing <> "" Then
        MsgBox "Enter the Following" + vbCr + strMissing, vbInformation + vbOKOnly, App.Title
        Exit Sub
    Else
        If flag = 1 Then
            rsusers.AddNew
                rsusers.Fields("LoginName").Value = Me.txtLoginName.Text
                rsusers.Fields("UserName").Value = Me.txtFirstName.Text
                rsusers.Fields("User_ID").Value = mdlUsers.GetNewID("tbl_Users", "User_ID", "U")
                rsusers.Fields("ContactNo").Value = Me.txtContactNumber.Text
                rsusers.Fields("Createdby").Value = "Administrator"
                rsusers.Fields("Address").Value = Me.txtAddress.Text
                rsusers.Fields("UserType").Value = LCase(Me.cboUserType.List(Me.cboUserType.ListIndex))
                rsusers.Fields("Password").Value = Me.txtPassword2.Text
            rsusers.Update
            Button_setting Me, rsusers, 3
            MsgBox "New User Created", vbInformation, App.Title
            flag = 2
        Else
            rsusers.Close
 rsusers.Open "select * from tbl_Users where User_ID =" & pubUserID
    rsusers.Fields("LoginName").Value = Me.txtLoginName.Text
    rsusers.Fields("UserName").Value = Me.txtFirstName.Text
    rsusers.Fields("ContactNo").Value = Me.txtContactNumber.Text
    rsusers.Fields("Createdby").Value = "Administrator"
    rsusers.Fields("Address").Value = Me.txtAddress.Text
    rsusers.Fields("UserType").Value = LCase(Me.cboUserType.List(Me.cboUserType.ListIndex))
    rsusers.Fields("Password").Value = Me.txtPassword2.Text
rsusers.Update
Button_setting Me, rsusers, 3
mdlFunctions.EnableInput Me, False
flag = 3
End If
End If
End Sub

Private Sub Form_Load()
refreshData
mdlFX.MakeGradient Me, 1

End Sub
Private Sub refreshData()
   Set rsusers = New ADODB.Recordset
   If flag = 1 Then
    rsusers.Open "select * from tbl_Users order by User_ID", pubcnn, 1, 2
   rsusers.AddNew
    pubUserID = GetNewID("tbl_Users", "User_ID", "U")
    Me.txtUserID.Text = pubUserID
    Button_setting frmNewUser, rsusers, flag
    Else
    rsusers.Open "select * from tbl_Users where User_ID =" + pubUserID, pubcnn, 1, 2
  Disp
  End If
End Sub

Private Sub txtFirstName_LostFocus()
txtFirstName.Text = mdlFX.cSentenceCase(txtFirstName.Text)
End Sub

Private Sub txtLoginName_LostFocus()
If flag = 1 Then
If mdlUsers.CheckDuplicates("tbl_Users", "LoginName", txtLoginName.Text, txtLoginName) = True Then
MsgBox "Login Name already exists." + vbCr + "Enter a new loging name.", vbInformation, "New User"
End If
End If
End Sub
Private Function Disp()
On Error Resume Next

  Me.txtLoginName.Text = rsusers.Fields("LoginName").Value
      Me.txtUserID = rsusers.Fields("User_ID").Value
   Me.txtFirstName.Text = rsusers.Fields("UserName").Value
    Me.txtContactNumber.Text = rsusers.Fields("ContactNo").Value
    Me.txtAddress.Text = rsusers.Fields("Address").Value
    Me.txtPassword2.Text = rsusers.Fields("Password").Value
    Me.txtPassword1.Text = rsusers.Fields("Password").Value
End Function

