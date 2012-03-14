VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditUser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit User"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   Icon            =   "frmEditUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   7440
   Begin VB.ComboBox cboUserDepartment 
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
      Left            =   2160
      TabIndex        =   16
      Top             =   4440
      Width           =   3135
   End
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
      ItemData        =   "frmEditUser.frx":076A
      Left            =   2160
      List            =   "frmEditUser.frx":0774
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
      Top             =   4920
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
      Top             =   5520
      Width           =   3135
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3810
      TabIndex        =   19
      Top             =   6240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
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
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   2280
      TabIndex        =   20
      Top             =   6240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
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
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit User."
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
      Width           =   1395
   End
   Begin VB.Label lblUserDepartment 
      BackStyle       =   0  'Transparent
      Caption         =   "User Department"
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
      TabIndex        =   17
      Top             =   4440
      Width           =   1695
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
      Top             =   4920
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
      Top             =   5520
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmEditUser.frx":078D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As VbMsgBoxResult
Dim rsusers As New ADODB.Recordset
Dim flag As Integer
Dim Valid As Boolean
Dim strItmX(10) As String
Dim x, i As Integer
Dim strMissing As String
Public Sub Form_Cancel()
Unload Me
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
MsgBox strMissing, vbInformation + vbOKOnly, "New User"
Else
    rsusers.Fields("LoginName").Value = Me.txtLoginName.Text
    rsusers.Fields("UserName").Value = Me.txtFirstName.Text
    rsusers.Fields("ContactNo").Value = Me.txtContactNumber.Text
    rsusers.Fields("Createdby").Value = "Administrator"
    rsusers.Fields("Address").Value = Me.txtAddress.Text
    rsusers.Fields("UserType").Value = LCase(Me.cboUserType.List(Me.cboUserType.ListIndex))
    rsusers.Fields("Password").Value = Me.txtPassword2.Text
rsusers.Update
mdlFunctions.EnableInput Me, False
End If

End Sub

Private Sub Form_Load()
refreshData
Disp
mdlFX.MakeGradient Me, 1
End Sub
Private Sub refreshData()
If pubUserLevel = "user" Then
MsgBox "Sorry you do not have enough privilages to edit this account.", vbInformation
Else
OpenDataBase
    Set rsusers = New ADODB.Recordset
    rsusers.Open "select * from tbl_Users where LoginName ='" + pubUserID + "'", pubcnn, 1, 2
End If
End Sub

Private Sub txtFirstName_LostFocus()
txtFirstName.Text = mdlFX.cSentenceCase(txtFirstName.Text)
End Sub

Private Function Disp()
'On Error Resume Next
  Me.txtLoginName.Text = rsusers.Fields("LoginName").Value
      Me.txtUserID = rsusers.Fields("User_ID").Value
   Me.txtFirstName.Text = rsusers.Fields("UserName").Value
    Me.txtContactNumber.Text = rsusers.Fields("ContactNo").Value
    Me.txtAddress.Text = rsusers.Fields("Address").Value
    Me.txtPassword2.Text = rsusers.Fields("Password").Value
    Me.txtPassword1.Text = rsusers.Fields("Password").Value
End Function
