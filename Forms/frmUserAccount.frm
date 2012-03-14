VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUserAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   620
   Begin VB.TextBox txtName 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Timer timerUSU 
      Interval        =   1
      Left            =   4050
      Top             =   2790
   End
   Begin MSComctlLib.ImageList ilRecordIcos 
      Left            =   3870
      Top             =   3510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":058A
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":0B24
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   3810
      Top             =   2820
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
            Picture         =   "frmUserAccount.frx":10BE
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":1F98
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Add"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   4350
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Edit"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   4740
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Delete"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdReload 
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   5130
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Reload"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin MSComctlLib.ListView lvwUsers 
      Height          =   4935
      Left            =   2400
      TabIndex        =   8
      Top             =   525
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8705
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ilRecordIcos"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User_ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Login Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "User Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Creation Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Creation Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Created By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Phone No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Type"
         Object.Width           =   38100
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   7440
      TabIndex        =   11
      Top             =   5520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Caption         =   "&Close"
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
   Begin VB.Image Image3 
      Height          =   525
      Left            =   0
      Picture         =   "frmUserAccount.frx":2E72
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   9585
   End
   Begin VB.Label lblFullName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2115
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1290
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   960
      Width           =   2085
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "frmUserAccount.frx":2F0F
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Accounts"
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
      Left            =   720
      TabIndex        =   4
      Top             =   180
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -30
      Picture         =   "frmUserAccount.frx":37D9
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   0
      Picture         =   "frmUserAccount.frx":46A3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9585
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   5085
      Left            =   0
      Top             =   510
      Width           =   2430
   End
End
Attribute VB_Name = "frmUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsusers As ADODB.Recordset
Dim UserType, pop, ico As String
Dim x, pos As Integer
Dim itmX As ListItem
Public LastSQL As String


Private Sub cmdSearch_Click()
Set itmX = lvwUsers.FindItem(pop, , , lvwPartial)
   If itmX Is Nothing Then  ' If no match, inform user and exit.
      'MsgBox "No match found"
      Exit Sub
   Else
       itmX.EnsureVisible ' Scroll ListView to show found ListItem.
       itmX.Selected = True   ' Select the ListItem.
      ' Return focus to the control to see selection.
       lst.SetFocus
   End If
End Sub
Private Sub Command1_Click()

End Sub

Private Sub cmdAdd_Click()
flag = 1
frmNewUser.Show 1, frmMain
Button_setting frmNewUser, rsusers, 1

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim rsUser As New ADODB.Recordset
Set rsUser = New ADODB.Recordset

pubUserID = Me.lvwUsers.ListItems.Item(Me.lvwUsers.SelectedItem.Index).SubItems(1)

If pubUserID <> "" Then
    If Me.lvwUsers.ListItems.Item(Me.lvwUsers.SelectedItem.Index).SubItems(8) = "root" Or pubUserLevel = "user" Then
        MsgBox "Cannot delete system account", vbExclamation, App.Title
        Exit Sub
End If
rsUser.Open "select * from tbl_Users where User_ID=" + pubUserID, pubcnn, adOpenDynamic, adLockOptimistic
res = MsgBox("Are you sure you want to delete this account?", vbYesNo, "User Account")
If res = vbYes Then
rsUser.Delete adAffectCurrent
Else
Exit Sub
End If
Call Display("SELECT * FROM tbl_Users")
Else
MsgBox "Please select user record first", vbInformation

End If
End Sub

Private Sub cmdEdit_Click()
pubUserID = Me.lvwUsers.ListItems.Item(Me.lvwUsers.SelectedItem.Index).SubItems(1)
If pubUserID <> "" Then
    If Me.lvwUsers.ListItems.Item(Me.lvwUsers.SelectedItem.Index).SubItems(8) = "root" Or pubUserLevel = "user" Then
        MsgBox "Cannot edit a system account", vbExclamation, App.Title
        Exit Sub
    End If
    flag = 2
    frmNewUser.Show 1
    Call Display("SELECT * FROM tbl_Users")
Else
    MsgBox "Please select user record first", vbInformation
End If
End Sub

Private Sub cmdReload_Click()
Call Display("SELECT * FROM tbl_Users")
End Sub

Private Sub Form_Load()
OpenDataBase
Call Display("SELECT * FROM tbl_Users")
End Sub

Private Sub lvwUsers_BeforeLabelEdit(Cancel As Integer)
Cancel = 1

End Sub



Private Function Display(ByVal pSQL As String)
Dim rsusers As ADODB.Recordset
Set rsusers = New ADODB.Recordset
On Error Resume Next
rsusers.Open pSQL, pubcnn, adOpenDynamic, adLockOptimistic
x = 0
Me.lvwUsers.ListItems.Clear
    LastSQL = pSQL
   
Do While Not rsusers.EOF
    x = x + 1
    
   If LCase(rsusers("UserType")) = "administrator" Then
ico = "admin"
Else
ico = "user"
End If

Me.lvwUsers.ListItems.Add , , rsusers("LoginName"), ico, ico

    Me.lvwUsers.ListItems.Item(x).SubItems(1) = rsusers("User_ID")
    lvwUsers.ListItems.Item(x).SubItems(2) = rsusers("UserName")
    lvwUsers.ListItems.Item(x).SubItems(3) = rsusers("CreationDate")
    lvwUsers.ListItems.Item(x).SubItems(4) = rsusers("CreationTime")
    lvwUsers.ListItems.Item(x).SubItems(5) = rsusers("Createdby")
    lvwUsers.ListItems.Item(x).SubItems(6) = rsusers("ContactNo")
    lvwUsers.ListItems.Item(x).SubItems(7) = rsusers("Address")
    lvwUsers.ListItems.Item(x).SubItems(8) = rsusers("UserType")

    rsusers.MoveNext
  Loop
End Function

Private Sub tlbView_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button = "Edit User" Then
 If Not lvwUsers.ListItems.Count > 0 Then
    MsgBox "No record seleceted", vbExclamation
    Exit Sub
  End If
  
  Load frmNewUser
'rsUsers.Find
 frmNewUser.Show 1 ', Me
 End If
End Sub

Private Sub txtValue_Change()
Call cmdSearch_Click
End Sub

Private Function RecDisp()
Me.txtUserName.Text = Me.lvwUsers.ListItems.Item(Me.lvwUsers.SelectedItem.Index)
Me.txtName = lvwUsers.ListItems.Item(Me.lvwUsers.SelectedItem.Index).SubItems(2)


End Function

Private Sub lvwUsers_Click()
RecDisp
    pubUserID = Me.lvwUsers.ListItems.Item(Me.lvwUsers.SelectedItem.Index).SubItems(1)
End Sub

Public Sub form_new()
cmdAdd_Click
End Sub
Public Sub Form_Edit()
cmdEdit_Click
End Sub
Public Sub Form_Delete()
cmdDelete_Click
End Sub
Public Sub Form_Refresh()
cmdReload_Click
End Sub

