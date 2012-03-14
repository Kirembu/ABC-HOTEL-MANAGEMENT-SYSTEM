VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewUsers 
   Caption         =   "View User Details"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   Icon            =   "frmViewUsers.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmViewUsers.frx":076A
   ScaleHeight     =   6855
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cboCreteria 
      Height          =   315
      ItemData        =   "frmViewUsers.frx":1DC8
      Left            =   1080
      List            =   "frmViewUsers.frx":1DD8
      TabIndex        =   3
      Text            =   "None"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin MSComctlLib.ImageList imglistUserSmall 
      Left            =   7800
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewUsers.frx":1E01
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewUsers.frx":381B
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlUsers 
      Left            =   7680
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewUsers.frx":52AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewUsers.frx":5A27
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbView 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   1058
      ButtonWidth     =   3307
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlUsers"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Create New User"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit User"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lvwUsers 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imglistUser"
      SmallIcons      =   "imglistUserSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
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
   End
   Begin MSComctlLib.ImageList imglistUser 
      Left            =   7680
      Top             =   4080
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
            Picture         =   "frmViewUsers.frx":61A1
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewUsers.frx":707B
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter By"
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
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Value"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1005
   End
End
Attribute VB_Name = "frmViewUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserType, pop, ico As String
Dim x, pos As Integer
Dim itmX As ListItem
Dim strItmX(10) As String
Public LastSQL As String


Private Sub cmdSearch_Click()
pop = Me.txtValue.Text

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

Private Sub cmdFilter_Click()
Dim mUserType As Boolean
   Select Case Trim(Me.cboCreteria.Text)
          Case "User ID"
                pubsql = "SELECT * FROM tbl_Users WHERE User_ID >='" & Trim(Me.txtValue.Text) & "'"
          Case "User Name"
                pubsql = "SELECT * FROM tbl_Users WHERE UserName >= '" & Trim(Me.txtValue.Text) & "'"
          Case "User Type"
                pubsql = "SELECT * FROM tbl_Users WHERE UserType >= '" & Trim(Me.txtValue.Text) & "'"
          Case Else
                pubsql = "SELECT * FROM tbl_Users"
   End Select
   Display pubsql
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_GotFocus()
If EditMode = True Then
Call Populate
EditMode = False
End If
End Sub

Private Sub Form_Load()
'Call Display("SELECT * FROM tbl_Users")
End Sub

Private Sub lvwUsers_BeforeLabelEdit(Cancel As Integer)
Cancel = 1

End Sub



Private Function Display(ByVal pSQL As String)
On Error Resume Next
x = 0
Me.lvwUsers.ListItems.Clear
    LastSQL = pSQL
   
Do While Not pubRst.EOF
    x = x + 1
   If LCase(pubRst("UserType")) = "administrator" Then
ico = "admin"
Else
ico = "user"
End If

Me.lvwUsers.ListItems.Add , , x, ico, ico

    Me.lvwUsers.ListItems.Item(x).SubItems(1) = pubRst("User_ID")
    lvwUsers.ListItems.Item(x).SubItems(2) = pubRst("LoginName")
    lvwUsers.ListItems.Item(x).SubItems(3) = pubRst("UserName")
    lvwUsers.ListItems.Item(x).SubItems(4) = pubRst("CreationDate")
    lvwUsers.ListItems.Item(x).SubItems(5) = pubRst("CreationTime")
    lvwUsers.ListItems.Item(x).SubItems(6) = pubRst("Createdby")
    lvwUsers.ListItems.Item(x).SubItems(7) = pubRst("ContactNo")
    lvwUsers.ListItems.Item(x).SubItems(8) = pubRst("Address")


    pubRst.MoveNext
  Loop
End Function

Private Sub tlbView_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button = "Edit User" Then
 If Not lvwUsers.ListItems.Count > 0 Then
    MsgBox "No record seleceted", vbExclamation
    Exit Sub
  End If
  
  Load frmNewUser
'pubRst.Find
 frmNewUser.Show 1 ', Me
 End If
End Sub

Private Sub txtValue_Change()
Call cmdSearch_Click
End Sub
