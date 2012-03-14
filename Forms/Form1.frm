VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmGuestView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guests"
   ClientHeight    =   6345
   ClientLeft      =   1590
   ClientTop       =   1515
   ClientWidth     =   9270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9270
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   2415
      TabIndex        =   8
      Top             =   3360
      Width           =   2415
      Begin lvButton.lvButtons_H cmdAdd 
         Height          =   405
         Left            =   0
         TabIndex        =   9
         Top             =   0
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
         TabIndex        =   10
         Top             =   390
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
         TabIndex        =   11
         Top             =   780
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
         TabIndex        =   12
         Top             =   1170
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
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   405
         Left            =   0
         TabIndex        =   13
         Top             =   0
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
      Begin lvButton.lvButtons_H cmdPrint 
         Height          =   405
         Left            =   0
         TabIndex        =   14
         Top             =   1560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   714
         Caption         =   "&Print"
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
   End
   Begin VB.TextBox txtGuestName 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1830
      Width           =   2175
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2790
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvwGuestInfo 
      Height          =   5055
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Guest ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Full Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Age"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Of Birth"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Gender"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Country"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Boarding"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Booking Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Passport No."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "National ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Phone No."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Check Out Date"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Check Out Time"
         Object.Width           =   2822
      EndProperty
      Picture         =   "Form1.frx":076A
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   405
      Left            =   7440
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Caption         =   "&OK"
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
   Begin VB.Label lblGuestName 
      BackStyle       =   0  'Transparent
      Caption         =   "Guest Name"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lbAge 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2430
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guests"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected guest"
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
      TabIndex        =   1
      Top             =   960
      Width           =   2085
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   0
      Picture         =   "Form1.frx":66A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9585
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   4965
      Left            =   0
      Top             =   600
      Width           =   2430
   End
End
Attribute VB_Name = "frmGuestView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim mItem As ListItem
Public LastSQL As String

Dim FullName As String
Dim x, i As Integer
Dim res As String

Private Function Display(ByVal pSQL As String)
'On Error Resume Next
x = 0
Me.lvwGuestInfo.ListItems.Clear
 If mdlData.DataBaseToForm(pSQL) = False Then
   Exit Function
 End If
    LastSQL = pSQL
   
Do While Not pubRst.EOF
    x = x + 1

Me.lvwGuestInfo.ListItems.Add , , x

    Me.lvwGuestInfo.ListItems.Item(x).SubItems(1) = pubRst("Guest_ID").Value

    FullName = pubRst("Title") + " " & pubRst("First_Name") + " " & pubRst("Second_Name")
    lvwGuestInfo.ListItems.Item(x).SubItems(2) = FullName
    lvwGuestInfo.ListItems.Item(x).SubItems(3) = pubRst("Age")
    lvwGuestInfo.ListItems.Item(x).SubItems(4) = pubRst("DOB")
    lvwGuestInfo.ListItems.Item(x).SubItems(5) = pubRst("Sex")
    lvwGuestInfo.ListItems.Item(x).SubItems(6) = pubRst("Country")
'    lvwGuestInfo.ListItems.Item(x).SubItems(7) = pubRst("Boarding")
    lvwGuestInfo.ListItems.Item(x).SubItems(8) = pubRst("BookinDate")
    lvwGuestInfo.ListItems.Item(x).SubItems(9) = pubRst("Passport_No")
    lvwGuestInfo.ListItems.Item(x).SubItems(10) = pubRst("National_ID")
    lvwGuestInfo.ListItems.Item(x).SubItems(11) = pubRst("Phone")
    lvwGuestInfo.ListItems.Item(x).SubItems(12) = pubRst("Address")
    lvwGuestInfo.ListItems.Item(x).SubItems(13) = IIf(pubRst("CheckOutDate") <> "", pubRst("CheckOutDate"), "Not yet")
'    lvwGuestInfo.ListItems.Item(x).SubItems(14) = pubRst("CheckOutTime") <<< this was removed due to lazyness
    
    pubRst.MoveNext
  Loop
End Function
Private Sub cboCreteria_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub
Private Sub cmdAdd_Click()
If frmGuest.Visible = False Then
    strsearch = ""
    frmGuest.Show 1, Me
Else
  Exit Sub
  End If
End Sub



Private Sub cmdDelete_Click()
Dim rsGuest As New ADODB.Recordset
Set rsGuest = New ADODB.Recordset
res = MsgBox("Are you sure you want to delete this record", vbYesNo)
If res = vbNo Then Exit Sub
If Me.lvwGuestInfo.ListItems.Count <> 0 Then
pubUserID = Me.lvwGuestInfo.ListItems.Item(Me.lvwGuestInfo.SelectedItem.Index).SubItems(1)
If pubUserID <> "" Then
If pubUserID = "Administrator" Then
MsgBox "Can not delete system account", vbExclamation, App.Title
Exit Sub
End If
rsGuest.Open "select * from tbl_Guest where Guest_ID=" + pubUserID, pubcnn, adOpenDynamic, adLockOptimistic
rsGuest.Delete adAffectCurrent
Call Display("SELECT * FROM tbl_Guest")
Else
MsgBox "Please select user record first", vbInformation

End If
End If
End Sub

Private Sub cmdEdit_Click()
 If Not lvwGuestInfo.ListItems.Count > 0 Then
    MsgBox "No record seleceted", vbExclamation
    Exit Sub
  End If
  strsearch = lvwGuestInfo.ListItems(lvwGuestInfo.SelectedItem.Index).ListSubItems(1).Text
flag = 2
frmGuest.Show
pubtempsql = "Select * from tbl_Guest where Guest_ID =" & strsearch
  frmGuest.Display (pubtempsql)
End Sub


Private Sub cmdRefresh_Click()
  If LastSQL <> "" Then
    Display LastSQL
  End If
End Sub


Private Sub cmdOK_Click()

strsearch = Me.lvwGuestInfo.ListItems.Item(Me.lvwGuestInfo.SelectedItem.Index).SubItems(1)
Unload Me

End Sub

Private Sub cmdPrint_Click()
mdlData.initDtEnv
strsearch = Me.lvwGuestInfo.ListItems.Item(Me.lvwGuestInfo.SelectedItem.Index).SubItems(1)
DataEnv.cmdGuestRecord strsearch
rptGuestRecord.Show


End Sub

Private Sub cmdReload_Click()
Display "select * from tbl_Guest"
End Sub

Private Sub Form_Load()

pubsql = "SELECT * FROM tbl_Guest"
Call Display(pubsql)

End Sub

Private Sub lblUserName_Click()

End Sub


Private Sub lvwGuestInfo_Click()
If Me.lvwGuestInfo.ListItems.Count <> 0 Then
x = lvwGuestInfo.SelectedItem.Index
Me.txtGuestName = lvwGuestInfo.ListItems.Item(x).SubItems(2)
Me.txtAge = lvwGuestInfo.ListItems.Item(x).SubItems(3)
End If
End Sub

Private Sub lvwGuestInfo_KeyUp(KeyCode As Integer, Shift As Integer)
x = lvwGuestInfo.SelectedItem.Index
Me.txtGuestName = lvwGuestInfo.ListItems.Item(x).SubItems(2)
Me.txtAge = lvwGuestInfo.ListItems.Item(x).SubItems(3)

End Sub
