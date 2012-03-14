VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewRooms 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Rooms"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewRooms.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8385
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwRooms 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Room NO"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Room Type"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   8819
      EndProperty
      Picture         =   "frmViewRooms.frx":058C
   End
End
Attribute VB_Name = "frmViewRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRooms As ADODB.Recordset
Dim rSql As String


Private Sub Form_Load()
OpenDataBase

If flag = "3" Then
RefreshReportList ("select * from tbl_Room where Status = 0 order by Room_No")
Else
RefreshReportList ("select * from tbl_Room order by Room_No")
End If
If lvwRooms.ListItems.Count = 0 Then
MsgBox "Sorry no rooms available", vbInformation + vbOKOnly
End If
mdlFX.MakeGradient Me, 1
End Sub
Public Sub RefreshReportList(rSql As String)
Dim rsRooms As ADODB.Recordset
Set rsRooms = New ADODB.Recordset
rsRooms.Open rSql, pubcnn, adOpenDynamic, adLockOptimistic
Me.lvwRooms.ListItems.Clear

Do While Not rsRooms.EOF
    x = x + 1
    Me.lvwRooms.ListItems.Add = rsRooms("Room_No")
    lvwRooms.ListItems.Item(x).SubItems(1) = rsRooms("Type")
    lvwRooms.ListItems.Item(x).SubItems(2) = rsRooms("Rates")
    lvwRooms.ListItems.Item(x).SubItems(3) = rsRooms("Description")
    rsRooms.MoveNext
  Loop

       
End Sub


Private Sub lvwRooms_DblClick()
   strsearch = Me.lvwRooms.ListItems.Item(lvwRooms.SelectedItem.Index).Text
   Unload Me
End Sub

