VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGuestList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guest List"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   Icon            =   "frmGuestCheckIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList icoHeader 
      Left            =   6720
      Top             =   1800
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
            Picture         =   "frmGuestCheckIn.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuestCheckIn.frx":05A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listRecord 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9763
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Guest ID"
         Object.Width           =   2540
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
         Text            =   "Booking Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Citizenship"
         Object.Width           =   2540
      EndProperty
      Picture         =   "frmGuestCheckIn.frx":0B40
   End
End
Attribute VB_Name = "frmGuestList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If flag = 3 Then
    RefreshReportList ("select * from tbl_Guest where GuestIn = 0")
Else
    RefreshReportList ("select * from tbl_Guest order by Guest_ID")
End If
    mdlFX.MakeGradient Me, 1
End Sub
Public Sub RefreshReportList(pSQL)
On Error Resume Next
If mdlData.DataBaseToForm(pSQL) = True Then

Do While Not pubRst.EOF
    x = x + 1
    listRecord.ListItems.Add , , x
    listRecord.ListItems.Item(x).SubItems(1) = pubRst("Guest_ID")
    FullName = pubRst("Title") & " " & pubRst("First_Name") & " " & pubRst("Second_Name")
    listRecord.ListItems.Item(x).SubItems(2) = FullName
    listRecord.ListItems.Item(x).SubItems(3) = pubRst("Age")
    listRecord.ListItems.Item(x).SubItems(4) = pubRst("BookinDate")
    If pubRst("Citizen").Value = 0 Then
    listRecord.ListItems.Item(x).SubItems(5) = "Foreign"
    Else
    listRecord.ListItems.Item(x).SubItems(5) = "Local"
    End If
    pubRst.MoveNext
  Loop
End If
       
End Sub

Private Sub listRecord_DblClick()
pubCitizen = True
If listRecord.ListItems.Count > 0 Then
   strsearch = Format(Me.listRecord.ListItems.Item(listRecord.SelectedItem.Index).SubItems(1), "0")
   If listRecord.ListItems.Item(listRecord.SelectedItem.Index).SubItems(5) = "Foreign" Then pubCitizen = False
   
   Unload Me
   End If
End Sub
