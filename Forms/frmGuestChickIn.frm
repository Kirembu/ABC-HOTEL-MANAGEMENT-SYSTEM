VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGuestChickIn 
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList icoHeader 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmGuestChickIn.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuestChickIn.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listRecord 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7858
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Guest ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Age"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Booking Date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmGuestChickIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
RefreshReportList
End Sub
Public Sub RefreshReportList()
If mdlData.DataBaseToForm("select * from tbl_Guest where Guest_ID <> ''") = True Then
Do While Not pubRst.EOF
    x = x + 1
    Me.listRecord.ListItems.Add = pubRst("Guest_ID")
    FullName = pubRst("Title") + " " + pubRst("First_Name") + " " + pubRst("Second_Name") + " " + pubRst("Other_Name")
    listRecord.ListItems.Item(x).SubItems(1) = FullName
    listRecord.ListItems.Item(x).SubItems(2) = pubRst("Age")
    pubRst.MoveNext
  Loop
End If
       
End Sub
