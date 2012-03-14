VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00808080&
   Caption         =   "ABC Hotel Management System"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8625
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":197A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AC58
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EF2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Accounts"
            Object.ToolTipText     =   "Manage Accounts"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sevices"
            Object.ToolTipText     =   "Manage Products"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Update_Sevices"
            Object.ToolTipText     =   "Update Product Profile"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sep1"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sales"
            Object.ToolTipText     =   "Create Bills"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Receipt"
            Object.ToolTipText     =   "Receipt Entry"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Payment"
            Object.ToolTipText     =   "Payment Entry"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sep2"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ledger"
            Object.ToolTipText     =   "Show Ledger"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Hotel Register"
            Object.ToolTipText     =   "Hotel Register"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Transaction_Summary"
            Object.ToolTipText     =   "Show Transaction Summary"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Billing_Monitor"
            Object.ToolTipText     =   "Billing Monitor"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sep3"
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Quit"
            Object.ToolTipText     =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuAddNew 
         Caption         =   "Add New"
         Begin VB.Menu mnuGuest 
            Caption         =   "Guest"
         End
         Begin VB.Menu mnuUser 
            Caption         =   "User"
         End
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEdGuest 
         Caption         =   "Guest Records"
      End
      Begin VB.Menu mnuEdUser 
         Caption         =   "User"
      End
      Begin VB.Menu mnuEdRoom 
         Caption         =   "Room Rates"
      End
      Begin VB.Menu mnuEdFacilites 
         Caption         =   "Facility Rates"
      End
      Begin VB.Menu mnuEdService 
         Caption         =   "Service Rates"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuGuestRec 
         Caption         =   "Guest Records"
      End
      Begin VB.Menu mnuViewUsers 
         Caption         =   "Users"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
Call mdlData.OpenDataBase
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuGuest_Click()
frmGuest.Show 0, frmMain
End Sub

Private Sub mnuGuestRec_Click()
frmGuestView.Show , frmMain
End Sub

Private Sub mnuUser_Click()
frmNewUser.Show , frmMain
End Sub

Private Sub mnuViewUsers_Click()
frmViewUsers.Show , frmMain
End Sub

Private Sub tlbMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)
If Button = "Guest" Then
Call mnuGuest_Click
End If
If Button = "User" Then
Call mnuUser_Click
End If
End Sub


