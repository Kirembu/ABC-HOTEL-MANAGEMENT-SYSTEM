VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00EFEFEF&
   Caption         =   "ABC Hotel Management System"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11580
   Icon            =   "frmMain.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   1535
      ButtonWidth     =   1746
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Guest"
            Object.ToolTipText     =   "Enter new guest details"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            Description     =   "Manage MIS Users"
            Object.ToolTipText     =   "Manage users"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Laundry"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Laundry"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Vehicle"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Services"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Billing Moitor"
            Description     =   "Billing Monitor"
            Object.ToolTipText     =   "Hotel Billing Monitor & management"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   0
      ScaleHeight     =   5010
      ScaleWidth      =   90
      TabIndex        =   2
      Top             =   1290
      Width           =   90
   End
   Begin VB.PictureBox tblMain 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   11580
      TabIndex        =   1
      Top             =   870
      Width           =   11580
      Begin VB.Label lblLoginName 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4680
      Top             =   3960
   End
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
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B09A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F36C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   6300
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   14764
            MinWidth        =   5292
            Picture         =   "frmMain.frx":1223E
            Text            =   "Today is:"
            TextSave        =   "Today is:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Picture         =   "frmMain.frx":129B8
            TextSave        =   "3:59 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Current User."
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglistUser 
      Left            =   2040
      Top             =   3120
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
            Picture         =   "frmMain.frx":13132
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1400C
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuAddNew 
         Caption         =   "Add New"
         Begin VB.Menu mnuAddGuest 
            Caption         =   "Guest"
         End
         Begin VB.Menu mnuAddUser 
            Caption         =   "User"
         End
         Begin VB.Menu mnuAddVehicle 
            Caption         =   "Vehicle"
         End
         Begin VB.Menu mnuAddService 
            Caption         =   "Service/Facility"
         End
         Begin VB.Menu mnuAddNewRoom 
            Caption         =   "Room"
         End
      End
      Begin VB.Menu mnuChk_In 
         Caption         =   "Check-In a guest"
      End
      Begin VB.Menu mnuViewBills 
         Caption         =   "View Guest Bills"
      End
      Begin VB.Menu mnu_checkout 
         Caption         =   "Check-Out a guest"
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
      Begin VB.Menu mnuEdService 
         Caption         =   "Service/Facility"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuGuestRec 
         Caption         =   "Guest Records"
      End
      Begin VB.Menu mnuViewUserAccounts 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu mnuServices 
         Caption         =   "Sevices and fascilities"
      End
      Begin VB.Menu mnuViewRooms 
         Caption         =   "Rooms"
      End
      Begin VB.Menu mnuVehicle 
         Caption         =   "Vehicle Management"
      End
   End
   Begin VB.Menu mnuOutput 
      Caption         =   "Output"
      Begin VB.Menu mnurptGuest 
         Caption         =   "Guest List"
      End
      Begin VB.Menu mnurptUser 
         Caption         =   "User List"
      End
      Begin VB.Menu mnurptPayments 
         Caption         =   "Payments"
         Begin VB.Menu mnuUnpaid 
            Caption         =   "Unpaid Bills"
         End
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu mnuServ 
         Caption         =   "Room Service"
      End
      Begin VB.Menu mnuBilling 
         Caption         =   "Billing Monitor"
      End
      Begin VB.Menu mnuLaundry 
         Caption         =   "Laundry"
      End
      Begin VB.Menu mnuVehicleRent 
         Caption         =   "Vehicle Rent"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      Begin VB.Menu mnuWCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu spacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowsList 
         Caption         =   ""
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuCtrl 
      Caption         =   "Ctrl"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEditm 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuSavem 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRefresh_Click()
On Error Resume Next
Call Me.ActiveForm.Form_Refresh
End Sub

Private Sub MDIForm_Click()
If Button = 2 Then
PopupMenu Me.mnuAddNew
End If
End Sub

Private Sub MDIForm_Load()
lblLoginName.Caption = IIf(pubLoginName <> "", pubLoginName, "Coder")
pubUserLevel = IIf(pubUserLevel <> "", pubUserLevel, "root")
OpenDataBase
mdlData.initDtEnv
End Sub

Private Sub mnu_checkout_Click()
If pubUserLevel <> "user" Then
    flag = 3
    frmGuestList.Show 1, Me
        If strsearch <> "" Then
        
            pubsql = "select * from tbl_Guest where Guest_ID = " & strsearch
                If mdlData.DataBaseToForm(pubsql) = False Then Exit Sub
                pubRst.Fields.Item("CheckOutDate") = Date
                pubRst.Fields.Item("GuestIn") = 1
                If MsgBox("Are you sure you want to check out Guest G-" & strsearch, vbYesNo, App.Title) = vbNo Then
                               MsgBox "Guest check out canceled", vbInformation, App.Title
                               Exit Sub
 
                End If
                pubRst.Update
                
                MsgBox "Guest has been checked out", vbInformation, App.Title
        End If

Else
    MsgBox "You do not have enough privillages to check out a guest", vbExclamation
    
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub
Private Sub mnuAdd_Click()
On Error Resume Next
Call Me.ActiveForm.form_new
End Sub

Private Sub mnuAddService_Click()
flag = 1
frmServicesAndFacilities.Show
End Sub

Private Sub mnuBilling_Click()
frmBillingMonitor.Show 1, Me
End Sub

Private Sub mnuChk_In_Click()
frmBookIn.Show
End Sub

Private Sub mnuDel_Click()
On Error Resume Next
Call Me.ActiveForm.Form_Delete
End Sub

Private Sub mnuEdGuest_Click()
flag = 4
frmGuestList.Show 1, Me
If strsearch = "" Then
    MsgBox "No record seleceted", vbExclamation
    Exit Sub
  End If
If frmGuest.Visible = False Then
flag = 2
pubtempsql = "select * from tbl_Guest where Guest_ID =" & strsearch
Load frmGuest
  frmGuest.Visible = True
End If
flag = 2
pubtempsql = "select * from tbl_Guest where Guest_ID =" & strsearch
  frmGuest.Display (pubtempsql)
End Sub

Private Sub mnuEditm_Click()
On Error Resume Next
Call Me.ActiveForm.Form_Edit
End Sub

Private Sub mnuEdRoom_Click()
frmAddRoom.Show
End Sub

Private Sub mnuEdService_Click()
flag = 2
frmServicesAndFacilities.Show
End Sub

Private Sub mnuEdUser_Click()
frmUserAccount.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuAddGuest_click()
flag = 1
frmGuest.Show
End Sub

Private Sub mnuGuestRec_Click()
frmGuestView.Show
End Sub

Private Sub mnuNewDepartment_Click()
frmAddDepartment.Show
End Sub

Private Sub mnuAddNewRoom_Click()
frmAddRoom.Show
End Sub



Private Sub mnuAddUser_Click()
flag = 1
frmNewUser.Show 1, Me
End Sub

Private Sub mnuAddVehicle_Click()
flag = 1
frmAddVehicle.Show
End Sub

Private Sub mnuLaundry_Click()
flag = 1
frmLaundry.Show
End Sub

Private Sub mnuPaid_Click()
rptPaidBills.Show 1, Me
End Sub

Private Sub mnurptGuest_Click()
rptGuests.Show 1, Me
End Sub

Private Sub mnurptUser_Click()
rptUsers.Show 1, Me
End Sub

Private Sub mnuSavem_Click()
On Error Resume Next
Call Me.ActiveForm.Form_Save
End Sub

Private Sub mnuServ_Click()
frmManageServices.Show 1, Me
End Sub

Private Sub mnuServices_Click()
frmServicesAndFacilities.Show
End Sub

Private Sub mnuUnpaid_Click()
rptUnpaidBills.Show 1, Me
End Sub

Private Sub mnuVehicle_Click()
frmTransport.Show
End Sub

Private Sub mnuVehicleRent_Click()
frmVehicleRent.Show
End Sub

Private Sub mnuViewBills_Click()
strsearch = ""
flag = 1
frmGuestList.Show 1, Me
If strsearch = "" Then Exit Sub
Load frmBillingMonitor
pubsql = "select * from tbl_Guest where Guest_ID = " & strsearch

 
 If mdlData.DataBaseToForm(pubsql) = False Then Exit Sub
    frmBillingMonitor.dtFrom = IIf(pubRst.Fields.Item("CheckInDate") <> "", pubRst.Fields.Item("CheckInDate"), Date)
    frmBillingMonitor.txtGuestID = strsearch
    frmBillingMonitor.chkOnlyOf.Value = 1
    frmBillingMonitor.RefreshLV
    frmBillingMonitor.Show , Me
End Sub

Private Sub mnuViewUserAccounts_Click()
frmUserAccount.Show
End Sub

Private Sub mnuWCascade_Click()
frmMain.Arrange vbCascade
End Sub

Private Sub mnuWTile_Click()
frmMain.Arrange vbTileHorizontal
End Sub

Private Sub Timer1_Timer()
    Me.sbrMain.Panels(1).Text = "Today is: " & FormatDateTime(Now, vbLongDate)
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
flag = 1
frmGuest.Show
Case 2
frmUserAccount.Show
Case 3

    frmLaundry.Show

Case 4
flag = 0
Load frmBillingMonitor
frmBillingMonitor.Show
End Select

End Sub
