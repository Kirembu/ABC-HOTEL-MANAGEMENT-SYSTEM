VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmBillingMonitor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Billing Monitor."
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   Icon            =   "frmPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEnv 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   1800
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox chkGrid 
         Caption         =   "Show Grid Lines"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   1680
         Width           =   1455
      End
      Begin lvButton.lvButtons_H btnExit 
         Height          =   375
         Left            =   4920
         TabIndex        =   22
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   4
         Shape           =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12632256
         cGradient       =   12632256
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmPayment.frx":076A
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H btnFont 
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&Font"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmPayment.frx":0D83
         ImgSize         =   32
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnColor 
         Height          =   495
         Left            =   1680
         TabIndex        =   24
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&Color"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmPayment.frx":41C9
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   3
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   5415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Your Display Environment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   4320
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton cmdGetGuestID 
         BackColor       =   &H00D8E9EC&
         Height          =   285
         Left            =   2760
         Picture         =   "frmPayment.frx":4CCF
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1560
         Width           =   345
      End
      Begin VB.TextBox txtGuestID 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   29
         Top             =   1560
         Width           =   1905
      End
      Begin VB.CheckBox chkOnlyOf 
         Caption         =   "Only of"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All Billing Records"
         Height          =   195
         Left            =   4080
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton optPaid 
         Caption         =   "Paid Bills Only"
         Height          =   195
         Left            =   4080
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optUnpaid 
         Caption         =   "Unpaid Bills Only"
         Height          =   195
         Left            =   4080
         TabIndex        =   8
         Top             =   720
         Value           =   -1  'True
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   53280771
         CurrentDate     =   39044
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   53280771
         CurrentDate     =   39044
      End
      Begin lvButton.lvButtons_H btnShow 
         Height          =   495
         Left            =   5760
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "&Show"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmPayment.frx":5259
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   3
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Your Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   2490
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To &Date"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&From Date"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CheckBox chkDblClick 
      Caption         =   "Use Mouse Double Click to Set the Bill as 'Paid'"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   3855
   End
   Begin MSComctlLib.ListView lvBills 
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   4210752
      BackColor       =   -2147483628
      Appearance      =   1
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
         Text            =   "Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Guest ID"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bill No"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Net Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Details"
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H btnOptions 
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&Options"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPayment.frx":5EAE
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnEnv 
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "E&nvironment"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPayment.frx":66D7
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnRefresh 
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Refresh/Default"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPayment.frx":6F93
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnPaid 
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Set As &Paid"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPayment.frx":7E94
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnUnpaid 
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Set As &Unpaid"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      LockHover       =   1
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPayment.frx":8AE9
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   6840
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H btnReprint 
      Height          =   495
      Left            =   8280
      TabIndex        =   30
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&Re-Print Bill"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPayment.frx":8DF1
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnViewGuest 
      Height          =   495
      Left            =   8280
      TabIndex        =   31
      Top             =   4920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&View Guest Details"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmPayment.frx":938A
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin VB.Label lblNet 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   7560
      TabIndex        =   26
      Top             =   6360
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View, Analyze && Manipulate All Your Billing Data."
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   3420
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Monitor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   225
      TabIndex        =   17
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmBillingMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsBills As ADODB.Recordset

Dim payStatus As String
Dim FromDate As String
Dim ToDate As String



Private Sub btnExit_Click()
    fraEnv.Visible = False
End Sub

Private Sub btnOptions_Click()
    fraOptions.Visible = Not fraOptions.Visible
End Sub

Private Sub btnPaid_Click()
    On Error GoTo handl
    
    Dim x As ListItem
    Dim t As ADODB.Recordset
    Dim rcpt As Long
    
    For Each x In lvBills.ListItems
        If x.Selected Then
            Set t = getBillDets(x.SubItems(2))
            
        If Not t!Paid Then
            pubcnn.Execute "Update tbl_Payment set Paid = 1 where Receipt_No = " & x.SubItems(2)
    
        End If
        End If
        Next
    t.Close
    Set t = Nothing
    Call RefreshLV


    Exit Sub
handl:
    'MsgBox Err.Description
End Sub

Private Sub btnRefresh_Click()
    Call Defaults
    Call RefreshLV
End Sub

Private Sub btnReprint_Click()
    On Error GoTo hand
    
    Dim vBillNo As Long
    vBillNo = lvBills.SelectedItem.SubItems(2)
        
    Call initDtEnv

    rptPaidBills.Show 1
    
    Exit Sub
hand:
    MsgBox Err.Description
End Sub


Private Sub btnShow_Click()
    
        If chkOnlyOf.Value And Trim(txtGuestID) = "" Then
        MsgBox "Select Any Guestr... Bacause You Checked the Only of Button"
        txtGuestID.SetFocus
        Exit Sub
    End If
    
    fraOptions.Visible = False
    Call RefreshLV
End Sub

Private Sub btnViewGuest_Click()
mdlData.initDtEnv
strsearch = Me.lvBills.ListItems.Item(Me.lvBills.SelectedItem.Index).SubItems(1)
DataEnv.cmdGuestRecord strsearch
rptGuestRecord.Show 1
End Sub

Private Sub chkGrid_Click()
    If chkGrid.Value Then
        lvBills.GridLines = True
    Else
        lvBills.GridLines = False
    End If
End Sub


Private Sub cmdGetGuestID_Click()
strsearch = ""
flag = 1
frmGuestList.Show 1, Me
txtGuestID = strsearch

End Sub

Private Sub dtFrom_Change()
    FromDate = Format(dtFrom, "dd-MMM-yyyy")
End Sub

Private Sub dtTo_Change()
    ToDate = Format(dtTo, "dd-MMM-yyyy")
End Sub

Private Sub Form_Load()
    Call Defaults
    Call RefreshLV
End Sub

Public Sub RefreshLV()
'    On Error GoTo handl
    
    Dim vReceipt_No As ListItem
    Dim amt As Single
    
      
    Set rsBills = New ADODB.Recordset
      If chkOnlyOf.Value = 0 Then
        Select Case True
            Case optUnpaid
            rsBills.Open "Select * from tbl_Payment where Paid = 0 and `Entry_Date` between #" & dtFrom & "# and #" & dtTo & "# order by Receipt_No", pubcnn, 1, 2
            
            Case optPaid
            rsBills.Open "Select * from tbl_Payment where Paid <> 0 and `Entry_Date` between #" & dtFrom & "# and #" & dtTo & "# order by Receipt_No", pubcnn, 1, 2
            
            Case optAll
            rsBills.Open "Select * from tbl_Payment where`Entry_Date` between # " & dtFrom & "# and #" & dtTo & "# order by Receipt_No", pubcnn, 1, 2
        End Select
        Else
        Select Case True
            Case optUnpaid
            rsBills.Open "Select * from tbl_Payment where " & txtGuestID & " = Guest_ID AND Paid = 0 and `Entry_Date` between #" & dtFrom & "# and #" & dtTo & "# order by Receipt_No", pubcnn, 1, 2

            Case optPaid
            rsBills.Open "Select * from tbl_Payment where " & txtGuestID & " = Guest_ID AND Paid = 0 and `Entry_Date` between #" & dtFrom & "# and #" & dtTo & "# order by Receipt_No", pubcnn, 1, 2

            Case optAll
            rsBills.Open "Select * from tbl_Payment where " & txtGuestID & " = Guest_ID  and `Entry_Date` between #" & dtFrom & "# and #" & dtTo & "# order by Receipt_No", pubcnn, 1, 2
        End Select
    End If
    lvBills.ListItems.Clear
    
    While Not rsBills.EOF
        Set vReceipt_No = lvBills.ListItems.Add(, , Format(rsBills!Entry_Date, "dd-MMM-yyyy"))
        If rsBills!Guest_ID <> "" Then
        vReceipt_No.SubItems(1) = rsBills!Guest_ID
        Else
        vReceipt_No.SubItems(1) = "<None>"
        End If
        vReceipt_No.SubItems(2) = rsBills!Receipt_No
        vReceipt_No.SubItems(3) = Format(rsBills!Amount, "0.00")
        vReceipt_No.SubItems(4) = IIf(rsBills!Paid, "Paid", "Not Paid")
        vReceipt_No.SubItems(5) = IIf(rsBills!Details <> "", rsBills!Details, "<None>")
        amt = amt + rsBills!Amount
        
        rsBills.MoveNext
    Wend
    rsBills.Close
    If amt > 0 Then lblNet = "Total: " & Format(amt, "0,0.00") Else lblNet = ""
If lvBills.ListItems.Count = 0 And dtFrom <> dtTo Then MsgBox "No billing data available for Dates:" & vbNewLine & dtFrom & " to " & dtTo, vbInformation
If lvBills.ListItems.Count = 0 And dtFrom = dtTo Then MsgBox "No billing data available for:" & vbNewLine & dtTo, vbInformation
    
    Exit Sub
End Sub


Private Sub btnUnpaid_Click()
    On Error GoTo handl
    
    Dim x  As ListItem
    Dim t As ADODB.Recordset
    
    For Each x In lvBills.ListItems
        If x.Selected Then
            Set t = getBillDets(x.SubItems(2))
 
                pubcnn.Execute "Update tbl_Payment set Paid = 0 where Receipt_No = " & x.SubItems(2)
 
            End If
    Next
    t.Close
    Set t = Nothing
    Call RefreshLV
    Exit Sub
handl:
    'MsgBox Err.Description
End Sub


Private Sub btnEnv_Click()
    fraEnv.Visible = Not fraEnv.Visible
End Sub

Private Sub btnFont_Click()
    Cd.FLAGS = cdlCFBoth
    
    Cd.FontBold = lvBills.Font.Bold
    Cd.FontName = lvBills.Font.Name
    Cd.FontSize = lvBills.Font.Size
    Cd.FontItalic = lvBills.Font.Italic
    
    Cd.ShowFont
    
    lvBills.Font.Bold = Cd.FontBold
    lvBills.Font.Name = Cd.FontName
    lvBills.Font.Size = Cd.FontSize
    lvBills.Font.Italic = Cd.FontItalic
End Sub

Private Sub btnColor_Click()
    Cd.Color = lvBills.ForeColor
    Cd.ShowColor
    lvBills.ForeColor = Cd.Color
End Sub

Private Sub lvBills_DblClick()
    If chkDblClick.Value Then Call btnPaid_Click
End Sub

Private Sub optAll_Click()
    payStatus = "All Bills"
End Sub

Private Sub optPaid_Click()
    payStatus = "Paid Bills"
End Sub

Private Sub optUnpaid_Click()
    payStatus = "Un-Paid Bills"
End Sub
Private Sub Defaults()
    dtFrom.Value = Date
    dtTo.Value = Date
    optUnpaid.Value = True
    
    payStatus = "Un-Paid Bills"
    FromDate = Format(dtFrom, "dd-MMM-yyyy")
    ToDate = Format(dtTo, "dd-MMM-yyyy")
    
    

End Sub

