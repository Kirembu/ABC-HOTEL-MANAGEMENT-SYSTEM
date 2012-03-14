VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLaundry 
   Caption         =   "Laundry"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   Icon            =   "frmLaundry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   11580
   Begin VB.Frame Frame3 
      Caption         =   "Transaction Details"
      Height          =   2535
      Left            =   480
      TabIndex        =   26
      Top             =   6720
      Width           =   8055
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Ksh ""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   37
         Text            =   "0"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtAmountDue 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Ksh ""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0"
         ToolTipText     =   "Amount due"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtAmountPaid 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Ksh""#,##00.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   30
         Text            =   "0"
         ToolTipText     =   "Amount paid by customer"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtBalance 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Ksh ""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         ToolTipText     =   "Balace"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtCommission 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Ksh ""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0"
         ToolTipText     =   "Charged on transactions made by cheque"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtTax 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Ksh ""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         ToolTipText     =   "Charge on forreingers"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   38
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Due"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Commision"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblTax 
         BackStyle       =   0  'Transparent
         Caption         =   "Other Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   32
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Number of each article type"
      Height          =   2535
      Left            =   480
      TabIndex        =   16
      Top             =   1440
      Width           =   7935
      Begin VB.TextBox txtHeavy 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   22
         Text            =   "0"
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox txtMedium 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "0"
         Top             =   1200
         Width           =   945
      End
      Begin VB.TextBox txtLight 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "0"
         Top             =   1920
         Width           =   945
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   375
         Left            =   2985
         TabIndex        =   17
         Top             =   1920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtLight"
         BuddyDispid     =   196625
         OrigLeft        =   3360
         OrigTop         =   3480
         OrigRight       =   3615
         OrigBottom      =   3855
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   375
         Left            =   2985
         TabIndex        =   18
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMedium"
         BuddyDispid     =   196624
         OrigLeft        =   3360
         OrigTop         =   2760
         OrigRight       =   3615
         OrigBottom      =   3135
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   2985
         TabIndex        =   19
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtHeavy"
         BuddyDispid     =   196623
         OrigLeft        =   3360
         OrigTop         =   2040
         OrigRight       =   3615
         OrigBottom      =   2415
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Heavy linen articles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Medium articles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of light articles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Service"
      Height          =   2415
      Left            =   3960
      TabIndex        =   9
      Top             =   4080
      Width           =   4455
      Begin VB.CheckBox chkWashing 
         Caption         =   "Washing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkIroning 
         Caption         =   "Ironing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.OptionButton optHotelWashed 
         Appearance      =   0  'Flat
         Caption         =   "Hotel Washed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   600
         TabIndex        =   11
         Top             =   960
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optNotHotelWashed 
         Appearance      =   0  'Flat
         Caption         =   "Not Hotel Washed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   600
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   600
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   240
         Y1              =   840
         Y2              =   1320
      End
   End
   Begin VB.Frame fmePayment 
      Caption         =   "Paymenet Type"
      Height          =   2535
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   3375
      Begin VB.Frame fmeCurrency 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   615
         Left            =   1080
         TabIndex        =   41
         Top             =   600
         Width           =   1695
         Begin VB.OptionButton optCurrency 
            Caption         =   "Local currency"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   43
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton optCurrency 
            Caption         =   "Foreignl currency"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtChequeNo 
         Appearance      =   0  'Flat
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
         Left            =   600
         TabIndex        =   14
         Top             =   1920
         Width           =   2415
      End
      Begin VB.OptionButton optPayment 
         Appearance      =   0  'Flat
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optPayment 
         Appearance      =   0  'Flat
         Caption         =   "Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Number"
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
         Left            =   600
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.TextBox txtRecipt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtGuestID 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   405
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Search through Hotel guest records"
      Top             =   840
      Width           =   375
   End
   Begin lvButton.lvButtons_H cmdExit 
      Cancel          =   -1  'True
      Height          =   480
      Left            =   8520
      TabIndex        =   39
      Top             =   2160
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   847
      Caption         =   "&Exit"
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
   Begin lvButton.lvButtons_H cmdTransaction 
      Default         =   -1  'True
      Height          =   480
      Left            =   8520
      TabIndex        =   40
      Top             =   1560
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   847
      Caption         =   "&Save Transaction"
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
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Recipt Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblGuestID 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Guest ID"
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
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Laundry Sevice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image imgGuest 
      Appearance      =   0  'Flat
      Height          =   9615
      Left            =   120
      Stretch         =   -1  'True
      Top             =   -8880
      Width           =   8955
   End
End
Attribute VB_Name = "frmLaundry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rslaundry As ADODB.Recordset
Dim mYsQL As String
Dim GX, GY
Dim Tx As Long
Dim Ty As Long
Dim x As Long
Dim Y As Long
Dim pubTax, Work As String

Private Sub chkCurrencyType_Click(Index As Integer)

End Sub

Private Sub chkIroning_Click()
If Me.chkIroning.Value = 1 Then
 If bleHotelWashed = True Then
 Beep
 End If
 End If
End Sub


Private Sub chkWashing_Click()
If Me.chkWashing.Value = 0 Then
Me.optHotelWashed.Enabled = True
Me.optNotHotelWashed.Enabled = True
bleHotelWashed = False
Else
Me.optHotelWashed.Enabled = False
Me.optNotHotelWashed.Enabled = False
bleHotelWashed = True
End If

End Sub



Private Sub cmdBack2_Click()

End Sub

Private Sub cmdBrowse_Click()
With frmGuestList
    flag = 3
    .Show 1
End With
pubTax = ""
GetDetails
End Sub


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdTransaction_Click()
Dim res As String
  If Me.optPayment(1).Value = True And Me.txtChequeNo = "" Then

MsgBox "Please enter the cheque number to complete this transaction", vbExclamation, App.Title
Exit Sub
End If
    If Me.txtAmountDue = "0" Then
    MsgBox "Please enter the trasansaction details", vbInformation, App.Title
    Exit Sub
    End If
If Val(Me.txtAmountDue - Me.txtAmountPaid) > 0 Then
res = MsgBox("The customer has not cleared his/her bill" & _
"Continue with this credit entry?", vbExclamation + vbYesNo, App.Title)
If res = vbNo Then
Exit Sub
Else
pubCredit = True
End If
End If
Transact
End Sub



Private Sub imgGuest_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
  GX = x
  GY = Y
  End If
End Sub

Private Sub imgGuest_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
      If Button = 1 Then
         imgGuest.Parent.Move imgGuest.Parent.Left + x - GX, imgGuest.Parent.Top + Y - GY
      End If
End Sub

Private Sub optPayment_Click(Index As Integer)
If optPayment(Index).Caption = "Cheque" Then
Me.txtChequeNo.Enabled = True
pubRate = 0
Else
Me.txtChequeNo.Enabled = False
pubRate = 0.05
Me.txtCommission.Text = Val(Me.txtAmountDue * pubRate)
End If
txtAmountDue.Text = Val(Me.txtLight * 100 + Me.txtHeavy * 300 + Me.txtMedium * 200)
Call Calculate(txtHeavy, txtLight, txtMedium, txtAmountDue, txtBalance, txtAmountPaid)
End Sub
Private Function Updatedb()
Dim rslaundry As ADODB.Recordset
Set rslaundry = New ADODB.Recordset
OpenDataBase
mYsQL = "select * from tbl_Laundry"
rslaundry.Open mYsQL, pubcnn, adOpenDynamic, adLockOptimistic

With rslaundry

    If flag = 1 Then
    .AddNew
    End If
    If txtGuestID.Text <> "" Then
    .Fields("Guest_ID").Value = txtGuestID
    Else
    .Fields("Guest_ID").Value = "<N/A>"
    End If
.Fields("Recipt_No").Value = Val(txtRecipt)
.Fields("Heavy_linen").Value = Me.txtHeavy
.Fields("Medium").Value = Me.txtMedium
.Fields("Light").Value = Me.txtLight
If Me.chkIroning.Value = 1 Then Work = "Ironing"
If Me.chkWashing.Value = 1 Then Work = Work + " & Washing"
.Fields("Work") = Work
.Update
End With

End Function

Private Sub txtAmountPaid_Change()
zero_length
txtAmountDue.Text = Val(Me.txtLight * 100 + Me.txtHeavy * 300 + Me.txtMedium * 200)
Call Calculate(txtHeavy, txtLight, txtMedium, txtAmountDue, txtBalance, Val(txtAmountPaid))
End Sub


Private Sub txtBalance_Change()
zero_length
End Sub

Private Sub txtBalance_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtChequeNo_KeyPress(KeyAscii As Integer)
Call mdldb.Only_Numbers(KeyAscii)
End Sub

Private Sub txtGuestID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
strsearch = txtGuestID.Text
GetDetails
End If
mdldb.Only_Numbers KeyAscii
End Sub

Private Sub txtGuestID_LostFocus()
strsearch = txtGuestID.Text
GetDetails
End Sub

Private Sub txtHeavy_Change()
zero_length
txtAmountDue.Text = Val(Me.txtLight * 100 + Val(Me.txtHeavy) * 300 + Me.txtMedium * 200)
Call Calculate(txtHeavy, txtLight, txtMedium, txtAmountDue, txtBalance, txtAmountPaid)
End Sub

Private Sub txtHeavy_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub


Private Sub txtLight_Change()
zero_length
txtAmountDue.Text = Val(Me.txtLight * 100 + Me.txtHeavy * 300 + Me.txtMedium * 200)
Call Calculate(txtHeavy, txtLight, txtMedium, txtAmountDue, txtBalance, txtAmountPaid)
End Sub

Private Sub txtLight_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub

Private Sub txtMedium_Change()
zero_length
txtAmountDue.Text = Val(Me.txtLight * 100 + Me.txtHeavy * 300 + Val(Me.txtMedium) * 200)
Call Calculate(txtHeavy, txtLight, txtMedium, txtAmountDue, txtBalance, txtAmountPaid)
End Sub

Private Sub txtMedium_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub

Private Sub Calculate(heavy As Long, light As Long, medium As Long, amtdue As Long, balance As Long, amtpaid As Long)
Dim commission As Long
Dim tax As Long
Dim GrossCost As Long
Call zero_length
If optPayment(1).Value = False Then
GrossCost = Val(light * 100 + heavy * 300 + medium * 200 + tax)
If Me.optCurrency(1).Value = True Then
commission = Val(amtdue * 0.05)
End If
Else
commission = Val(amtdue * 0.05)
End If
GrossCost = Val(light * 100 + heavy * 300 + medium * 200)
If pubTax <> "" Then
tax = Val(amtdue * pubTax)
txtTax = tax
Else
tax = Val(0)
End If
amtdue = Val(GrossCost + commission + tax)
balance = Val(amtpaid - amtdue)
 If balance > 0 Then
 balance = Val(amtpaid - amtdue)
 txtBalance = balance
 pubCredit = False
 Else
 pubCredit = True
End If
 
 If balance < 0 Then
 balance = Val(amtpaid - amtdue)
 txtBalance = 0
End If
 txtCommission = commission
 txtAmountDue.Text = amtdue
Me.txtTax = tax
 
 End Sub
Private Sub zero_length()
If Me.txtHeavy = "" Then Me.txtHeavy = 0
If Me.txtLight = "" Then Me.txtLight = 0
If Me.txtMedium = "" Then Me.txtMedium = 0
If Me.txtAmountDue = "" Then Me.txtAmountDue = 0
If Me.txtAmountPaid = "" Then Me.txtAmountPaid = 0
If Me.txtBalance = "" Then Me.txtBalance = 0

End Sub
Private Function GetDetails()
If strsearch <> "" Then
mdlData.DataBaseToForm "SELECT * FROM tbl_Guest where Guest_ID =" & Val(strsearch)
    If pubRst.EOF = True And pubRst.BOF = True Then
    MsgBox "GuestID not found. It may be changed or deleted." _
    + vbCr + "Please enter a valid ID", vbInformation, App.Title
      Exit Function
    Else
    Me.txtGuestID.Text = pubRst("Guest_ID")
    
    If pubRst("Citizen") = 0 Then
    pubTax = "0.1"
    Else
    pubTax = "0"
    End If
    End If
    End If

End Function
Private Function Transact()
If cmdTransaction.Caption = "&Save Transaction" Then
mdlPayment.recTransact txtGuestID.Text, Val(txtAmountDue.Text), IIf(Me.optPayment(1).Value = True, "Cheque", "Cash"), Val(txtAmountPaid.Text), Me.txtChequeNo.Text, pubCredit, "Laundry"
flag = 1
Updatedb
cmdTransaction.Caption = "&New Transaction"
flag = 1
Else
mdldb.Clear_All Me
cmdTransaction.Caption = "&Save Transaction"
flag = 2
End If
End Function
