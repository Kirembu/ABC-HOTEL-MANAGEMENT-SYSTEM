VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmBookIn 
   Caption         =   "Guest Book-In"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   Icon            =   "frmBookIn.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   12705
   WindowState     =   2  'Maximized
   Begin VB.Frame fmePayment 
      Caption         =   "Paymenet Type"
      Height          =   2535
      Left            =   120
      TabIndex        =   42
      Top             =   6120
      Width           =   6135
      Begin VB.Frame fmeCurrency 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   615
         Left            =   1080
         TabIndex        =   46
         Top             =   600
         Width           =   1695
         Begin VB.OptionButton optCurrency 
            Caption         =   "Local currency"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   48
            Top             =   120
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optCurrency 
            Caption         =   "Foreignl currency"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   47
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtChequeNo 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   600
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label19 
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
         TabIndex        =   49
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   6600
      TabIndex        =   29
      Top             =   840
      Width           =   5655
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
            Size            =   12
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
         TabIndex        =   34
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   33
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
            Size            =   12
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
         TabIndex        =   32
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0"
         ToolTipText     =   "Charged on transactions made by cheque"
         Top             =   2280
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         ToolTipText     =   "Charge on forreingers"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         TabIndex        =   39
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label16 
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
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label17 
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
         TabIndex        =   37
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label18 
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
         Left            =   120
         TabIndex        =   36
         Top             =   2280
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
         Left            =   120
         TabIndex        =   35
         Top             =   3000
         Width           =   1335
      End
   End
   Begin VB.Frame fmeDetails 
      Caption         =   "Guest details and preference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6135
      Begin VB.CommandButton cmdGetGuestID 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2865
         Picture         =   "frmBookIn.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   750
         Width           =   345
      End
      Begin VB.CommandButton cmdGetRoom 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2865
         Picture         =   "frmBookIn.frx":0CF4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   345
      End
      Begin VB.ComboBox cboStayMode 
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
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   2400
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   2
         Top             =   4440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   52559875
         CurrentDate     =   40037
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   3480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   52559875
         CurrentDate     =   40037
         MinDate         =   39833
      End
      Begin VB.TextBox txtGuestID 
         Height          =   375
         Left            =   600
         MaxLength       =   20
         TabIndex        =   7
         Top             =   720
         Width           =   2625
      End
      Begin VB.TextBox txtRoom 
         Height          =   375
         Left            =   600
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1530
         Width           =   2625
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   600
         TabIndex        =   13
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room"
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
         Left            =   615
         TabIndex        =   12
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Guest Boarding type."
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of CheckIn"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of check out"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   4200
         Width           =   1575
      End
   End
   Begin prjXTab.XTab XGuestDetailes 
      Height          =   3135
      Left            =   6600
      TabIndex        =   14
      Top             =   4680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5530
      TabCount        =   2
      TabCaption(0)   =   "Guest Detailes"
      TabContCtrlCnt(0)=   8
      Tab(0)ContCtrlCap(1)=   "txtCitizenship"
      Tab(0)ContCtrlCap(2)=   "lblDateOfBookin"
      Tab(0)ContCtrlCap(3)=   "lbGuestAge"
      Tab(0)ContCtrlCap(4)=   "lblGuestFullName"
      Tab(0)ContCtrlCap(5)=   "Label11"
      Tab(0)ContCtrlCap(6)=   "Label5"
      Tab(0)ContCtrlCap(7)=   "Label6"
      Tab(0)ContCtrlCap(8)=   "Label7"
      TabCaption(1)   =   "Room Detailes"
      TabContCtrlCnt(1)=   6
      Tab(1)ContCtrlCap(1)=   "txtDesc"
      Tab(1)ContCtrlCap(2)=   "Label15"
      Tab(1)ContCtrlCap(3)=   "txtRoomCategory"
      Tab(1)ContCtrlCap(4)=   "Label10"
      Tab(1)ContCtrlCap(5)=   "Label12"
      Tab(1)ContCtrlCap(6)=   "Label13"
      TabTheme        =   1
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.TextBox txtCitizenship 
         BorderStyle     =   0  'None
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
         Left            =   1440
         TabIndex        =   28
         Text            =   "Citizenship"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtDesc 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -73560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmBookIn.frx":127E
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Label15 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   405
         Left            =   -73560
         TabIndex        =   22
         Text            =   "Room Number"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtRoomCategory 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         TabIndex        =   21
         Text            =   "Category"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox lblDateOfBookin 
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
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Text            =   "Date"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox lbGuestAge 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   16
         Text            =   "Age"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox lblGuestFullName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Guest Name"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   26
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   24
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of birth"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   675
         Width           =   465
      End
   End
   Begin lvButton.lvButtons_H cmdCheckIn 
      Height          =   495
      Left            =   8880
      TabIndex        =   40
      Top             =   8160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "Check In"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483638
   End
   Begin lvButton.lvButtons_H cmdCalculate 
      Height          =   495
      Left            =   6600
      TabIndex        =   41
      Top             =   8160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "Calculate charges"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483638
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   10
      X1              =   0
      X2              =   15240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking In"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F556A&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1710
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmBookIn.frx":128A
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmBookIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBooking As New ADODB.Recordset
Dim mYsQL(2) As String

Private Sub cmdCalculate_Click()
mYsQL(0) = "select * from tbl_Room where Room_No =" + Me.Label15.Text
mYsQL(1) = "select * from tbl_Guest where Guest_ID =" + Me.txtGuestID

    If Me.txtGuestID = "" Then
    MsgBox "Please select a guest.", vbExclamation, App.Title
    Exit Sub
    End If

    If Me.txtRoomCategory.Text = "Category" Or txtRoomCategory = "" Then
    MsgBox "Please select a room.", vbExclamation, App.Title
    Exit Sub
    End If
    
    If Me.cboStayMode.Text = "" Then
    MsgBox "Please select a boarding type.", vbExclamation, App.Title
    Me.cboStayMode.SetFocus
    Exit Sub
    End If
    Calculate
End Sub

Private Sub cmdCheckIn_Click()
If pubCredit = True Then
MsgBox "Guest has not cleared payment!", vbExclamation, App.Title
Exit Sub
End If
mYsQL(0) = "select * from tbl_Room where Room_No =" + Me.Label15.Text
mYsQL(1) = "select * from tbl_Guest where Guest_ID =" + Me.txtGuestID
mdlPayment.recTransact txtGuestID.Text, Val(txtAmountDue.Text), IIf(Me.optPayment(1).Value = True, "Cheque", "Cash"), Val(txtAmountPaid.Text), Me.txtChequeNo.Text, pubCredit, "Accomodation"
Guest_Checkin
End Sub

Private Sub cmdGetGuestID_Click()
strsearch = ""
Me.cmdCheckIn.Enabled = False
flag = 2
frmGuestList.Show 1
mYsQL(0) = "select * from tbl_Guest where Guest_ID = " & strsearch
Call Form_Refresh
End Sub

Private Sub cmdGetRoom_Click()
strsearch = ""
Me.cmdCheckIn.Enabled = False
flag = "3"
frmViewRooms.Show 1
If strsearch <> "" Then
mYsQL(0) = "select * from tbl_Room where Room_No =" & strsearch
Room_Details
End If
End Sub

Private Sub Form_Load()
OpenDataBase

mdlFunctions.Fill_Area "Boarding", "tbl_Boarding", Me.cboStayMode
Me.dtDate(0).MaxDate = Date
Me.dtDate(1).MinDate = Date
Me.dtDate(0).Value = Date
Me.dtDate(1).Value = Me.dtDate(1).Value + 1
End Sub
Private Sub Form_Refresh()
Dim FullName As String
Dim rsBooking As New ADODB.Recordset
On Error Resume Next
rsBooking.Open mYsQL(0), pubcnn, adOpenDynamic, adLockOptimistic
Me.txtGuestID = rsBooking.Fields("Guest_ID").Value
With rsBooking.Fields
FullName = !Title + " " + !First_Name + " " + !Second_Name + " " + !Other_Name
Me.lblGuestFullName = FullName
Me.lbGuestAge.Text = !Age
Me.lblDateOfBookin = !BookinDate
If !Citizen = 0 Then
 Me.txtCitizenship = "Foreign"
 Else
 Me.txtCitizenship = "Local"
 End If
End With

End Sub
Private Sub Room_Details()
Dim rsBooking As New ADODB.Recordset
rsBooking.Open mYsQL(0), pubcnn, adOpenDynamic, adLockOptimistic
Me.txtRoom = rsBooking.Fields("Room_No").Value
With rsBooking.Fields
Me.Label15 = !Room_No
Me.txtRoomCategory = !Type
Me.txtDesc = !Description
End With
End Sub
Private Sub Guest_Checkin()
Dim rsBooking As New ADODB.Recordset
Set rsBooking = New ADODB.Recordset
rsBooking.Open mYsQL(0), pubcnn, adOpenDynamic, adLockOptimistic
If Me.txtGuestID.Text <> "" And Me.txtRoom <> "" Then
rsBooking.MoveFirst
    With rsBooking.Fields
    !Guest_ID = Me.txtGuestID
    !Status = 1
    End With
    rsBooking.Update
    rsBooking.Close

rsBooking.Open mYsQL(1), pubcnn, adOpenDynamic, adLockOptimistic
rsBooking.MoveFirst
With rsBooking.Fields
!GuestIn = 1
!CheckInDate = Me.dtDate(0).Value
!CheckOutDate = Me.dtDate(1).Value
End With
rsBooking.Update
MsgBox "Guest checked in", vbInformation, "Check In"
mdldb.Clear_All Me
cmdCheckIn.Enabled = False
Else
MsgBox "Select a Guest and room", vbInformation, "Guest Check In"
End If

End Sub

Private Sub Calculate()
Dim intDuration, intRate As Integer
Dim strMode, strRoom As String
Dim commission As Long
Dim tax As Long
Dim GrossCost As Long
Dim pubTax As String

intDuration = Val(dtDate(1).Value - dtDate(0).Value)
strMode = Me.cboStayMode.Text
strRoom = txtRoomCategory.Text
If strMode = "Full board" And strRoom = "Single room" Then intRate = 5100
If strMode = "Full board" And strRoom = "Double room" Then intRate = 5500
If strMode = "Full board" And strRoom = "Single room self contained" Then intRate = 5700
If strMode = "Full board" And strRoom = "Double room self contained" Then intRate = 6300

If strMode = "Half board" And strRoom = "Single room" Then intRate = 3900
If strMode = "Half board" And strRoom = "Double room" Then intRate = 4300
If strMode = "Half board" And strRoom = "Single room self contained" Then intRate = 4500
If strMode = "Half board" And strRoom = "Double room self contained" Then intRate = 5100

If strMode = "Bed and breakfast" And strRoom = "Single room" Then intRate = 2500
If strMode = "Bed and breakfast" And strRoom = "Double room" Then intRate = 2900
If strMode = "Bed and breakfast" And strRoom = "Single room self contained" Then intRate = 3100
If strMode = "Bed and breakfast" And strRoom = "Double room self contained" Then intRate = 3700

pubCitizen = IIf(Me.txtCitizenship = "Local", True, False)
GrossCost = Str(intDuration * intRate)

pubTax = IIf(pubCitizen = False, "0.02", "0")
tax = Str(GrossCost * Val(pubTax))
amtdue = Val(GrossCost + commission + tax)
amtpaid = Val(txtAmountPaid.Text)
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
If Me.optCurrency(1).Value = True Or Me.optPayment(1).Value = True Then
commission = Val(amtdue * 0.05)

End If
 txtCommission = commission
 Me.txtAmountDue = amtdue
Me.txtTax = tax
 

Me.cmdCheckIn.Enabled = True

End Sub


Private Sub optCurrency_Click(Index As Integer)
Calculate
End Sub

Private Sub optPayment_Click(Index As Integer)
Calculate
If optPayment(1).Value = True Then
Me.txtChequeNo.Enabled = True
Me.txtChequeNo.Text = ""
Else
Me.txtChequeNo.Text = ""
Me.txtChequeNo.Enabled = False
End If
End Sub

Private Sub txtAmountPaid_Change()
Calculate
End Sub
