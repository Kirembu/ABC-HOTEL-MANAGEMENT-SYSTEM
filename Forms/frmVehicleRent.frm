VERSION 5.00
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmVehicleRent 
   Caption         =   "Vehicle Rent"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmVehicleRent.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame fmePayment 
      Caption         =   "Paymenet Type"
      Height          =   2535
      Left            =   120
      TabIndex        =   41
      Top             =   5040
      Width           =   5775
      Begin VB.Frame fmeCurrency 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   615
         Left            =   1080
         TabIndex        =   45
         Top             =   600
         Width           =   1695
         Begin VB.OptionButton optCurrency 
            Caption         =   "Local currency"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   47
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton optCurrency 
            Caption         =   "Foreignl currency"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   46
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
         TabIndex        =   44
         Tag             =   "ww"
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         TabIndex        =   48
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Transaction Details"
      Height          =   4335
      Left            =   7320
      TabIndex        =   28
      Top             =   360
      Width           =   4335
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         ToolTipText     =   "Charge on forreingers"
         Top             =   3000
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0"
         ToolTipText     =   "Charged on transactions made by cheque"
         Top             =   2280
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
         TabIndex        =   31
         Text            =   "0"
         ToolTipText     =   "Balace"
         Top             =   1680
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
      Begin VB.TextBox txtAmount 
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
         TabIndex        =   29
         Text            =   "0"
         ToolTipText     =   "Amount due"
         Top             =   960
         Width           =   1575
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
         TabIndex        =   38
         Top             =   3000
         Width           =   1335
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
         TabIndex        =   37
         Top             =   2280
         Width           =   1215
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
         TabIndex        =   36
         Top             =   1800
         Width           =   1095
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
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
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
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame fmeDetails 
      Caption         =   "Guest details and preference"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.TextBox txtDistance 
         Height          =   375
         Left            =   600
         TabIndex        =   39
         Text            =   "0"
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox txtNewMilage 
         Height          =   375
         Left            =   600
         MaxLength       =   6
         TabIndex        =   10
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtOldMilage 
         Height          =   375
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton cmdGetVehicle 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2865
         Picture         =   "frmVehicleRent.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   345
      End
      Begin VB.CommandButton cmdGetGuestID 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2865
         Picture         =   "frmVehicleRent.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   510
         Width           =   345
      End
      Begin VB.TextBox txtGuestID 
         Height          =   375
         Left            =   600
         MaxLength       =   20
         TabIndex        =   3
         Top             =   480
         Width           =   2625
      End
      Begin VB.TextBox txtVehicle 
         Height          =   375
         Left            =   600
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1290
         Width           =   2625
      End
      Begin VB.Label Label19 
         Caption         =   "Distance"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "New Mileage"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Old milleage"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle"
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
         TabIndex        =   6
         Top             =   1080
         Width           =   675
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
         TabIndex        =   5
         Top             =   240
         Width           =   195
      End
   End
   Begin prjXTab.XTab XGuestDetailes 
      Height          =   2655
      Left            =   7320
      TabIndex        =   11
      Top             =   5040
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4683
      TabCount        =   2
      TabCaption(0)   =   "Guest Detailes"
      TabContCtrlCnt(0)=   8
      Tab(0)ContCtrlCap(1)=   "lblGuestFullName"
      Tab(0)ContCtrlCap(2)=   "lbGuestAge"
      Tab(0)ContCtrlCap(3)=   "lblDateOfBookin"
      Tab(0)ContCtrlCap(4)=   "txtCitizenship"
      Tab(0)ContCtrlCap(5)=   "Label7"
      Tab(0)ContCtrlCap(6)=   "Label6"
      Tab(0)ContCtrlCap(7)=   "Label5"
      Tab(0)ContCtrlCap(8)=   "Label11"
      TabCaption(1)   =   "Vehicle Detailes"
      TabContCtrlCnt(1)=   8
      Tab(1)ContCtrlCap(1)=   "TXTRATE"
      Tab(1)ContCtrlCap(2)=   "txtVehicleType"
      Tab(1)ContCtrlCap(3)=   "Label15"
      Tab(1)ContCtrlCap(4)=   "txtDesc"
      Tab(1)ContCtrlCap(5)=   "Label3"
      Tab(1)ContCtrlCap(6)=   "Label13"
      Tab(1)ContCtrlCap(7)=   "Label12"
      Tab(1)ContCtrlCap(8)=   "Label10"
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
      Begin VB.TextBox TXTRATE 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Rate"
         Top             =   1680
         Width           =   2535
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
         TabIndex        =   18
         Text            =   "Guest Name"
         Top             =   600
         Width           =   2655
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
         TabIndex        =   17
         Text            =   "Age"
         Top             =   1080
         Width           =   1215
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
         TabIndex        =   16
         Text            =   "Date"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtVehicleType 
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
         TabIndex        =   15
         Text            =   "Category"
         Top             =   960
         Width           =   1935
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
         TabIndex        =   14
         Text            =   "Vehicle Regisration"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtDesc 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -73560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmVehicleRent.frx":0F56
         Top             =   1320
         Width           =   2535
      End
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
         TabIndex        =   12
         Text            =   "Citizenship"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate per KM:"
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
         TabIndex        =   49
         Top             =   1680
         Width           =   945
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
         TabIndex        =   25
         Top             =   675
         Width           =   465
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
         TabIndex        =   24
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of birth"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   975
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
         TabIndex        =   22
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
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
         TabIndex        =   21
         Top             =   960
         Width           =   420
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
         TabIndex        =   20
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1920
         Width           =   855
      End
   End
   Begin lvButton.lvButtons_H cmdCalculate 
      Height          =   495
      Left            =   600
      TabIndex        =   26
      Top             =   7920
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
   Begin lvButton.lvButtons_H cmdRentOut 
      Height          =   495
      Left            =   3480
      TabIndex        =   27
      Top             =   7920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "Save Transaction"
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
End
Attribute VB_Name = "frmVehicleRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsVehicle As New ADODB.Recordset
Dim mYsQL(1), res As String

Private Sub cmdCalculate_Click()
Me.txtAmount = Val(Me.txtDistance * Val(pubRate) + tax)
Calculate
End Sub

Private Sub cmdGetGuestID_Click()
strsearch = ""
Me.cmdRentOut.Enabled = False
flag = 3
frmGuestList.Show 1
mYsQL(0) = "select * from tbl_Guest where Guest_ID = " & strsearch
Me.txtGuestID = strsearch
Call Form_Refresh
End Sub


Private Sub cmdGetVehicle_Click()
flag = 4
frmTransport.Show 1
Me.txtVehicle = strsearch
Me.txtOldMilage = pubMileage
Me.Label15 = Me.txtVehicle
Me.txtVehicleType = pubType
Me.txtDesc = pubDesc
End Sub

Private Sub cmdRentOut_Click()
If pubCredit = True Then
    res = MsgBox("The bill has not been cleared. Continue with credit entry?", vbQuestion + vbYesNo, App.Title)
        If res = vbNo Then
            Exit Sub
        
        Else
        
                mdlPayment.recTransact txtGuestID.Text, Val(txtAmount.Text), IIf(Me.optPayment(1).Value = True, "Cheque", "Cash"), Val(txtAmountPaid.Text), Me.txtChequeNo.Text, pubCredit, "Vehicle"
                mdldb.Clear_All Me
        
        End If
Else
    If mdlFunctions.Check(Me) = True Then
        mdlPayment.recTransact txtGuestID.Text, Val(txtAmount.Text), IIf(Me.optPayment(1).Value = True, "Cheque", "Cash"), Val(txtAmountPaid.Text), Me.txtChequeNo.Text, pubCredit, "Vehicle"
        mdldb.Clear_All Me
End If
End If
End Sub

Private Sub txtDiscount_Change()

End Sub

Private Sub optCurrency_Click(Index As Integer)
Calculate
End Sub

Private Sub optPayment_Click(Index As Integer)
Calculate
End Sub

Private Sub txtAmountPaid_Change()
Calculate
End Sub

Private Sub txtDistance_Change()
If Val(txtDistance.Text) > 0 Then
Calculate
cmdRentOut.Enabled = True
Else
cmdRentOut.Enabled = False
End If
End Sub

Private Sub txtNewMilage_Change()
If Me.txtOldMilage <> "" Then
Me.txtDistance = Val(Val(Me.txtNewMilage) - Me.txtOldMilage)
Else
MsgBox "Please select a vehicle", vbInformation, App.Title
End If
End Sub

Private Sub txtNewMilage_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub

Private Sub txtNewMilage_LostFocus()
If txtNewMilage <= txtOldMilage And txtoldmillage <> 0 Then
MsgBox "New Milage cannot be less than Previous milage", vbInformation, App.Title
txtNewMilage.SetFocus
        SendKeys "{Home}+{End}"
        End If
End Sub

Private Sub txtOldMilage_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub
Private Sub Form_Refresh()
Dim FullName As String
Dim rsVehicle As New ADODB.Recordset
On Error Resume Next
rsVehicle.Open mYsQL(0), pubcnn, adOpenDynamic, adLockOptimistic
With rsVehicle.Fields
FullName = !Title & " " & !First_Name & " " & !Second_Name
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
Private Sub Vehicle_Details()
Dim rsVehicle As New ADODB.Recordset
On Error Resume Next
rsVehicle.Open mYsQL(1), pubcnn, adOpenDynamic, adLockOptimistic
With rsVehicle.Fields
Me.Label15.Text = !Vehicle_Reg
Me.txtDesc = !Vehicle_Description
Me.txtOldMilage = !Vehicle_Mileage
Me.txtVehicleType = !Vehicle_Type
Me.TXTRATE = IIf(!Type_Rate <> "", !Type_Rate, 1)
pubRate = TXTRATE
End With
End Sub
Private Sub Calculate()
Dim commission As Long
Dim tax As Long
Dim GrossCost As Long
Dim pubTax As String

pubCitizen = IIf(Me.txtCitizenship = "Local", True, False)
GrossCost = Val(Me.txtDistance * Val(pubRate) + tax)
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
 Me.txtAmount = amtdue
Me.txtTax = tax
 

 
 End Sub


