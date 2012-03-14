VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmGuest 
   AutoRedraw      =   -1  'True
   Caption         =   " Guest"
   ClientHeight    =   8550
   ClientLeft      =   4365
   ClientTop       =   1710
   ClientWidth     =   10755
   Icon            =   "frmGuest.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   10755
   WindowState     =   2  'Maximized
   Begin VB.Frame fme2 
      Caption         =   "Guest Entry Succesfully Created!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   3600
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
      Begin lvButton.lvButtons_H cmdViewAllGuests 
         Height          =   435
         Left            =   600
         TabIndex        =   31
         Top             =   2340
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
         Caption         =   "Manage Guest List"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmGuest.frx":000C
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdAddNewEntry 
         Height          =   435
         Left            =   600
         TabIndex        =   32
         Top             =   1740
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
         Caption         =   "Add Another New Entry"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmGuest.frx":05A6
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdCheckIn 
         Height          =   435
         Left            =   600
         TabIndex        =   33
         Top             =   1200
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
         Caption         =   "Check In This Guest"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmGuest.frx":0B40
         cBack           =   16185592
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "New Guest entry successfull created!!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   600
         TabIndex        =   34
         Top             =   360
         Width           =   4245
      End
   End
   Begin VB.Frame fmeMain 
      Height          =   8295
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   9975
      Begin VB.Frame fme1 
         Caption         =   "Other Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   360
         TabIndex        =   19
         Top             =   3720
         Width           =   9375
         Begin VB.TextBox txtPassportNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   9
            Top             =   2640
            Width           =   2655
         End
         Begin VB.ComboBox cboGender 
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
            ItemData        =   "frmGuest.frx":10DA
            Left            =   1680
            List            =   "frmGuest.frx":10E4
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox txtAge 
            Appearance      =   0  'Flat
            DataField       =   "Age"
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Top             =   810
            Width           =   735
         End
         Begin VB.TextBox txtNatioalID 
            Appearance      =   0  'Flat
            DataField       =   "National_ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   20
            Top             =   3120
            Width           =   2655
         End
         Begin VB.ComboBox cboCountry 
            Appearance      =   0  'Flat
            DataField       =   "Country"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "ddd, dd MMMM yyyy"
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
            Height          =   360
            ItemData        =   "frmGuest.frx":10F6
            Left            =   1680
            List            =   "frmGuest.frx":1100
            TabIndex        =   7
            Text            =   "cboCountry"
            Top             =   1680
            Width           =   2775
         End
         Begin VB.CheckBox chkCitizen 
            Appearance      =   0  'Flat
            Caption         =   "Citizen"
            DataField       =   "Citizen"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   1680
            TabIndex        =   8
            Top             =   2040
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dobPicker 
            Bindings        =   "frmGuest.frx":1111
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   375
            Left            =   5520
            TabIndex        =   5
            Tag             =   "Date of birth"
            ToolTipText     =   "Day/Month/Year"
            Top             =   840
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmGuest.frx":1131
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115146755
            CurrentDate     =   33041
            MaxDate         =   2952895
            MinDate         =   2
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
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
            Left            =   360
            TabIndex        =   29
            Top             =   1260
            Width           =   675
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
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
            Left            =   360
            TabIndex        =   26
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblPassport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Passport NO."
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
            Left            =   360
            TabIndex        =   25
            Top             =   2640
            Width           =   1200
         End
         Begin VB.Label lblNationalID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NationalID No. "
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
            Left            =   360
            TabIndex        =   24
            Top             =   3120
            Width           =   1350
         End
         Begin VB.Label lblDOB 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Birth"
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
            Left            =   4320
            TabIndex        =   23
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label lblCountry 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
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
            Left            =   360
            TabIndex        =   22
            Top             =   1800
            Width           =   675
         End
      End
      Begin VB.Frame fmeGuestDetails 
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   4575
         Begin VB.ComboBox cboTitle 
            Height          =   315
            ItemData        =   "frmGuest.frx":144B
            Left            =   1560
            List            =   "frmGuest.frx":1458
            TabIndex        =   41
            Text            =   "Mr"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtFirstName 
            Appearance      =   0  'Flat
            DataField       =   "First_Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   1
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtSecondName 
            Appearance      =   0  'Flat
            DataField       =   "Second_Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   2
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Title"
            Height          =   375
            Left            =   480
            TabIndex        =   42
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblFirstName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   960
            Width           =   750
         End
         Begin VB.Label lblSecondName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name"
            Height          =   195
            Left            =   480
            TabIndex        =   17
            Top             =   1440
            Width           =   930
         End
      End
      Begin VB.Frame fmeGuestDetails 
         Caption         =   "Contacts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   1
         Left            =   5160
         TabIndex        =   13
         Top             =   720
         Width           =   4575
         Begin VB.ComboBox cboContryCode 
            Appearance      =   0  'Flat
            DataField       =   "Country"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "ddd, dd MMMM yyyy"
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
            Height          =   360
            ItemData        =   "frmGuest.frx":146E
            Left            =   1440
            List            =   "frmGuest.frx":1478
            TabIndex        =   38
            Text            =   "cboCountry"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            DataField       =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3360
            MaxLength       =   20
            TabIndex        =   37
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            DataField       =   "Phone"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   36
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            DataField       =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   3
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  'Flat
            DataField       =   "Phone"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   2640
            TabIndex        =   40
            Top             =   960
            Width           =   195
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3000
            TabIndex        =   39
            Top             =   480
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail Address"
            Height          =   195
            Left            =   360
            TabIndex        =   35
            Top             =   1680
            Width           =   1050
         End
         Begin VB.Label lblAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   360
            TabIndex        =   15
            Top             =   600
            Width           =   570
         End
         Begin VB.Label lblPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone NO."
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   1080
            Width           =   795
         End
      End
      Begin VB.TextBox txtGuestID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         DataField       =   "Guest_ID"
         DataMember      =   "Guest_ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "Guest Id"
         Top             =   360
         Width           =   855
      End
      Begin lvButton.lvButtons_H cmdSave 
         Default         =   -1  'True
         Height          =   360
         Left            =   7800
         TabIndex        =   27
         Top             =   7680
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   635
         Caption         =   "&Save"
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
      Begin lvButton.lvButtons_H cmdExit 
         Cancel          =   -1  'True
         Height          =   360
         Left            =   6360
         TabIndex        =   28
         Top             =   7680
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   635
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
      Begin VB.Label lblID 
         BackStyle       =   0  'Transparent
         Caption         =   "Guest ID."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   9360
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuest.frx":1489
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuest.frx":1C03
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   7440
      Width           =   975
   End
End
Attribute VB_Name = "frmGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Country, dat, res As String, x As Integer
Dim Complete As Boolean
Dim strMissing As String
Dim flag As Integer
Dim i As Integer
Dim Y As Long
Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub cboCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
If Me.chkCitizen.Value = 0 Then
Me.txtNatioalID.SetFocus
Else
Me.txtPassportNo.SetFocus
End If
End If
End Sub

Private Sub cboTitle_GotFocus()
        SendKeys "{Home}+{End}"
End Sub

Private Sub cboTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
KeyAscii = 0
SendKeys "{Tab}"

End If
End Sub

Private Sub chkCitizen_Click()
If chkCitizen.Value = 0 And flag = 1 Or flag = 2 Then
cboCountry.Enabled = True
txtPassportNo.Enabled = True
Else
cboCountry.Enabled = False
txtPassportNo.Enabled = False
txtPassportNo.Text = "-"
cboCountry.Text = "Kenya (Republic of)"
End If
End Sub

Private Sub chkCitizen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then SendKeys "{Tab}"
End Sub

Public Sub Form_Cancel()
flag = 3
    Call Clear_All(Me)
 Button_setting frmGuest, pubRst, flag

End Sub

Public Sub Form_Delete()
flag = 3
If pubRst.RecordCount Then
res = MsgBox("Do you relly want to delete this guest record?", vbQuestion + vbYesNo, "Delete")
If res = vbYes Then
    pubRst.Delete

 Call Clear_All(Me)
Else
Exit Sub
End If
End If
 Button_setting frmGuest, pubRst, 3
End Sub

Public Sub Form_Edit()
flag = 2
 Button_setting frmGuest, pubRst, flag
Call frmEnable(True)
Me.cboTitle.SetFocus
End Sub
Public Sub form_new()
res = MsgBox("Add New Record?", vbOKCancel + vbQuestion, "Guests")
If res = vbOK Then
Call frmEnable(True)
Me.dobPicker.Value = Date - 365 * 18
flag = 1
 Button_setting frmGuest, pubRst, flag
 mdldb.Clear_All Me
fmeMain.Visible = True
fme2.Visible = False
txtGuestID.Text = GetNewID("tbl_Guest", "Guest_ID", "G")
mdlFunctions.Button_setting Me, pubRst, flag
End If
End Sub
Public Sub Form_Save()
Call Save_Guest
End Sub
Private Sub cmdAddNewEntry_Click()
form_new
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If Check(Me) = True Then
Save_Guest
fme2.Left = Me.ScaleWidth / 2 - 1000
fme2.Top = Me.ScaleHeight / 2 - 1000
fme2.Visible = True
fmeMain.Visible = False
End If
txtGuestID.BackColor = &HC0FFFF
End Sub
Private Sub cmdCheckIn_Click()
If DataBaseToForm("select distinct GuestIn from tbl_Guest where Guest_ID = " & txtGuestID.Text) = True Then
If pubRst.Fields("GuestIn") = 0 Then
strsearch = txtGuestID.Text
Unload Me
frmBookIn.Show
    End If
    End If

'res = MsgBox("The guest aready checked in." + vbCr + "Check Out guest?", vbQuestion, App.Title)
'    If res = vbYes Then
'    frmBookIn.Show
'    flag = 4
'    strsearch = txtGuestID.Text
'    Unload Me
'    frmBookIn.Show
'    ElseIf res = vbNo Then
'    Exit Sub
'    End If
'    strsearch = txtGuestID.Text
'Unload Me
'frmBookIn.Show
'End If
'    Else
'strsearch = txtGuestID.Text
'Unload Me
'frmBookIn.Show
'End If
End Sub

Private Sub cmdViewAllGuests_Click()
Unload Me
frmGuestView.Show
End Sub

Private Sub dobPicker_Change()
pubAge = Str(Int((Date - dobPicker.Value) / 365))
txtAge.Text = pubAge
End Sub
Private Sub dobPicker_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    txtGuestID.Text = GetNewID("tbl_Guest", "Guest_ID", "G")
If strsearch <> "" Then
    pubsql = "select * from tbl_Guest where Guest_ID = " & strsearch
    Display (pubsql)
    flag = 2
Else
flag = 1
    txtGuestID.Text = GetNewID("tbl_Guest", "Guest_ID", "G")
End If
mdlFunctions.Fill_Area "Country", "tbl_Country", Me.cboContryCode
dat = App.path + "\Text\Count.txt"
x = 0
Open dat For Input As 1
Do Until EOF(1)
Input #1, Country
cboCountry.List(x) = Country
x = x + 1
Loop
Close 1
End Sub

Private Function frmEnable(Kweli As Boolean)
Me.cboCountry.Enabled = Kweli
Me.cboCountry.Enabled = Kweli
Me.cboTitle.Enabled = Kweli
Me.chkCitizen.Enabled = Kweli
Me.dobPicker.Enabled = Kweli
Me.txtAddress.Enabled = Kweli
Me.txtFirstName.Enabled = Kweli
Me.txtNatioalID.Enabled = Kweli
'Me.txtOtherName.Enabled = Kweli
Me.txtPhone.Enabled = Kweli
Me.txtSecondName.Enabled = Kweli
Me.txtPassportNo.Enabled = Kweli

End Function
Private Function Save_Guest()
    If flag = 1 Then
        txtGuestID.Text = GetNewID("tbl_Guest", "Guest_ID", "G")
        Call Updatedb
        fmeMain.Visible = False
        fme2.Visible = True
        flag = 2
        Button_setting frmGuest, pubRst, flag
    End If
    
    
    If flag = 2 Then
        Call Updatedb
        flag = 2
        Button_setting frmGuest, pubRst, flag
    End If
End Function

Private Sub txtNatID_KeyPress(KeyAscii As Integer)
Only_Numbers KeyAscii
End Sub
Private Sub txtPassport_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case Asc(1) To Asc(9)
  Case Asc(0)
  Case Else
  'Inform user of invalid entry
  Beep
  KeyAscii = 0
  End Select

End Sub

Public Function Display(ByVal pSQL As String)
On Error Resume Next
    LastSQL = pSQL
     If mdlData.DataBaseToForm(pSQL) = False Then
   Exit Function
 End If
 With pubRst
    Me.txtGuestID.Text = !Guest_ID
    Me.cboTitle.Text = !Title
    Me.txtFirstName.Text = !First_Name
    Me.txtSecondName.Text = !Second_Name
   ' Me.txtOtherName = !Other_Name
    Me.txtAge = !Age
    Me.dobPicker.Value = !DOB
    Sex = !Sex
    If Sex = "Male" Then
    Me.cboGender = "Male"
    Else
    Me.cboGender = "Female"
    End If
    Me.cboCountry.Text = !Country
 
    Me.txtPassportNo.Text = !Passport_No
   Me.txtNatioalID.Text = !National_ID
    Me.txtPhone.Text = !Phone
    Me.txtAddress.Text = !Address
    End With
        bleDataChanged = False
txtGuestID.BackColor = &HC0FFFF

End Function

Private Sub Form_Unload(Cancel As Integer)

Button_setting frmMain, pubRst, 3
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then SendKeys "{Tab}"
End Sub

Private Sub txtAge_GotFocus()
dobPicker.SetFocus
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
KeyAscii = 0
SendKeys "{Tab}"
End If
End Sub

Private Sub txtFirstName_LostFocus()
Me.txtFirstName = cSentenceCase(txtFirstName.Text)
End Sub

Private Sub txtNatioalID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
KeyAscii = 0
SendKeys "{Tab}"
End If
End Sub

Private Sub txtOtherName_GotFocus()
        SendKeys "{Home}+{End}"
End Sub

Private Sub txtOtherName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
KeyAscii = 0
SendKeys "{Tab}"
End If
End Sub

Private Sub txtOtherName_LostFocus()
txtOtherName.Text = mdlFX.cSentenceCase(txtOtherName.Text)
End Sub

Private Sub txtPassportNo_GotFocus()
        SendKeys "{Home}+{End}"
End Sub

Private Sub txtPassportNo_KeyPress(KeyAscii As Integer)
Call mdldb.Only_Numbers(KeyAscii)

If Me.chkCitizen.Value = 0 Then
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
KeyAscii = 0
SendKeys "{Tab}"
End If
Else
KeyAscii = 0
Exit Sub
End If
End Sub

Private Sub txtPassportNo_LostFocus()
mdldb.CheckAdd_Primary Me, "tbl_Guest", txtPassportNo, "Passport_No"
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
Call mdldb.Only_Numbers(KeyAscii)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
KeyAscii = 0
SendKeys "{Tab}"
End If
End Sub
Private Function Updatedb()
Dim pubRst As New ADODB.Recordset
Set pubRst = New ADODB.Recordset

pubRst.Open "select * from tbl_Guest", pubcnn, adOpenDynamic, adLockOptimistic

With pubRst

If flag = 1 Then
    .AddNew
    .Fields("Guest_ID").Value = txtGuestID.Text
End If
    .Fields("Title").Value = cboTitle.Text
    .Fields("First_Name").Value = Me.txtFirstName.Text
    .Fields("Second_Name").Value = Me.txtSecondName.Text
    .Fields("Citizen").Value = Me.chkCitizen.Value
    .Fields("Sex").Value = Me.cboGender.Text
    
If Me.txtPassportNo.Text <> "" Then .Fields("Passport_No").Value = Me.txtPassportNo.Text
If Me.txtNatioalID.Text <> "" Then .Fields("National_ID").Value = Me.txtNatioalID.Text
.Fields("Address").Value = Me.txtAddress.Text
.Fields("Phone").Value = Me.txtPhone.Text
.Fields("Country").Value = Me.cboCountry.Text
.Fields("DOB").Value = Me.dobPicker.Value
.Fields("Age") = txtAge.Text
.Update

End With
txtGuestID.BackColor = &HC0FFFF
End Function

Private Sub txtSecondName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = vbKeySpace Or KeyAscii = vbKeyTab Then
KeyAscii = 0
SendKeys "{Tab}"
End If
End Sub

Private Sub txtSecondName_LostFocus()
txtSecondName.Text = mdlFX.cSentenceCase(txtSecondName.Text)
End Sub
Private Sub ClearMe()
Dim mycontrol As Control
For Each mycontrol In Me.Controls
    If TypeOf mycontrol Is TextBox Then
        mycontrol.Text = ""
    ElseIf TypeOf mycontrol Is ComboBox Then
            mycontrol.Text = ""
    ElseIf TypeOf mycontrol Is DTPicker Then
            mycontrol.Value = Date - 365 * 18

        End If
    
Next
End Sub

