VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManageServices 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Services"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   Icon            =   "frmManageServices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Room Service"
      TabPicture(0)   =   "frmManageServices.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Other Services"
      TabPicture(1)   =   "frmManageServices.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Other Services"
         Height          =   2535
         Left            =   -74760
         TabIndex        =   13
         Top             =   600
         Width           =   4215
         Begin VB.ComboBox cboServices 
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
            Left            =   480
            TabIndex        =   14
            Text            =   "Combo1"
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Select a service"
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Room service"
         Height          =   2175
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   4095
         Begin VB.CheckBox chkMeal 
            Caption         =   "Teas"
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
            Index           =   3
            Left            =   480
            TabIndex        =   12
            Tag             =   "250"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CheckBox chkMeal 
            Caption         =   "Dinner"
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
            Index           =   2
            Left            =   480
            TabIndex        =   11
            Tag             =   "1500"
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox chkMeal 
            Caption         =   "Lunch"
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
            Index           =   1
            Left            =   480
            TabIndex        =   10
            Tag             =   "700"
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox chkMeal 
            Caption         =   "Breakfast"
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
            Index           =   0
            Left            =   480
            TabIndex        =   9
            Tag             =   "1000"
            Top             =   360
            Width           =   2175
         End
      End
   End
   Begin VB.Frame fmeDetails 
      Caption         =   "Guest"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4095
      Begin VB.CommandButton cmdGetGuestID 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2865
         Picture         =   "frmManageServices.frx":0044
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   510
         Width           =   345
      End
      Begin VB.TextBox txtGuestID 
         Height          =   375
         Left            =   600
         MaxLength       =   20
         TabIndex        =   4
         Top             =   480
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
         TabIndex        =   5
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Charges"
      Height          =   1335
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin VB.TextBox txtAmountDue 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Ksh""#,##0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
      End
   End
   Begin lvButton.lvButtons_H cmdSaveTransaction 
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   5760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "Save transaction"
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
Attribute VB_Name = "frmManageServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim intCharge, i, ItemX As Integer
Dim pubRst As New ADODB.Recordset
Dim strServ, strFormat As String

Private Sub cboServices_Change()
ItemX = Me.cboServices.ListIndex
Me.txtAmountDue = Me.cboServices.ItemData(Me.cboServices.ListIndex)
strServ = "Service_ID" & Me.cboServices.List(Me.cboServices.ListIndex)
End Sub

Private Sub cboServices_Click()
Me.txtAmountDue = Format(Me.cboServices.ItemData(Me.cboServices.ListIndex), "0,0.00")
ItemX = strServ
End Sub

Private Sub chkMeal_Click(Index As Integer)
Calculate
End Sub
Private Sub cmdGetGuestID_Click()
strsearch = ""
flag = 1
frmGuestList.Show 1, Me
mYsQL = "select * from tbl_Guest where Guest_ID = " & strsearch
Me.txtGuestID.Text = strsearch

End Sub

Private Function Calculate()
intCharge = 0
If pubCitizen = True Then
pubRate = 0
Else
pubRate = 0.02
End If
For i = 0 To 3
If Me.chkMeal(i).Value = 1 Then
intCharge = intCharge + Val(Me.chkMeal(i).Tag) + Val(Me.chkMeal(i).Tag * pubRate)
End If
Next i
strFormat = "0,0.00"
Me.txtAmountDue = Format(intCharge, strFormat)
End Function

Private Sub cmdSaveTransaction_Click()
Set pubRst = New ADODB.Recordset
If Me.txtGuestID <> "" Then
    If Me.txtAmountDue.Text <> "" Then
        mdlPayment.recTransact txtGuestID.Text, Me.txtAmountDue, "", 0, "", False, "Room Service" & ItemX
    Else
        MsgBox "Please select a service", vbInformation, App.Title
    End If
    
txtGuestID.Text = ""
MsgBox "Transaction saved", vbInformation, App.Title
Else
MsgBox "Please select a Guest", vbInformation, App.Title
End If
End Sub

Private Sub Form_Load()
mdlFunctions.fillCombo cboServices, "tbl_Services", "Name"
End Sub

