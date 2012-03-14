VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmVehicleTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Types"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTypeID 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   2610
      TabIndex        =   6
      Top             =   2520
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
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   1080
      TabIndex        =   7
      Top             =   2520
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Cancel"
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New vehicle type."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2550
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmVehicleTypes.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label lblTypeID 
      BackStyle       =   0  'Transparent
      Caption         =   "Type ID"
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
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblRate 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate For Type"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblTypeName 
      BackStyle       =   0  'Transparent
      Caption         =   "Types Name"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmVehicleTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim flag As Integer
Dim rsDisp As ADODB.Recordset
Dim rsAcct As ADODB.Recordset

Public Sub Form_Save()
If flag = 1 Or flag = 2 Then
Call Save
 
End If
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

If mdlUsers.CheckDuplicates("tbl_VehicleTypes", "Vehicle_Type", Me.txtType.Text, txtType) = False Then
Call Save
End If
End Sub

Private Sub Form_Load()
OpenDataBase
Me.txtTypeID.Text = mdlUsers.GetNewID("tbl_VehicleTypes", "Type_ID", "VT")
flag = 3
End Sub

Public Sub Form_Edit()
flag = 2
Button_setting frmAddRoom, rsDisp, flag
mdlFunctions.EnableInput Me, True
Me.txtRate.SetFocus
End Sub

Private Function Save()
    Set rsDisp = New ADODB.Recordset

    mYsQL = "select * from tbl_VehicleTypes order by Type_ID"
    rsDisp.Open mYsQL, pubcnn, adOpenDynamic, adLockOptimistic
Me.txtTypeID.Text = mdlUsers.GetNewID("tbl_VehicleTypes", "Type_ID", "VT")
If mdlFunctions.Check(Me) = True Then
Me.txtTypeID.BackColor = &HC0FFFF
With rsDisp
    .AddNew
    .Fields("Type_ID") = Me.txtTypeID.Text
    .Fields("Vehicle_Type") = Me.txtType
    .Fields("Type_Rate") = Me.txtRate
    .Update
    End With
 mdlFunctions.EnableInput Me, False
Else
Beep
Exit Function
End If
End Function
Private Sub txtRate_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub
