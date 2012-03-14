VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmTransport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mange Vehicles"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTransport.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVehicleReg 
      BackColor       =   &H00F6F8F8&
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
      Height          =   360
      Left            =   150
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtVtype 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   150
      TabIndex        =   0
      Top             =   2355
      Width           =   2175
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Add"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   5070
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Edit"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   5460
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Delete"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdReload 
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   5850
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Reload"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin MSComctlLib.ListView lvwVehicles 
      Height          =   6150
      Left            =   2430
      TabIndex        =   6
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10848
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ilRecordIcos"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Vehicle_Reg"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vehicle_Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vehicle_Mileage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   3810
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   7920
      TabIndex        =   11
      Top             =   6720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Caption         =   "&Close"
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
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdRent 
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Top             =   6240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Rent Out Vehicle"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmTransport.frx":0442
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   11145
   End
   Begin VB.Image imgVehicle 
      Height          =   1335
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   30
      Picture         =   "frmTransport.frx":04DF
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transport management"
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
      Left            =   750
      TabIndex        =   10
      Top             =   180
      Width           =   3345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Vehicle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2085
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Registration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   1290
      Width           =   2295
   End
   Begin VB.Label lblFullName 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   2040
      Width           =   2235
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   0
      Picture         =   "frmTransport.frx":0921
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10905
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   6165
      Left            =   0
      Top             =   600
      Width           =   2430
   End
End
Attribute VB_Name = "frmTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
flag = 1
frmAddVehicle.Show
End Sub

Private Sub cmdDelete_Click()
'On Error Resume Next
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
DelImg (App.path & "/Images/Cars/" & Me.lvwVehicles.ListItems(Me.lvwVehicles.SelectedItem.Index).SubItems(5))
rs.Open "select * from tbl_Vehicles where Vehicle_Reg = '" & Me.lvwVehicles.ListItems(Me.lvwVehicles.SelectedItem.Index).Text & "'", pubcnn, adOpenDynamic, adLockOptimistic
rs.Delete adAffectCurrent
cmdReload_Click
End Sub

Private Sub cmdEdit_Click()
With frmAddVehicle
strsearch = Me.lvwVehicles.ListItems(Me.lvwVehicles.SelectedItem.Index).Text
 flag = 2
 .Show
 End With
End Sub

Private Sub cmdReload_Click()
If flag <> 4 Then
Display ("select * from tbl_Vehicles")
Else
Display ("select * from tbl_Vehicles where Status = 0")
End If
End Sub
Private Sub Form_Load()
If flag <> 4 Then
Display ("select distinct Vehicle_Type ,Vehicle_Reg ,Vehicle_Mileage ,Vehicle_Description ,Type_Rate ,Img_Path  from tbl_Vehicles")
Else
Display ("select distinct Vehicle_Type ,Vehicle_Reg ,Vehicle_Mileage ,Vehicle_Description ,Type_Rate ,Img_Path from tbl_Vehicles where Status = 0")
End If
End Sub
Private Function Display(ByVal pSQL As String)
Dim x As Integer
Dim rsTransport As New ADODB.Recordset
On Error Resume Next
Set rsTransport = New ADODB.Recordset
rsTransport.Open pSQL, pubcnn, adOpenDynamic, adLockOptimistic
x = 0
Me.lvwVehicles.ListItems.Clear
    LastSQL = pSQL
   
Do While Not rsTransport.EOF
    x = x + 1

Me.lvwVehicles.ListItems.Add , , rsTransport("Vehicle_Reg")

    lvwVehicles.ListItems.Item(x).SubItems(1) = rsTransport("Vehicle_Type")
    lvwVehicles.ListItems.Item(x).SubItems(2) = rsTransport("Vehicle_Mileage")
    lvwVehicles.ListItems.Item(x).SubItems(3) = rsTransport("Vehicle_Description")
      lvwVehicles.ListItems.Item(x).SubItems(4) = rsTransport("Type_Rate")
    lvwVehicles.ListItems.Item(x).SubItems(5) = rsTransport("Img_Path")

    rsTransport.MoveNext
  Loop
End Function

Private Sub lvButtons_H1_Click()

End Sub

Private Sub lvwVehicles_Click()
With Me.lvwVehicles.ListItems.Item(Me.lvwVehicles.SelectedItem.Index)
 Me.txtVehicleReg = .SubItems(2)
 Me.txtVtype = .SubItems(1)
 imgBox App.path & "\Images\Cars\" & .SubItems(5), Me.imgVehicle
 End With
End Sub

Private Sub lvwVehicles_DblClick()
If flag = 4 Then
strsearch = Me.lvwVehicles.ListItems.Item(Me.lvwVehicles.SelectedItem.Index).Text
With Me.lvwVehicles.ListItems.Item(Me.lvwVehicles.SelectedItem.Index)
 pubMileage = .SubItems(2)
 pubRate = .SubItems(4)
 pubType = .SubItems(1)
 pubDesc = .SubItems(3)
 
 End With
 Unload Me
 End If
End Sub
