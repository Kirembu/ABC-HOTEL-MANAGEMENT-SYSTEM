VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Vehicle"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   Icon            =   "frmEditVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7485
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      DataField       =   "Description"
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
      MaxLength       =   200
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtMileage 
      Appearance      =   0  'Flat
      DataField       =   "VMileage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2040
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.ComboBox cboVehicleType 
      Appearance      =   0  'Flat
      DataField       =   "VType"
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
      Left            =   2040
      TabIndex        =   5
      Text            =   "cboVehicleType"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtVehicleRegNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "VReg"
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
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.PictureBox btnExit 
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   360
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdBrowse 
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Browse.."
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdClear 
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Clear"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   5730
      TabIndex        =   12
      Top             =   3960
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
      Left            =   4200
      TabIndex        =   13
      Top             =   3960
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
      Caption         =   "Edit Vehicle"
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
      Left            =   60
      TabIndex        =   14
      Top             =   180
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   5040
      Top             =   840
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblMilage1 
      BackStyle       =   0  'Transparent
      Caption         =   "KM"
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
      Left            =   3960
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblVehicleReg 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Registration no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblMileage 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Current Mileage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2370
      Width           =   1815
   End
   Begin VB.Label lblVehicleType 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehice Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmEditVehicle.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmEditVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rsVehicle As ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim pubRst As ADODB.Recordset
Dim flag As Integer
Dim i As Integer
Dim res, pop  As String
Dim strDest, strImg, strPic As String
Dim fso  As New FileSystemObject
Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub cboVehicleType_LostFocus()
Dim found As Boolean
For i = 0 To Me.cboVehicleType.ListCount
If Me.cboVehicleType.List(i) = Me.cboVehicleType.Text Then
found = True
End If
Next i
If found = False Then
res = MsgBox("The vehicle type does not exist in current Database" + vbCr + "Do you wish to add a new type", vbQuestion + vbYesNo, Vehicles)
If res = vbYes Then
Call cmdType_Click
Me.cboVehicleType.SetFocus
Else
Me.cboVehicleType.Text = ""
Me.cboVehicleType.SetFocus
End If
End If
End Sub

Private Sub cmdBrowse_Click()
CommonDialog1.Filter = "All Images|*.Bmp;*.gif;*.jpg|Bitmaps|*.BMP|Gif Images|*.Gif|Jpeg Images|*.jpg"
CommonDialog1.ShowOpen
If Not Trim(CommonDialog1.FileName) = "" Then
strImg = CommonDialog1.FileName
Image1.Picture = LoadPicture(strImg)
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
Set Image1 = Nothing
strImg = ""
End Sub

Public Sub Form_Delete()
flag = 3
If rsTemp.RecordCount Then
res = MsgBox("Do you relly want to delete this guest record?", vbQuestion + vbYesNo, "Delete")
If res = vbYes Then
    rsTemp.Delete
Else
Exit Sub
End If
End If

Call frmEnable(False)
If rsTemp.RecordCount Then
        rsTemp.MoveFirst
        Call Disp
    Else
        Call Clear_All(Me)
End If
  Button_setting frmAddVehicle, rsTemp, flag

End Sub


Public Sub Form_Save()
flag = 1
Call Save

End Sub

Private Sub cmdSearch_Click()
frmTransport.Show 1, Me
End Sub


Public Sub Form_Cancel()
On Error Resume Next
flag = 3
If rsTemp.RecordCount Then
    rsTemp.MoveFirst
    Call Disp
Else
    Call Clear_All(Me)
End If
     Button_setting frmAddVehicle, rsTemp, flag

End Sub


Public Sub Form_Edit()
flag = 2
    Button_setting frmAddVehicle, rsTemp, flag
Call frmEnable(True)
Me.txtVehicleRegNo.SetFocus
End Sub


Public Sub form_new()
Unload Me
frmAddVehicle.Show
End Sub

Private Sub cmdType_Click()
frmVehicleTypes.Show 1
mdlFunctions.Fill_Area "Vehicle_Type", "tbl_VehicleTypes", cboVehicleType
End Sub

Private Sub cmdSave_Click()
Call Save
End Sub

Private Sub Form_Load()
mdlFX.MakeGradient Me, 1
    Set rsTemp = New ADODB.Recordset
    OpenDataBase
    strsearch = "KCB 355J"
    rsTemp.Open "select * from tbl_Vehicles where Vehicle_Reg ='" & strsearch + "'", pubcnn, adOpenForwardOnly, adLockOptimistic
'Call NewVehicleNo
flag = 3
 Button_setting frmAddVehicle, rsTemp, flag
mdlFunctions.Fill_Area "Vehicle_Type", "tbl_VehicleTypes", cboVehicleType
Disp
End Sub

Private Sub frmEnable(Kweli As Boolean)
Me.txtDescription.Enabled = Kweli
Me.txtVehicleRegNo.Enabled = Kweli
Me.cboVehicleType.Enabled = Kweli
Me.txtMileage.Enabled = Kweli
End Sub


Private Function Save()
'On Error Resume Next
        '*************************************************************************************
        '*********Image Copy Script **********************************************************
        If Not Trim(strImg) = "" Then
        strDest = App.path & "/Images"
        If Not fso.FolderExists(strDest) Then
            fso.CreateFolder (strDest)
        End If
        strDest = strDest & "/Cars"
        If Not fso.FolderExists(strDest) Then
            fso.CreateFolder (strDest)
        End If
            strDest = strDest & "/" & Trim(Me.txtVehicleRegNo.Text) & Trim(Right(strImg, 4))
            strSource = strImg
            strImageName = Trim(txtBID) & Trim(Right(strImg, 4))
            If fso.FileExists(strDest) Then
                SetAttr strDest, vbNormal
            End If
            fso.CopyFile strSource, strDest, True
        End If
        '******************************************************************************

With rsTemp
.Fields("Vehicle_Reg").Value = Me.txtVehicleRegNo
.Fields("Vehicle_Mileage").Value = Me.txtMileage.Text
.Fields("Vehicle_Description").Value = Me.txtDescription.Text
If strDest <> "" Then
.Fields("Img_Path").Value = fso.GetFileName(strDest)
End If
'.Fields("CreatedBy").Value = Me.txtOtherName.Text
.Update
End With
    flag = 3
 Button_setting frmAddVehicle, rsTemp, flag
    

End Function

Private Function Updatedb()

End Function
Private Function NewVehicleNo() As Long
On Error Resume Next
With rsTemp
            If .BOF Then
                txtVehicleID = 1
            Else
    .MoveLast
    If Not IsNull(rsTemp("Vehicle_ID")) Then
                txtVehicleID = !Vehicle_ID + 1
Else
txtVehicleID = 1
      End If
      End If
            NewVehicleNo = txtVehicleID
End With
End Function
Private Sub refreshData()
    Set rsDisp = New ADODB.Recordset
    rsDisp.Open "select * from tbl_Vehicles order by CreationDate", pubcnn, 1, 2
  Button_setting frmAddVehicle, rsTemp, flag
End Sub
Private Sub Disp()

Call Clear_All(Me)
With rsTemp
    If .BOF = False And .EOF = False Then

    If Not IsNull(rsTemp("Vehicle_Reg")) Then
    Me.txtVehicleRegNo = !Vehicle_Reg
    End If
    If Not IsNull(rsTemp("Vehicle_Description")) Then
    Me.txtDescription = !Vehicle_Description
    End If
    If Not IsNull(rsTemp("Vehicle_Mileage")) Then
    Me.txtMileage = !Vehicle_Mileage
    End If
    If Not IsNull(rsTemp("Vehicle_Type")) Then
    Me.cboVehicleType.Text = !Vehicle_Type
    End If
    If Not IsNull(rsTemp("Img_Path")) Then
        strPic = App.path & "/Images/Cars/" & Trim(!Img_Path)
        If fso.FileExists(strPic) Then
        Me.Image1.Picture = LoadPicture(strPic)
        End If
        Else
        Me.Image1.Picture = LoadPicture(App.path & "/Images/Cars/None.bmp")
End If
End If
    End With
End Sub

Private Sub frmClear()
Me.txtDescription = ""
Me.txtMileage = ""

Me.txtVehicleRegNo = ""

End Sub

Private Sub txtMileage_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub

Private Sub txtVehicleRegNo_Change()
If flag <> 3 Then
If mdlUsers.CheckDuplicates("tbl_Vehicles", "Vehicle_Reg", Me.txtVehicleRegNo.Text, Me.txtVehicleRegNo) = True Then
MsgBox "Registration Number already exists." + vbCr + "Enter a new Registration.", vbInformation, "New User"
End If
txtVehicleRegNo.BackColor = &HC0FFFF
Else
flag = 2
End If
End Sub

