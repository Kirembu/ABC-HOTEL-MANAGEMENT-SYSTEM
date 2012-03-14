VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicles"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   Icon            =   "frmAddVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7710
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   14
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtVehicleType 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdBrowse 
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2760
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
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   40
      cBack           =   -2147483633
   End
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
      TabIndex        =   6
      Top             =   3840
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
      MaxLength       =   7
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
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
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin lvButton.lvButtons_H btnExit 
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   360
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
      Image           =   "frmAddVehicle.frx":000C
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdClear 
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   2760
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
      TabIndex        =   11
      Top             =   4440
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
      TabIndex        =   12
      Top             =   4440
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
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
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate per KM"
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
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddVehicle.frx":0625
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   5040
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1200
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
      TabIndex        =   7
      Top             =   3120
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
      Top             =   1080
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
      Top             =   3960
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
      Top             =   3090
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
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rsVehicle As ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim pubRst As ADODB.Recordset
Dim i As Integer
Dim res, pop  As String
Dim strDest, strImg, strPic As String
Dim fso  As New FileSystemObject
Private Sub btnExit_Click()
Unload Me
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
DelImg (App.path & "/Images/Cars/" & rsTemp.Fields("Img_Path").Value)
End Sub

Public Sub Form_Delete()
flag = 3
If rsTemp.RecordCount Then
res = MsgBox("Do you relly want to delete this guest record?", vbQuestion + vbYesNo, "Delete")
If res = vbYes Then
DelImg (App.path & "/Images/Cars/" & rsTemp.Fields("Img_Path").Value)
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
res = MsgBox("Add New Record?", vbOKCancel + vbQuestion, "Guests")
If res = vbOK Then
Call frmEnable(True)
Call frmClear
flag = 1

 Button_setting frmAddVehicle, rsTemp, flag
End If
End Sub

Private Sub cmdSave_Click()
If mdlFunctions.Check(Me) = False Then
MsgBox "Complete vehicle entry", vbInformation
Else
Call Save
End If
txtVehicleRegNo.BackColor = &HC0FFFF
End Sub

Private Sub Form_Load()
mdlFX.MakeGradient Me, 1
    Set rsTemp = New ADODB.Recordset
    OpenDataBase
    If flag = 1 Then
    rsTemp.Open "select * from tbl_Vehicles order by CreationDate", pubcnn, adOpenForwardOnly, adLockOptimistic
    Else
    rsTemp.Open "select * from tbl_Vehicles where Vehicle_Reg = '" & strsearch & "'", pubcnn, adOpenForwardOnly, adLockOptimistic
    Disp
    End If
FormatMe
 Button_setting frmAddVehicle, rsTemp, flag
End Sub

Private Sub frmEnable(Kweli As Boolean)
Me.txtDescription.Enabled = Kweli
Me.txtVehicleRegNo.Enabled = Kweli
Me.txtVehicleType = Kweli
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

If flag = 1 Then .AddNew
.Fields("Vehicle_Reg").Value = Me.txtVehicleRegNo
.Fields("Vehicle_Mileage").Value = Me.txtMileage.Text
.Fields("Vehicle_Type").Value = Me.txtVehicleType
.Fields("Type_Rate").Value = Me.TXTRATE
.Fields("Vehicle_Description").Value = Me.txtDescription.Text
If strDest <> "" Then
.Fields("Img_Path").Value = fso.GetFileName(strDest)
End If
.Fields("CreatedBy").Value = pubLoginName
.Update
End With
    flag = 3
 Button_setting frmAddVehicle, rsTemp, flag
    

End Function

Private Function Updatedb()

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
    Me.txtVehicleType = !Vehicle_Type
    End If
        If Not IsNull(rsTemp("Type_Rate")) Then
    Me.TXTRATE = !Type_Rate
    End If
    If Not IsNull(rsTemp("Img_Path")) Then
        strPic = App.path & "/Images/Cars/" & Trim(!Img_Path)
        If fso.FileExists(strPic) Then
        Me.Image1.Picture = LoadPicture(strPic)
        End If
        Else
        Me.Image1.Picture = LoadPicture(App.path & "/Images/noimg.gif")
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

Private Sub txtRate_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub

Private Sub txtVehicleRegNo_Change()
If flag = 1 Then
If mdlUsers.CheckDuplicates("tbl_Vehicles", "Vehicle_Reg", Me.txtVehicleRegNo.Text, Me.txtVehicleRegNo) = True Then
MsgBox "Registration Number already exists." + vbCr + "Enter a new Registration.", vbInformation, "New User"
End If
End If
txtVehicleRegNo.BackColor = &HC0FFFF
End Sub
Private Sub FormatMe()
If flag = 1 Then
Me.Caption = "Add Vehicle"
'Me.Label1.Caption = "Add vehicle"
Else
Me.Caption = "Edit Vehicle record"
'Me.Label1.Caption = "Edit vehicle record"
End If
End Sub

