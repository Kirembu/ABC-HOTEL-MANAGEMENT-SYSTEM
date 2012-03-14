VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmServicesAndFacilities 
   Caption         =   "Services and Facilities"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   Icon            =   "frmSevicesAndFacilities.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   8790
   Begin VB.TextBox txtServiceNo 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1095
      Width           =   2055
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   1680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtCapacity"
      BuddyDispid     =   196610
      OrigLeft        =   2880
      OrigTop         =   2880
      OrigRight       =   3135
      OrigBottom      =   3255
      Max             =   1000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtCapacity 
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   1680
      Width           =   705
   End
   Begin VB.TextBox txtRate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """KSH""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtServiceName 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   6570
      TabIndex        =   8
      Top             =   6240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Left            =   5040
      TabIndex        =   9
      Top             =   6240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Exit"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin MSComctlLib.ListView lvwList 
      Height          =   3615
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Service No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Capacity"
         Object.Width           =   2540
      EndProperty
      Picture         =   "frmSevicesAndFacilities.frx":000C
   End
   Begin VB.Label Label3 
      Caption         =   "No."
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   " Name"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Rate"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Capacity"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Services and Facilities"
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
      TabIndex        =   0
      Top             =   180
      Width           =   3165
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmSevicesAndFacilities.frx":5F4A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "frmServicesAndFacilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsServ As New ADODB.Recordset
Dim pubRst As New ADODB.Recordset
Dim flag As Integer
Dim i As Integer
Dim res, pop  As String

Private Sub btnExit_Click()
Unload Me
End Sub


Public Sub Form_Delete()
flag = 3
If rsServ.RecordCount Then
res = MsgBox("Do you relly want to delete this guest record?", vbQuestion + vbYesNo, "Delete")
If res = vbYes Then
    rsServ.Delete
Else
Exit Sub
End If
End If

Call frmEnable(False)
If rsServ.RecordCount Then
        rsServ.MoveFirst
        Call Disp("select * from tbl_Services order by Service_No")
    Else
        Call Clear_All(Me)
End If
  Button_setting frmServicesAndFacilities, rsServ, flag

End Sub


Public Sub Form_Save()
flag = 2
Call Save

End Sub

Private Sub cmdSearch_Click()

End Sub


Public Sub Form_Cancel()
On Error Resume Next
flag = 3
If rsServ.RecordCount Then
    rsServ.MoveFirst
    Call Disp("select * from tbl_Services order by Service_No")
Else
    Call Clear_All(Me)
End If
     Button_setting frmServicesAndFacilities, rsServ, flag

End Sub


Public Sub Form_Edit()
flag = 2
    Button_setting frmServicesAndFacilities, rsServ, flag
Call frmEnable(True)
Me.txtServiceName.SetFocus
End Sub


Public Sub form_new()
res = MsgBox("Add New Record?", vbOKCancel + vbQuestion, App.Title)
If res = vbOK Then
Call frmEnable(True)
Call frmClear
flag = 1
 Button_setting frmServicesAndFacilities, rsServ, flag
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Call Save
End Sub

Private Sub Form_Load()
Fill_List
mdlFX.MakeGradient Me, 1
    Set rsServ = New ADODB.Recordset
    OpenDataBase
    rsServ.Open "select * from tbl_Services order by Service_No", pubcnn, adOpenForwardOnly, adLockOptimistic
txtServiceNo = mdlUsers.GetNewID("tbl_Services", "Service_No", "S")

End Sub

Private Sub frmEnable(Kweli As Boolean)
Me.txtServiceName.Enabled = Kweli
End Sub


Private Function Save()
Dim sid As Integer
If mdlFunctions.Check(Me) = False Then
MsgBox "Fill in all the required fields"
Exit Function
Else

With rsServ

If flag <> 2 Then
.AddNew
End If
.Fields("Service_No") = Me.txtServiceNo.Text
.Fields("Name").Value = Me.txtServiceName
.Fields("Rate").Value = Me.TXTRATE
.Fields("Capacity").Value = Me.txtCapacity
.Update

End With
 Button_setting frmServicesAndFacilities, rsServ, 3
Fill_List
End If

End Function

Private Sub refreshData()
    Set rsDisp = New ADODB.Recordset
    rsDisp.Open "select * from tbl_Sevices order by Service_No", pubcnn, 1, 2
  Button_setting frmServicesAndFacilities, rsServ, flag
End Sub
Private Sub Disp(pSQL As String)

Call Clear_All(Me)
With rsServ
.Close
.Open (pSQL)
If .BOF = False And .EOF = False Then

    If Not IsNull(rsServ("Name")) Then
    Me.txtServiceName = !Name
    End If
    If Not IsNull(rsServ("Rate")) Then
    Me.TXTRATE = !Rate
    End If
    If Not IsNull(rsServ("Capacity")) Then
    Me.txtCapacity = !capacity
    End If
    If Not IsNull(rsServ("Service_No")) Then
    Me.txtServiceNo = !Service_No
    End If
End If
    End With
End Sub

Private Sub frmClear()
Me.txtCapacity = ""
Me.TXTRATE = ""
Me.txtServiceName = ""
Me.txtServiceNo = ""
End Sub

Private Sub txtMileage_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
Dim mnupop As Menu
Set mnupop = frmMain.mnuCtrl
PopupMenu mnupop
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.lvwList.Width = Me.ScaleWidth - 75
Me.lvwList.Height = Me.ScaleHeight - 225

Me.cmdCancel.Top = Me.ScaleHeight - Me.cmdCancel.Height - Me.cmdCancel.Height
Me.cmdSave.Top = Me.cmdCancel.Top
Me.cmdCancel.Left = Me.ScaleWidth - Me.cmdCancel.Width * 2 - Me.cmdCancel.Width
Me.cmdSave.Left = Me.cmdCancel.Left + Me.cmdCancel.Width + 10
End Sub

Private Sub lvwList_Click()
flag = 2
Disp ("select * from tbl_Services where Service_No = " & Me.lvwList.ListItems.Item(lvwList.SelectedItem.Index).SubItems(1))

End Sub

Private Sub lvwList_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
Dim mnupop As Menu
Set mnupop = frmMain.mnuCtrl
PopupMenu mnupop
End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
mdldb.Only_Numbers KeyAscii
End Sub
Private Function Fill_List()
On Error Resume Next
OpenDataBase
Me.lvwList.ListItems.Clear
Set pubRst = Nothing
If mdlData.DataBaseToForm("select * from tbl_Services order by Service_No") = True Then

pubRst.Open "select * from tbl_Services order by Service_No", pubcnn, adOpenDynamic, adLockOptimistic
Do While Not pubRst.EOF
    x = x + 1
lvwList.ListItems.Add , , x
With lvwList.ListItems.Item(x)
.SubItems(1) = pubRst.Fields("Service_No")
.SubItems(2) = pubRst.Fields("Name")
.SubItems(3) = pubRst.Fields("Rate")
.SubItems(4) = pubRst.Fields("Capacity")
pubRst.MoveNext
End With
  Loop
End If
End Function
