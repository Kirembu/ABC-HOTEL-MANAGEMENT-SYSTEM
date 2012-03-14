VERSION 5.00
Begin VB.Form frmAddRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rooms"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "frmRooms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10365
   Begin VB.CheckBox chkIncrement 
      Caption         =   "Auto Increase"
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
      Left            =   3000
      TabIndex        =   29
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame fmeGuest 
      Caption         =   "Current Guest Details"
      Height          =   2535
      Left            =   5400
      TabIndex        =   20
      Top             =   2880
      Width           =   4215
      Begin VB.TextBox txtGuestFullName 
         BackColor       =   &H8000000F&
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         Tag             =   "p"
         Text            =   "Guest Name"
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtGuestAge 
         BackColor       =   &H8000000F&
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
         Left            =   1200
         TabIndex        =   23
         Tag             =   "p"
         Text            =   "Age"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtDateOfBookin 
         BackColor       =   &H8000000F&
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
         Left            =   1200
         TabIndex        =   22
         Tag             =   "p"
         Text            =   "Date"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtCitizenship 
         BackColor       =   &H8000000F&
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
         Left            =   1200
         TabIndex        =   21
         Tag             =   "p"
         Text            =   "Citizenship"
         Top             =   1680
         Width           =   1815
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
         Left            =   120
         TabIndex        =   28
         Top             =   435
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
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of birth"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame fmeCatrgory 
      Caption         =   "Room Catrgory"
      Height          =   2535
      Left            =   1920
      TabIndex        =   15
      Top             =   2880
      Width           =   3015
      Begin VB.OptionButton optRoomCategory 
         Caption         =   "Double room self contained"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   19
         Top             =   1680
         Width           =   2415
      End
      Begin VB.OptionButton optRoomCategory 
         Caption         =   "Single room self contained"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton optRoomCategory 
         Caption         =   "Double room"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton optRoomCategory 
         Caption         =   "Single room"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      Caption         =   "Next"
      Height          =   615
      Left            =   1830
      Picture         =   "frmRooms.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      Appearance      =   0  'Flat
      Caption         =   "First"
      Height          =   615
      Left            =   360
      Picture         =   "frmRooms.frx":0776
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      Appearance      =   0  'Flat
      Caption         =   "Previous"
      Height          =   615
      Left            =   1095
      Picture         =   "frmRooms.frx":0EE0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      Appearance      =   0  'Flat
      Caption         =   "Last"
      Height          =   615
      Left            =   2565
      Picture         =   "frmRooms.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      Caption         =   "Search"
      Height          =   615
      Left            =   3300
      Picture         =   "frmRooms.frx":1DB4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtDescription 
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
      Height          =   405
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txtRate 
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
      Height          =   450
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtRoomNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R-"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
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
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Tag             =   "Room No"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create new rooms, manage existing ones and edit or delete old room records"
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000009&
      Caption         =   "Room management"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   360
      Width           =   4575
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
      Left            =   360
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblRates 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description/Note"
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
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblRoomNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Room no."
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
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "frmAddRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim rsRoom As ADODB.Recordset
Dim rsDisp As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim flag As Integer
Dim i As Integer
Dim res, pop As String

Private Sub btnExit_Click()
Unload Me
End Sub

Public Sub Form_Cancel()
flag = 3
If rsTemp.RecordCount Then
    rsTemp.MoveFirst
    Disp
    
Else
    Call Clear_All(Me)
End If
Button_setting frmAddRoom, rsTemp, flag

End Sub

Public Sub Form_Delete()
flag = 3
If rsTemp.EOF = False Then
res = MsgBox("Do you relly want to delete this guest record?", vbQuestion + vbYesNo, "Delete")
If res = vbYes Then
rsTemp.Requery
    rsTemp.Delete

Else
Exit Sub
End If
End If

 frmEnable (False)
If rsTemp.RecordCount Then
        rsTemp.MoveFirst
        Call Disp
    Else
        Call Clear_All(Me)
End If
End Sub
Private Sub refreshData()
    Set rsDisp = New ADODB.Recordset
    rsDisp.Open "select * from tbl_Guest order by Guest_ID", pubcnn, 1, 2

End Sub


Private Sub cmdEdit_Click()
flag = 2
Button_setting frmAddRoom, rsTemp, flag
 frmEnable (True)
Me.TXTRATE.SetFocus
End Sub

Private Sub chkIncrement_Click()
If Me.chkIncrement.Value = 1 Then
txtRoomNo.Locked = True
Else
txtRoomNo.Locked = False
End If
End Sub

Private Sub cmdFirst_Click()
rsTemp.MoveFirst
Call Disp
Button_setting frmAddRoom, rsTemp, flag
End Sub

Private Sub cmdLast_Click()
rsTemp.MoveLast
Disp
flag = 3
Button_setting frmAddRoom, rsTemp, flag

End Sub

Private Sub cmdNew_Click()
res = MsgBox("Add New Record?", vbOKCancel + vbQuestion, "Guests")
If res = vbOK Then
 frmEnable (True)
 frmClear
flag = 1
  Button_setting frmAddRoom, rsTemp, flag
End If
End Sub

Private Sub cmdNext_Click()

If rsTemp.RecordCount > 0 Then
  If rsTemp.EOF = True Then
 rsTemp.MoveLast
    Else
   rsTemp.MoveNext
  Disp
  flag = 3
    Button_setting frmAddRoom, rsTemp, flag
  End If
End If
End Sub

Private Sub cmdPrevious_Click()

   If rsTemp.BOF = False Then
    rsTemp.MovePrevious
    Disp
    flag = 3
     Button_setting frmAddRoom, rsTemp, flag
  End If
End Sub

Public Sub Form_Save()
Save_Room
flag = 3
Button_setting frmAddRoom, rsTemp, flag
End Sub

Public Sub form_new()
frmClear
rsTemp.AddNew
Me.txtDescription.Enabled = True
Me.txtRoomNo.Enabled = True
For i = 0 To 3
Me.optRoomCategory(i).Enabled = True
Next
Button_setting Me, rsTemp, 1
flag = 1
End Sub

Private Sub cmdSearch_Click()
flag = "4"
frmViewRooms.Show 1
If strsearch <> "" Then
rsTemp.MoveFirst
rsTemp.Find "Room_No =" + "'" + strsearch + "'"
Disp
End If
End Sub

Private Sub Form_Load()

    Set rsTemp = New ADODB.Recordset
    OpenDataBase
    rsTemp.Open "select * from tbl_Room order by Room_No", pubcnn, 1, 2
Call mdlData.OpenDataBase
    Call refreshData
    Call Disp
flag = 3
 Button_setting frmAddRoom, rsTemp, flag
Call frmEnable(False)
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
Dim mnupop As Menu
Set mnupop = frmMain.mnuCtrl
PopupMenu mnupop
End If
End Sub

Private Sub optRoomCategory_Click(Index As Integer)

Select Case Index
 Case 0
 TXTRATE = "1000"
 Case 1
 TXTRATE = "2000"
 Case 2
 TXTRATE = "5000"
 Case 3
 TXTRATE = "8000"
End Select
End Sub

Private Sub Disp()
On Error Resume Next
Call Clear_All(Me)
With rsTemp
    If .BOF = False And .EOF = False Then
    Me.txtRoomNo.Text = !Room_No
    Me.TXTRATE.Text = !Rates
    Me.txtDescription = !Description
    If Not IsNull(rsTemp("Type")) Then
    If !Type = "Single room" Then Me.optRoomCategory(0).Value = True
    If !Type = "Double room" Then Me.optRoomCategory(1).Value = True
    If !Type = "Single room self contained" Then Me.optRoomCategory(2).Value = True
    If !Type = "Double room self contained" Then Me.optRoomCategory(3).Value = True

    End If
    End If
        lblPosition = .AbsolutePosition & " Of " & .RecordCount
    End With
    
    refreshData
    With rsDisp
    If .BOF = False And .EOF = False Then
    Me.txtCitizenship = IIf(!Citizen, "Local", "Foreign")
    Me.txtGuestFullName = !Title + " " + !First_Name + " " + !Second_Name + " " + !Other_Name
    Me.txtDateOfBookin = !BookinDate
    Me.txtGuestAge = !Age
    Else
        Me.txtCitizenship = "<None>"
    Me.txtGuestFullName = "<None>"
    Me.txtDateOfBookin = "<None>"
    Me.txtGuestAge = "<None>"
    End If
    End With

        End Sub
Private Function Save_Room()
If mdlFunctions.Check(frmAddRoom) = False Then
MsgBox "Please fill the highlighted regions", vbInformation, App.Title
Exit Function
Else
pubsql = "select * from tbl_Room"
    If mdlData.DataBaseToForm("select * from tbl_Room") = True Then
    For i = 0 To 3
    If Me.optRoomCategory(i).Value = True Then pubBoard = Me.optRoomCategory(i).Caption
    Next i
    With rsTemp
    .Fields("Room_No").Value = IIf(Me.chkIncrement.Value = 1, mdlUsers.GetNewID("tbl_Room", "Room_No", "R"), txtRoomNo.Text)
    .Fields("Rates").Value = TXTRATE.Text
    .Fields("Description").Value = Me.txtDescription.Text
    .Fields("Type").Value = pubBoard
    .Fields("Status").Value = "0"
    .Update
    End With
    End If
    mdldb.Clear_All frmAddRoom
    Disp
End If

End Function
Private Function frmEnable(Kweli As Boolean)
Me.txtDescription.Enabled = Kweli
Me.txtRoomNo.Enabled = Kweli
For i = 0 To 3
Me.optRoomCategory(i).Enabled = Kweli
Next i
End Function
Private Function frmClear()

Clear_All Me
Me.optRoomCategory(0).Value = True
End Function

Private Sub txtRoomNo_LostFocus()
mdldb.CheckAdd_Primary Me, "tbl_Room", txtRoomNo, "Room_No"
End Sub
Public Sub Form_Edit()
flag = 2
frmEnable True
 Button_setting frmAddRoom, rsTemp, flag
End Sub
