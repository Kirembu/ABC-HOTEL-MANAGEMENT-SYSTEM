VERSION 5.00
Begin VB.Form frmRoomService 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCharge 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtGuestID 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox cboMeal 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblCost 
      Caption         =   "Cost of service"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblMeal 
      Caption         =   "Meal to be served in guest's room"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblGuestID 
      Caption         =   "GuestID"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmRoomService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
