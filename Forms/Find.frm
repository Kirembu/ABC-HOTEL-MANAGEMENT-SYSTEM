VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find."
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbEntity 
      Height          =   5250
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type Contents to Match Your Criteria"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   2880
      Picture         =   "Find.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Engine."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vKey As Integer

Private Sub cmbEntity_DblClick()
    cmbEntity_KeyPress 13
End Sub

Private Sub cmbEntity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13     'enter
            If Not cmbEntity.ListIndex = -1 Then
                vKey = cmbEntity.ItemData(cmbEntity.ListIndex)
                Unload Me
            End If
        
        Case 27     'Esc
            vKey = -1
            Unload Me
    End Select
End Sub

Public Function getKey(ByVal pTblName As String, ByVal pFldName As String, Optional pCriteria As String)
    vKey = -1
    Call fillCombo(cmbEntity, pTblName, pFldName, pCriteria)
    Me.Show vbModal
    
    getKey = vKey
End Function

