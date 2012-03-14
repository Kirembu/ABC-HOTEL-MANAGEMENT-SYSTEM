Attribute VB_Name = "mdlFunctions"
Public intAge As String
Public Age As String
Public rs As New ADODB.Recordset
Dim cnn As ADODB.Connection
Option Explicit
Public Enum AccessMode
AddMode
EditMode
ShowMode
End Enum
Public pubmissing As String

Public Sub Fill_Area(strFieldName As String, strTableName As String, cmbBox As ComboBox)

'Fills a combo box with whatever you like

Dim area1 As ADODB.Recordset
Set area1 = New ADODB.Recordset
With area1
    .ActiveConnection = pubcnn
    .Source = "Select Distinct " & strFieldName & " From " & strTableName & " Where " & strFieldName & " is not null "
    .Open
End With

cmbBox.Clear
Do Until area1.EOF
    cmbBox.AddItem area1(0)
    area1.MoveNext
Loop
End Sub

Public Function Check(Frm As Form) As Boolean

Dim myctrl As Control
Check = True
For Each myctrl In Frm.Controls
    If TypeOf myctrl Is TextBox Then
        If myctrl.Text = "" And myctrl.Tag = "" Then
        myctrl.BackColor = vbYellow
        Check = False
        Else
        myctrl.BackColor = vbWhite
    End If
End If
Next
For Each myctrl In Frm.Controls
    If TypeOf myctrl Is ComboBox Then
        If myctrl.Text = "" And myctrl.Tag = "" Then
        myctrl.BackColor = vbYellow
        Check = False
        ElseIf myctrl.Text <> "" Then
        myctrl.BackColor = vbWhite
        End If
 
    End If
Next
End Function
Public Function EnableInput(Frm As Form, Kweli As Boolean) As Boolean
Dim myctrl As Control
For Each myctrl In Frm.Controls
    If TypeOf myctrl Is TextBox Then
         If myctrl.Tag = "" Then
         myctrl.Locked = Kweli
    End If
    End If

Next
End Function
Public Sub Button_setting(Frm As Form, RecSource As ADODB.Recordset, flag As Integer)
On Error Resume Next
Dim rc As Integer
Dim ap As Integer
rc = RecSource.RecordCount
If rc Then
ap = RecSource.AbsolutePosition
End If
With frmMain

    End With
       With Frm
       If rc = 0 Then

        .cmdFirst.Enabled = False
        .cmdPrevious.Enabled = False
        .cmdNext.Enabled = False
        .cmdLast.Enabled = False
        .cmdDelete.Enabled = False
        .cmdEdit.Enabled = False
    ElseIf rc = 1 Then
        .cmdFirst.Enabled = False
        .cmdPrevious.Enabled = False
        .cmdNext.Enabled = False
        .cmdLast.Enabled = False
    ElseIf ap = 1 Then
        .cmdFirst.Enabled = False
        .cmdPrevious.Enabled = False
        .cmdNext.Enabled = True
        .cmdLast.Enabled = True
    ElseIf ap = rc Then
        .cmdFirst.Enabled = True
        .cmdPrevious.Enabled = True
        .cmdNext.Enabled = False
        .cmdLast.Enabled = False
    Else
        .cmdFirst.Enabled = True
        .cmdPrevious.Enabled = True
        .cmdNext.Enabled = True
        .cmdLast.Enabled = True
            .cmdBrowse.Enabled = True
    .cmdClear.Enabled = True
    End If
End With


End Sub
Public Sub Fill_Area2(strFieldName As String, strTableName As String, lvw As ListView)
Dim area1 As ADODB.Recordset
Dim x, ico As Integer
Set area1 = New ADODB.Recordset
With area1
    .ActiveConnection = pubcnn
    .Source = "Select * From " & strTableName & " Where " & strFieldName & " is not null "
    .Open
End With

lvw.ListItems.Clear
x = 1
Do Until area1.EOF
If area1.Fields("UserType").Value = "administrator" Then
ico = 1
Else
ico = 2
End If
    lvw.ListItems.Add x, , area1(strFieldName), ico
    area1.MoveNext
Loop
End Sub

Public Sub fillCombo(cmbName As ComboBox, tblName As String, fldName As String, Optional criteria As String)
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from " & tblName & criteria & " order by " & fldName, pubcnn, adOpenForwardOnly, adLockOptimistic
    
    cmbName.Clear
    
    Do Until rs.EOF
        cmbName.AddItem rs.Fields(fldName)
        cmbName.ItemData(cmbName.NewIndex) = rs.Fields(3)
        rs.MoveNext
    Loop
    rs.Close
End Sub
Public Function DelImg(strImg As String)
Dim fso As New FileSystemObject
            If fso.FileExists(strImg) Then
                SetAttr strImg, vbNormal
            fso.DeleteFile strImg, True
            End If
End Function
