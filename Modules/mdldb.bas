Attribute VB_Name = "mdldb"
Sub Main()
On Error GoTo errhandler
Set cnn = New ADODB.Connection
cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
cnn.Open "Data Source =" & App.path & "\ABCdb.mdb"

errhandler:

If Err Then
    MsgBox "Error in Sub_Main" & _
    vbCrLf & Err.Number & " - " & Err.Description
End If

End Sub

Public Sub Clear_All(Frm As Form)
Dim mycontrol As Control
For Each mycontrol In Frm.Controls
    If TypeOf mycontrol Is TextBox Then
        mycontrol.Text = ""
'    ElseIf TypeOf mycontrol Is ComboBox Then
'            mycontrol.Text = ""
    ElseIf TypeOf mycontrol Is DTPicker Then
            mycontrol.Value = Date
'    ElseIf TypeOf mycontrol Is OptionButton Then
'            mycontrol.Value = False
        End If
    
Next
End Sub

Public Sub CheckAdd_Primary(Frm As Form, myRec As String, txt As TextBox, myField As String)
On Error GoTo exx:
Dim tempRs As New ADODB.Recordset
Set tempRs = New ADODB.Recordset
tempRs.Open "select distinct " & myField & " from " & myRec & " where " & myField & " =" & txt.Text, pubcnn, adOpenDynamic, adLockOptimistic
    If Not tempRs.BOF = False And tempRs.BOF = False Then
        b = MsgBox("Duplicate Value Not Allowed ...")
        If b = vbOK Then
        txt.SetFocus
        End If
        End If
Set tempRs = Nothing
exx:
Exit Sub
End Sub

Public Sub checkEdit_Primary(Frm As Form, mrs As Recordset, txt As TextBox, fd As String, flag As Integer)
Dim tempRs As Recordset

    Set tempRs = mrs.Clone
    tempRs.MoveFirst
    Do Until tempRs.EOF
        If Val(txt.Text) = tempRs(fd) Then
        b = MsgBox("Duplicate Value Not Allowed ...")
        If b = vbOK Then
        txt.SetFocus
        End If
       
        End If
        tempRs.MoveNext
    Loop
    tempRs.Close
Set tempRs = Nothing

End Sub

Public Sub Only_Numbers(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
Exit Sub
ElseIf KeyAscii = vbKeyClear Then Exit Sub
ElseIf KeyAscii = vbKeyReturn Then Exit Sub
ElseIf KeyAscii = vbKeyShift Then Exit Sub
ElseIf KeyAscii = vbKeyEnd Then Exit Sub
ElseIf KeyAscii = vbKeyLeft Then Exit Sub
ElseIf KeyAscii = vbKeyRight Then Exit Sub
ElseIf KeyAscii = vbKeyDelete Then Exit Sub
ElseIf KeyAscii = vbKeyNumlock Then Exit Sub
End If
If KeyAscii >= 48 And KeyAscii <= 57 Then
If KeyAscii = 46 Then
KeyAscii = 0
End If
Else
KeyAscii = 0
End If
End Sub

Public Sub Only_Text(Obj As Object, KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
Exit Sub
ElseIf KeyAscii = vbKeyClear Then Exit Sub
ElseIf KeyAscii = vbKeyReturn Then Exit Sub
ElseIf KeyAscii = vbKeyShift Then Exit Sub
ElseIf KeyAscii = vbKeyEnd Then Exit Sub
ElseIf KeyAscii = vbKeyLeft Then Exit Sub
ElseIf KeyAscii = vbKeyRight Then Exit Sub
ElseIf KeyAscii = vbKeyDelete Then Exit Sub
ElseIf KeyAscii = vbKeyNumlock Then Exit Sub
End If
Select Case KeyAscii
 Case LCase(vbKeyA) To LCase(vbKeyZ)
 Case Else
KeyAscii = 0
End Select

End Sub

