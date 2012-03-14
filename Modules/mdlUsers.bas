Attribute VB_Name = "mdlUsers"

Public Function GetNewID(strTable As String, strFieldName As String, strLetter As String) As String
Dim rsSource As ADODB.Recordset
Set rsSource = New ADODB.Recordset
Dim strSQL As String
Dim NewID As String
strSQL = "select * from " & strTable & " order by " & strFieldName
rsSource.Open strSQL, pubcnn, adOpenDynamic, adLockBatchOptimistic
If rsSource.EOF = False And rsSource.BOF = False Then
   rsSource.MoveLast
   NewID = Val(rsSource.Fields(strFieldName).Value) + 1
    GetNewID = NewID
    Else
    GetNewID = 1
End If

End Function

Public Function CheckDuplicates(strTable As String, strFieldName As String, txtValue As String, TxtBox As TextBox) As Boolean
Dim rsSource As New ADODB.Recordset
Dim chkSQL As String
chkSQL = "SELECT * FROM " & strTable & " WHERE " & strFieldName & " = '" & Trim(txtValue) & "'"
rsSource.Open chkSQL, pubcnn, adOpenDynamic, adLockOptimistic

    If (rsSource.EOF And rsSource.BOF) Then
        CheckDuplicates = False
If Not IsNull(TxtBox) = True Then
TxtBox.BackColor = vbWhite
End If
    Else
        CheckDuplicates = True
If Not IsNull(TxtBox) = True Then
TxtBox.BackColor = vbYellow
'TxtBox.SetFocus
        SendKeys "{Home}+{End}"
End If
    End If

End Function
