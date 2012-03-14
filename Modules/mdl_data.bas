Attribute VB_Name = "mdlData"
Option Explicit
Public pubRst As ADODB.Recordset
Public pubcnn As New ADODB.Connection
Public pubcmd As ADODB.Command
Public pubsql As String
Public pubtempsql As String
'**********************************************
'Variables to help store information abt currently loggedin user
'*******************************************
Public pubUserID As String
Public pubUserName As String
Public pubGuestID As String
Public bleDataChanged As Boolean

Public Function OpenDataBase() As Boolean
 OpenDataBase = False
  If pubcnn.State = adStateClosed Then
   pubcnn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\DataBase\ABCdb.mdb"
   pubcnn.Open

  End If
  OpenDataBase = True
End Function

Public Function DataBaseToForm(ByVal pSQL As String) As Boolean
DataBaseToForm = False
 If OpenDataBase = True Then
  Set pubRst = Nothing
  Set pubRst = New ADODB.Recordset
   pubRst.Open pSQL, pubcnn, adOpenKeyset, adLockOptimistic
  DataBaseToForm = True
 End If
End Function

Public Function dbForm(ByVal mYsQL As String, recset As Recordset, Frm As Form)
     Set recset = Nothing
    Set recset = New ADODB.Recordset
    recset.Open mYsQL, pubcnn, adOpenKeyset, adLockOptimistic
    
End Function
'this sub will reset data environment
Public Sub initDtEnv()
Set dataenv = New dataenv
       dataenv.DataCnn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\DataBase\ABCdb.mdb"
End Sub
Public Function getBillDets(ByVal pBillNo As Long) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from tbl_Payment where Receipt_No=" & pBillNo, pubcnn
    
    Set getBillDets = rs
End Function
