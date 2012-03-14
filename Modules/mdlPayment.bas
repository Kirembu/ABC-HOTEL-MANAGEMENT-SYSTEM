Attribute VB_Name = "mdlPayment"
Public Function recTransact(GuestID As String, Amount As Long, PaymentMode As String, AmountPaid As Long, ChequeNo As String, Credit As Boolean, Details As String)

'On Error Resume Next
    If mdlData.DataBaseToForm("select * from tbl_Payment order by Receipt_No") = True Then
    pubTax = 0
    If pubRst.EOF = False And pubRst.BOF = False Then
    pubRst.MoveLast
     End If
     With pubRst.Fields
    pubRst.AddNew
   !Receipt_No = mdlUsers.GetNewID("tbl_Payment", "Receipt_No", "P")
    !Amount = Amount
    !Amount_Paid = AmountPaid
    !Paid = Not pubCredit
    !Payment_Mode = IIf(PaymentMode <> "", PaymentMode, "N/A")
    !Cheque_No = IIf(ChequeNo <> "", ChequeNo, "N/A")
    !Details = Details
    If GuestID <> "" Then
    !Guest_ID = GuestID
    Else
    !Guest_ID = "0"
    End If
    If pubLoginName <> "" Then
    !LoginName = pubLoginName
    End If
End With
pubRst.Update
flag = 1
End If
'Else
End Function
