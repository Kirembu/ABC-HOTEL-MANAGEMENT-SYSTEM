Attribute VB_Name = "mdlDBDepartment"

Public Function GetNewDepartmentID(ByRef sNewDepartmentID As String) As String
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Integer
    
    'default
    GetNewDepartmentID = Failed
    
    sSQL = "SELECT 'D-' & String$(2-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tbl_Department;"
        
        sNewDepartmentID = vRS.Fields("NewID").Value
        
        While DepartmentExistByID(sNewDepartmentID) = Success
            NewDNumber = Val(Right(sNewDepartmentID, 2)) + 1
            sNewDepartmentID = "D-" & String(2 - Len(NewDNumber), "0") & NewDNumber
        Wend
        
        GetNewDepartmentID = Success
    
    Else
    
        GetNewDepartmentID = Failed
    End If
    
    
    Set vRS = Nothing

End Function


