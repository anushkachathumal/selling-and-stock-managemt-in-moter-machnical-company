Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class DAL_frmMecilleniusrecipt
   
    Public Overloads Shared Function UpdateT24RecevingHeader(ByVal ncQryType As String, ByVal vcCode As String, ByVal nvcFieldList As String, ByVal ncWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_RecievingHeaderNew", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcCode", vcCode), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
        New SqlParameter("@vcWhereClause", ncWhereClause)), Guid)
    End Function
    Public Overloads Shared Function UpdateT53MiscellaneousIssue(ByVal ncQryType As String, ByVal vcCode As String, ByVal nvcFieldList As String, ByVal ncWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT53MiscellaneousIssue", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcCode", vcCode), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
        New SqlParameter("@vcWhereClause", ncWhereClause)), Guid)
    End Function
End Class
