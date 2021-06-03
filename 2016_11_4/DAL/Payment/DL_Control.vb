
Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class DL_Control
    Public Overloads Shared Function InsertTempDistributorct(ByVal ncQryType As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDistributorProcCtl", _
         New SqlParameter("@cQryType", ncQryType), _
        New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function InsertDistributorsAdvance(ByVal ncQryType As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDistributorAdvanceInfo", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
        New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function
    Public Overloads Shared Function InsertSalesReceving(ByVal ncQryType As String, ByVal nciRcvNo As String, ByVal vcItemNo As String, ByVal vcStartNo As String, ByVal vcEndNo As String, ByVal vcTxnDate As Date, ByVal vcFrom As String, ByVal vcTo As String, ByVal vcQty As Integer, ByVal vcTransCode As String, ByVal vcStatus As String, ByVal vcUser As String, ByVal vcTime As Date, ByVal vcWrkID As String, ByVal vcmDAdv As Double, ByVal vcTrfNo As Integer, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_SalesReceiving", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@iRcvNo", nciRcvNo), _
         New SqlParameter("@vcItemNo", vcItemNo), _
         New SqlParameter("@iStartNo", vcStartNo), _
         New SqlParameter("@iEndNo", vcEndNo), _
         New SqlParameter("@dTxnDate", vcTxnDate), _
         New SqlParameter("@vcFrom", vcFrom), _
         New SqlParameter("@vcTo", vcTo), _
         New SqlParameter("@iQty", vcQty), _
         New SqlParameter("@iTransCode", vcTransCode), _
         New SqlParameter("@iStatus", vcStatus), _
         New SqlParameter("@vcUser", vcUser), _
         New SqlParameter("@dTime", vcTime), _
        New SqlParameter("@cWrkID", vcWrkID), _
         New SqlParameter("@mDAdv", vcmDAdv), _
         New SqlParameter("@iTrfNo", vcTrfNo)), Guid)

        'New SqlParameter("@dTxnDate", vcTxnDate), _
        'New SqlParameter("@vcFrom", vcFrom), _
        'New SqlParameter("@vcTo", vcTo), _
        'New SqlParameter("@iTransCode", vcTransCode), _
        'New SqlParameter("@iStatus", vcStatus), _
        'New SqlParameter("@vcUser", vcUser), _
        'New SqlParameter("@dTime", vcTime), _
        'New SqlParameter("@cWrkID", vcWrkID), _
        ' New SqlParameter("@mDAdv", vcmDAdv), _
    End Function
    'Public Overloads Shared Function InsertT33StockTransfer(ByVal ncQryType As String, ByVal vcTrfNo As Integer, ByVal vcItemNo As String, ByVal vcTxnDate As String, ByVal vcStartNo As String, ByVal vcEndNo As String, ByVal vcFrom As String, ByVal vcTo As String, ByVal vcQty As Integer, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
    '    '//Execute the query and return the new Guid
    '    Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_SalesReceiving", _
    '     New SqlParameter("@cQryType", ncQryType), _
    '     'New SqlParameter("@iTrfNo", vcTrfNo), _
    '    'New SqlParameter("@vcItemNo", vcItemNo), _
    '    ' New SqlParameter("@dTxnDate", vcTxnDate), _
    '    'New SqlParameter("@iStartNo", vcStartNo), _
    '    'New SqlParameter("@iEndNo", vcEndNo), _
    '      New SqlParameter("@vcFrom", vcFrom), _
    '      New SqlParameter("@vcTo", vcTo), _
    '      New SqlParameter("@iQty", vcQty), _
    '     New SqlParameter("@iTrfNo", vcTrfNo)), Guid)

    '    'New SqlParameter("@dTxnDate", vcTxnDate), _
    '    'New SqlParameter("@vcFrom", vcFrom), _
    '    'New SqlParameter("@vcTo", vcTo), _
    '    'New SqlParameter("@iTransCode", vcTransCode), _
    '    'New SqlParameter("@iStatus", vcStatus), _
    '    'New SqlParameter("@vcUser", vcUser), _
    '    'New SqlParameter("@dTime", vcTime), _
    '    'New SqlParameter("@cWrkID", vcWrkID), _
    '    ' New SqlParameter("@mDAdv", vcmDAdv), _
    'End Function
    Public Overloads Shared Function InsertT60SalesProcess(ByVal ncQryType As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_T60SalesProcess", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
        New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function
End Class
