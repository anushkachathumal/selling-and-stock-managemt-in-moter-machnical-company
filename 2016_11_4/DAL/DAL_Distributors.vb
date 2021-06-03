Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class DAL_Distributors
    Public Overloads Shared Function Insert(ByVal cM09DistributorNo As String, ByVal cM09CrLimi As Decimal, ByVal cM09LocationNo As String, ByVal cM09Status As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "Ins_Dis", _
         New SqlParameter("@M09DistributorNo", cM09DistributorNo), _
         New SqlParameter("@M09CrLimit", cM09CrLimi), _
         New SqlParameter("@M09LocationNo", cM09LocationNo), _
         New SqlParameter("@M09Status", cM09Status)), Guid)
    End Function

    Public Overloads Shared Function up_GetSettmpGriege_Stock(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetUse_Griege_Qty", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetDelivary_Planning(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Planning", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function


    Public Overloads Shared Function up_GetSetBlock_KntMC(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Planning", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetYarn_Bookingtmp(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Quatation", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function
    Public Overloads Shared Function up_GetSetBlock_Projection(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Quatation", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetT21Pack_Allocation(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT21Pack_Allocation", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSettmp_ProjectAllocation(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Quatation", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetYarn_Request(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Planning", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
                   New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSettmp_Knt_Mc(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSettmp_Knt_Mc", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetConsume_Grige(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetUse_Griege_Qty", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function
    Public Overloads Shared Function up_GetSetCAPACITY(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Quatation", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function


    Public Overloads Shared Function up_GetSetProjection(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Planning", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetLAB_DIP_DateConf(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Planning", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetBlock_KntPlanning_Boad(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Planning", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetBlock_DyeMC(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDyeMC_Process", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetYarn_DyeMCPln(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDyeMC_Process", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetYarn_StockCode(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDelivary_Planning", _
         New SqlParameter("@cQryType", qcType), _
          New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause1", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetT21PrizeDetail(ByVal qcType As String, ByVal vcFieldlist As Decimal, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT21PrizeDetail", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function


    Public Overloads Shared Function up_GetSetMEETING(ByVal qcType As String, ByVal vcCode As String, ByVal vcWharer As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetMEETING", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcCode", vcCode), _
                  New SqlParameter("@vcWhereClause", vcWharer)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetM09ZPL_ORDER(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcWharer As String, ByVal vcLine As String, ByVal vcDate As String, ByVal vcCode As String, ByVal vcBatch As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetM09ZPL_ORDER", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
          New SqlParameter("@vcWhereClause", vcWharer), _
          New SqlParameter("@vcLine", vcLine), _
             New SqlParameter("@vcDate", vcDate), _
              New SqlParameter("@vcBATCH", vcBatch), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetup_GetSetM19Time_Taken(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcWharer As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetM19Time_Taken", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcWhereClause", vcWharer)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetM04Lot(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetM04Lot", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetM04Lottmp(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetM04Lottmp", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetPRODUCTION_FIGURES(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetPRODUCTION_FIGURES", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetDAILY_UTILITY(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDAILY_UTILITY", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetT17(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT17LocationStock", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSettmpGoodReceived(ByVal qcType As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_tmpGoodReceived", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSettmpGoodReceived(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_tmpGoodReceived", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function InsertT53MiscellaneousIssue(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT53MiscellaneousIssue", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetM08Stock(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcWharer As String, ByVal vcLine As String, ByVal vcDate As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetM08Stock", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
          New SqlParameter("@vcWhereClause", vcWharer), _
          New SqlParameter("@vcLine", vcLine), _
             New SqlParameter("@vcDate", vcDate), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetM07TobeDelivered(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcWharer As String, ByVal vcLine As String, ByVal vcDate As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetM07TobeDelivered", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
          New SqlParameter("@vcWhereClause", vcWharer), _
          New SqlParameter("@vcLine", vcLine), _
             New SqlParameter("@vcDate", vcDate), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function up_GetSetM06Delivary_Qty(ByVal qcType As String, ByVal vcFieldlist As String, ByVal vcWharer As String, ByVal vcLine As String, ByVal vcDate As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid

        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetM06Delivary_Qty", _
         New SqlParameter("@cQryType", qcType), _
         New SqlParameter("@vcFieldList", vcFieldlist), _
          New SqlParameter("@vcWhereClause", vcWharer), _
          New SqlParameter("@vcLine", vcLine), _
             New SqlParameter("@vcDate", vcDate), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function
End Class
