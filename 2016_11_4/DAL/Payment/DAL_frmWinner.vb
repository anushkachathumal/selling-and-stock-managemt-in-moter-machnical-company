Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class DAL_frmWinner

    Public Overloads Shared Function InsertDeatails(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT21PrizeDetail", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function InsertDeatailsT58(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT58PrizeDetailNonwinning", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function

    ''========================================================================
    Public Overloads Shared Function InsertHEDER(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT20PaymentHeader", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function InsertWinner(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT35Winners", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function InsertCreditT37(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT37CreditVoucher", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function InsertCreditT59(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT59PrizeDetail", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function

    Public Overloads Shared Function UpdateM03Agent(ByVal ncQryType As String, ByVal vcCode As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetM03Agent", _
            New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
        New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function
    Public Overloads Shared Function InserttmpDestributors(ByVal ncQryType As String, ByVal nvcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDistributorProcCtl", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcCode", nvcCode)), Guid)
    End Function

    Public Overloads Shared Function UpdateT37tbl(ByVal ncQryType As String, ByVal nvcFieldList As String, ByVal nvcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT37CreditVoucher", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldListe", nvcFieldList), _
         New SqlParameter("@vcCode", nvcCode)), Guid)
    End Function

    Public Overloads Shared Function UpdateT45(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDistributorAdvanceInfo", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function

    Public Overloads Shared Function InsertDeposit(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT04BCDR", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function InsertDISTRIBUTERSINFO(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDistributorAdvanceInfo", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function InsertDISTRIBUTERSINFO(ByVal ncQryType As String, ByVal vcCode As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid @cQryType,@vcFieldList, @vcWhereClause
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetDistributorAdvanceInfo", _
         New SqlParameter("@cQryType", ncQryType), _
        New SqlParameter("@vcCode", vcCode), _
        New SqlParameter("@vcFieldList", nvcFieldList), _
         New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function

    Public Overloads Shared Function InsertDistributor(ByVal ncQryType As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_M09DistributorAdvance", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
         New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function
    Public Overloads Shared Function InsertT44TABLE(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_T44AgentAdvance", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function

    Public Overloads Shared Function InsertT24RecevingHeader(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_RecievingHeaderNew", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function InsertT24RecevingHeader(ByVal ncQryType As String, ByVal vcCode As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_RecievingHeaderNew", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcCode", vcCode), _
        New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function T25SalesRecieving(ByVal ncQryType As String, ByVal nciRcvNo As String, ByVal ncvcItemNo As String, ByVal nciStartNo As String, ByVal nciEndNo As String, ByVal nciQty As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_ManualSalesReceiving", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@iRcvNo", nciRcvNo), _
         New SqlParameter("@vcItemNo", ncvcItemNo), _
         New SqlParameter("@iStartNo", nciStartNo), _
         New SqlParameter("@iEndNo", nciEndNo), _
         New SqlParameter("@iQty", nciQty)), Guid)
    End Function
    Public Overloads Shared Function T12FundRecieving(ByVal ncQryType As String, ByVal nciRcvNo As String, ByVal ncvcMethod As String, ByVal ncBCDR As String, ByVal ncAmount As String, ByVal ncSeqNo As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_FundRecieving", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@iRcvNo", nciRcvNo), _
         New SqlParameter("@vcMethod", ncvcMethod), _
         New SqlParameter("@vcBCDR", ncBCDR), _
         New SqlParameter("@mAmount", ncAmount), _
         New SqlParameter("@iSeqNo", ncSeqNo)), Guid)
    End Function

    Public Overloads Shared Function InsertCreditT37Table(ByVal ncQryType As String, ByVal nvcFieldList As String, ByVal vcCode As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT37CreditVoucher", _
         New SqlParameter("@cQryType", ncQryType), _
          New SqlParameter("@vcFieldList", nvcFieldList), _
         New SqlParameter("@vcCode", vcCode)), Guid)
    End Function

    Public Overloads Shared Function InsertT44TABLEtest(ByVal ncQryType As String, ByVal vcCode As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_T44AgentAdvance", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcCode", vcCode), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
        New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function

    Public Overloads Shared Function Insert04BCDR(ByVal ncQryType As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT04BCDR", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
        New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function
    Public Overloads Shared Function InserT17LocationStock(ByVal ncQryType As String, ByVal nvcFieldList As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetLocationStockInfo", _
         New SqlParameter("@cQryType", ncQryType), _
        New SqlParameter("@vcFieldList", nvcFieldList)), Guid)
    End Function
    Public Overloads Shared Function UpdateT17LocationStock(ByVal ncQryType As String, ByVal vcCode As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        '//Execute the query and return the new Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetLocationStockInfo", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcCode", vcCode), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
        New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function
    Public Overloads Shared Function UpdateT37Table(ByVal ncQryType As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_GetSetT37CreditVoucher", _
         New SqlParameter("@cQryType", ncQryType), _
          New SqlParameter("@vcFieldList", nvcFieldList), _
         New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function
    Public Overloads Shared Function M09DistributorTable(ByVal ncQryType As String, ByVal nvcCode As String, ByVal nvcFieldList As String, ByVal vcWhereClause As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "up_M09DistributorAdvance", _
         New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@vcCode", ncQryType), _
         New SqlParameter("@vcFieldList", nvcFieldList), _
         New SqlParameter("@vcWhereClause", vcWhereClause)), Guid)
    End Function

    Public Overloads Shared Function WinnerFileUpLoad1Dgt(ByVal ncQryType As String, ByVal nItem1 As String, ByVal nT36Date As Date, ByVal ndBonusPrize As Double, ByVal ndJackpotPrize As Double, ByVal VPrize As Double, ByVal nDrawNo As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "sp_WinnerUpload1Dgt", _
        New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@Item1", nItem1), _
         New SqlParameter("@T36Date", nT36Date), _
         New SqlParameter("@dBonusPrize", ndBonusPrize), _
         New SqlParameter("@dJackpotPrize", ndJackpotPrize), _
         New SqlParameter("@VPrize", VPrize), _
         New SqlParameter("@DrawNo", nDrawNo)), Guid)

    End Function
    Public Overloads Shared Function WinnerFileUpLoad2Dgt(ByVal ncQryType As String, ByVal nItem1 As String, ByVal nT36Date As Date, ByVal ndBonusPrize As Double, ByVal ndJackpotPrize As Double, ByVal VPrize As Double, ByVal nDrawNo As String, ByRef connection As SqlConnection, ByRef transaction As SqlTransaction) As Guid
        Return CType(DBEngin.ExecuteScalar(connection, transaction, "sp_WinnerUpload2Dgt", _
        New SqlParameter("@cQryType", ncQryType), _
         New SqlParameter("@Item1", nItem1), _
         New SqlParameter("@T36Date", nT36Date), _
         New SqlParameter("@dBonusPrize", ndBonusPrize), _
         New SqlParameter("@dJackpotPrize", ndJackpotPrize), _
         New SqlParameter("@VPrize", VPrize), _
         New SqlParameter("@DrawNo", nDrawNo)), Guid)

    End Function
End Class
