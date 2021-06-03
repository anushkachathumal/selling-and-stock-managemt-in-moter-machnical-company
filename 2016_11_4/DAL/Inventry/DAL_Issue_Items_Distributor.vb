Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class DAL_Issue_Items_Distributor


    'This function is use to get, set and unset Distributor's Process control
    'data entry routines must chk for this process lock to to their tasks
    Public Shared Function GetSetDistributorInProcess(ByVal sType As String, ByVal sCode As String) As Boolean
        Dim dsUser As DataSet
        Dim con = New SqlConnection()
        GetSetDistributorInProcess = False
        '=======================================================================
        Try
            con = DBEngin.GetConnection()
            Select Case sType
                Case "LST"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetDistributorProcCtl", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    If common.isValidDataset(dsUser) Then
                        GetSetDistributorInProcess = True
                    End If
                Case "ADD"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetDistributorProcCtl", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    If common.isValidDataset(dsUser) Then
                        GetSetDistributorInProcess = False
                    Else
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetDistributorProcCtl", New SqlParameter("@cQryType", "ADD"), New SqlParameter("@vcCode", sCode))
                        GetSetDistributorInProcess = True
                    End If
                Case "DEL"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetDistributorProcCtl", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", sCode))
                    If common.isValidDataset(dsUser) = False Then
                        GetSetDistributorInProcess = False
                    Else
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "dbo.up_GetSetDistributorProcCtl", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcCode", sCode))
                        GetSetDistributorInProcess = True
                    End If
            End Select
            '===================================================================
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
        '=========================================================================

    End Function

    Public Function MakeDataTable1() As DataTable

        ' declare a DataTable to contain the program generated data
        Dim dataTable As New DataTable("StkItem")
        ' create and add a Code column
        Dim colWork As New DataColumn("Item No", GetType(String))
        dataTable.Columns.Add(colWork)
        '' add CustomerID column to key array and bind to DataTable
        Dim Keys(0) As DataColumn
        Keys(0) = colWork
        ' dataTable.PrimaryKey = Keys
        ' create and add a Description column
        colWork = New DataColumn("Description", GetType(String))
        colWork.MaxLength = 250
        dataTable.Columns.Add(colWork)
        ' create and add a Qty column
        colWork = New DataColumn("Qty", GetType(Integer))
        dataTable.Columns.Add(colWork)
        ' create and add a Start column
        colWork = New DataColumn("Start No", GetType(Integer))
        dataTable.Columns.Add(colWork)
        ' create and add a End column
        colWork = New DataColumn("End No", GetType(Integer))
        dataTable.Columns.Add(colWork)
        colWork = New DataColumn("Value", GetType(Double))
        dataTable.Columns.Add(colWork)

        Return dataTable

    End Function



End Class
