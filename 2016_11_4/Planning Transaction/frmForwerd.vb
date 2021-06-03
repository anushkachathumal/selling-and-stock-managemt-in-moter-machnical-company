
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Public Class frmForwerd
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable

    Private Sub frmForwerd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDate.Text = Today
        Call Load_Combo()
        Call Load_Gride()
    End Sub

    Function Load_Combo()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim Sql As String
        Dim M01 As DataSet

        Try
            Sql = "select T01Sales_Order   from T01Delivary_Request where T01Status='A' group by T01Sales_Order "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            With cboSales_Order
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                '.Rows.Band.Columns(1).Width = 260
            End With
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try


    End Function

    Function Search_Records() As Boolean
        Dim vcWhere As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)

        Try
            vcWhere = "T01Sales_Order='" & cboSales_Order.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "T01D"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                txtDate.Text = M01.Tables(0).Rows(0)("T01Date")
                txtMerchant.Text = M01.Tables(0).Rows(0)("T01Planner")
            End If
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cboSales_Order_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSales_Order.AfterCloseUp
        Call Search_Records()
        Call Allocate_Planer()
        Call Load_Gride()
    End Sub

    Function Allocate_Planer()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim Sql As String
        Dim M01 As DataSet

        Try
            Sql = "select T01Planner  as [##] from T01Delivary_Request where T01Planner<>'" & txtMerchant.Text & "' group by T01Planner "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            With cboMerchant
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                '.Rows.Band.Columns(1).Width = 260
            End With
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function

    Private Sub cboSales_Order_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSales_Order.LostFocus
        Call Search_Records()
        Call Allocate_Planer()
        Call Load_Gride()
    End Sub

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            
            Sql = "select T01Line_Item as [Line Items],T01Qty as [Qty],T01Bulk as [##],T01Lab_Dye as [Lab Dye] from T01Delivary_Request where T01Sales_Order='" & cboSales_Order.Text & "' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 110
            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim i As Integer
        Try
            nvcFieldList1 = "update T01Delivary_Request set T01Planner='" & cboMerchant.Text & "' where T01Sales_Order='" & cboSales_Order.Text & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from T20Forwerd_Merchant where T20Sales_Order='" & cboSales_Order.Text & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert Into T20Forwerd_Merchant(T20Sales_Order,T20Last_Merchant,T20New_Merchant,T20Date)" & _
                                                          " values('" & UCase(Trim(cboSales_Order.Text)) & "', '" & (Trim(txtMerchant.Text)) & "','" & cboMerchant.Text & "','" & Now & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            connection.Close()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR05)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            OPR05.Enabled = True


            Call Load_Gride()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try

    End Sub
End Class