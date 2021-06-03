Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Invoice_Job
    Dim _Print_Status As String
    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_Grid_Invoice(ByVal strStatus As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T08ID ) as  ##,t08date as [Date],T08Invo_No as [Inv.No],T08Job_No as [Job No],T08V_No as [Vehicle No],M06Name as [Customer Name],CAST(T08Net_Amount AS DECIMAL(16,2)) as [Invoice Value]  from T08Sales_Header inner join M06Customer_Master  on T08Cus_NO=M06Code where T08Status='A' and T08Tr_Type='" & strStatus & "' order by T08ID desc   "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 220
            UltraGrid2.Rows.Band.Columns(6).Width = 110
            'UltraGrid2.Rows.Band.Columns(7).Width = 110
            'UltraGrid2.Rows.Band.Columns(8).Width = 90

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Private Sub frmView_Invoice_Job_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_Invoice("JOB_INVOICE")
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid2.ActiveRow.Index
        frmItem_Issue_Uniq.UltraTabControl1.Tabs(1).Selected = True
        frmItem_Issue_Uniq.txtEnter1.Text = Trim(UltraGrid2.Rows(_Row).Cells(2).Text)
        frmItem_Issue_Uniq.Search_Invoice()
        Me.Close()
    End Sub
End Class