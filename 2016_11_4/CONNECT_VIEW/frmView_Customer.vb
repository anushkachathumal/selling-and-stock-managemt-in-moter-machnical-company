Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Customer
    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M06Code) as  ##,M06Code as [Customer Code],M06Name as [Customer Name]  from M06Customer_Master where M06Status='A' order by M06ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 280
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function Load_Grid_Find()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M06Code) as  ##,M06Code as [Customer Code],M06Name as [Customer Name]  from M06Customer_Master where M06Status='A' and M06Name like '%" & txtSearch.Text & "%' order by M06ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 280
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Private Sub frmView_Customer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid()
    End Sub

    Private Sub txtSearch_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.ValueChanged
        Call Load_Grid_Find()
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer
        _Row = UltraGrid2.ActiveRow.Index
        If strWindowName = "frmDiscount_2" Then
            frmDiscount_2.cboCus_Code.Text = UltraGrid2.Rows(_Row).Cells(1).Text
            frmDiscount_2.Search_Cus_Code(frmDiscount_2.cboCus_Code.Text)
            Me.Close()
        ElseIf strWindowName = "frmJob_Card_Uniq" Then
            frmJob_Card_Uniq.cboCus_Name.Text = UltraGrid2.Rows(_Row).Cells(2).Text
            frmJob_Card_Uniq.search_Customer()
            Me.Close()
        End If
    End Sub

   
End Class