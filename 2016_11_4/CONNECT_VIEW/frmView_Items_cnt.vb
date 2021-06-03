Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Items_cnt
    Function Load_Grid_PRODUCT()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05ID ) as  ##,M05Ref_no as [##],M05Item_Code as [iItem Code],tmpDescription as [Item Name]  from View_Product_Item where  M05Status='A' order by M05ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 280
            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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


    Function sEARCH_Grid_PRODUCT()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select ROW_NUMBER() OVER (ORDER BY m05item_code ) as  ##, m05item_code as [Part No],MAX(tmpDescription) as [Description],max(CAST(Retail AS DECIMAL(16,2))) as [Retail Price],sum(qty) as [Current Stock],max(rack) as [Rack No],max(cell) as [Cell No] from View_Product_Stock  where tmpDescription like '%" & Trim(txtSearch.Text) & "%' group by m05item_code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 40

            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 280
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            UltraGrid2.Rows.Band.Columns(6).Width = 80
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function SEARCH_Grid_ROW()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05ID ) as  ##,M05Item_Code as [iItem Code],tmpDescription as [Item Name]  from View_Product_Item where  M05Status='A' AND tmpDescription LIKE '%" & Trim(txtSearch.Text) & "%' order by M05ID"
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

    Function Load_Grid_ROW()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select ROW_NUMBER() OVER (ORDER BY m05item_code ) as  ##, m05item_code as [Part No],MAX(tmpDescription) as [Description],max(CAST(Retail AS DECIMAL(16,2))) as [Retail Price],sum(qty) as [Current Stock],max(rack) as [Rack No],max(cell) as [Cell No] from View_Product_Stock group by m05item_code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 40

            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 280
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            UltraGrid2.Rows.Band.Columns(6).Width = 80
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer
        _Row = UltraGrid2.ActiveRow.Index
        If strWindowName = "frmWastage_cnt" Then
            frmWastage_cnt.cboCode.Text = UltraGrid2.Rows(_Row).Cells(1).Text
            frmWastage_cnt.Serch_Item()
            Me.Close()
        
        ElseIf strWindowName = "frmMK_Return_cnt" Then
            frmMK_Return_cnt.cboCode.Text = UltraGrid2.Rows(_Row).Cells(1).Text
            frmMK_Return_cnt.Serch_Item()
            Me.Close()
        ElseIf strWindowName = "frmGRN_uniq" Then
            frmGRN_uniq.cboCode.Text = UltraGrid2.Rows(_Row).Cells(1).Text
            frmGRN_uniq.Serch_Item()
            Me.Close()
        End If
    End Sub

    Private Sub txtSearch_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.ValueChanged
        Call sEARCH_Grid_PRODUCT()
       
    End Sub

    Private Sub frmView_Items_cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_ROW()
    End Sub
End Class