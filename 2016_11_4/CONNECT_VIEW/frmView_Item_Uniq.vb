Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Item_Uniq
    Dim _Print_Status As String


    Function Load_Grid_Category_1()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05Ref_No ) as  ##,MAX(M01Description) as [Category],M05Ref_No as [#Ref.No],max(M05Item_Code) as [Part No],MAX(tmpDescription) as [Item Name],CAST(MAX(Retail) AS DECIMAL(16,2)) as [Retail Price],SUM(qTY) as [Stock Qty],MAX(Rack) as [#Rack No],MAX(Cell) as [#Cell No],MAX(M05Use_For) as [Use For]  from View_Product_Stock where M01Description = '" & Trim(cboCategory1.Text) & "' and M05Use_For like '%" & txtFind1.Text & "%' GROUP BY M05Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 90
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            UltraGrid2.Rows.Band.Columns(6).Width = 70
            UltraGrid2.Rows.Band.Columns(7).Width = 60
            UltraGrid2.Rows.Band.Columns(8).Width = 60
            UltraGrid2.Rows.Band.Columns(9).Width = 270
            UltraGrid2.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function Load_Grid_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05Ref_No ) as  ##,MAX(M01Description) as [Category],M05Ref_No as [#Ref.No],max(M05Item_Code) as [Part No],MAX(tmpDescription) as [Item Name],CAST(MAX(Retail) AS DECIMAL(16,2)) as [Retail Price],SUM(qTY) as [Stock Qty],MAX(Rack) as [#Rack No],MAX(Cell) as [#Cell No],MAX(M05Use_For) as [Use For]  from View_Product_Stock where M01Description = '" & Trim(cboCategory.Text) & "' GROUP BY M05Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 90
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            UltraGrid2.Rows.Band.Columns(6).Width = 70
            UltraGrid2.Rows.Band.Columns(7).Width = 60
            UltraGrid2.Rows.Band.Columns(8).Width = 60
            UltraGrid2.Rows.Band.Columns(9).Width = 270
            UltraGrid2.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function Load_Grid_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05Ref_No ) as  ##,MAX(M01Description) as [Category],M05Ref_No as [#Ref.No],max(M05Item_Code) as [Part No],MAX(tmpDescription) as [Item Name],CAST(MAX(Retail) AS DECIMAL(16,2)) as [Retail Price],SUM(qTY) as [Stock Qty],MAX(Rack) as [#Rack No],MAX(Cell) as [#Cell No],MAX(M05Use_For) as [Use For]  from View_Product_Stock GROUP BY M05Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 90
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            UltraGrid2.Rows.Band.Columns(6).Width = 70
            UltraGrid2.Rows.Band.Columns(7).Width = 60
            UltraGrid2.Rows.Band.Columns(8).Width = 60
            UltraGrid2.Rows.Band.Columns(9).Width = 270
            UltraGrid2.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function Load_Grid_Data_partNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05Ref_No ) as  ##,MAX(M01Description) as [Category],M05Ref_No as [#Ref.No],max(M05Item_Code) as [Part No],MAX(tmpDescription) as [Item Name],CAST(MAX(Retail) AS DECIMAL(16,2)) as [Retail Price],SUM(qTY) as [Stock Qty],MAX(Rack) as [#Rack No],MAX(Cell) as [#Cell No],MAX(M05Use_For) as [Use For]  from View_Product_Stock where M05Item_Code like '%" & txtSearch.Text & "%' GROUP BY M05Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 90
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            UltraGrid2.Rows.Band.Columns(6).Width = 70
            UltraGrid2.Rows.Band.Columns(7).Width = 60
            UltraGrid2.Rows.Band.Columns(8).Width = 60
            UltraGrid2.Rows.Band.Columns(9).Width = 270
            UltraGrid2.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function Load_Grid_Data_ItemName()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05Ref_No ) as  ##,MAX(M01Description) as [Category],M05Ref_No as [#Ref.No],max(M05Item_Code) as [Part No],MAX(tmpDescription) as [Item Name],CAST(MAX(Retail) AS DECIMAL(16,2)) as [Retail Price],SUM(qTY) as [Stock Qty],MAX(Rack) as [#Rack No],MAX(Cell) as [#Cell No],MAX(M05Use_For) as [Use For]  from View_Product_Stock where M01Description like '%" & txtSearch.Text & "%' GROUP BY M05Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 90
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            UltraGrid2.Rows.Band.Columns(6).Width = 70
            UltraGrid2.Rows.Band.Columns(7).Width = 60
            UltraGrid2.Rows.Band.Columns(8).Width = 60
            UltraGrid2.Rows.Band.Columns(9).Width = 270
            UltraGrid2.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Private Sub frmView_Item_Uniq_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_Data()
        Call Load_Brand()
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Grid_Data()
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Panel3.Visible = False
    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        _Print_Status = "A1"
        Panel1.Visible = False
        txtSearch.Text = ""
        Panel3.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub txtSearch_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.ValueChanged
        If _Print_Status = "A1" Then
            Call Load_Grid_Data_partNo()
        ElseIf _Print_Status = "A2" Then
            Call Load_Grid_Data_ItemName()
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        _Print_Status = "A2"
        Panel3.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        txtSearch.Text = ""
        txtSearch.Focus()
    End Sub

    Private Sub CategoryOnlyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CategoryOnlyToolStripMenuItem.Click
        _Print_Status = "A3"
        Panel1.Visible = True
        Panel3.Visible = False
        Panel2.Visible = False
        ' cboCategory.Text = ""
        ' Call Load_Brand()
    End Sub

    Function Load_Brand()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M01Description as [##] from M01Category WHERE M01Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275
           
            End With
            With cboCategory1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275

            End With
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

 

    Private Sub UltraButton1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Grid_Category()
        Panel1.Visible = False
    End Sub

    Private Sub CategoryVehicalBrandNameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CategoryVehicalBrandNameToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        txtFind1.Text = ""
        Call Load_Brand()
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Panel2.Visible = False
    End Sub

    Private Sub txtFind1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFind1.ValueChanged
        Call Load_Grid_Category_1()
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid2.ActiveRow.Index
        If strWindowName = "frmItem_Issue_Uniq" Then
            strItem_Code = UltraGrid2.Rows(_Row).Cells(2).Text
            frmItem_Issue_Uniq.Search_Item()
            Me.Close()
        ElseIf strWindowName = "frmDirect_Sales" Then
            strItem_Code = UltraGrid2.Rows(_Row).Cells(2).Text
            frmDirect_Sales.Search_Item()
            Me.Close()
        End If
    End Sub
End Class