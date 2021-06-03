Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_GRN_cnt
    Dim _Print_Status As String

    Function Load_Grid_GRN()
        Dim Sql As String
        Dim con = New SqlConnection()
        '   SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Com_Invoice as [Com.Inv.No],T01Date as [Date],M04Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount],CAST(Discount AS DECIMAL(16,2)) as [Discount],CAST(Gross AS DECIMAL(16,2)) as [Gross Amount]  from View_GRN_Header where  T01Status='A' and T01Tr_Type='GRN'  order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 100
            UltraGrid2.Rows.Band.Columns(6).Width = 100
            UltraGrid2.Rows.Band.Columns(7).Width = 100
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Function Load_Grid_GRN_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Com_Invoice as [Com.Inv.No],T01Date as [Date],M04Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount],CAST(Discount AS DECIMAL(16,2)) as [Discount],CAST(Gross AS DECIMAL(16,2)) as [Gross Amount]  from View_GRN_Header where  T01Status='A' and T01Date between '" & txtA1.Text & "' and '" & txtA2.Text & "' and M04Name='" & Trim(cboSupplier.Text) & "' AND T01Tr_Type='GRN' order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 100
            UltraGrid2.Rows.Band.Columns(6).Width = 100
            UltraGrid2.Rows.Band.Columns(7).Width = 100
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Function Load_Grid_GRN_Rtn_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Com_Invoice as [Com.Inv.No],T01Date as [Date],M04Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount],CAST(Discount AS DECIMAL(16,2)) as [Discount],CAST(Gross AS DECIMAL(16,2)) as [Gross Amount]  from View_GRN_Header where  T01Status='CLOSE' and T01Date between '" & txtA1.Text & "' and '" & txtA2.Text & "' and M04Name='" & Trim(cboSupplier.Text) & "' and T01Tr_Type='GRN' order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 100
            UltraGrid2.Rows.Band.Columns(6).Width = 100
            UltraGrid2.Rows.Band.Columns(7).Width = 100
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Function Load_Grid_GRN_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Com_Invoice as [Com.Inv.No],T01Date as [Date],M04Name as [Supplier Name],CAST(T02Cost AS DECIMAL(16,2)) as [Cost Price],CAST(T02Qty AS DECIMAL(16,2)) as [Qty],CAST(Total AS DECIMAL(16,2)) as [Total Amount]  from View_GRN_Flutter where  T01Status='A' and T01Date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and T02Part_No='" & Trim(cboItem.Text) & "' and T01Tr_Type='GRN' order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 100
            UltraGrid2.Rows.Band.Columns(6).Width = 100
            UltraGrid2.Rows.Band.Columns(7).Width = 100
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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


    Function Load_Grid_GRN_Date()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Com_Invoice as [Com.Inv.No],T01Date as [Date],M04Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount],CAST(Discount AS DECIMAL(16,2)) as [Discount],CAST(Gross AS DECIMAL(16,2)) as [Gross Amount]  from View_GRN_Header where  T01Status='A' and T01Date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and T01Tr_Type='GRN'  order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 100
            UltraGrid2.Rows.Band.Columns(6).Width = 100
            UltraGrid2.Rows.Band.Columns(7).Width = 100
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Function Load_Grid_Return_Date()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Com_Invoice as [Com.Inv.No],T01Date as [Date],M04Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount],CAST(Discount AS DECIMAL(16,2)) as [Discount],CAST(Gross AS DECIMAL(16,2)) as [Gross Amount]  from View_GRN_Header where  T01Status='CLOSE' and T01Date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and T01Tr_Type='GRN'  order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 100
            UltraGrid2.Rows.Band.Columns(6).Width = 100
            UltraGrid2.Rows.Band.Columns(7).Width = 100
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Private Sub frmView_GRN_cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_GRN()
        Call Load_Item_Code()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Grid_GRN()
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
    End Sub

    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M04Name as [##] from M04Supplier where M04Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupplier
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 252
                ' .Rows.Band.Columns(1).Width = 180


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

    Function Load_Item_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M05Item_Code as [##],tmpDescription as [Description] from View_Product_Item where M05Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 252
                .Rows.Band.Columns(1).Width = 310


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


    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        _Print_Status = "A"
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel1.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _Print_Status = "A" Then
            Call Load_Grid_GRN_Date()
            Panel1.Visible = False
        ElseIf _Print_Status = "X1" Then
            Call Load_Grid_Return_Date()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub BySupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem.Click
        _Print_Status = "B"
        txtA1.Text = Today
        txtA2.Text = Today
        Panel2.Visible = True
        Panel1.Visible = False
        Call Load_Supplier()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _Print_Status = "B" Then
            Call Load_Grid_GRN_Supplier()
            Panel2.Visible = False
        ElseIf _Print_Status = "X2" Then
            Call Load_Grid_GRN_Rtn_Supplier()
            Panel2.Visible = False
        End If
    End Sub

    Private Sub ByPartNoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByPartNoToolStripMenuItem.Click
        _Print_Status = "C"
        txtC1.Text = Today
        txtC2.Text = Today
        Panel2.Visible = False
        Panel1.Visible = False
        Panel3.Visible = True
        Call Load_Item_Code()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Load_Grid_GRN_Item()
        Panel3.Visible = False
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer
        _Row = UltraGrid2.ActiveRow.Index
        frmGRN_uniq.txtEntry.Text = Trim(UltraGrid2.Rows(_Row).Cells(1).Text)
        frmGRN_uniq.Load_Gride2()
        frmGRN_uniq.Search_Records()
        Me.Close()
    End Sub


    Private Sub ByDateToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem1.Click
        _Print_Status = "X1"
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel1.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub BySupplierToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem1.Click
        _Print_Status = "X2"
        txtA1.Text = Today
        txtA2.Text = Today
        Panel2.Visible = True
        Panel1.Visible = False
        Call Load_Supplier()
    End Sub
End Class