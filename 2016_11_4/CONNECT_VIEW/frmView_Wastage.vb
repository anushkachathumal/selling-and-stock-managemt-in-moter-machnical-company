Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Wastage
    Dim _Print_Status As String

    Function Load_Wastege()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Date as [Date],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount]  from View_Wastage_Header where  T01Status='A' and T01Tr_Type='WST' order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 120
            'UltraGrid2.Rows.Band.Columns(4).Width = 180
            'UltraGrid2.Rows.Band.Columns(5).Width = 100

            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

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

    Function Load_Grid_WST_Date()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Date as [Date],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount]  from View_Wastage_Header where  T01Status='A' and T01Date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and T01Tr_Type='WST' order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 120
            'UltraGrid2.Rows.Band.Columns(4).Width = 180
            'UltraGrid2.Rows.Band.Columns(5).Width = 100

            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

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

    Function Load_Grid_DEAC_WST_Date()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Date as [Date],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount]  from View_Wastage_Header where  T01Status='CLOSE' and T01Date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and T01Tr_Type='WST' order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 120
            'UltraGrid2.Rows.Band.Columns(4).Width = 180
            'UltraGrid2.Rows.Band.Columns(5).Width = 100

            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

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


    Private Sub frmView_Wastage_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Wastege()
    End Sub

    Function Load_Grid_WST_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Date as [Date],T02Part_No AS [Part No],tmpDescription as [Description],CAST(T02Cost AS DECIMAL(16,2)) as [Cost Price],CAST(T02Qty AS DECIMAL(16,2)) as [Qty],CAST(Total AS DECIMAL(16,2)) as [Total Amount]  from View_Wastage_Fluter where  T01Status='A' and T01Date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and T02Part_No='" & Trim(cboItem.Text) & "' and T01Tr_Type='WST' order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 110
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 70
            UltraGrid2.Rows.Band.Columns(7).Width = 100
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Function Load_Grid_Deac_WST_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T01ID ) as  ##,T01Ref_No as [Ref No],T01Date as [Date],T02Part_No AS [Part No],tmpDescription as [Description],CAST(T02Cost AS DECIMAL(16,2)) as [Cost Price],CAST(T02Qty AS DECIMAL(16,2)) as [Qty],CAST(Total AS DECIMAL(16,2)) as [Total Amount]  from View_Wastage_Fluter where  T01Status='CLOSE' and T01Date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and T02Part_No='" & Trim(cboItem.Text) & "' and T01Tr_Type='WST' order by T01ID desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 110
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 180
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 70
            UltraGrid2.Rows.Band.Columns(7).Width = 100
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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
        con = DBEngin.GetConnection(True)
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

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Wastege()
        Panel1.Visible = False
        Panel3.Visible = False
        cboItem.Text = ""
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        _Print_Status = "A"
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel1.Visible = True
        Panel3.Visible = False
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _Print_Status = "A" Then
            Call Load_Grid_WST_Date()
            Panel1.Visible = False
        ElseIf _Print_Status = "B" Then
            Call Load_Grid_DEAC_WST_Date()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub ByPartNoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByPartNoToolStripMenuItem.Click
        Call Load_Item_Code()
        _Print_Status = "A2"
        txtC1.Text = Today
        txtC2.Text = Today
        Panel1.Visible = False
        Panel3.Visible = True
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If _Print_Status = "A2" Then
            Call Load_Grid_WST_Item()
            Panel3.Visible = False
        ElseIf _Print_Status = "B2" Then
            Call Load_Grid_Deac_WST_Item()
            Panel3.Visible = False
        End If
    End Sub

    Private Sub ByDateToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem1.Click
        _Print_Status = "B"
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel1.Visible = True
        Panel3.Visible = False
    End Sub

    Private Sub BySupplierToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem1.Click
        Call Load_Item_Code()
        _Print_Status = "B2"
        txtC1.Text = Today
        txtC2.Text = Today
        Panel1.Visible = False
        Panel3.Visible = True
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _rOW As Integer
        _rOW = UltraGrid2.ActiveRow.Index
        frmWastage_cnt.txtEntry.Text = Trim(UltraGrid2.Rows(_rOW).Cells(1).Text)
        frmWastage_cnt.Load_Gride2()
        frmWastage_cnt.Search_Records()
        Me.Close()
    End Sub

   
End Class