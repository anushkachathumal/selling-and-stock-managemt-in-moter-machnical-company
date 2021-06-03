Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Direct_Sales
    Dim _Print_Status As String
    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_Grid_DIRECT()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T08ID ) as  ##,T08Invo_No as [Invo.No],T08Date as [Date],Cus_Name as [Customer Name],CAST(T08Net_Amount AS DECIMAL(16,2)) as [Total Amount] from View_DIRECT_SALES where T08Status='A' order by T08ID  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 220
            UltraGrid2.Rows.Band.Columns(4).Width = 110

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid_DIRECT_Date()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T08ID ) as  ##,T08Invo_No as [Invo.No],T08Date as [Date],Cus_Name as [Customer Name],CAST(T08Net_Amount AS DECIMAL(16,2)) as [Total Amount] from View_DIRECT_SALES where T08Status='A' and T08Date between '" & txtC1.Text & "' and '" & txtC2.Text & "' order by T08ID  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 220
            UltraGrid2.Rows.Band.Columns(4).Width = 110

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid_DIRECT_Cancel_Date()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T08ID ) as  ##,T08Invo_No as [Invo.No],T08Date as [Date],Cus_Name as [Customer Name],CAST(T08Net_Amount AS DECIMAL(16,2)) as [Total Amount] from View_DIRECT_SALES where T08Status='CANCEL' and T08Date between '" & txtC1.Text & "' and '" & txtC2.Text & "' order by T08ID  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 220
            UltraGrid2.Rows.Band.Columns(4).Width = 110

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid_DIRECT_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T08ID ) as  ##,T08Invo_No as [Invo.No],T08Date as [Date],Cus_Name as [Customer Name],CAST(T08Net_Amount AS DECIMAL(16,2)) as [Total Amount] from View_DIRECT_SALES where T08Status='A' and T08Date between '" & txtD1.Text & "' and '" & txtD2.Text & "' and Cus_Name='" & Trim(cboCustomer.Text) & "' order by T08ID  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 220
            UltraGrid2.Rows.Band.Columns(4).Width = 110

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid_DIRECT_PartNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T08ID ) as  ##,T08Invo_No as [Invo.No],t08date as [Date],T09Item_Code as [Part No],T09Item_Name as [Item Name],CAST(T09Retail AS DECIMAL(16,2)) as [Rate],T09Qty as [Qty],T09Discount as [Discount%],CAST((T09Qty*T09Retail) -((T09Qty*T09Retail)*(T09Discount/100)) AS DECIMAL(16,2)) as [Total Amount]  from View_DIRECT_SALES inner join T09Sales_Flutter on T08Invo_No=T09Inv_No where T08Status='A' and t08date between '" & txtE1.Text & "' and '" & txtE2.Text & "' and T09Item_Code='" & Trim(cboPartNo.Text) & "' order by  T08ID "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 70
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 190
            UltraGrid2.Rows.Band.Columns(5).Width = 70
            UltraGrid2.Rows.Band.Columns(6).Width = 70
            UltraGrid2.Rows.Band.Columns(7).Width = 90

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Function Load_PartNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M05Item_Code as [##],tmpDescription as [Distribution] from View_Product_Item where M05Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPartNo
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 249
                .Rows.Band.Columns(0).Width = 349


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


    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M06Name as [##] from M06Customer_Master where M06Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 249



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

    Private Sub frmView_Direct_Sales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_DIRECT()
    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Panel3.Visible = True
        Panel1.Visible = False
        txtC1.Text = Today
        txtC2.Text = Today
        _Print_Status = "A1"
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Call Load_Grid_DIRECT()

    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If _Print_Status = "A1" Then
            Call Load_Grid_DIRECT_Date()
            Panel3.Visible = False
        ElseIf _Print_Status = "A2" Then
            Call Load_Grid_DIRECT_Cancel_Date()
            Panel3.Visible = False
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Panel1.Visible = True
        Panel3.Visible = False
        cboCustomer.Text = ""
        Call Load_Customer()
        cboCustomer.ToggleDropdown()
        txtD1.Text = Today
        txtD2.Text = Today
        _Print_Status = "B"
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Grid_DIRECT_Customer()
        Panel1.Visible = False
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        _Print_Status = "C"
        txtE1.Text = Today
        txtE2.Text = Today
        Call Load_PartNo()
        cboPartNo.Text = ""

    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Grid_DIRECT_PartNo()
        Panel2.Visible = False
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid2.ActiveRow.Index
        frmDirect_Sales.txtEnter1.Text = Trim(UltraGrid2.Rows(_Row).Cells(1).Text)
        frmDirect_Sales.Search_Records()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Panel3.Visible = True
        Panel1.Visible = False
        txtC1.Text = Today
        txtC2.Text = Today
        _Print_Status = "A2"
    End Sub
End Class