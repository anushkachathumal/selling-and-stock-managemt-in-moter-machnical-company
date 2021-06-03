Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Vehicle_History


    Function Load_Vehicle_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05ID ) as [##],T05Job_No as [Job No],T05Date as [Date],T05Department as [Department],T05Vehi_No as [Vehicle No],T05Mtr as [Meter Reading],M06Name  as [Customer Name] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No  where T05Status='CLOSE' order by T05Id desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 90
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 290

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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

    Function Load_Vehicle_Records_vehi()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05ID ) as [##],T05Job_No as [Job No],T05Date as [Date],T05Department as [Department],T05Vehi_No as [Vehicle No],T05Mtr as [Meter Reading],M06Name  as [Customer Name] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No  where T05Status='CLOSE' and T05Vehi_No='" & Trim(cboV_No.Text) & "' order by T05Id desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 90
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 290

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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


    Function Load_Vehicle_Records_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05ID ) as [##],T05Job_No as [Job No],T05Date as [Date],T05Department as [Department],T05Vehi_No as [Vehicle No],T05Mtr as [Meter Reading],M06Name  as [Customer Name] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No  where T05Status='CLOSE' and M06Name='" & Trim(cboCustomer.Text) & "' order by T05Id desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 90
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 290

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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

    Function Load_Vehicle_Records_Department()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05ID ) as [##],T05Job_No as [Job No],T05Date as [Date],T05Department as [Department],T05Vehi_No as [Vehicle No],T05Mtr as [Meter Reading],M06Name  as [Customer Name] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No  where T05Status='CLOSE' and T05Department='" & Trim(cboDepartment.Text) & "' order by T05Id desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 90
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 290

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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

    Function Load_VNO()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select T05Vehi_No as [##],max(M06Name)  as [Customer Name] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No where T05Status='CLOSE'  group by T05Vehi_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboV_No
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 90
                .Rows.Band.Columns(1).Width = 210

            End With
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
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
            Sql = "select M06Name  as [##] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No where T05Status='CLOSE'  group by M06Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 290
                '.Rows.Band.Columns(1).Width = 210

            End With
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Load_Department()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select T05Department  as [##] from T05Job_Card inner join M06Customer_Master on M06Code=T05Cus_No where T05Status='CLOSE'  group by T05Department"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDepartment
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 234
                '.Rows.Band.Columns(1).Width = 210

            End With
            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Private Sub frmView_Vehicle_History_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Vehicle_Records()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Vehicle_Records()
        Panel3.Visible = False
        Panel1.Visible = False
        Panel3.Visible = False
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        Call Load_VNO()
        Panel3.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        cboV_No.ToggleDropdown()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Load_Vehicle_Records_vehi()
        Panel3.Visible = False
    End Sub

    Private Sub ByPartNoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByPartNoToolStripMenuItem.Click
        Call Load_Customer()
        cboCustomer.Text = ""
        Panel1.Visible = True
        Panel3.Visible = False
        Panel2.Visible = False
        cboCustomer.ToggleDropdown()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Vehicle_Records_Customer()
        Panel1.Visible = False
    End Sub

    Private Sub ByDepartmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDepartmentToolStripMenuItem.Click
        Call Load_Department()
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        cboDepartment.Text = ""

    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Vehicle_Records_Department()
        Panel2.Visible = False
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid2.ActiveRow.Index
        frmView_Invoice.Close()
        frmView_Invoice.Show()
        With frmView_Invoice
            .txtJob.Text = UltraGrid2.Rows(_Row).Cells(1).Text
            .Search_Invoice()
        End With
    End Sub

   
End Class