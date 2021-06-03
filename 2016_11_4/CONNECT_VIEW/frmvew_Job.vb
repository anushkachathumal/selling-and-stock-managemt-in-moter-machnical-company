Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmvew_Job
    Dim _Print_Status As String
    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_Grid_JOB()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05Id ) as  ##, T05Department as [Department],T05Job_No as [Job No],T05Ref_No as [Ref No],T05Date as [Date],T05Vehi_No as [Vehicle No],M06Name as [Customer Name] FROM T05Job_Card INNER JOIN M06Customer_Master ON M06Code=T05Cus_No where T05Status='A' order by T05Id "

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 110
            UltraGrid2.Rows.Band.Columns(6).Width = 210

            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            'UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Function Load_Grid_JOB_Close()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05Id ) as  ##, T05Department as [Department],T05Job_No as [Job No],T05Ref_No as [Ref No],T05Date as [Date],T05Vehi_No as [Vehicle No],M06Name as [Customer Name] FROM T05Job_Card INNER JOIN M06Customer_Master ON M06Code=T05Cus_No where T05Status='CLOSE' AND T05Date BETWEEN '" & txtDate1.Text & "' AND '" & txtDate2.Text & "' order by T05Id "

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 110
            UltraGrid2.Rows.Band.Columns(6).Width = 210

            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function Load_Grid_JOB_Cancel()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05Id ) as  ##, T05Department as [Department],T05Job_No as [Job No],T05Ref_No as [Ref No],T05Date as [Date],T05Vehi_No as [Vehicle No],M06Name as [Customer Name] FROM T05Job_Card INNER JOIN M06Customer_Master ON M06Code=T05Cus_No where T05Status='CANCEL' AND T05Date BETWEEN '" & txtDate1.Text & "' AND '" & txtDate2.Text & "' order by T05Id "

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 110
            UltraGrid2.Rows.Band.Columns(6).Width = 210

            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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


    Function Load_Grid_JOB_Close_customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05Id ) as  ##, T05Department as [Department],T05Job_No as [Job No],T05Ref_No as [Ref No],T05Date as [Date],T05Vehi_No as [Vehicle No],M06Name as [Customer Name] FROM T05Job_Card INNER JOIN M06Customer_Master ON M06Code=T05Cus_No where T05Status='CLOSE' AND T05Date BETWEEN '" & txtA1.Text & "' AND '" & txtA2.Text & "' and M06Name='" & Trim(cboCustomer.Text) & "' order by T05Id desc"

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 110
            UltraGrid2.Rows.Band.Columns(6).Width = 210

            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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


    Function Load_Grid_JOB_CANCEL_customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05Id ) as  ##, T05Department as [Department],T05Job_No as [Job No],T05Ref_No as [Ref No],T05Date as [Date],T05Vehi_No as [Vehicle No],M06Name as [Customer Name] FROM T05Job_Card INNER JOIN M06Customer_Master ON M06Code=T05Cus_No where T05Status='CLOSE' AND T05Date BETWEEN '" & txtA1.Text & "' AND '" & txtA2.Text & "' and M06Name='" & Trim(cboCustomer.Text) & "' order by T05Id desc"

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 110
            UltraGrid2.Rows.Band.Columns(6).Width = 210

            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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


    Function Load_Grid_JOB_Close_VNO()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05Id ) as  ##, T05Department as [Department],T05Job_No as [Job No],T05Ref_No as [Ref No],T05Date as [Date],T05Vehi_No as [Vehicle No],M06Name as [Customer Name] FROM T05Job_Card INNER JOIN M06Customer_Master ON M06Code=T05Cus_No where T05Status='CLOSE' AND T05Date BETWEEN '" & txtB1.Text & "' AND '" & txtB2.Text & "' and T05Vehi_No='" & Trim(cboVNo.Text) & "' order by T05Id desc"

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 110
            UltraGrid2.Rows.Band.Columns(6).Width = 210

            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Function Load_Grid_JOB_CANCEL_VNO()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY T05Id ) as  ##, T05Department as [Department],T05Job_No as [Job No],T05Ref_No as [Ref No],T05Date as [Date],T05Vehi_No as [Vehicle No],M06Name as [Customer Name] FROM T05Job_Card INNER JOIN M06Customer_Master ON M06Code=T05Cus_No where T05Status='CANCEL' AND T05Date BETWEEN '" & txtB1.Text & "' AND '" & txtB2.Text & "' and T05Vehi_No='" & Trim(cboVNo.Text) & "' order by T05Id desc"

            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 110
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 80
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 110
            UltraGrid2.Rows.Band.Columns(6).Width = 210

            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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

    Private Sub frmvew_Job_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_JOB()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Call Load_Grid_JOB()
    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Call Load_Grid_JOB()
    End Sub

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        txtDate1.Text = Today
        txtDate2.Text = Today
        _Print_Status = "A1"

    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _Print_Status = "A1" Then
            Call Load_Grid_JOB_Close()
            Panel1.Visible = False
        ElseIf _Print_Status = "B1" Then
            Call Load_Grid_JOB_Cancel()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub ByCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByCustomerToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        txtA1.Text = Today
        txtA2.Text = Today
        cboCustomer.Text = ""
        Call Load_Customer_name()
        _Print_Status = "A2"
    End Sub

    Function Load_Customer_name()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select  (M06Name) as [##] from M06Customer_Master WHERE M06Status='A' order by M06ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 252
                '  .Rows.Band.Columns(1).Width = 160


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

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _Print_Status = "A2" Then
            Call Load_Grid_JOB_Close_customer()
            Panel2.Visible = False
        ElseIf _Print_Status = "B2" Then
            Call Load_Grid_JOB_CANCEL_customer()
            Panel2.Visible = False
        End If
    End Sub

    Function Load_VNO()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M07V_No as [##] from M07Vehicle_Master inner join M06Customer_Master on M06Code=M07Cus_Code where M07Status='A'  order by M07ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboVNo
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 252
                ' .Rows.Band.Columns(1).Width = 210

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

    Private Sub ByVehicleNoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByVehicleNoToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = True
        txtB1.Text = Today
        txtB2.Text = Today
        cboVNo.Text = ""
        Call Load_VNO()
        _Print_Status = "A3"
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If _Print_Status = "A3" Then
            Call Load_Grid_JOB_Close_VNO()
            Panel3.Visible = False
        ElseIf _Print_Status = "B3" Then
            Call Load_Grid_JOB_CANCEL_VNO()
            Panel3.Visible = False
        End If
    End Sub

    Private Sub ByDateToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem1.Click
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        txtDate1.Text = Today
        txtDate2.Text = Today
        _Print_Status = "B1"
    End Sub

    Private Sub BySupplierToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem1.Click
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        txtA1.Text = Today
        txtA2.Text = Today
        cboCustomer.Text = ""
        Call Load_Customer_name()
        _Print_Status = "B2"
    End Sub

    Private Sub ByVehicleNoToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByVehicleNoToolStripMenuItem1.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = True
        txtB1.Text = Today
        txtB2.Text = Today
        cboVNo.Text = ""
        Call Load_VNO()
        _Print_Status = "B3"
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _ROW As Integer

        _ROW = UltraGrid2.ActiveRow.Index
        frmJob_Card_Uniq.txtEntry.Text = Trim(UltraGrid2.Rows(_ROW).Cells(2).Text)
        frmJob_Card_Uniq.SEARCH_RECORDS()
        Me.Close()
    End Sub
End Class