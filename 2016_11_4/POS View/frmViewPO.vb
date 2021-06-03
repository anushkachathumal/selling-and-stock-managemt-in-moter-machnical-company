Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmViewPO
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub frmViewPO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Grid()

        txtFrom.Text = Today
        txtTo.Text = Today
        Call Load_Combo()
    End Sub
    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Active='A' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupp
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 270
                '  .Rows.Band.Columns(1).Width = 160


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Load_Grid_Sup2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select Status as [##],T12PO_No as [PO No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where T12Loc_Code='" & _Comcode & "' order by Supplier desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 70
            UltraGrid3.Rows.Band.Columns(2).Width = 70
            UltraGrid3.Rows.Band.Columns(3).Width = 170
            UltraGrid3.Rows.Band.Columns(4).Width = 90
            UltraGrid3.Rows.Band.Columns(5).Width = 70
            UltraGrid3.Rows.Band.Columns(6).Width = 90
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Grid_Sup1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select Status as [##],T12PO_No as [PO No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where T12Loc_Code='" & _Comcode & "' order by Supplier "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 70
            UltraGrid3.Rows.Band.Columns(2).Width = 70
            UltraGrid3.Rows.Band.Columns(3).Width = 170
            UltraGrid3.Rows.Band.Columns(4).Width = 90
            UltraGrid3.Rows.Band.Columns(5).Width = 70
            UltraGrid3.Rows.Band.Columns(6).Width = 90
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select Status as [##],T12PO_No as [PO No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where T12Loc_Code='" & _Comcode & "' order by Date  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 70
            UltraGrid3.Rows.Band.Columns(2).Width = 70
            UltraGrid3.Rows.Band.Columns(3).Width = 170
            UltraGrid3.Rows.Band.Columns(4).Width = 90
            UltraGrid3.Rows.Band.Columns(5).Width = 70
            UltraGrid3.Rows.Band.Columns(6).Width = 90
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select Status as [##],T12PO_No as [PO No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where T12Loc_Code='" & _Comcode & "' order by Date desc "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 70
            UltraGrid3.Rows.Band.Columns(2).Width = 70
            UltraGrid3.Rows.Band.Columns(3).Width = 170
            UltraGrid3.Rows.Band.Columns(4).Width = 90
            UltraGrid3.Rows.Band.Columns(5).Width = 70
            UltraGrid3.Rows.Band.Columns(6).Width = 90
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select Status as [##],T12PO_No as [PO No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where T12Loc_Code='" & _Comcode & "' order by T12PO_No desc "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 70
            UltraGrid3.Rows.Band.Columns(2).Width = 70
            UltraGrid3.Rows.Band.Columns(3).Width = 170
            UltraGrid3.Rows.Band.Columns(4).Width = 90
            UltraGrid3.Rows.Band.Columns(5).Width = 70
            UltraGrid3.Rows.Band.Columns(6).Width = 90
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub AZToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AZToolStripMenuItem.Click
        Call Load_Grid1()
        Panel1.Visible = False
        cboSupp.Text = ""
    End Sub

    Private Sub ZAToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZAToolStripMenuItem.Click
        Call Load_Grid2()
        Panel1.Visible = False
        cboSupp.Text = ""
    End Sub

    Private Sub AZToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AZToolStripMenuItem1.Click
        Call Load_Grid_Sup1()
        Panel1.Visible = False
        cboSupp.Text = ""
    End Sub

    Private Sub ZAToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZAToolStripMenuItem1.Click
        Call Load_Grid_Sup2()
        Panel1.Visible = False
        cboSupp.Text = ""
    End Sub

    Private Sub ByPONoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByPONoToolStripMenuItem.Click
        Call Load_Grid()
        Panel1.Visible = False
        cboSupp.Text = ""
    End Sub



    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _Rowcount As Integer

        _Rowcount = UltraGrid3.ActiveRow.Index
        frmPO.txtEntry.Text = UltraGrid3.Rows(_Rowcount).Cells(1).Text
        frmPO.Search_Records()
        Me.Close()
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        cboSupp.Text = ""
        Panel1.Visible = False
        txtFrom.Text = Today
        txtTo.Text = Today
    End Sub

    Private Sub FindBySupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindBySupplierToolStripMenuItem.Click
        Panel1.Visible = True
        cboSupp.ToggleDropdown()
    End Sub

    Function Load_Grid_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select Status as [##],T12PO_No as [PO No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where Supplier='" & Trim(cboSupp.Text) & "' and date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and T12Loc_Code='" & _Comcode & "' order by T12PO_No desc "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 70
            UltraGrid3.Rows.Band.Columns(2).Width = 70
            UltraGrid3.Rows.Band.Columns(3).Width = 170
            UltraGrid3.Rows.Band.Columns(4).Width = 90
            UltraGrid3.Rows.Band.Columns(5).Width = 70
            UltraGrid3.Rows.Band.Columns(6).Width = 90
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Load_Grid_Supplier()
        Panel1.Visible = False
    End Sub
End Class