Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmViewMarkert
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _PrintStatus As String
    Dim _Comcode As String
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Grid1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select T01GRN_No as [Return No],T01Date as [Date],M01Acc_Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from View_MkReturnnw where T01Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and T01Com_Code='" & _Comcode & "' order by T01Grn_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 90
            UltraGrid3.Rows.Band.Columns(2).Width = 170
            UltraGrid3.Rows.Band.Columns(3).Width = 90
           
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            UltraGrid3.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            'UltraGrid3.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            'UltraGrid3.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            'UltraGrid3.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            'UltraGrid3.Rows.Band.Columns(7).CellActivation = Activation.NoEdit

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
            Sql = "select T01GRN_No as [Return No],T01Date as [Date],M01Acc_Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from View_MkReturnnw where  T01Com_Code='" & _Comcode & "' order by T01Grn_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 90
            UltraGrid3.Rows.Band.Columns(2).Width = 170
            UltraGrid3.Rows.Band.Columns(3).Width = 90

            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            UltraGrid3.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(3).CellActivation = Activation.NoEdit

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
            Sql = "select T01GRN_No as [Return No],T01Date as [Date],M01Acc_Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from View_MkReturnnw where T01Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and M01Acc_Name='" & cboSupp.Text & "' and T01Com_Code='" & _Comcode & "' order by T01Grn_No "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 90
            UltraGrid3.Rows.Band.Columns(2).Width = 170
            UltraGrid3.Rows.Band.Columns(3).Width = 90

            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            UltraGrid3.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid3()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select T01GRN_No as [Return No],T01Date as [Date],M01Acc_Name as [Supplier Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from View_MkReturnnw where T01Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and M01Acc_Name='" & cboSupp.Text & "' and T01Com_Code='" & _Comcode & "' order by T01Grn_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 90
            UltraGrid3.Rows.Band.Columns(2).Width = 170
            UltraGrid3.Rows.Band.Columns(3).Width = 90

            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            UltraGrid3.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(3).CellActivation = Activation.NoEdit

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub frmViewMarkert_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Grid()
        txtFrom.Text = Today
        txtTo.Text = Today
        txtDate1.Text = Today
        txtDate2.Text = Today
        Call Load_Combo()
    End Sub


    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        Panel1.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Load_Grid1()
        Panel1.Visible = False
        txtFrom.Text = Today
        txtTo.Text = Today
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

    Private Sub AZToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AZToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = True
        cboSupp.ToggleDropdown()
        _PrintStatus = "A1"
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If Trim(cboSupp.Text) <> "" Then
            If _PrintStatus = "A1" Then
                Call Load_Grid2()
            ElseIf _PrintStatus = "A2" Then
                Call Load_Grid3()
            ElseIf _PrintStatus = "A3" Then
                ' Call Load_GridVAT()
            End If

            txtDate1.Text = Today
            txtDate2.Text = Today
            Panel2.Visible = False
            cboSupp.Text = ""
        Else
            MsgBox("Please select the supplier", MsgBoxStyle.Information, "Information ......")
        End If
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Call Load_Grid()
    End Sub

    Private Sub ZAToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZAToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = True
        cboSupp.ToggleDropdown()
        _PrintStatus = "A2"
    End Sub

    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer

        _Rowindex = UltraGrid3.ActiveRow.Index
        With frmMK_Return
            .txtEntry.Text = UltraGrid3.Rows(_Rowindex).Cells(0).Text
            .Search_RecordsUsing_Entry()
        End With
        Me.Close()
    End Sub

    Private Sub UltraGrid3_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid3.InitializeLayout

    End Sub
End Class