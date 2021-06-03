Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmViewWastage
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
            Sql = "select T01GRN_No as [Wastage No],T01Date as [Date],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from View_Wastage_Header where T01Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and T01FromLoc_Code='" & _Comcode & "' order by T01Grn_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 90
            ' UltraGrid3.Rows.Band.Columns(2).Width = 170
            UltraGrid3.Rows.Band.Columns(2).Width = 90

            UltraGrid3.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            UltraGrid3.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            ' UltraGrid3.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
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
            Sql = "select T01GRN_No as [Wastage No],T01Date as [Date],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from View_Wastage_Header where T01FromLoc_Code='" & _Comcode & "' order by T01Grn_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 90
            UltraGrid3.Rows.Band.Columns(1).Width = 90
            UltraGrid3.Rows.Band.Columns(2).Width = 90
            ' UltraGrid3.Rows.Band.Columns(3).Width = 90

            UltraGrid3.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid3.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            UltraGrid3.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid3.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            ' UltraGrid3.Rows.Band.Columns(3).CellActivation = Activation.NoEdit

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub frmViewWastage_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Grid()
        txtFrom.Text = Today
        txtTo.Text = Today
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Load_Grid1()
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call Load_Grid()
    End Sub

    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer

        _Rowindex = UltraGrid3.ActiveRow.Index
        With frmWastage
            .txtEntry.Text = UltraGrid3.Rows(_Rowindex).Cells(0).Text
            .Search_RecordsUsing_Entry()
        End With
        Me.Close()
    End Sub
End Class