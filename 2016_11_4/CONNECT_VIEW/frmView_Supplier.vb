Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmView_Supplier
    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim i As Integer
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M04Code  ) as  ##,M04Code as [Supplier Code],M04Name as [Supplier Name],case when M04Status='A'then 'Active' else 'Deactive' end as [Active Status]  from M04Supplier order by M04ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 280
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(3).Text) = "Deactive" Then
                    UltraGrid2.Rows(i).Cells(0).Appearance.BackColor = Color.Red
                    UltraGrid2.Rows(i).Cells(1).Appearance.BackColor = Color.Red
                    UltraGrid2.Rows(i).Cells(2).Appearance.BackColor = Color.Red
                    UltraGrid2.Rows(i).Cells(3).Appearance.BackColor = Color.Red
                End If
                i = i + 1
            Next
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
        frmSupplier_Cnt.txtCode.Text = Trim(UltraGrid2.Rows(_Row).Cells(1).Text)
        frmSupplier_Cnt.Search_Records()
        Me.Close()
    End Sub

    Private Sub frmView_Supplier_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid()
    End Sub
End Class