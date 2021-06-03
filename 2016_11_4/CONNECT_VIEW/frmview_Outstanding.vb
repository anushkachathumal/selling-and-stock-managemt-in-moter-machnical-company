Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmview_Outstanding
    Dim _Print_Status As String

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub


    Function Load_Grid_date(ByVal strStatus As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T16Pay_No ) as  ##,MAX(T16Date) as [Date],T16Pay_No as [Pay.No],MAX(M06Name) as [Customer Name],CAST(MAX(T16Net_Amount) AS DECIMAL(16,2)) as [Pay Amount]    from T16Outstanding_Pay_Summery inner join T15Outstanding_Collection on T15Pay_No=T15Pay_No inner join M06Customer_Master on M06Code=T15Cus_Code where T16Status='" & strStatus & "' and T16Date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' GROUP BY T16Pay_No "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 270
            UltraGrid2.Rows.Band.Columns(4).Width = 110

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Function Load_Grid_Customer(ByVal strStatus As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T16Pay_No ) as  ##,MAX(T16Date) as [Date],T16Pay_No as [Pay.No],MAX(M06Name) as [Customer Name],CAST(MAX(T16Net_Amount) AS DECIMAL(16,2)) as [Pay Amount]    from T16Outstanding_Pay_Summery inner join T15Outstanding_Collection on T15Pay_No=T15Pay_No inner join M06Customer_Master on M06Code=T15Cus_Code where T16Status='" & strStatus & "' and M06Name='" & Trim(cboCustomer.Text) & "' GROUP BY T16Pay_No "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 270
            UltraGrid2.Rows.Band.Columns(4).Width = 110

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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


    Function Load_Grid(ByVal strStatus As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T16Pay_No ) as  ##,MAX(T16Date) as [Date],T16Pay_No as [Pay.No],MAX(M06Name) as [Customer Name],CAST(MAX(T16Net_Amount) AS DECIMAL(16,2)) as [Pay Amount]    from T16Outstanding_Pay_Summery inner join T15Outstanding_Collection on T15Pay_No=T15Pay_No inner join M06Customer_Master on M06Code=T15Cus_Code where T16Status='" & strStatus & "' GROUP BY T16Pay_No "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 270
            UltraGrid2.Rows.Band.Columns(4).Width = 110
            
            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ' UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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
                .Rows.Band.Columns(0).Width = 358
                '.Rows.Band.Columns(1).Width = 310


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

    Private Sub frmview_Outstanding_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid("A")
        Call Load_Customer()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Grid("A")
        Panel1.Visible = False
        Panel2.Visible = False
    End Sub

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        Panel1.Visible = True
        Panel2.Visible = False
        txtDate1.Text = Today
        txtDate2.Text = Today
        _Print_Status = "A"

    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _Print_Status = "A" Then
            Call Load_Grid_date("A")
            Panel1.Visible = False
        End If
    End Sub

    Private Sub BySupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem.Click
        _Print_Status = "A1"
        Panel1.Visible = False
        Panel2.Visible = True
        cboCustomer.Text = ""
        Call Load_Customer()
        cboCustomer.ToggleDropdown()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _Print_Status = "A1" Then
            Call Load_Grid_Customer("A")
            Panel2.Visible = False
        End If
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _row As Integer

        _row = UltraGrid2.ActiveRow.Index
        With frmOutstanding_Collection
            .OPR10.Visible = True
            .txtRef1.Text = Trim(UltraGrid2.Rows(_row).Cells(2).Text)
            .Search_Records()
            .cmdRecord.Text = "Cancel"
        End With
        Me.Close()
    End Sub

   
End Class