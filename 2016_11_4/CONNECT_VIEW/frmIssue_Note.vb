Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmIssue_Note
    Dim _Print_Status As String
    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Load_Grid_Issue(ByVal strStatus As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T06Ref_No ) as  ##,T06Ref_No as [Ref.Doc.No],max(T06Date) as [Date],max(T06Department) as [Department],max(T06Job_No) as [Job No],max(T06V_no) as [Vehicle No],max(M06Name) as [Customer Name] ,SUM(T07Rate*T07Qty)-(SUM(T07Rate*T07Qty)*(SUM(T07Discount))/100) as [Total Amount] from T06Item_Issue_Header  inner join T07Item_Issue_Fluter on T06Ref_No=T07Ref_No inner join T05Job_Card on T06Job_No=T05Job_No inner join M06Customer_Master on M06Code=T06Cus_No where T06Status='" & strStatus & "' and T05Status='A' group by T06Ref_No   "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 90
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 220
            UltraGrid2.Rows.Band.Columns(7).Width = 110
            'UltraGrid2.Rows.Band.Columns(8).Width = 90

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid_Issue_Date(ByVal strStatus As String, ByVal STRStatus1 As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T06Ref_No ) as  ##,T06Ref_No as [Ref.Doc.No],max(T06Date) as [Date],max(T06Department) as [Department],max(T06Job_No) as [Job No],max(T06V_no) as [Vehicle No],max(M06Name) as [Customer Name] ,SUM(T07Rate*T07Qty)-(SUM(T07Rate*T07Qty)*(SUM(T07Discount))/100) as [Total Amount] from T06Item_Issue_Header  inner join T07Item_Issue_Fluter on T06Ref_No=T07Ref_No inner join T05Job_Card on T06Job_No=T05Job_No inner join M06Customer_Master on M06Code=T06Cus_No where T06Status='" & strStatus & "' and T06Date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and T05Status='" & STRStatus1 & "' group by T06Ref_No   "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 90
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 220
            UltraGrid2.Rows.Band.Columns(7).Width = 110

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            'UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid_Issue_vEHICLE(ByVal strStatus As String, ByVal strStatus1 As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try

            Sql = "select ROW_NUMBER() OVER (ORDER BY T06Ref_No ) as  ##,T06Ref_No as [Ref.Doc.No],max(T06Date) as [Date],max(T06Department) as [Department],max(T06Job_No) as [Job No],max(T06V_no) as [Vehicle No],max(M06Name) as [Customer Name] ,SUM(T07Rate*T07Qty)-(SUM(T07Rate*T07Qty)*(SUM(T07Discount))/100) as [Total Amount] from T06Item_Issue_Header  inner join T07Item_Issue_Fluter on T06Ref_No=T07Ref_No inner join T05Job_Card on T06Job_No=T05Job_No inner join M06Customer_Master on M06Code=T06Cus_No where T06Status='" & strStatus & "' and T06Date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and T05Status='" & strStatus1 & "' AND T06V_no='" & Trim(cboVNo.Text) & "' group by T06Ref_No   "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 80
            UltraGrid2.Rows.Band.Columns(2).Width = 80
            UltraGrid2.Rows.Band.Columns(3).Width = 110
            UltraGrid2.Rows.Band.Columns(4).Width = 90
            UltraGrid2.Rows.Band.Columns(5).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 90
            UltraGrid2.Rows.Band.Columns(6).Width = 220
            UltraGrid2.Rows.Band.Columns(7).Width = 110

            UltraGrid2.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            'UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmIssue_Note_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_Issue("ISSUE")
    End Sub

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        Panel1.Visible = True
        Panel3.Visible = False
        txtDate1.Text = Today
        txtDate2.Text = Today
        _Print_Status = "A1"
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _Print_Status = "A" Then
            Call Load_Grid_Issue_Date("ISSUE", "A")
            Panel1.Visible = False
        ElseIf _Print_Status = "B1" Then
            Call Load_Grid_Issue_Date("ISSUE", "CLOSE")
            Panel1.Visible = False
        ElseIf _Print_Status = "C1" Then
            Call Load_Grid_Issue_Date("CANCEL", "A")
            Panel1.Visible = False
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Panel1.Visible = False
    End Sub

    Private Sub BySupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem.Click
        Panel1.Visible = False
        Panel3.Visible = True
        txtB1.Text = Today
        txtB2.Text = Today
        _Print_Status = "A2"
        Call Load_VNO()
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
            End If
        End Try
    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If _Print_Status = "A2" Then
            Call Load_Grid_Issue_vEHICLE("ISSUE", "A")
            Panel3.Visible = False
        ElseIf _Print_Status = "B2" Then
            Call Load_Grid_Issue_vEHICLE("ISSUE", "CLOSE")
            Panel3.Visible = False
        ElseIf _Print_Status = "C2" Then
            Call Load_Grid_Issue_vEHICLE("CANCEL", "A")
            Panel3.Visible = False
        End If
    End Sub

    Private Sub ByDateToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem1.Click
        Panel1.Visible = True
        Panel3.Visible = False
        txtDate1.Text = Today
        txtDate2.Text = Today
        _Print_Status = "B1"
    End Sub

    Private Sub BySupplierToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem1.Click
        Panel1.Visible = False
        Panel3.Visible = True
        txtB1.Text = Today
        txtB2.Text = Today
        _Print_Status = "B2"
        Call Load_VNO()
    End Sub

    Private Sub ByDateToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem2.Click
        Panel1.Visible = True
        Panel3.Visible = False
        txtDate1.Text = Today
        txtDate2.Text = Today
        _Print_Status = "C1"
    End Sub

    Private Sub ByVehicleNoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByVehicleNoToolStripMenuItem.Click
        Panel1.Visible = False
        Panel3.Visible = True
        txtB1.Text = Today
        txtB2.Text = Today
        _Print_Status = "C2"
        Call Load_VNO()
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        On Error Resume Next
        Dim _rOW As Integer

        _rOW = UltraGrid2.ActiveRow.Index
        frmItem_Issue_Uniq.UltraTabControl1.Tabs(0).Selected = True
        frmItem_Issue_Uniq.txtEntry.Text = Trim(UltraGrid2.Rows(_rOW).Cells(1).Text)
        frmItem_Issue_Uniq.SEARCH_RECORDS()
        Me.Close()
    End Sub
End Class