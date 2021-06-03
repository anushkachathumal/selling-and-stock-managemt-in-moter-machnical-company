Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration

Public Class frmView_Sales
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim _Comcode As String
    Dim _Acctype As String
    Dim _CusCode As String
    Dim _Sales_Invoice As String
    Dim _Transaction_Ref As Integer

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select T01Ref_No as [Ref.No],t01date as [Date],T01PO_NO as [Loading No],T01Invoice_No as [Sys.No],T01Grn_No as [Company Inv.No],M17Name as [Customer Name],CAST(T01Net_Amount +T01Com_Discount AS DECIMAL(16,2)) as [Net Amount],CAST(T01Com_Discount AS DECIMAL(16,2)) as [Discount],CAST(T01Net_Amount  AS DECIMAL(16,2)) as [Gross Amount],T01User as [User] from T01Transaction_Header  inner join M17Customer on T01Customer=M17Code where T01Date='" & Today & "' and T01Trans_Type='DR' and  T01status='A' and T01Com_Code='" & _Comcode & "'  order by T01Ref_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 60
            ' UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 90
            UltraGrid1.Rows.Band.Columns(3).Width = 90
            UltraGrid1.Rows.Band.Columns(4).Width = 90
            UltraGrid1.Rows.Band.Columns(5).Width = 170
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 90
            UltraGrid1.Rows.Band.Columns(8).Width = 90

            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Panal4()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            ' Sql = "select T01Invoice_No as [Invoice No],M17Name as [Customer Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from T01Transaction_Header  inner join M17Customer on T01Customer=M17Code where T01Date between '" & txtD1.Text & "' and '" & txtD2.Text & "' and T01Trans_Type='DR' order by T01Ref_No desc"
            Sql = "select T01Ref_No as [Ref.No],t01date as [Date],T01PO_NO as [Loading No],T01Invoice_No as [Sys.No],T01Grn_No as [Company Inv.No],M17Name as [Customer Name],CAST(Net_Amount  AS DECIMAL(16,2)) as [Net Amount],CAST(Discount AS DECIMAL(16,2)) as [Discount],CAST(Gross_Amount  AS DECIMAL(16,2)) as [Gross Amount],T01User as [User] from View_Sales_new where T01Date between '" & txtD1.Text & "'  and '" & txtD2.Text & "' and T01Trans_Type='DR' and T01Com_Code='" & _Comcode & "'  order by T01Ref_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 60
            ' UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 90
            UltraGrid1.Rows.Band.Columns(3).Width = 90
            UltraGrid1.Rows.Band.Columns(4).Width = 90
            UltraGrid1.Rows.Band.Columns(5).Width = 170
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 90
            UltraGrid1.Rows.Band.Columns(8).Width = 90

            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Panal1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            'Sql = "select T01Invoice_No as [Invoice No],M17Name as [Customer Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from T01Transaction_Header  inner join M17Customer on T01Customer=M17Code where T01Date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and T01Trans_Type='DR' and M17Name='" & Trim(cboCustomer.Text) & "' order by T01Ref_No desc"
            'dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            Sql = "select T01Ref_No as [Ref.No],t01date as [Date],T01PO_NO as [Loading No],T01Invoice_No as [Sys.No],T01Grn_No as [Company Inv.No],M17Name as [Customer Name],CAST(T01Net_Amount +T01Com_Discount AS DECIMAL(16,2)) as [Net Amount],CAST(T01Com_Discount AS DECIMAL(16,2)) as [Discount],CAST(T01Net_Amount  AS DECIMAL(16,2)) as [Gross Amount],T01User as [User] from T01Transaction_Header  inner join M17Customer on T01Customer=M17Code where T01Date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and T01Trans_Type='DR' and M17Name='" & Trim(cboCustomer.Text) & "' and T01status='A' and T01Com_Code='" & _Comcode & "' order by T01Ref_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 60
            ' UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 90
            UltraGrid1.Rows.Band.Columns(3).Width = 90
            UltraGrid1.Rows.Band.Columns(4).Width = 90
            UltraGrid1.Rows.Band.Columns(5).Width = 170
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 90
            UltraGrid1.Rows.Band.Columns(8).Width = 90

            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Panal2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            ' Sql = "select T01Invoice_No as [Invoice No],M17Name as [Customer Name],CAST(T01Net_Amount AS DECIMAL(16,2)) as [Net Amount] from T01Transaction_Header  inner join M17Customer on T01Customer=M17Code where T01Date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and T01Trans_Type='RP' and T14Vehicle_No='" & Trim(cboVehicle.Text) & "' order by T01Ref_No desc"
            Sql = "select T01Ref_No as [Ref.No],t01date as [Date],T01PO_NO as [Loading No],T01Invoice_No as [Sys.No],T01Grn_No as [Company Inv.No],M17Name as [Customer Name],CAST(Net_Amount  AS DECIMAL(16,2)) as [Net Amount],CAST(Discount AS DECIMAL(16,2)) as [Discount],CAST(Gross_Amount  AS DECIMAL(16,2)) as [Gross Amount],T01User as [User] from View_Sales_new where T01Date between '" & txtB1.Text & "'  and '" & txtB2.Text & "' and T01Trans_Type='DR' and T20Vehicle_No='" & Trim(cboVehicle.Text) & "'  order by T01Ref_No desc"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 60
            ' UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 90
            UltraGrid1.Rows.Band.Columns(3).Width = 90
            UltraGrid1.Rows.Band.Columns(4).Width = 90
            UltraGrid1.Rows.Band.Columns(5).Width = 170
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 90
            UltraGrid1.Rows.Band.Columns(8).Width = 90

            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub frmView_Sales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid()
        txtD1.Text = Today
        txtD2.Text = Today
        txtC1.Text = Today
        txtC2.Text = Today
        Call Load_Customer()
        Call Load_Vehicle()

    End Sub

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim I As Integer
        I = UltraGrid1.ActiveRow.Index
        frmRapi_Invoice.txtEntry.Text = UltraGrid1.Rows(I).Cells(0).Text
        strSales_Status = True
        Me.Close()
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Panel4.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
    End Sub

    Private Sub UsingSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Panel4.Visible = False
        Panel1.Visible = False
        Panel2.Visible = True
    End Sub

    Private Sub UsingCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingCategoryToolStripMenuItem.Click
        Panel4.Visible = False
        Panel1.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Grid()
        Panel4.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Call Load_Grid_Panal4()
        Panel4.Visible = False
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Grid_Panal1()
        Panel1.Visible = False
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub
    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M17Name as [##] from M17Customer where M17Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 220
                ' .Rows.Band.Columns(1).Width = 180


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
            End If
        End Try
    End Function

    Function Load_Vehicle()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M31Vehicle_No as [##] from M31Vehicle_Master  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboVehicle
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 220
                ' .Rows.Band.Columns(1).Width = 180


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
            End If
        End Try
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Grid_Panal2()
        Panel2.Visible = False
    End Sub
End Class