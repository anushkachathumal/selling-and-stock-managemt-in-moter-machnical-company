Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptItems
    Dim _Print_Status As String
    Dim _cATEGORY As String

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub


    Function Load_Grid_PRODUCT()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05ID ) as  ##,M01Description as [Category],M05Item_Code as [iItem Code],tmpDescription as [Item Name],M05Use_For as [Use For]  from View_Product_Item where  M05Status='A' order by M05ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 110
            UltraGrid2.Rows.Band.Columns(3).Width = 280
            UltraGrid2.Rows.Band.Columns(4).Width = 320
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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

    Function Load_Grid_PRODUCT_CATEGORY()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M05ID ) as  ##,M01Description as [Category],M05Item_Code as [iItem Code],tmpDescription as [Item Name],M05Use_For as [Use For]  from View_Product_Item where  M05Status='A' AND M01Description='" & Trim(cboCategory.Text) & "' order by M05ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 30
            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 110
            UltraGrid2.Rows.Band.Columns(3).Width = 280
            UltraGrid2.Rows.Band.Columns(4).Width = 320
            ' UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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


    Private Sub frmrptItems_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid_PRODUCT()
        _Print_Status = "A"
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Grid_PRODUCT()
        _Print_Status = "A"
        Panel1.Visible = False
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Dim B As New ReportDocument
        Dim A As String
        Try
            If _Print_Status = "A" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Items.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                'B.SetParameterValue("Total_cost", _Total_Cost)
                'B.SetParameterValue("Total_Rate", _Total_Rate)
                ''  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Product_Item.M05Status} ='A' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Print_Status = "B1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Items.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Product_Item.M05Status} ='A' AND {View_Product_Item.M01Description} ='" & _cATEGORY & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
                'ElseIf _Print_Status = "B1" Then
                '    A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                '    B.Load(A.ToString)
                '    B.SetDatabaseLogon("sa", "sainfinity")
                '    'B.SetParameterValue("To", _To)
                '    'B.SetParameterValue("From", _From)
                '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '    frmReport.CrystalReportViewer1.DisplayToolbar = True
                '    'frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {M01Category.M01Description}='" & _Dis & "' "
                '    frmReport.Refresh()
                '    ' frmReport.CrystalReportViewer1.PrintReport()
                '    ' B.PrintToPrinter(1, True, 0, 0)
                '    frmReport.MdiParent = MDIMain
                '    frmReport.Show()
                'ElseIf _Print_Status = "B2" Then
                '    A = ConfigurationManager.AppSettings("ReportPath") + "\Stock1.rpt"
                '    B.Load(A.ToString)
                '    B.SetDatabaseLogon("sa", "sainfinity")
                '    B.SetParameterValue("Total_cost", _Total_Cost)
                '    B.SetParameterValue("Total_Rate", _Total_Rate)
                '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '    frmReport.CrystalReportViewer1.DisplayToolbar = True
                '    frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {M01Category.M01Description}='" & _Dis & "' "
                '    frmReport.Refresh()
                '    ' frmReport.CrystalReportViewer1.PrintReport()
                '    ' B.PrintToPrinter(1, True, 0, 0)
                '    frmReport.MdiParent = MDIMain
                '    frmReport.Show()
                'ElseIf _Print_Status = "C1" Then
                '    A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                '    B.Load(A.ToString)
                '    B.SetDatabaseLogon("sa", "sainfinity")
                '    'B.SetParameterValue("To", _To)
                '    'B.SetParameterValue("From", _From)
                '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '    frmReport.CrystalReportViewer1.DisplayToolbar = True
                '    frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {M05Item_Master.M05Brand_Name}='" & _Dis & "'"
                '    frmReport.Refresh()
                '    ' frmReport.CrystalReportViewer1.PrintReport()
                '    ' B.PrintToPrinter(1, True, 0, 0)
                '    frmReport.MdiParent = MDIMain
                '    frmReport.Show()
                'ElseIf _Print_Status = "C2" Then
                '    A = ConfigurationManager.AppSettings("ReportPath") + "\Stock1.rpt"
                '    B.Load(A.ToString)
                '    B.SetDatabaseLogon("sa", "sainfinity")
                '    B.SetParameterValue("Total_cost", _Total_Cost)
                '    B.SetParameterValue("Total_Rate", _Total_Rate)
                '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '    frmReport.CrystalReportViewer1.DisplayToolbar = True
                '    frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {M05Item_Master.M05Brand_Name}='" & _Dis & "' "
                '    frmReport.Refresh()
                '    ' frmReport.CrystalReportViewer1.PrintReport()
                '    ' B.PrintToPrinter(1, True, 0, 0)
                '    frmReport.MdiParent = MDIMain
                '    frmReport.Show()
                'ElseIf _Print_Status = "D" Then
                '    Call Save_Stock_Movement()
                '    A = ConfigurationManager.AppSettings("ReportPath") + "\Srock_Movement.rpt"
                '    B.Load(A.ToString)
                '    B.SetDatabaseLogon("sa", "sainfinity")
                '    B.SetParameterValue("From", _From)
                '    B.SetParameterValue("To", _To)
                '    B.SetParameterValue("OB", "O/B on " & Year(_From) & "/" & Month(_From) & "/" & Microsoft.VisualBasic.Day(_From))
                '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '    frmReport.CrystalReportViewer1.DisplayToolbar = True
                '    frmReport.CrystalReportViewer1.SelectionFormula = "{View_Product_Item.M05Status} ='A' "
                '    frmReport.Refresh()
                '    ' frmReport.CrystalReportViewer1.PrintReport()
                '    ' B.PrintToPrinter(1, True, 0, 0)
                '    frmReport.MdiParent = MDIMain
                '    frmReport.Show()
                'ElseIf _Print_Status = "D1" Then
                '    Call Save_Stock_Movement()
                '    A = ConfigurationManager.AppSettings("ReportPath") + "\Srock_Movement.rpt"
                '    B.Load(A.ToString)
                '    B.SetDatabaseLogon("sa", "sainfinity")
                '    B.SetParameterValue("From", _From)
                '    B.SetParameterValue("To", _To)
                '    B.SetParameterValue("OB", "O/B on " & Year(_From) & "/" & Month(_From) & "/" & Microsoft.VisualBasic.Day(_From))
                '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '    frmReport.CrystalReportViewer1.DisplayToolbar = True
                '    frmReport.CrystalReportViewer1.SelectionFormula = "{View_Product_Item.M01Description}='" & _Dis & "' AND {View_Product_Item.M05Status}='A'"
                '    frmReport.Refresh()
                '    ' frmReport.CrystalReportViewer1.PrintReport()
                '    ' B.PrintToPrinter(1, True, 0, 0)
                '    frmReport.MdiParent = MDIMain
                '    frmReport.Show()
                'ElseIf _Print_Status = "X1" Then
                '    A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                '    B.Load(A.ToString)
                '    B.SetDatabaseLogon("sa", "sainfinity")
                '    'B.SetParameterValue("To", _To)
                '    'B.SetParameterValue("From", _From)
                '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                '    frmReport.CrystalReportViewer1.DisplayToolbar = True
                '    frmReport.CrystalReportViewer1.SelectionFormula = "{M05Item_Master.M05Status} ='A' and {View_Stock_Balance.Qty} < 0"
                '    frmReport.Refresh()
                '    ' frmReport.CrystalReportViewer1.PrintReport()
                '    ' B.PrintToPrinter(1, True, 0, 0)
                '    frmReport.MdiParent = MDIMain
                '    frmReport.Show()
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)


            End If
        End Try
    End Sub

    Private Sub CategoryOnlyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CategoryOnlyToolStripMenuItem.Click
        _Print_Status = "B1"
        Panel1.Visible = True
        Call Load_Category()
        cboCategory.Text = ""
        cboCategory.ToggleDropdown()
    End Sub

    Function Load_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M01Description as [##] from M01Category WHERE M01Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275

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
        Call Load_Grid_PRODUCT_CATEGORY()
        _cATEGORY = Trim(cboCategory.Text)
        Panel1.Visible = False
    End Sub
End Class