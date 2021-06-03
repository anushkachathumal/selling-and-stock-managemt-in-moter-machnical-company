Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptPO
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Comcode As String

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub frmrptPO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Grid()
        _PrintStatus = "A1"
        txtDate5.Text = Today
        txtDate6.Text = Today
        Call Load_Supplier()
        Call Load_Item()
    End Sub


    Function Load_Gride_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name] from M03Item_Master where M03Status='A' and M03Location='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 130
            UltraGrid3.Rows.Band.Columns(1).Width = 370
            ' UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where date between '" & txtDate5.Text & "' and '" & txtDate6.Text & "' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA3()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where date between '" & txtDate5.Text & "' and '" & txtDate6.Text & "' and Status='Approved' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA4()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where date between '" & txtDate5.Text & "' and '" & txtDate6.Text & "' and Status='Pending' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA5()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where date between '" & txtDate5.Text & "' and '" & txtDate6.Text & "' and Status='Reject' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String

        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"

            If _PrintStatus = "A1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.T12Loc_Code}='" & _Comcode & "'  "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_PO.T12Loc_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A3" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_PO.Status} = 'Approved' and {View_PO.T12Loc_Code} ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A4" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_PO.Status} = 'Pending' and {View_PO.T12Loc_Code} ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A5" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_PO.Status} = 'Reject' and {View_PO.T12Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A6" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_PO.supplier} ='" & Trim(cboSupplier.Text) & "' and {View_PO.T12Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A7" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_PO.supplier} ='" & Trim(cboSupplier.Text) & "' and  {View_PO.Status} = 'Approved' and {View_PO.T12Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A8" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_PO.supplier} ='" & Trim(cboSupplier.Text) & "' and  {View_PO.Status} = 'Reject' and {View_PO.T12Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A9" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_PO.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_PO.supplier} ='" & Trim(cboSupplier.Text) & "' and  {View_PO.Status} = 'Pending' and {View_PO.T12Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A10" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport3.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_POReport.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_POReport.M03Item_Code}  ='" & Trim(cboItem.Text) & "' and {View_PO.T12Loc_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A11" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport3.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_POReport.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_POReport.M03Item_Code}  ='" & Trim(cboItem.Text) & "' and {View_POReport.Status} = 'Approved' and {View_PO.T12Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()

            ElseIf _PrintStatus = "A12" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\POreport3.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_POReport.Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_POReport.M03Item_Code}  ='" & Trim(cboItem.Text) & "' and {View_POReport.Status} = 'Pending' and {View_PO.T12Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub


    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        _PrintStatus = "A2"
        txtDate5.Text = Today
        txtDate6.Text = Today
        Panel3.Visible = True
        Panel2.Visible = False
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If _PrintStatus = "A2" Then
            _From = txtDate5.Text
            _To = txtDate6.Text
            Call Load_GridA2()
            Panel3.Visible = False
        ElseIf _PrintStatus = "A3" Then
            _From = txtDate5.Text
            _To = txtDate6.Text
            Call Load_GridA3()
            Panel3.Visible = False
        ElseIf _PrintStatus = "A4" Then
            _From = txtDate5.Text
            _To = txtDate6.Text
            Call Load_GridA4()
            Panel3.Visible = False
        ElseIf _PrintStatus = "A5" Then
            _From = txtDate5.Text
            _To = txtDate6.Text
            Call Load_GridA5()
            Panel3.Visible = False
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Grid()
        Panel3.Visible = False
        txtDate5.Text = Today
        txtDate6.Text = Today
        Panel2.Visible = False
        Panel3.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub UsingSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingSupplierToolStripMenuItem.Click
        _PrintStatus = "A3"
        txtDate5.Text = Today
        txtDate6.Text = Today
        Panel3.Visible = True
        Panel2.Visible = False
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub UsingCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingCategoryToolStripMenuItem.Click
        _PrintStatus = "A4"
        txtDate5.Text = Today
        txtDate6.Text = Today
        Panel3.Visible = True
        Panel2.Visible = False
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub UsingItemNameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingItemNameToolStripMenuItem.Click
        _PrintStatus = "A5"
        txtDate5.Text = Today
        txtDate6.Text = Today
        Panel3.Visible = True
        Panel2.Visible = False
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub AllPOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllPOToolStripMenuItem.Click
        _PrintStatus = "A6"
        txtDate3.Text = Today
        txtDate4.Text = Today
        Panel3.Visible = False
        Panel2.Visible = True
        cboSupplier.Text = ""
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub


    Function Load_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Code as [##] from M03Item_Master where m03Status='A' and M03Location='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
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


    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Active='A' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupplier
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

    Function Load_GridA6()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and Supplier='" & Trim(cboSupplier.Text) & "' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA7()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and Supplier='" & Trim(cboSupplier.Text) & "' and Status='Approved' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA8()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and Supplier='" & Trim(cboSupplier.Text) & "' and Status='Reject' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA9()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],app as [Approved by],Requser as [User],CAST(Total AS DECIMAL(16,2)) as [Net Amount] from View_PO where date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and Supplier='" & Trim(cboSupplier.Text) & "' and Status='Pending' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA10()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = " select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],m03item_Name as [Item Name],CAST(t13Rate AS DECIMAL(16,2)) as [Rate],CAST(t13Qty AS DECIMAL(16,2)) as [Qty],CAST(t13Rate*t13qty AS DECIMAL(16,2)) as [Total] from View_PO inner join View_PO_Fluter on T13ref_no=t12ref_no where date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M03Item_Code='" & Trim(cboItem.Text) & "' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 150
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 70
            UltraGrid1.Rows.Band.Columns(8).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(8).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA11()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = " select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],m03item_Name as [Item Name],CAST(t13Rate AS DECIMAL(16,2)) as [Rate],CAST(t13Qty AS DECIMAL(16,2)) as [Qty],CAST(t13Rate*t13qty AS DECIMAL(16,2)) as [Total] from View_PO inner join View_PO_Fluter on T13ref_no=t12ref_no where date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M03Item_Code='" & Trim(cboItem.Text) & "' and Status='Approved' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 150
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 70
            UltraGrid1.Rows.Band.Columns(8).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(8).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_GridA12()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = " select Status as [##],T12PO_No as [PO No],GRN as [GRN No],date as [Req Date],Supplier as [Supplier Name],m03item_Name as [Item Name],CAST(t13Rate AS DECIMAL(16,2)) as [Rate],CAST(t13Qty AS DECIMAL(16,2)) as [Qty],CAST(t13Rate*t13qty AS DECIMAL(16,2)) as [Total] from View_PO inner join View_PO_Fluter on T13ref_no=t12ref_no where date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M03Item_Code='" & Trim(cboItem.Text) & "' and Status='Pending' and T12Loc_Code='" & _Comcode & "' order by T12PO_No  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 70
            UltraGrid1.Rows.Band.Columns(2).Width = 70
            UltraGrid1.Rows.Band.Columns(4).Width = 170
            UltraGrid1.Rows.Band.Columns(5).Width = 150
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 70
            UltraGrid1.Rows.Band.Columns(8).Width = 110
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            UltraGrid1.Rows.Band.Columns(0).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(1).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(2).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(3).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(4).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(5).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(6).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(7).CellActivation = Activation.NoEdit
            UltraGrid1.Rows.Band.Columns(8).CellActivation = Activation.NoEdit
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _PrintStatus = "A6" Then
            _From = txtDate3.Text
            _To = txtDate4.Text
            Call Load_GridA6()
            Panel2.Visible = False
        ElseIf _PrintStatus = "A7" Then
            _From = txtDate3.Text
            _To = txtDate4.Text
            Call Load_GridA7()
            Panel2.Visible = False
        ElseIf _PrintStatus = "A8" Then
            _From = txtDate3.Text
            _To = txtDate4.Text
            Call Load_GridA8()
            Panel2.Visible = False
        ElseIf _PrintStatus = "A9" Then
            _From = txtDate3.Text
            _To = txtDate4.Text
            Call Load_GridA9()
            Panel2.Visible = False
        End If
    End Sub

    Private Sub ApprovedPOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApprovedPOToolStripMenuItem.Click
        _PrintStatus = "A7"
        txtDate3.Text = Today
        txtDate4.Text = Today
        Panel3.Visible = False
        Panel2.Visible = True
        cboSupplier.Text = ""
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub RejectPOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RejectPOToolStripMenuItem.Click
        _PrintStatus = "A8"
        txtDate3.Text = Today
        txtDate4.Text = Today
        Panel3.Visible = False
        Panel2.Visible = True
        cboSupplier.Text = ""
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub PendingPOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PendingPOToolStripMenuItem.Click
        _PrintStatus = "A9"
        txtDate3.Text = Today
        txtDate4.Text = Today
        Panel3.Visible = False
        Panel2.Visible = True
        cboSupplier.Text = ""
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub AllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllToolStripMenuItem.Click
        _PrintStatus = "A10"
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        Panel2.Visible = False
        cboItem.Text = ""
        Panel1.Visible = True
    End Sub



    Private Sub cboItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItem.KeyUp
        If e.KeyCode = Keys.F1 Then
            Call Load_Gride_Item()
            OPR5.Visible = True
            txtFind.Text = ""
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
        End If
    End Sub




    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _PrintStatus = "A10" Then
            _From = txtDate1.Text
            _To = txtDate2.Text
            Call Load_GridA10()
            Panel1.Visible = False
        ElseIf _PrintStatus = "A11" Then
            _From = txtDate1.Text
            _To = txtDate2.Text
            Call Load_GridA11()
            Panel1.Visible = False
        ElseIf _PrintStatus = "A12" Then
            _From = txtDate1.Text
            _To = txtDate2.Text
            Call Load_GridA12()
            Panel1.Visible = False
        End If
    End Sub





    Private Sub UltraGrid3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid3.KeyUp
        On Error Resume Next
        If e.KeyCode = 13 Then

            Dim _Rowindex As Integer
            _Rowindex = UltraGrid3.ActiveRow.Index
            cboItem.Text = UltraGrid3.Rows(_Rowindex).Cells(1).Text
            OPR5.Visible = False
        End If
    End Sub

    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer
        _Rowindex = UltraGrid3.ActiveRow.Index
        cboItem.Text = UltraGrid3.Rows(_Rowindex).Cells(1).Text
        OPR5.Visible = False
    End Sub

    Function Load_Gride_Item3()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Item_Name  like '%" & txtFind.Text & "%' and M03Status='A' and M03Location='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 130
            UltraGrid3.Rows.Band.Columns(1).Width = 370
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtFind_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFind.ValueChanged
        Call Load_Gride_Item3()
    End Sub

    Private Sub ApprovedPOToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApprovedPOToolStripMenuItem1.Click
        _PrintStatus = "A11"
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        Panel2.Visible = False
        cboItem.Text = ""
        Panel1.Visible = True
    End Sub

    Private Sub PendingPOToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PendingPOToolStripMenuItem1.Click
        _PrintStatus = "A12"
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        Panel2.Visible = False
        cboItem.Text = ""
        Panel1.Visible = True
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class