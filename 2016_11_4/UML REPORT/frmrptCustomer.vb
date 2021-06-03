Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptCustomer
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Gride_Customer()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_rptCustomer
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 260
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 160
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 110
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            
        End With
    End Function

    Private Sub frmrptCustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride_Customer()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Load_Gride_Customer()
    End Sub


    Private Sub ActiveCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ActiveCustomerToolStripMenuItem.Click
        _PrintStatus = "A"
        Call Load_Gride_Customer()
        Call Load_Date1()

    End Sub

    Function Load_Date1()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Try
            Sql = "select * from M17Customer where M17Active='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Cus No") = M01.Tables(0).Rows(i)("M17Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17Name")
                '  newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                newRow("Address") = M01.Tables(0).Rows(i)("M17Address")
                newRow("Contact No") = M01.Tables(0).Rows(i)("M17TP")
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Date2()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Try
            Sql = "select * from M17Customer where M17Active='I'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Cus No") = M01.Tables(0).Rows(i)("M17Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17Name")
                '  newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                newRow("Address") = M01.Tables(0).Rows(i)("M17Address")
                newRow("Contact No") = M01.Tables(0).Rows(i)("M17TP")
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub InactiveCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InactiveCustomerToolStripMenuItem.Click
        _PrintStatus = "B"
        Call Load_Gride_Customer()
        Call Load_Date2()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Try
            If _PrintStatus = "A" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\rptCustomer.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M17Customer.M17Active}='A'  "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "B" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\rptCustomer.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{M17Customer.M17Active}='I'  "
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
End Class