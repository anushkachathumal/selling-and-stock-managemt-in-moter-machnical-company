Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmPending_Orders
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        txtTodate.Text = Today
        txtFromDate.Text = Today

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String
        Dim _Todate As Date
        Try
            If Trim(strUGroup) = "MERCHANT" Then

                A = ConfigurationManager.AppSettings("ReportPath") + "\Pending_Status_Merchant.rpt"

                _Todate = CDate(txtFromDate.Text).AddDays(-365)

                StrFromDate = "(" & Year(_Todate) & ", " & VB6.Format(Month(_Todate), "0#") & ", " & VB6.Format(CDate(_Todate).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtFromDate.Value) & ", " & VB6.Format(Month(txtFromDate.Value), "0#") & ", " & VB6.Format(CDate(txtFromDate.Text).Day, "0#") & ", 00, 00, 00)"

                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True

                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Delivary_Request.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T01Delivary_Request.T01Status} = 'A' and {T01Delivary_Request.T01User} ='" & strDisname & "'"
                frmReport.MdiParent = MDIMain
                'frmReport.Show()
                frmReport.Show()

            ElseIf Trim(strUGroup) = "PLN" Then

                A = ConfigurationManager.AppSettings("ReportPath") + "\Pending_Status_Planner.rpt"

                _Todate = CDate(txtFromDate.Text).AddDays(-365)

                StrFromDate = "(" & Year(_Todate) & ", " & VB6.Format(Month(_Todate), "0#") & ", " & VB6.Format(CDate(_Todate).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtFromDate.Value) & ", " & VB6.Format(Month(txtFromDate.Value), "0#") & ", " & VB6.Format(CDate(txtFromDate.Text).Day, "0#") & ", 00, 00, 00)"

                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True

                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Delivary_Request.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T01Delivary_Request.T01Status} = 'A' and {T01Delivary_Request.T01Planner} ='" & strDisname & "'"
                frmReport.MdiParent = MDIMain
                'frmReport.Show()
                frmReport.Show()

            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub frmPending_Orders_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtTodate.Text = Today
        txtFromDate.Text = Today
    End Sub
End Class