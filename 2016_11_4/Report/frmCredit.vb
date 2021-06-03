Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmCredit
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        'OPR3.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        cmdSave.Enabled = True
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        OPR0.Enabled = False
        'OPR3.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String
        Try
            A = ConfigurationManager.AppSettings("ReportPath") + "\Credit.rpt"


            StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

            B.Load(A.ToString)
            B.SetParameterValue("Todate", txtTo.Value)
            B.SetParameterValue("Fromdate", txtDate.Value)
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            'frmReport.CrystalReportViewer1.SelectionFormula = "{T01MotocycleSales.T01Date} in '" & cdate(txtDate.Text) & "' to '" & cdate(txtTo.Text) & "'"
            'UPGRADE_WARNING: Couldn't resolve default property of object frmReport.CrystalReportViewer1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object txtTo.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object txtDate.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'frmReport.CrystalReportViewer1.SelectionFormula = "{T01MotocycleSales.T01Date} in " & CDate(txtDate.Text) & " to " & CDate(txtTo.Text) & ""
            frmReport.CrystalReportViewer1.SelectionFormula = "{T08EasyPayment.T08Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T08EasyPayment.T08Status} = 'A'"
            frmReport.MdiParent = MDIMain
            'frmReport.Show()
            frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class