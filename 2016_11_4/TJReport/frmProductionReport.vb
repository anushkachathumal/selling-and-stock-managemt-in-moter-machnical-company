
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmProductionReport
    Dim Clicked As String
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True

        txtDate.Text = Today
        txtTo.Text = Today
        chk0.Checked = True
        cmdEdit.Enabled = True
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        'cmdSave.Enabled = False


        cmdAdd.Focus()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim _FromTime As String
        Dim _ToTime As String
        Dim _Shift As Integer

        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String
        Try

            If chk0.Checked = False And chk1.Checked = False Then
                MsgBox("Please select the shift", MsgBoxStyle.Information, "Textured Jersey ........")

            Else
                If chk0.Checked = True Then
                    _Shift = 1

                    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                    StrToDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"

                    _FromTime = txtDate.Text & " " & "07:30:00"
                    _ToTime = txtDate.Text & " " & "19:30:00"


                    A = ConfigurationManager.AppSettings("ReportPath") + "\Production1.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("Shift", _Shift)
                    B.SetParameterValue("From", txtDate.Value)
                    ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime" & P01 & ""
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T04Summery.T04Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T04Summery.T04Shift}=" & _Shift & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()

                Else
                    _Shift = 2

                    _FromTime = txtDate.Text & " " & "07:30:00"
                    _ToTime = System.DateTime.FromOADate(CDate(txtDate.Text).ToOADate + 1)
                    _ToTime = _ToTime & " " & "07:30:00"

                    txtTo.Text = _ToTime

                    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                    StrToDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"

                    _FromTime = txtDate.Text & " " & "07:30:00"
                    _ToTime = txtDate.Text & " " & "19:30:00"


                    A = ConfigurationManager.AppSettings("ReportPath") + "\Production2.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "tommya")
                    B.SetParameterValue("Shift", _Shift)
                    B.SetParameterValue("From", txtDate.Value)
                    ' frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime" & P01 & ""
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T04Summery.T04Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T04Summery.T04Shift}=" & _Shift & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()

                End If



            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class