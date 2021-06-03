Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmExaminer_Down
    Dim Clicked As String
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' cboDep.ToggleDropdown()

        txtDate.Text = Today
        txtTo.Text = Today
        cmdEdit.Enabled = True
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        'cmdSave.Enabled = False
        cmdAdd.Focus()
    End Sub


    Private Sub frmExaminer_Down_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'load NSL Combo box
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select EPFNo as [EPF No],FirstName as [Name] from Users"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboEmp
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
                .Rows.Band.Columns(1).Width = 160
                ' .Rows.Band.Columns(2).Width = 90
                '.Rows.Band.Columns(3).Width = 240
                ' .Rows.Band.Columns(4).Width = 110
                '  .Rows.Band.Columns(5).Width = 110

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
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim _Cutofwt As Double
        Dim _Reasone As String
        Dim Sql As String
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String

        Dim ncQryType As String
        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVccode As String
        Dim i As Integer

        'Dim con = New SqlConnection()
        'con = DBEngin.GetConnection()
        Dim recGRNheader As DataSet
        Dim recStockBalance As DataSet
        ' Dim A As String
        '  Dim B As New ReportDocument

        'Dim connection As SqlClient.SqlConnection
        'Dim transaction As SqlClient.SqlTransaction
        'Dim transactionCreated As Boolean
        'Dim connectionCreated As Boolean

        'connection = DBEngin.GetConnection(True)
        'connectionCreated = True
        'transaction = connection.BeginTransaction()
        'transactionCreated = True

        Dim strInvo As String
        Dim strChqvalue As Double
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _QReasons As String
        Dim _FromTime As String
        Dim _ToTime As String
        Dim _QReasons1 As String

        Try

            If chk1.Checked = True Then
                Dim Y As Date
                Dim X As Date
                Y = txtDate.Text & " " & txtTime1.Text
                X = txtTo.Text & " " & txtToTime.Text
                ' MsgBox(Format(Y.ToString("HH:mm:ss")))
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\examiner_Down.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Text)
                B.SetParameterValue("From", txtDate.Text)
                B.SetParameterValue("Name", cboEmp.Text)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T18Downtime.T18Timein} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T18Downtime.T18Time} >0 and {T18Downtime.T18User}='" & Trim(cboEmp.Text) & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()


            Else

                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\examiner_Down.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Text)
                B.SetParameterValue("From", txtDate.Text)
                B.SetParameterValue("Name", cboEmp.Text)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T18Downtime.T18Timein} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T18Downtime.T18OrderNo} <> '' and {T18Downtime.T18User}='" & Trim(cboEmp.Text) & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class