Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmQurantineGrf
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


    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String


        'Dim con = New SqlConnection()
        'con = DBEngin.GetConnection()
        Dim recGRNheader As DataSet
        Dim recStockBalance As DataSet
        ' Dim A As String
        '  Dim B As New ReportDocument

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim _FromTime As String
        Dim _ToTime As String
        Dim T01 As DataSet
        Dim nvcFieldList As String
        Dim i As Integer
        Dim _Shift As String

        Try
            StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

            nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            If chk1.Checked = True Then
                _FromTime = txtDate.Text & " " & txtTime1.Text
                _ToTime = txtTo.Text & " " & txtToTime.Text


                nvcFieldList = "select M02Name,count(M02Name) as Fcount from T07Q_Reason inner join T01Transaction_Header on T01RefNo=T07RefNo inner join M02Fault on T07F_Code=M02Code where T01Time between '" & _FromTime & "' and '" & _ToTime & "' group by M02Name"
                T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                i = 0
                For Each DTRow1 As DataRow In T01.Tables(0).Rows

                    nvcFieldList = "Insert Into R01Report(R01OrderNo,R01CutWhight,R01WorkStation)" & _
                                                           " values('" & T01.Tables(0).Rows(i)("M02Name") & "', '" & T01.Tables(0).Rows(i)("Fcount") & "','" & netCard & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    i = i + 1
                Next

                MsgBox("Report Genarating Sucessfully", MsgBoxStyle.Information, "Textured Jersey ......")
                transaction.Commit()

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""

                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineGrf.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("From", _FromTime)
                B.SetParameterValue("To", _ToTime)
                ' B.SetParameterValue("Fname", cboFrom.Text)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub
End Class