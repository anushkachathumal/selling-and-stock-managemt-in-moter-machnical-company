Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmKnittingReport
    Dim Clicked As String
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' cboDep.ToggleDropdown()

       
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

    Private Sub frmKnittingReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m03 As DataSet
        Dim Sql As String

        Try
            'Load Production Order No

            Sql = "select M03OrderNo as [Order No] from M03Knittingorder where M03Status='A'"
            m03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboFrom
                .DataSource = m03
                .Rows.Band.Columns(0).Width = 140
            End With

            With cboTo
                .DataSource = m03
                .Rows.Band.Columns(0).Width = 140
            End With


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click

        Dim B As New ReportDocument
        Dim A As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim Sql As String

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim i As Integer

        Dim M03 As DataSet
        Dim T01 As DataSet

        Dim _Scrap As Double
        Dim _Qty As Double
        Dim _Quarantine As Double
        Dim _Reject As Double
        Dim nvcFieldList As String

        i = 0
        Try
            Sql = "delete from R03Report where R03Status='" & netCard & "' "
            ExecuteNonQueryText(connection, transaction, Sql)

            Sql = "select M03MCNo,M03OrderNo,M03Orderqty from M03Knittingorder  where M03OrderNo between '" & cboFrom.Text & "' and '" & cboTo.Text & "'"
            M03 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            For Each DTRow1 As DataRow In M03.Tables(0).Rows
                _Scrap = 0
                _Qty = 0
                _Quarantine = 0
                _Reject = 0
                Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & M03.Tables(0).Rows(i)("M03OrderNo") & "' and T01Status in ('P','QP') group by T01OrderNo"
                T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                If isValidDataset(T01) Then
                    _Qty = T01.Tables(0).Rows(0)("T0Rollweight")
                End If

                'QUARANTING QTY

                Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & M03.Tables(0).Rows(i)("M03OrderNo") & "' and T01Status in ('RP','QR','Q') group by T01OrderNo"
                T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                If isValidDataset(T01) Then
                    _Quarantine = T01.Tables(0).Rows(0)("T0Rollweight")
                End If

                'SCRAP

                Sql = "SELECT SUM(T05Weight) AS T05Weight FROM T05Scrab INNER JOIN T01Transaction_Header ON T01RefNo=T05RefNo WHERE T01OrderNo='" & M03.Tables(0).Rows(i)("M03OrderNo") & "' GROUP BY T05RefNo"
                T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                If isValidDataset(T01) Then
                    _Scrap = T01.Tables(0).Rows(0)("T05Weight")
                End If

                'REJECT
                Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & M03.Tables(0).Rows(i)("M03OrderNo") & "' and T01Status in ('R') group by T01OrderNo"
                T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                If isValidDataset(T01) Then
                    _Reject = T01.Tables(0).Rows(0)("T0Rollweight")
                End If


                nvcFieldList = "Insert Into R03Report(R03OrderNo,R03Qty,R03Usable,R03Quarantine,R03Reject,R03Scrap,R03Balance,R03Status,R03MC)" & _
                                                            " values('" & M03.Tables(0).Rows(i)("M03OrderNo") & "', " & M03.Tables(0).Rows(i)("M03Orderqty") & "," & _Qty & "," & _Quarantine & "," & _Reject & "," & _Scrap & "," & (_Qty + _Quarantine) - Val(M03.Tables(0).Rows(i)("M03Orderqty")) & ",'" & netCard & "','" & M03.Tables(0).Rows(i)("M03MCNo") & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)


                i = i + 1
            Next
            MsgBox("Report Genarating Successfully", MsgBoxStyle.Information, "Report Genarating ........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            A = ConfigurationManager.AppSettings("ReportPath") + "\Knitting.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            B.SetParameterValue("To", cboTo.Text)
            B.SetParameterValue("From", cboFrom.Text)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{R03Report.R03Status}='" & netCard & "'"
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class