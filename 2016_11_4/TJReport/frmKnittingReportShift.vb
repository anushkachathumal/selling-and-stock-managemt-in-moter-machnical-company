
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmKnittingReportShift

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
        Dim _FromTime As String
        Dim _ToTime As String
        Dim R05 As DataSet
        Dim T05 As DataSet
        Dim TX As DataSet

        Dim n_Usableqty As Double
        i = 0
        Try
            Sql = "delete from R05Report where R05id='" & netCard & "' "
            ExecuteNonQueryText(connection, transaction, Sql)


            _FromTime = txtDate.Text & " " & txtTime1.Text
            _ToTime = txtTo.Text & " " & txtToTime.Text


            If Trim(txtM1.Text) <> "" Then
                If Trim(txtM2.Text) <> "" Then
                    nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('P','QP') AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
                    T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                    i = 0

                    'Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & M03.Tables(0).Rows(i)("M03OrderNo") & "' and T01Status in ('P','QP') group by T01OrderNo"
                    'T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    'If isValidDataset(T01) Then
                    '    _Qty = T01.Tables(0).Rows(0)("T0Rollweight")
                    'End If

                    'USABLE QTY
                    For Each DTRow1 As DataRow In T01.Tables(0).Rows
                        _Qty = 0
                        Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and T01Status in ('P','QP') group by T01OrderNo"
                        TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(TX) Then
                            _Qty = TX.Tables(0).Rows(0)("T0Rollweight")
                        End If
                        ' n_Usableqty = T01.Tables(0).Rows(0)("T0Rollweight")
                        nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
                                                              " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','" & _Qty & "','0','0','0','" & netCard & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)
                        i = i + 1
                    Next
                    '---------------------------------------------------------
                    'SCAP
                    nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('R','QR') AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
                    T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                    i = 0
                    For Each DTRow1 As DataRow In T01.Tables(0).Rows
                        _Scrap = 0
                        Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and T01Status in ('QR','R') group by T01OrderNo"
                        TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(TX) Then
                            _Scrap = TX.Tables(0).Rows(0)("T0Rollweight")
                        End If

                        Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
                        R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(R05) Then
                            nvcFieldList = "update R05Report set R05Scrap='" & _Scrap & "' where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList)
                        Else
                            nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
                                                             " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','" & _Scrap & "','0','0','" & netCard & "')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList)
                        End If



                        i = i + 1
                    Next

                    i = 0
                    '--------------------------------------------------------
                    nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status<>'I' AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
                    T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                    For Each DTRow1 As DataRow In T01.Tables(0).Rows
                        _Scrap = 0
                        Sql = "SELECT SUM(T05Weight) AS T05Weight FROM T05Scrab INNER JOIN T01Transaction_Header ON T01RefNo=T05RefNo WHERE T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' GROUP BY T05RefNo"
                        TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(TX) Then
                            _Scrap = TX.Tables(0).Rows(0)("T05Weight")
                        End If

                        Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
                        R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(R05) Then
                            nvcFieldList = "update R05Report set R05Scrap=R05Scrap +" & _Scrap & " where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList)
                        Else
                            nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
                                                             " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','" & _Scrap & "','0','0','" & netCard & "')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList)
                        End If

                        'End If
                        i = i + 1
                    Next
                    '---------------------------------------------------------------------------------
                    i = 0
                    '--------------------------------------------------------
                    'nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status<>'I' AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
                    'T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                    'For Each DTRow1 As DataRow In T01.Tables(0).Rows
                    '    Sql = "select SUM(T04Weight) as T04Weight from T01Transaction_Header inner join T04Cutoff on T04RefNo=T01RefNo where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' group by T01OrderNo "
                    '    T05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    '    If isValidDataset(T05) Then

                    '        Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
                    '        R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    '        If isValidDataset(R05) Then
                    '            nvcFieldList = "update R05Report set R05Scrap=R05Scrap +" & T05.Tables(0).Rows(0)("T04Weight") & " where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
                    '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    '        Else
                    '            nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
                    '                                             " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','" & T05.Tables(0).Rows(0)("T04Weight") & "','0','0','" & netCard & "')"
                    '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    '        End If
                    '    End If
                    '    'End If
                    '    i = i + 1
                    'Next
                    '--------------------------------------------------------------------------------------
                    'QUARANTINE
                    nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo where  T01Status in ('Q','RP') AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' and T01Time between '" & _FromTime & "' and '" & _ToTime & "' group by M03MCNo,T01OrderNo"
                    T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                    i = 0
                    For Each DTRow1 As DataRow In T01.Tables(0).Rows
                        _Quarantine = 0

                        Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and T01Status in ('RP','QR','Q') group by T01OrderNo"
                        TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(TX) Then
                            _Quarantine = TX.Tables(0).Rows(0)("T0Rollweight")
                        End If


                        Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
                        R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                        If isValidDataset(R05) Then
                            nvcFieldList = "update R05Report set R05Quarantine='" & _Quarantine & "' where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList)
                        Else
                            nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
                                                             " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','0','" & _Quarantine & "','0','" & netCard & "')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList)
                        End If



                        i = i + 1
                    Next

                    '  MsgBox("Records Update sucessfully", MsgBoxStyle.Information, "Information .....")
                    transaction.Commit()
                    DBEngin.CloseConnection(connection)
                    connection.ConnectionString = ""
                Else
                    MsgBox("Please enter the To Machine", MsgBoxStyle.Information, "Information ....")
                    Exit Sub

                End If
            Else
                MsgBox("Please enter the form Machine", MsgBoxStyle.Information, "Information ....")
                Exit Sub
            End If
            Sql = ""

            MsgBox("Report Genarating Successfully", MsgBoxStyle.Information, "Report Genarating ........")
            ' transaction.Commit()
            A = ConfigurationManager.AppSettings("ReportPath") + "\KPShift.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            B.SetParameterValue("To", _ToTime)
            B.SetParameterValue("From", _FromTime)
            B.SetParameterValue("M/C", txtM1.Text & " - " & txtM2.Text)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{R05Report.R05id}='" & netCard & "'"
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub frmKnittingReportShift_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDate.Text = Today
        txtTo.Text = Today
    End Sub
End Class