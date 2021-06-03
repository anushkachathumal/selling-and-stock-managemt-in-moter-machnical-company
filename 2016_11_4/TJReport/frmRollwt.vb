Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmRollwt
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

        ' Dim con = New SqlConnection()
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

        Dim M01 As DataSet
        Dim M02 As DataSet


        Dim weekNum As Integer
        Dim weekNum2 As Integer
        weekNum = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(txtDate.Text, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)
        weekNum2 = System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(txtTo.Text, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)

        Dim _Count As Integer
        Dim _Count1 As Integer
        Dim _Count2 As Integer

        Dim _Total1 As Integer
        Dim _Total2 As Integer
        Dim _Total3 As Integer
        Dim _Total4 As Integer
        Dim _Total5 As Integer
        Dim _Total6 As Integer
        Dim _Total7 As Integer
        Dim _Total8 As Integer

        Try
            i = 0
            nvcFieldList = "delete from R02Report where R02Status='" & netCard & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)


            Sql = "select * from T01Transaction_Header where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' order by T01Date"
            M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Sql = "select * from R02Report where R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                If isValidDataset(M02) Then

                    'If Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 0 Then
                    '    _Total1 = _Total1 + 1
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 1 Then
                    '    _Total2 = _Total2 + 1
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 2 Then
                    '    _Total3 = _Total3 + 1
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 3 Then
                    '    _Total4 = _Total4 + 1
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 4 Then
                    '    _Total5 = _Total5 + 1
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 6 Then
                    '    _Total6 = _Total6 + 1
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 7 Then
                    '    _Total7 = _Total7 + 1
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 8 Then
                    '    _Total8 = _Total8 + 1
                    'End If
                    If M01.Tables(0).Rows(i)("T0Rollweight") < 10 Then
                        nvcFieldList = "update R02Report set R02W1=R02W1 + " & 1 & ",R02WT=R02WT + " & M01.Tables(0).Rows(i)("T0Rollweight") & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                        _Count = _Count + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 10 And M01.Tables(0).Rows(i)("T0Rollweight") < 20 Then
                        nvcFieldList = "update R02Report set R02W2=R02W2 + " & 1 & ",R02WT=R02WT + " & M01.Tables(0).Rows(i)("T0Rollweight") & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                        _Count = _Count + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 20 And M01.Tables(0).Rows(i)("T0Rollweight") < 25 Then
                        nvcFieldList = "update R02Report set R02W3=R02W3 + " & 1 & ",R02WT=R02WT + " & M01.Tables(0).Rows(i)("T0Rollweight") & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                        _Count = _Count + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 25 And M01.Tables(0).Rows(i)("T0Rollweight") < 28 Then
                        nvcFieldList = "update R02Report set R02W4=R02W4 + " & 1 & ",R02WT=R02WT + " & M01.Tables(0).Rows(i)("T0Rollweight") & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                        _Count1 = _Count1 + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 28 And M01.Tables(0).Rows(i)("T0Rollweight") < 30 Then
                        nvcFieldList = "update R02Report set R02W5=R02W5 + " & 1 & ",R02WT=R02WT + " & M01.Tables(0).Rows(i)("T0Rollweight") & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                        _Count2 = _Count2 + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 30 And M01.Tables(0).Rows(i)("T0Rollweight") <= 39 Then
                        nvcFieldList = "update R02Report set R02W6=R02W6 + " & 1 & ",R02WT=R02WT + " & M01.Tables(0).Rows(i)("T0Rollweight") & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                        _Count2 = _Count2 + 1
                    End If

                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    'If Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 0 Then
                    '    nvcFieldList = "update R02Report set R02W7= " & _Total1 & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 1 Then
                    '    nvcFieldList = "update R02Report set R02W7=" & _Total2 & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 2 Then
                    '    nvcFieldList = "update R02Report set R02W7= " & _Total3 & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 3 Then
                    '    nvcFieldList = "update R02Report set R02W7=" & _Total4 & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 4 Then
                    '    nvcFieldList = "update R02Report set R02W7= " & _Total5 & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 6 Then
                    '    nvcFieldList = "update R02Report set R02W7=" & _Total6 & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 7 Then
                    '    nvcFieldList = "update R02Report set R02W7=" & _Total7 & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    'ElseIf Val(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) - weekNum = 8 Then
                    '    nvcFieldList = "update R02Report set R02W7= " & _Total8 & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    'End If
                    nvcFieldList = "update R02Report set R02W7=R02W7 + " & 1 & "  where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    'LOW WEIGHT
                    nvcFieldList = "update R02Report set R02W8=" & (_Count / i) & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    'STANDED WEIGHT
                    nvcFieldList = "update R02Report set R02W9=" & (_Count1 / i) & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    'OVER WEIGHT
                    nvcFieldList = "update R02Report set R02W10=" & (_Count2 / i) & " where  R02Dis='" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "' and R02Status='" & netCard & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)


                Else

                    If M01.Tables(0).Rows(i)("T0Rollweight") < 10 Then
                        nvcFieldList = "Insert Into R02Report(R02Dis,R02W1,R02Status,R02W7,R02WT,R02W2,R02W3,R02W4,R02W5,R02W6,R02W8,R02W9,R02W10)" & _
                                                             " values('" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "','1','" & netCard & "','1'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'0','0','0','0','0','0','0','0')"
                        _Count = _Count + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 10 And M01.Tables(0).Rows(i)("T0Rollweight") < 20 Then
                        nvcFieldList = "Insert Into R02Report(R02Dis,R02W2,R02Status,R02W7,R02WT,R02W1,R02W3,R02W4,R02W5,R02W6,R02W8,R02W9,R02W10)" & _
                                                             " values('" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "','1','" & netCard & "','1'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'0','0','0','0','0','0','0','0')"
                        _Count = _Count + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 20 And M01.Tables(0).Rows(i)("T0Rollweight") < 25 Then
                        nvcFieldList = "Insert Into R02Report(R02Dis,R02W3,R02Status,R02W7,R02WT,R02W1,R02W2,R02W4,R02W5,R02W6,R02W8,R02W9,R02W10)" & _
                                                             " values('" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "','1','" & netCard & "','1'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'0','0','0','0','0','0','0','0')"
                        _Count = _Count + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 25 And M01.Tables(0).Rows(i)("T0Rollweight") < 28 Then
                        nvcFieldList = "Insert Into R02Report(R02Dis,R02W4,R02Status,R02W7,R02WT,R02W1,R02W2,R02W3,R02W5,R02W6,R02W8,R02W9,R02W10)" & _
                                                             " values('" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "','1','" & netCard & "','1'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'0','0','0','0','0','0','0','0')"
                        _Count1 = _Count1 + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 28 And M01.Tables(0).Rows(i)("T0Rollweight") < 30 Then
                        nvcFieldList = "Insert Into R02Report(R02Dis,R02W5,R02Status,R02W7,R02WT,R02W1,R02W2,R02W3,R02W4,R02W6,R02W8,R02W9,R02W10)" & _
                                                             " values('" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "','1','" & netCard & "','1'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'0','0','0','0','0','0','0','0')"
                        _Count2 = _Count2 + 1
                    ElseIf M01.Tables(0).Rows(i)("T0Rollweight") >= 30 And M01.Tables(0).Rows(i)("T0Rollweight") <= 39 Then
                        nvcFieldList = "Insert Into R02Report(R02Dis,R02W6,R02Status,R02W7,R02WT,R02W1,R02W2,R02W3,R02W4,R02W5,R02W8,R02W9,R02W10)" & _
                                                              " values('" & System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(Trim(M01.Tables(0).Rows(i)("T01Date")), System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) & "','1','" & netCard & "','1'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'0','0','0','0','0','0','0','0')"
                        _Count2 = _Count2 + 1
                    End If

                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                End If
                i = i + 1
            Next

            i = 0
            Sql = "select (R02W1+R02W2+R02W3) as Low,R02W7,R02W4,(R02W5+R02W6) as Over1 ,R02Dis from R02Report where R02Status='" & netCard & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                nvcFieldList = "update R02Report set R02W8=" & (Val(M01.Tables(0).Rows(i)("Low")) / Val(M01.Tables(0).Rows(i)("R02W7"))) * 100 & ",R02W9=" & (Val(M01.Tables(0).Rows(i)("R02W4")) / Val(M01.Tables(0).Rows(i)("R02W7"))) * 100 & ",R02W10=" & (Val(M01.Tables(0).Rows(i)("Over1")) / Val(M01.Tables(0).Rows(i)("R02W7"))) * 100 & " where  R02Dis='" & M01.Tables(0).Rows(i)("R02Dis") & "' and R02Status='" & netCard & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)
                i = i + 1
            Next
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            MsgBox("Record genarated successfully", MsgBoxStyle.Information, "Textued Jersey .........")
            A = ConfigurationManager.AppSettings("ReportPath") + "\Rollwt.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            B.SetParameterValue("To", txtTo.Value)
            B.SetParameterValue("From", txtDate.Value)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{R02Report.R02Status}='" & netCard & "'"
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