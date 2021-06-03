Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmQurantine
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
        Dim _Regectwt As Double
        Dim _FromTime As String
        Dim _ToTime As String
        Dim _Department As String

        Try
            If chk1.Checked = False And chk2.Checked = False And chk3.Checked = False And chk4.Checked = False Then
                Dim Y As Date
                Dim X As Date
                'Y = txtDate.Text & " " & txtTime1.Text
                'X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                'StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                'StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"

                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineT.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and not ({rptReport_Detailes.rptStatus} like ['P', 'R'])"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = False And chk2.Checked = True And chk3.Checked = False And chk4.Checked = False Then

                'Y = txtDate.Text & " " & txtTime1.Text
                'X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                'StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                'StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"

                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineT.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptStatus}='Q'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = False And chk2.Checked = False And chk3.Checked = True And chk4.Checked = False Then

                'Y = txtDate.Text & " " & txtTime1.Text
                'X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                'StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                'StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"

                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineT.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptStatus}='RP'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = False And chk2.Checked = False And chk3.Checked = False And chk4.Checked = True Then

                'Y = txtDate.Text & " " & txtTime1.Text
                'X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                'StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                'StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"

                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineT.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptStatus}='QP'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = False And chk4.Checked = False Then
                Dim Y As Date
                Dim X As Date

                Y = txtDate.Text & " " & txtTime1.Text
                X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"



                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineT.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and not ({rptReport_Detailes.rptStatus} like ['P', 'R'])"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = True And chk2.Checked = True And chk3.Checked = False And chk4.Checked = False Then
                Dim Y As Date
                Dim X As Date

                Y = txtDate.Text & " " & txtTime1.Text
                X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"



                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineT.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T01Transaction_Header.T01Status}='Q'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = True And chk4.Checked = False Then
                Dim Y As Date
                Dim X As Date

                Y = txtDate.Text & " " & txtTime1.Text
                X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"



                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineT.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptStatus}='RP'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = False And chk4.Checked = True Then
                Dim Y As Date
                Dim X As Date

                Y = txtDate.Text & " " & txtTime1.Text
                X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"



                A = ConfigurationManager.AppSettings("ReportPath") + "\QuarantineTQP.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptStatus}='QP'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            'If Trim(strDisname) = "KNT" Then

            '    nvcFieldList = "delete from tmpR01Report where R01WorkStation='" & netCard & "'"
            '    ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '    Dim X As Integer
            '    If chk1.Checked = False Then
            '        Sql = "select max(T01RefNo) as T01RefNo,max(T01R_Reason) as T01R_Reason,max(T0Rollweight) as T0Rollweight,T01OrderNo ,T01RollNo from T01Transaction_Header where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and  T01Status in('Q','QP','QR','RP')  group by T01OrderNo,T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select M02Name,T07Dis from T07Q_Reason inner join M02Fault on M02Code=T07F_Code where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("M02Name"))
            '                    If IsDBNull(Trim(M02.Tables(0).Rows(X)("T07Dis"))) Then
            '                    Else
            '                        _Department = Trim(M02.Tables(0).Rows(X)("T07Dis"))
            '                    End If
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("M02Name"))
            '                    If IsDBNull(Trim(M02.Tables(0).Rows(X)("T07Dis"))) Then
            '                    Else
            '                        _Department = _Department & "/" & Trim(M02.Tables(0).Rows(X)("T07Dis"))
            '                    End If
            '                End If
            '                X = X + 1

            '            Next
            '            Dim _CPI As String
            '            _CPI = ""
            '            Sql = "select * from T03CPI_Reading where T03RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            If isValidDataset(M02) Then

            '                _CPI = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '                End If
            '            End If

            '            nvcFieldList = "Insert Into tmpR01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01Department)" & _
            '                                                        " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'" & netCard & "','" & _Department & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            i = i + 1
            '        Next
            '    Else
            '        'Dim _FromTime As String
            '        'Dim _ToTime As String

            '        _FromTime = txtDate.Text & " " & txtTime1.Text
            '        _ToTime = txtTo.Text & " " & txtToTime.Text

            '        _Department = ""
            '        Sql = "select max(T01RefNo) as T01RefNo,max(T01R_Reason) as T01R_Reason,max(T0Rollweight) as T0Rollweight,T01OrderNo ,T01RollNo from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in('Q','QP','QR','RP') group by T01OrderNo,T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select M02Name,T07Dis from T07Q_Reason inner join M02Fault on M02Code=T07F_Code where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("M02Name"))
            '                    If IsDBNull(Trim(M02.Tables(0).Rows(X)("T07Dis"))) Then
            '                    Else
            '                        _Department = Trim(M02.Tables(0).Rows(X)("T07Dis"))
            '                    End If
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("M02Name"))
            '                    If IsDBNull(Trim(M02.Tables(0).Rows(X)("T07Dis"))) Then
            '                    Else
            '                        _Department = _Department & "/" & Trim(M02.Tables(0).Rows(X)("T07Dis"))
            '                    End If
            '                End If
            '                X = X + 1

            '            Next
            '            Dim _CPI As String
            '            _CPI = ""
            '            Sql = "select * from T03CPI_Reading where T03RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            If isValidDataset(M02) Then

            '                _CPI = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '                End If
            '            End If
            '            '    If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) >= Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) Then
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '            '    Else
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))
            '            '    End If

            '            '    If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) Then
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '            '    End If

            '            '    If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) Then
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '            '    End If

            '            '    If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) Then
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '            '    End If

            '            '    If Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) = 0 Then
            '            '        _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '            '    Else
            '            '        If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) >= Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) Then
            '            '            _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))
            '            '        Else
            '            '            _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '            '        End If
            '            '    End If


            '            '    If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) <> 0 Then
            '            '        _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '            '    End If

            '            '    If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) <> 0 Then
            '            '        _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '            '    End If

            '            '    If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) <> 0 Then
            '            '        _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '            '    End If

            '            '    If _Max = 0 And _Min = 0 Then
            '            '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01R_Whight,R01Status)" & _
            '            '                                              " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','0','" & _Max - _Min & "')"
            '            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            '    Else
            '            '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01R_Whight,R01Status)" & _
            '            '                                              " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _AvgCPI / _COUNT & "','" & _Max - _Min & "')"
            '            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            '    End If
            '            'End If

            '            Dim n_QPassUser As String
            '            Dim T06 As DataSet
            '            nvcFieldList = "select * from T06QuarantinePass where T06RefNo='" & M01.Tables(0).Rows(i)("T01RefNo") & "'"
            '            T06 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
            '            If isValidDataset(T06) Then
            '                n_QPassUser = T06.Tables(0).Rows(0)("T06User")
            '            Else
            '                n_QPassUser = ""
            '            End If

            '            nvcFieldList = "Insert Into tmpR01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01Department,R01Status,R01CPI)" & _
            '                                                             " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'" & netCard & "','" & _Department & "','" & n_QPassUser & "','" & _CPI & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '            i = i + 1
            '        Next

            '    End If
            '    MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            '    transaction.Commit()
            '    DBEngin.CloseConnection(connection)
            '    If chk1.Checked = False And chk2.Checked = False And chk3.Checked = False And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReportTMP6.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", txtTo.Value)
            '        B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{tmpR01Report.R01WorkStation}='" & netCard & "'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = False And chk2.Checked = True And chk3.Checked = False And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReportTMP6.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", txtTo.Value)
            '        B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{tmpR01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'Q'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = False And chk2.Checked = False And chk3.Checked = True And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReportTMP6.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", txtTo.Value)
            '        B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{tmpR01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'RP'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = False And chk2.Checked = False And chk3.Checked = False And chk4.Checked = True Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReportTMP6.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", txtTo.Value)
            '        B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{tmpR01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'QR'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()
            '    ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = False And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReportTMP5.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("From", _FromTime)
            '        B.SetParameterValue("To", _ToTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{tmpR01Report.R01WorkStation}='" & netCard & "'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()
            '    ElseIf chk1.Checked = True And chk2.Checked = True And chk3.Checked = False And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport5.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("From", _FromTime)
            '        B.SetParameterValue("To", _ToTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'Q'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = True And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport5.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("From", _FromTime)
            '        B.SetParameterValue("To", _ToTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'RP'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = False And chk4.Checked = True Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport5.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("From", _FromTime)
            '        B.SetParameterValue("To", _ToTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'QP'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()
            '    End If
            'Else

            '    nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            '    ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '    Dim X As Integer
            '    If chk1.Checked = False Then
            '        Sql = "select max(T01RefNo) as T01RefNo,max(T01R_Reason) as T01R_Reason,max(T0Rollweight) as T0Rollweight,T01OrderNo ,T01RollNo from T01Transaction_Header where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and  T01Status in('Q','QP','QR','RP')  group by T01OrderNo,T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select M02Name,T07Dis from T07Q_Reason inner join M02Fault on M02Code=T07F_Code where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("M02Name"))
            '                    If IsDBNull(Trim(M02.Tables(0).Rows(X)("T07Dis"))) Then
            '                    Else
            '                        _Department = Trim(M02.Tables(0).Rows(X)("T07Dis"))
            '                    End If
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("M02Name"))
            '                    If IsDBNull(Trim(M02.Tables(0).Rows(X)("T07Dis"))) Then
            '                    Else
            '                        _Department = _Department & "/" & Trim(M02.Tables(0).Rows(X)("T07Dis"))
            '                    End If
            '                End If
            '                X = X + 1

            '            Next
            '            Dim _CPI As String
            '            _CPI = ""
            '            Sql = "select * from T03CPI_Reading where T03RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            If isValidDataset(M02) Then

            '                _CPI = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '                End If
            '            End If

            '            nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01Department)" & _
            '                                                        " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'" & netCard & "','" & _Department & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            i = i + 1
            '        Next
            '    Else
            '        'Dim _FromTime As String
            '        'Dim _ToTime As String

            '        _FromTime = txtDate.Text & " " & txtTime1.Text
            '        _ToTime = txtTo.Text & " " & txtToTime.Text

            '        _Department = ""
            '        Sql = "select max(T01RefNo) as T01RefNo,max(T01R_Reason) as T01R_Reason,max(T0Rollweight) as T0Rollweight,T01OrderNo ,T01RollNo from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in('Q','QP','QR','RP') group by T01OrderNo,T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select M02Name,T07Dis from T07Q_Reason inner join M02Fault on M02Code=T07F_Code where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("M02Name"))
            '                    If IsDBNull(Trim(M02.Tables(0).Rows(X)("T07Dis"))) Then
            '                    Else
            '                        _Department = Trim(M02.Tables(0).Rows(X)("T07Dis"))
            '                    End If
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("M02Name"))
            '                    If IsDBNull(Trim(M02.Tables(0).Rows(X)("T07Dis"))) Then
            '                    Else
            '                        _Department = _Department & "/" & Trim(M02.Tables(0).Rows(X)("T07Dis"))
            '                    End If
            '                End If
            '                X = X + 1

            '            Next
            '            Dim _CPI As String
            '            _CPI = ""
            '            Sql = "select * from T03CPI_Reading where T03RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            If isValidDataset(M02) Then

            '                _CPI = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '                End If

            '                If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) Then
            '                    _CPI = _CPI & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '                    ' _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '                End If
            '            End If
            '            '    If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) >= Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) Then
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '            '    Else
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))
            '            '    End If

            '            '    If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) Then
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '            '    End If

            '            '    If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) Then
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '            '    End If

            '            '    If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) Then
            '            '        _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '            '    End If

            '            '    If Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) = 0 Then
            '            '        _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '            '    Else
            '            '        If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) >= Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) Then
            '            '            _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))
            '            '        Else
            '            '            _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '            '        End If
            '            '    End If


            '            '    If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) <> 0 Then
            '            '        _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '            '    End If

            '            '    If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) <> 0 Then
            '            '        _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '            '    End If

            '            '    If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) <> 0 Then
            '            '        _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '            '    End If

            '            '    If _Max = 0 And _Min = 0 Then
            '            '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01R_Whight,R01Status)" & _
            '            '                                              " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','0','" & _Max - _Min & "')"
            '            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            '    Else
            '            '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01R_Whight,R01Status)" & _
            '            '                                              " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _AvgCPI / _COUNT & "','" & _Max - _Min & "')"
            '            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            '    End If
            '            'End If

            '            Dim n_QPassUser As String
            '            Dim T06 As DataSet
            '            nvcFieldList = "select * from T06QuarantinePass where T06RefNo='" & M01.Tables(0).Rows(i)("T01RefNo") & "'"
            '            T06 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
            '            If isValidDataset(T06) Then
            '                n_QPassUser = T06.Tables(0).Rows(0)("T06User")
            '            Else
            '                n_QPassUser = ""
            '            End If

            '            nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01Department,R01Status,R01CPI)" & _
            '                                                             " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'" & netCard & "','" & _Department & "','" & n_QPassUser & "','" & _CPI & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '            i = i + 1
            '        Next

            '    End If
            '    MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            '    transaction.Commit()
            '    DBEngin.CloseConnection(connection)
            '    If chk1.Checked = False And chk2.Checked = False And chk3.Checked = False And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport6.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", txtTo.Value)
            '        B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = False And chk2.Checked = True And chk3.Checked = False And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport6.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", txtTo.Value)
            '        B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'Q'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = False And chk2.Checked = False And chk3.Checked = True And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport6.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", txtTo.Value)
            '        B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'RP'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = False And chk2.Checked = False And chk3.Checked = False And chk4.Checked = True Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport6.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", txtTo.Value)
            '        B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'QR'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()
            '    ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = False And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport5.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("From", _FromTime)
            '        B.SetParameterValue("To", _ToTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()
            '    ElseIf chk1.Checked = True And chk2.Checked = True And chk3.Checked = False And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport5.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("From", _FromTime)
            '        B.SetParameterValue("To", _ToTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'Q'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = True And chk4.Checked = False Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport5.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("From", _FromTime)
            '        B.SetParameterValue("To", _ToTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'RP'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk1.Checked = True And chk2.Checked = False And chk3.Checked = False And chk4.Checked = True Then
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport5.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("From", _FromTime)
            '        B.SetParameterValue("To", _ToTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {T01Transaction_Header.T01Status} = 'QP'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()
            '    End If

            'End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub chk2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk2.CheckedChanged
        If chk2.Checked = True Then
            chk3.Checked = False
            chk4.Checked = False
        Else

        End If
    End Sub

    Private Sub chk3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk3.CheckedChanged
        If chk3.Checked = True Then
            chk2.Checked = False
            chk4.Checked = False
        Else

        End If
    End Sub

    Private Sub chk4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk4.CheckedChanged
        If chk4.Checked = True Then
            chk3.Checked = False
            chk2.Checked = False
        Else

        End If
    End Sub
End Class