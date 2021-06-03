Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmCutoff
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

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim strInvo As String
        Dim strChqvalue As Double
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _QReasons As String
        Dim _Regectwt As Double
        Dim _FromTime As String
        Dim _ToTime As String
        Dim _3Mtrcutoff As Double
        Dim _Department As String

        Try
            'If chk1.Checked = True Then

            '    Dim Y As Date
            '    Dim X As Date
            '    Y = txtDate.Text & " " & txtTime1.Text
            '    X = txtTo.Text & " " & txtToTime.Text
            '    ' MsgBox(Format(Y.ToString("HH:mm:ss")))
            '    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
            '    StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"

            '    A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport2T.rpt"
            '    B.Load(A.ToString)
            '   B.SetDatabaseLogon("sa", "tommya")
            '    B.SetParameterValue("To", txtTo.Value)
            '    B.SetParameterValue("From", txtDate.Value)
            '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '    frmReport.CrystalReportViewer1.DisplayToolbar = True
            '    frmReport.CrystalReportViewer1.SelectionFormula = "({%Reject} > $0.00  or {rptReport_Detailes.rpt3mcutoff} > $0.00) and{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
            '    frmReport.Refresh()
            '    ' frmReport.CrystalReportViewer1.PrintReport()
            '    ' B.PrintToPrinter(1, True, 0, 0)
            '    frmReport.MdiParent = MDIMain
            '    frmReport.Show()
            'Else
            '    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            '    StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

            '    A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport2T.rpt"
            '    B.Load(A.ToString)
            '   B.SetDatabaseLogon("sa", "tommya")
            '    B.SetParameterValue("To", txtTo.Value)
            '    B.SetParameterValue("From", txtDate.Value)
            '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '    frmReport.CrystalReportViewer1.DisplayToolbar = True
            '    frmReport.CrystalReportViewer1.SelectionFormula = "({%Reject} > $0.00  or {rptReport_Detailes.rpt3mcutoff} > $0.00) and {T01Transaction_Header.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  "
            '    frmReport.Refresh()
            '    ' frmReport.CrystalReportViewer1.PrintReport()
            '    ' B.PrintToPrinter(1, True, 0, 0)
            '    frmReport.MdiParent = MDIMain
            '    frmReport.Show()
            'End If



            nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            Dim X As Integer
            _Cutofwt = 0
            _3Mtrcutoff = 0
            _Department = ""
            _QReasons = ""
            If chk1.Checked = False Then
                Sql = "select max(T01RefNo) as T01RefNo,max(T01R_Reason) as T01R_Reason,max(T01Reject)as T01Reject,T01OrderNo ,T01RollNo from T01Transaction_Header where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "'  group by T01OrderNo,T01RollNo"
                M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                i = 0
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    _QReasons = ""
                    'If M01.Tables(0).Rows(i)("T01OrderNo") = "1485910" Then
                    '    MsgBox("Ok")
                    'End If
                    Sql = "select * from T05Scrab where T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
                    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    X = 0
                    For Each DTRow2 As DataRow In M02.Tables(0).Rows
                        If X = 0 Then
                            _QReasons = Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
                            _Department = Trim(M02.Tables(0).Rows(X)("T05Department"))
                        Else
                            _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
                            _Department = _Department & "/" & Trim(M02.Tables(0).Rows(X)("T05Department"))
                        End If
                        X = X + 1

                    Next
                    '-----------------------------------------------------
                    _3Mtrcutoff = 0
                    _Cutofwt = 0
                    Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
                    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M02) Then
                        _Cutofwt = M02.Tables(0).Rows(0)("QTY")
                    End If
                    '---------------------------------------------------------

                    Sql = "select sum(T04Weight) as Qty from T04Cutoff where T04RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " group by T04RefNo"
                    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M02) Then
                        _3Mtrcutoff = M02.Tables(0).Rows(0)("QTY")
                    End If


                    If (M01.Tables(0).Rows(i)("T01Reject")) > 0 Then

                        _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
                        If Trim(M01.Tables(0).Rows(i)("T01R_Reason")) <> "" Then
                            If _QReasons <> "" Then
                                _QReasons = _QReasons & "/" & (M01.Tables(0).Rows(i)("T01R_Reason"))
                            Else
                                _QReasons = (M01.Tables(0).Rows(i)("T01R_Reason"))
                            End If
                        End If
                    End If

                    'If Val(M01.Tables(0).Rows(i)("T01Reject")) > 0 Then
                    '    _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
                    '    If _QReasons <> "" Then
                    '        _QReasons = _QReasons & "/" & (M01.Tables(0).Rows(i)("T01R_Reason"))
                    '    Else
                    '        _QReasons = (M01.Tables(0).Rows(i)("T01R_Reason"))
                    '    End If
                    'End If

                    nvcFieldList = "select * from T06QuarantinePass where T06RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
                    M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                    If isValidDataset(M02) Then
                        _Cutofwt = _Cutofwt + Val(M02.Tables(0).Rows(0)("t06Scrab"))
                    End If

                    If _Cutofwt > 0 Or _3Mtrcutoff > 0 Then
                        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,T01R_Reason,R01R_Whight,R01Department)" & _
                                                                    " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _Reasone & "'," & _3Mtrcutoff & ",'" & _Department & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    End If
                    i = i + 1
                Next
            Else
                'Dim _FromTime As String
                'Dim _ToTime As String

                _FromTime = txtDate.Text & " " & txtTime1.Text
                _ToTime = txtTo.Text & " " & txtToTime.Text


                Sql = "select max(T01status) as T01status,max(T01RefNo) as T01RefNo,max(T01R_Reason) as T01R_Reason,max(T01Reject) as T01Reject,T01OrderNo ,T01RollNo from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "'  group by T01OrderNo,T01RollNo"
                M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                i = 0
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    _QReasons = ""
                    _Department = ""
                    If i = 55 Then
                        ' MsgBox("")
                    End If
                    Sql = "select * from T05Scrab where T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
                    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    X = 0
                    For Each DTRow2 As DataRow In M02.Tables(0).Rows
                        If X = 0 Then
                            _QReasons = Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
                            _Department = Trim(M02.Tables(0).Rows(X)("T05Department"))
                        Else
                            _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
                            _Department = _Department & "/" & Trim(M02.Tables(0).Rows(X)("T05Department"))
                        End If
                        X = X + 1
                    Next

                    _3Mtrcutoff = 0
                    _Cutofwt = 0

                    Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
                    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M02) Then
                        _Cutofwt = M02.Tables(0).Rows(0)("QTY")
                    End If


                    Sql = "select sum(T04Weight) as Qty from T04Cutoff where T04RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " group by T04RefNo"
                    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M02) Then
                        _3Mtrcutoff = M02.Tables(0).Rows(0)("QTY")
                    End If

                    Sql = "select * from T04Cutoff where T04RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
                    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    X = 0
                    For Each DTRow2 As DataRow In M02.Tables(0).Rows
                        If Microsoft.VisualBasic.Len(_QReasons) = 0 Then
                            _QReasons = Trim(M02.Tables(0).Rows(X)("T04Fualcode"))
                        Else
                            _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T04Fualcode"))
                        End If
                        X = X + 1
                    Next

                    '---------------------------------------------------------
                    If IsDBNull(M01.Tables(0).Rows(i)("T01Reject")) Then
                    Else
                        _Cutofwt = _Cutofwt

                        If IsDBNull(Val(M01.Tables(0).Rows(i)("T01Reject"))) Then
                        Else
                            _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
                        End If


                        If M01.Tables(0).Rows(i)("T01status") = "R" Then
                            If Trim(M01.Tables(0).Rows(i)("T01R_Reason")) <> "" Then
                                If _QReasons <> "" Then
                                    _QReasons = _QReasons & "/" & (M01.Tables(0).Rows(i)("T01R_Reason"))
                                Else
                                    _QReasons = (M01.Tables(0).Rows(i)("T01R_Reason"))
                                End If
                            End If
                        End If
                    End If

                    nvcFieldList = "select * from T06QuarantinePass where T06RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
                    M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
                    If isValidDataset(M02) Then
                        _Cutofwt = _Cutofwt + Val(M02.Tables(0).Rows(0)("t06Scrab"))
                    End If

                    If _Cutofwt > 0 Or _3Mtrcutoff > 0 Then
                        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,T01R_Reason,R01R_Whight,R01Department)" & _
                                                                         " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _Reasone & "'," & _3Mtrcutoff & ",'" & _Department & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    End If
                    i = i + 1
                Next

            End If
            MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            If chk1.Checked = False Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "' and {R01Report.R01CutWhight} > 0 or {R01Report.R01R_Whight} > 0"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            Else
                A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport3.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("From", _FromTime)
                B.SetParameterValue("To", _ToTime)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'" ' and {R01Report.R01CutWhight} > 0 or {R01Report.R01R_Whight} > 0"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            ' End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub UltraGroupBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox5.Click

    End Sub

    Private Sub UltraLabel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraLabel3.Click

    End Sub

    Private Sub frmCutoff_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class