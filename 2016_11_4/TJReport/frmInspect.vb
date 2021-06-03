Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmInspect
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

                A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport4T.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()

            ElseIf chk2.Checked = True Then
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport7T.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01OrderNo}='" & cboOrder.Text & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            Else

                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport1T.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If


            'Deactivate Code
            'If Trim(strDisname) = "KNT" Then
            '    If chk1.Checked = True Then
            '        _FromTime = txtDate.Text & " " & txtTime1.Text
            '        _ToTime = txtTo.Text & " " & txtToTime.Text

            '        nvcFieldList = "delete from tmpR01Report where R01WorkStation='" & netCard & "'"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '        Dim X As Integer
            '        Sql = "select * from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "'  order by T01OrderNo,T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select * from T05Scrab where T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons1 = Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                Else
            '                    _QReasons1 = _QReasons1 & "/" & Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                End If
            '                X = X + 1

            '            Next

            '            Sql = "select * from T02Trans_Fault where T02Ref=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                End If
            '                X = X + 1

            '            Next
            '            '-----------------------------------------------------
            '            _Cutofwt = 0
            '            Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            If isValidDataset(M02) Then
            '                _Cutofwt = M02.Tables(0).Rows(0)("QTY")
            '            End If
            '            '---------------------------------------------------------
            '            If IsDBNull(M01.Tables(0).Rows(i)("T01Reject")) Then
            '            Else
            '                _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
            '            End If
            '            '-----------------------------------------------------------------
            '            'LESS THAN 3M CUTOFF DEVELOPED BY SURANGA ON 2012.07.14

            '            nvcFieldList = "Insert Into tmpR01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,T01R_Reason)" & _
            '                                                        " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _QReasons1 & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            i = i + 1
            '        Next
            '        MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            '        transaction.Commit()
            '        DBEngin.CloseConnection(connection)
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReporttmp4.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", _ToTime)
            '        B.SetParameterValue("From", _FromTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{tmpR01Report.R01WorkStation}='" & netCard & "'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk2.Checked = True Then
            '        nvcFieldList = "delete from tmpR01Report where R01WorkStation='" & netCard & "'"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '        Dim X As Integer
            '        Sql = "select * from T01Transaction_Header where T01OrderNo = '" & Trim(cboOrder.Text) & "'  order by T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select * from T05Scrab where T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons1 = Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                Else
            '                    _QReasons1 = _QReasons1 & "/" & Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                End If
            '                X = X + 1

            '            Next

            '            Sql = "select * from T02Trans_Fault where T02Ref=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                End If
            '                X = X + 1

            '            Next
            '            '-----------------------------------------------------
            '            _Cutofwt = 0
            '            Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            If isValidDataset(M02) Then
            '                _Cutofwt = M02.Tables(0).Rows(0)("QTY")
            '            End If
            '            '---------------------------------------------------------
            '            If IsDBNull(M01.Tables(0).Rows(i)("T01Reject")) Then
            '            Else
            '                _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
            '            End If

            '            Dim L_Qreason As String
            '            L_Qreason = ""
            '            Sql = "select * from T07Q_Reason where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    L_Qreason = Trim(M02.Tables(0).Rows(X)("T07F_Code"))
            '                Else
            '                    L_Qreason = L_Qreason & "/" & Trim(M02.Tables(0).Rows(X)("T07F_Code"))
            '                End If
            '                X = X + 1

            '            Next

            '            nvcFieldList = "Insert Into tmpR01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,T01R_Reason,R01Q_Reson)" & _
            '                                                        " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _QReasons1 & "','" & L_Qreason & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            i = i + 1
            '        Next

            '        MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            '        transaction.Commit()

            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport7.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        '  B.SetParameterValue("To", txtTo.Value)
            '        '  B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()
            '    Else
            '        nvcFieldList = "delete from tmpR01Report where R01WorkStation='" & netCard & "'"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '        Dim X As Integer
            '        Sql = "select * from T01Transaction_Header where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "'  order by T01OrderNo,T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select * from T05Scrab where T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons1 = Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                Else
            '                    _QReasons1 = _QReasons1 & "/" & Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                End If
            '                X = X + 1

            '            Next

            '            Sql = "select * from T02Trans_Fault where T02Ref=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                End If
            '                X = X + 1

            '            Next
            '            '-----------------------------------------------------
            '            _Cutofwt = 0

            '            Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            If isValidDataset(M02) Then
            '                _Cutofwt = M02.Tables(0).Rows(0)("QTY")
            '            End If
            '            '---------------------------------------------------------
            '            If IsDBNull(M01.Tables(0).Rows(i)("T01Reject")) Then
            '            Else
            '                _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
            '            End If
            '            Dim L_Qreason As String

            '            Sql = "select * from T07Q_Reason where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    L_Qreason = Trim(M02.Tables(0).Rows(X)("T07F_Code"))
            '                Else
            '                    L_Qreason = L_Qreason & "/" & Trim(M02.Tables(0).Rows(X)("T07F_Code"))
            '                End If
            '                X = X + 1

            '            Next


            '            nvcFieldList = "Insert Into tmpR01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,T01R_Reason,R01Q_Reson)" & _
            '                                                        " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _QReasons1 & "','" & L_Qreason & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            i = i + 1
            '        Next
            '        MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            '        transaction.Commit()

            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport1.rpt"
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
            '    End If

            'Else

            '    If chk1.Checked = True Then
            '        _FromTime = txtDate.Text & " " & txtTime1.Text
            '        _ToTime = txtTo.Text & " " & txtToTime.Text

            '        nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '        Dim X As Integer
            '        Sql = "select * from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "'  order by T01OrderNo,T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select * from T05Scrab where T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons1 = Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                Else
            '                    _QReasons1 = _QReasons1 & "/" & Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                End If
            '                X = X + 1

            '            Next

            '            Sql = "select * from T02Trans_Fault where T02Ref=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                End If
            '                X = X + 1

            '            Next
            '            '-----------------------------------------------------
            '            _Cutofwt = 0
            '            Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            If isValidDataset(M02) Then
            '                _Cutofwt = M02.Tables(0).Rows(0)("QTY")
            '            End If
            '            '---------------------------------------------------------
            '            If IsDBNull(M01.Tables(0).Rows(i)("T01Reject")) Then
            '            Else
            '                _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
            '            End If
            '            '-----------------------------------------------------------------
            '            'LESS THAN 3M CUTOFF DEVELOPED BY SURANGA ON 2012.07.14

            '            nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,T01R_Reason)" & _
            '                                                        " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _QReasons1 & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            i = i + 1
            '        Next
            '        MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            '        transaction.Commit()
            '        DBEngin.CloseConnection(connection)
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport4.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        B.SetParameterValue("To", _ToTime)
            '        B.SetParameterValue("From", _FromTime)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()

            '    ElseIf chk2.Checked = True Then
            '        nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '        Dim X As Integer
            '        Sql = "select * from T01Transaction_Header where T01OrderNo = '" & Trim(cboOrder.Text) & "'  order by T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select * from T05Scrab where T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons1 = Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                Else
            '                    _QReasons1 = _QReasons1 & "/" & Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                End If
            '                X = X + 1

            '            Next

            '            Sql = "select * from T02Trans_Fault where T02Ref=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                End If
            '                X = X + 1

            '            Next
            '            '-----------------------------------------------------
            '            _Cutofwt = 0
            '            Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            If isValidDataset(M02) Then
            '                _Cutofwt = M02.Tables(0).Rows(0)("QTY")
            '            End If
            '            '---------------------------------------------------------
            '            If IsDBNull(M01.Tables(0).Rows(i)("T01Reject")) Then
            '            Else
            '                _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
            '            End If

            '            Dim L_Qreason As String
            '            L_Qreason = ""
            '            Sql = "select * from T07Q_Reason where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    L_Qreason = Trim(M02.Tables(0).Rows(X)("T07F_Code"))
            '                Else
            '                    L_Qreason = L_Qreason & "/" & Trim(M02.Tables(0).Rows(X)("T07F_Code"))
            '                End If
            '                X = X + 1

            '            Next

            '            nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,T01R_Reason,R01Q_Reson)" & _
            '                                                        " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _QReasons1 & "','" & L_Qreason & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            i = i + 1
            '        Next

            '        MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            '        transaction.Commit()
            '        DBEngin.CloseConnection(connection)
            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport7.rpt"
            '        B.Load(A.ToString)
            '       B.SetDatabaseLogon("sa", "tommya")
            '        '  B.SetParameterValue("To", txtTo.Value)
            '        '  B.SetParameterValue("From", txtDate.Value)
            '        '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '        frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '        frmReport.CrystalReportViewer1.DisplayToolbar = True
            '        frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            '        frmReport.Refresh()
            '        ' frmReport.CrystalReportViewer1.PrintReport()
            '        ' B.PrintToPrinter(1, True, 0, 0)
            '        frmReport.MdiParent = MDIMain
            '        frmReport.Show()
            '    Else
            '        nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '        Dim X As Integer
            '        Sql = "select * from T01Transaction_Header where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "'  order by T01OrderNo,T01RollNo"
            '        M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        i = 0
            '        For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '            Sql = "select * from T05Scrab where T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons1 = Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                Else
            '                    _QReasons1 = _QReasons1 & "/" & Trim(M02.Tables(0).Rows(X)("T05Fualcode"))
            '                End If
            '                X = X + 1

            '            Next

            '            Sql = "select * from T02Trans_Fault where T02Ref=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    _QReasons = Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                Else
            '                    _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("T02FualtCode"))
            '                End If
            '                X = X + 1

            '            Next
            '            '-----------------------------------------------------
            '            _Cutofwt = 0

            '            Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            If isValidDataset(M02) Then
            '                _Cutofwt = M02.Tables(0).Rows(0)("QTY")
            '            End If
            '            '---------------------------------------------------------
            '            If IsDBNull(M01.Tables(0).Rows(i)("T01Reject")) Then
            '            Else
            '                _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
            '            End If
            '            Dim L_Qreason As String

            '            Sql = "select * from T07Q_Reason where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '            M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '            X = 0
            '            For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '                If X = 0 Then
            '                    L_Qreason = Trim(M02.Tables(0).Rows(X)("T07F_Code"))
            '                Else
            '                    L_Qreason = L_Qreason & "/" & Trim(M02.Tables(0).Rows(X)("T07F_Code"))
            '                End If
            '                X = X + 1

            '            Next


            '            nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,T01R_Reason,R01Q_Reson)" & _
            '                                                        " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _QReasons1 & "','" & L_Qreason & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '            i = i + 1
            '        Next
            '        MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            '        transaction.Commit()

            '        A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport1.rpt"
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
            '    End If

            'End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
    Function Load_Order()
        'load NSL Combo box
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03OrderNo as [Order No],max(M03Quality) as [Quality No],max(M03Material) as [Material],max(M03Yarnstock) as [Yarn Stock Code] from M03Knittingorder  group by M03OrderNo"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboOrder
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 190
                .Rows.Band.Columns(1).Width = 90
                .Rows.Band.Columns(2).Width = 90
                .Rows.Band.Columns(3).Width = 240
                ' .Rows.Band.Columns(4).Width = 110
                '  .Rows.Band.Columns(5).Width = 110

            End With

            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub chk1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk1.CheckedChanged
        If chk1.Checked = True Then
            chk2.Checked = False
        Else
            chk2.Checked = True
        End If
    End Sub

    Private Sub chk2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk2.CheckedChanged
        If chk2.Checked = True Then
            chk1.Checked = False
        Else
            chk1.Checked = True
        End If
    End Sub

    Private Sub frmInspect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Order()
    End Sub
End Class