Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmReprocess
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

        Try
            If chk1.Checked = True Then
                Dim Y As Date
                Dim X As Date

                Y = txtDate.Text & " " & txtTime1.Text
                X = txtTo.Text & " " & txtToTime.Text
                '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"



                A = ConfigurationManager.AppSettings("ReportPath") + "\ReprocessT.rpt"
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
            Else
                Dim Y As Date
                Dim X As Date

                'Y = txtDate.Text & " " & txtTime1.Text
                'X = txtTo.Text & " " & txtToTime.Text
                ' '' MsgBox(Format(Y.ToString("HH:mm:ss")))
                'StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & "," & Hour(Y) & "," & Minute(Y) & ", 00)"
                'StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & "," & Hour(X) & "," & Minute(X) & ", 00)"

                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\ReprocessT.rpt"
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
            End If
            'nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList)

            'Dim X As Integer
            'If chk1.Checked = False Then
            '    Sql = "select * from T01Transaction_Header where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and  T01Status ='RP'  order by T01OrderNo,T01RollNo"
            '    M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '    i = 0
            '    For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '        Sql = "select M02Name from T07Q_Reason inner join M02Fault on M02Code=T07F_Code where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        X = 0
            '        For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '            If X = 0 Then
            '                _QReasons = Trim(M02.Tables(0).Rows(X)("M02Name"))
            '            Else
            '                _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("M02Name"))
            '            End If
            '            X = X + 1

            '        Next

            '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation)" & _
            '                                                    " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'" & netCard & "')"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '        i = i + 1
            '    Next
            'Else
            '    'Dim _FromTime As String
            '    'Dim _ToTime As String

            '    _FromTime = txtDate.Text & " " & txtTime1.Text
            '    _ToTime = txtTo.Text & " " & txtToTime.Text


            '    Sql = "select * from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status ='RP' order by T01OrderNo,T01RollNo"
            '    M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '    i = 0
            '    For Each DTRow1 As DataRow In M01.Tables(0).Rows

            '        Sql = "select M02Name from T07Q_Reason inner join M02Fault on M02Code=T07F_Code where T07RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '        X = 0
            '        For Each DTRow2 As DataRow In M02.Tables(0).Rows
            '            If X = 0 Then
            '                _QReasons = Trim(M02.Tables(0).Rows(X)("M02Name"))
            '            Else
            '                _QReasons = _QReasons & "/" & Trim(M02.Tables(0).Rows(X)("M02Name"))
            '            End If
            '            X = X + 1

            '        Next



            '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation)" & _
            '                                                         " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & M01.Tables(0).Rows(i)("T0Rollweight") & ",'" & netCard & "')"
            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

            '        i = i + 1
            '    Next

            'End If
            'MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            'transaction.Commit()
            'DBEngin.CloseConnection(connection)
            'If chk1.Checked = False Then
            '    A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport8.rpt"
            '    B.Load(A.ToString)
            '   B.SetDatabaseLogon("sa", "tommya")
            '    B.SetParameterValue("To", txtTo.Value)
            '    B.SetParameterValue("From", txtDate.Value)
            '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '    frmReport.CrystalReportViewer1.DisplayToolbar = True
            '    frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            '    frmReport.Refresh()
            '    ' frmReport.CrystalReportViewer1.PrintReport()
            '    ' B.PrintToPrinter(1, True, 0, 0)
            '    frmReport.MdiParent = MDIMain
            '    frmReport.Show()
            'Else
            '    A = ConfigurationManager.AppSettings("ReportPath") + "\InspectionReport8.rpt"
            '    B.Load(A.ToString)
            '   B.SetDatabaseLogon("sa", "tommya")
            '    B.SetParameterValue("From", _FromTime)
            '    B.SetParameterValue("To", _ToTime)
            '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '    frmReport.CrystalReportViewer1.DisplayToolbar = True
            '    frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            '    frmReport.Refresh()
            '    ' frmReport.CrystalReportViewer1.PrintReport()
            '    ' B.PrintToPrinter(1, True, 0, 0)
            '    frmReport.MdiParent = MDIMain
            '    frmReport.Show()
            'End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

End Class