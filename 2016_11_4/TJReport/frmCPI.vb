Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Public Class frmCPI
    Dim Clicked As String
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' cboDep.ToggleDropdown()
        chkS1.Checked = True

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


        'Dim nvcFieldList1 As String

        'Dim connection As SqlClient.SqlConnection
        'Dim transaction As SqlClient.SqlTransaction
        'Dim transactionCreated As Boolean
        'Dim connectionCreated As Boolean

        'connection = DBEngin.GetConnection(True)
        'connectionCreated = True
        'transaction = connection.BeginTransaction()
        'transactionCreated = True
        'Dim sql As String
        'Dim T01 As DataSet

        ''Dim Sql As String
        'Dim con = New SqlConnection()
        'con = DBEngin.GetConnection()

        'sql = "select * from T01Transaction_Header3"
        'T01 = DBEngin.ExecuteDataset(con, Nothing, sql)
        'Dim i As Integer
        'i = 0
        'For Each DTRow2 As DataRow In T01.Tables(0).Rows

        '    nvcFieldList1 = "Insert Into T01Transaction_Header(T01RefNo,T01OrderNo,T01RollNo,T01Barcode,T01Date,T01Time,T01KnitterEPF,T01InsEPF,T01MC,T0Rollweight,T01Rollwidth,T01RollLenth,T01R_Reason,T01Reject,T01Cutoff,T01Cutoffreason,T01Otherweight,T01TotalPoiint,T01Frate,T01OtherReason,T01Comment,T01Audit,T01Workstation,T01Status,T01Stop,T01Roll)" & _
        '                                                  " values(" & T01.Tables(0).Rows(i)("T01RefNo") & ", '" & Trim(T01.Tables(0).Rows(i)("T01OrderNo")) & "','" & Trim(T01.Tables(0).Rows(i)("T01RollNo")) & "','" & T01.Tables(0).Rows(i)("T01Barcode") & "','" & T01.Tables(0).Rows(i)("T01Date") & "','" & T01.Tables(0).Rows(i)("T01Time") & "','" & T01.Tables(0).Rows(i)("T01KnitterEPF") & "','" & T01.Tables(0).Rows(i)("T01InsEPF") & "','" & T01.Tables(0).Rows(i)("T01MC") & "','" & T01.Tables(0).Rows(i)("T0Rollweight") & "','" & T01.Tables(0).Rows(i)("T01Rollwidth") & "','" & T01.Tables(0).Rows(i)("T01RollLenth") & "','" & T01.Tables(0).Rows(i)("T01R_Reason") & "','" & T01.Tables(0).Rows(i)("T01Reject") & "','" & T01.Tables(0).Rows(i)("T01Cutoff") & "','" & T01.Tables(0).Rows(i)("T01Cutoffreason") & "','" & T01.Tables(0).Rows(i)("T01Otherweight") & "'," & T01.Tables(0).Rows(i)("T01TotalPoiint") & ",'" & T01.Tables(0).Rows(i)("T01Frate") & "','" & T01.Tables(0).Rows(i)("T01OtherReason") & "','" & T01.Tables(0).Rows(i)("T01Comment") & "','" & T01.Tables(0).Rows(i)("T01Audit") & "','" & T01.Tables(0).Rows(i)("T01Workstation") & "','" & T01.Tables(0).Rows(i)("T01Status") & "','" & T01.Tables(0).Rows(i)("T01Stop") & "'," & T01.Tables(0).Rows(i)("T01Roll") & ")"
        '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

        '    i = i + 1
        'Next

        'transaction.Commit()
        'MsgBox("")
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
        Dim _Shift As Integer
        Dim _FromTime As String
        Dim _ToTime As String

        Try

            If chk1.Checked = True And chkS1.Checked = True Then
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 7, 30, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 19, 30, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\CPIT1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptCpI} <> ''"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = True And chkS2.Checked = True Then
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 19, 30, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 7, 30, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\CPIT1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptCpI} <> ''"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = True And chkAll.Checked = True Then
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\CPIT1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01Date} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptCpI} <> ''"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = False And chkS1.Checked = True Then
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 7, 30, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 19, 30, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\CPIT1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptCpI} <> '' and {M03Knittingorder.M03Quality}='" & cboQuality.Text & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = False And chkS2.Checked = True Then
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 19, 30, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 7, 30, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\CPIT1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptCpI} <> '' and {M03Knittingorder.M03Quality}='" & cboQuality.Text & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf chk1.Checked = False And chkAll.Checked = True Then
                StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
                StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

                A = ConfigurationManager.AppSettings("ReportPath") + "\CPIT1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", txtTo.Value)
                B.SetParameterValue("From", txtDate.Value)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01time} in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {rptReport_Detailes.rptCpI} <> '' and {M03Knittingorder.M03Quality}='" & cboQuality.Text & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            'nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList)

            'Dim X As Integer
            'Dim _AvgCPI As Double


            'If chkS1.Checked = True Then
            '    _Shift = 1

            '    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            '    StrToDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"

            '    _FromTime = txtDate.Text & " " & "07:30:00"
            '    _ToTime = txtDate.Text & " " & "19:30:00"

            'ElseIf chkS2.Checked = True Then

            '    _Shift = 2

            '    _FromTime = txtDate.Text & " " & "19:30:00"
            '    _ToTime = System.DateTime.FromOADate(CDate(txtDate.Text).ToOADate + 1)
            '    _ToTime = _ToTime & " " & "07:30:00"

            '    txtTo.Text = _ToTime

            '    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            '    StrToDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            'ElseIf chkAll.Checked = True Then
            '    _FromTime = txtDate.Text & " " & "7:30:00"
            '    _ToTime = txtTo.Text
            '    _ToTime = _ToTime & " " & "7:30:00"

            '    txtTo.Text = _ToTime

            '    StrFromDate = "(" & Year(txtDate.Value) & ", " & VB6.Format(Month(txtDate.Value), "0#") & ", " & VB6.Format(CDate(txtDate.Text).Day, "0#") & ", 00, 00, 00)"
            '    StrToDate = "(" & Year(txtTo.Value) & ", " & VB6.Format(Month(txtTo.Value), "0#") & ", " & VB6.Format(CDate(txtTo.Text).Day, "0#") & ", 00, 00, 00)"

            '    ' _FromTime = txtDate.Text & " " & "07:30:00"
            '    '_ToTime = txtDate.Text & " " & "19:30:00"
            'End If
            'If chk1.Checked = True Then
            '    Sql = "select * from T01Transaction_Header where T01time between '" & _FromTime & "' and '" & _ToTime & "'  order by T01OrderNo,T01RollNo"
            'Else
            '    If cboQuality.Text <> "" Then
            '        Sql = "select * from T01Transaction_Header inner join M03Knittingorder on M03OrderNo=T01OrderNo  where T01time between '" & _FromTime & "' and '" & _ToTime & "' and M03Quality='" & cboQuality.Text & "' order by T01OrderNo,T01RollNo"
            '    End If
            'End If
            'M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            'i = 0
            'Dim _Max As Double
            'Dim _Min As Double
            'Dim _COUNT As Integer
            'For Each DTRow1 As DataRow In M01.Tables(0).Rows
            '    _COUNT = 0
            '    X = 0

            '    Sql = "select * from T03CPI_Reading where T03RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " "
            '    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '    If isValidDataset(M02) Then
            '        If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) > 0 Then
            '            _COUNT = 1
            '        End If
            '        If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) > 0 Then
            '            _COUNT = _COUNT + 1
            '        End If
            '        If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) > 0 Then
            '            _COUNT = _COUNT + 1
            '        End If
            '        If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) > 0 Then
            '            _COUNT = _COUNT + 1
            '        End If
            '        If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) > 0 Then
            '            _COUNT = _COUNT + 1
            '        End If

            '        X = X + 1
            '    End If

            '    Sql = "select * from T03CPI_Reading where T03RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & ""
            '    M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '    X = 0
            '    If isValidDataset(M02) Then

            '        _QReasons = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))
            '        _AvgCPI = Trim(M02.Tables(0).Rows(X)("T03CPIV1"))
            '        If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) Then
            '            _QReasons = _QReasons & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '            _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '        End If

            '        If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) Then
            '            _QReasons = _QReasons & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '            _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '        End If

            '        If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) Then
            '            _QReasons = _QReasons & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '            _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '        End If

            '        If Not IsDBNull(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) Then
            '            _QReasons = _QReasons & "/" & CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '            _AvgCPI = _AvgCPI + CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '        End If

            '        If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) >= Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) Then
            '            _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '        Else
            '            _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))
            '        End If

            '        If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) Then
            '            _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '        End If

            '        If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) Then
            '            _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '        End If

            '        If _Max < CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) Then
            '            _Max = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '        End If

            '        If Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) = 0 Then
            '            _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '        Else
            '            If CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2"))) >= Val(Trim(M02.Tables(0).Rows(X)("T03CPIV1"))) Then
            '                _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV1")))
            '            Else
            '                _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV2")))
            '            End If
            '        End If


            '        If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3"))) <> 0 Then
            '            _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV3")))
            '        End If

            '        If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4"))) <> 0 Then
            '            _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV4")))
            '        End If

            '        If _Min > CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) And CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5"))) <> 0 Then
            '            _Min = CInt(Trim(M02.Tables(0).Rows(X)("T03CPIV5")))
            '        End If

            '        If _Max = 0 And _Min = 0 Then
            '            nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01R_Whight,R01Status)" & _
            '                                                  " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','0','" & _Max - _Min & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '        Else
            '            nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01RollNo,R01Reason,R01CutWhight,R01WorkStation,R01R_Whight,R01Status)" & _
            '                                                  " values('" & M01.Tables(0).Rows(i)("T01RefNo") & "', '" & M01.Tables(0).Rows(i)("T01OrderNo") & "','" & M01.Tables(0).Rows(i)("T01RollNo") & "','" & _QReasons & "'," & _Cutofwt & ",'" & netCard & "','" & _AvgCPI / _COUNT & "','" & _Max - _Min & "')"
            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            '        End If
            '    End If
            '    ''-----------------------------------------------------
            '    'Sql = "SELECT SUM(T05Weight) AS QTY FROM T05Scrab  WHERE T05RefNo=" & M01.Tables(0).Rows(i)("T01RefNo") & " GROUP BY T05RefNo"
            '    'M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            '    'If isValidDataset(M02) Then
            '    '    _Cutofwt = M02.Tables(0).Rows(0)("QTY")
            '    'End If
            '    ''---------------------------------------------------------
            '    'If IsDBNull(M01.Tables(0).Rows(i)("T01Reject")) Then
            '    'Else
            '    '    _Cutofwt = _Cutofwt + Val(M01.Tables(0).Rows(i)("T01Reject"))
            '    'End If

            '    i = i + 1
            'Next
            'MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
            'transaction.Commit()
            'DBEngin.CloseConnection(connection)
            'A = ConfigurationManager.AppSettings("ReportPath") + "\CPI.rpt"
            'B.Load(A.ToString)
            'B.SetDatabaseLogon("sa", "sainfinity")
            'B.SetParameterValue("To", txtTo.Value)
            'B.SetParameterValue("From", txtDate.Value)
            ''  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            'frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            'frmReport.CrystalReportViewer1.DisplayToolbar = True
            'If chk1.Checked = True Or chkAll.Checked = True Then
            '    frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01WorkStation}='" & netCard & "'"
            'Else
            '    frmReport.CrystalReportViewer1.SelectionFormula = "{M03Knittingorder.M03Quality}='" & Trim(cboQuality.Text) & "' and {R01Report.R01WorkStation}='" & netCard & "'"
            'End If
            'frmReport.Refresh()
            '' frmReport.CrystalReportViewer1.PrintReport()
            '' B.PrintToPrinter(1, True, 0, 0)
            'frmReport.MdiParent = MDIMain
            'frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                '   MsgBox(i)
            End If
        End Try
    End Sub

    Private Sub frmCPI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim recArea As DataSet
        Dim M01 As DataSet

        Try
            'SET COMPANY
            Sql = "select M03Quality as [Quality] from M03Knittingorder group by M03Quality"
            recArea = DBEngin.ExecuteDataset(con, Nothing, Sql)
            cboQuality.DataSource = recArea
            cboQuality.Rows.Band.Columns(0).Width = 370
            ' cboSupp.Rows.Band.Columns(1).Width = 170

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

            txtDate.Text = Today
            txtTo.Text = Today

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub chkS1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkS1.CheckedChanged
        If chkAll.Checked = True Then
        Else
            If chkS1.Checked = True Then
                chkS2.Checked = False
            Else
                chkS2.Checked = True
            End If
        End If
    End Sub

    Private Sub chkS2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkS2.CheckedChanged
        If chkAll.Checked = True Then
        Else
            If chkS2.Checked = True Then
                chkS1.Checked = False
            Else
                chkS1.Checked = True
            End If
        End If
    End Sub

    Private Sub cboQuality_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboQuality.InitializeLayout

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
        Dim T01 As DataSet


        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True


        Sql = "select * from tmpSAP_Transfer inner join T01Transaction_Header on tmpRef=T01RefNo where T01Date='6/10/2013' and tmpStatus='P'"
        T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        For Each DTRow1 As DataRow In T01.Tables(0).Rows
            nvcFieldList = "Update T01Transaction_Header set T01Status='P' where T01RefNo=" & T01.Tables(0).Rows(i)("T01RefNo") & " and T01Status='QX'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)
            i = i + 1
        Next
        MsgBox("OK")
        transaction.Commit()
    End Sub
End Class