Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Imports System.IO.StreamWriter
Imports Microsoft.Office.Interop.Excel
Public Class frmKnittingProduction
    Dim Clicked As String
    Dim exc As New Application

    Dim workbooks As Workbooks = exc.Workbooks
    Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    Dim sheets As Sheets = Workbook.Worksheets
    Dim worksheet As _Worksheet = CType(Sheets.Item(1), _Worksheet)
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
        'Dim B As New ReportDocument
        'Dim A As String

        'Dim connection As SqlClient.SqlConnection
        'Dim transaction As SqlClient.SqlTransaction
        'Dim transactionCreated As Boolean
        'Dim connectionCreated As Boolean
        'Dim Sql As String

        'connection = DBEngin.GetConnection(True)
        'connectionCreated = True
        'transaction = connection.BeginTransaction()
        'transactionCreated = True
        'Dim i As Integer

        'Dim M03 As DataSet
        'Dim T01 As DataSet

        'Dim _Scrap As Double
        'Dim _Qty As Double
        'Dim _Quarantine As Double
        'Dim _Reject As Double
        'Dim nvcFieldList As String
        'Dim _FromTime As String
        'Dim _ToTime As String
        'Dim R05 As DataSet
        'Dim T05 As DataSet
        'Dim TX As DataSet

        'Dim n_Usableqty As Double
        'i = 0
        'Try
        '    Sql = "delete from R05Report where R05id='" & netCard & "' "
        '    ExecuteNonQueryText(connection, transaction, Sql)


        '    _FromTime = txtDate.Text & " " & txtTime1.Text
        '    _ToTime = txtTo.Text & " " & txtToTime.Text


        '    If Trim(txtM1.Text) <> "" Then
        '        If Trim(txtM2.Text) <> "" Then
        '            nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('P','QP') AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
        '            T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
        '            i = 0

        '            'Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & M03.Tables(0).Rows(i)("M03OrderNo") & "' and T01Status in ('P','QP') group by T01OrderNo"
        '            'T01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            'If isValidDataset(T01) Then
        '            '    _Qty = T01.Tables(0).Rows(0)("T0Rollweight")
        '            'End If

        '            'USABLE QTY
        '            For Each DTRow1 As DataRow In T01.Tables(0).Rows
        '                _Qty = 0
        '                'Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and T01Status in ('P','QP') group by T01OrderNo"
        '                'TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                'If isValidDataset(TX) Then
        '                '    _Qty = TX.Tables(0).Rows(0)("T0Rollweight")
        '                'End If
        '                ' n_Usableqty = T01.Tables(0).Rows(0)("T0Rollweight")
        '                nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
        '                                                      " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','" & T01.Tables(0).Rows(i)("T0Rollweight") & "','0','0','0','" & netCard & "')"
        '                ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '                i = i + 1
        '            Next
        '            '---------------------------------------------------------
        '            'SCAP
        '            nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('R','QR') AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
        '            T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
        '            i = 0
        '            For Each DTRow1 As DataRow In T01.Tables(0).Rows
        '                '_Scrap = 0
        '                'Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and T01Status in ('QR','R') group by T01OrderNo"
        '                'TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                'If isValidDataset(TX) Then
        '                '    _Scrap = TX.Tables(0).Rows(0)("T0Rollweight")
        '                'End If

        '                Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
        '                R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                If isValidDataset(R05) Then
        '                    nvcFieldList = "update R05Report set R05Scrap='" & T01.Tables(0).Rows(i)("T0Rollweight") & "' where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
        '                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '                Else
        '                    nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
        '                                                     " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','" & T01.Tables(0).Rows(i)("T0Rollweight") & "','0','0','" & netCard & "')"
        '                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '                End If



        '                i = i + 1
        '            Next

        '            i = 0
        '            '--------------------------------------------------------
        '            nvcFieldList = "select M03MCNo,T01OrderNo,SUM(T05Weight) AS T05Weight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo INNER JOIN T05Scrab ON T01RefNo=T05RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status<>'I' AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
        '            T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
        '            For Each DTRow1 As DataRow In T01.Tables(0).Rows
        '                '_Scrap = 0
        '                'Sql = "SELECT SUM(T05Weight) AS T05Weight FROM T05Scrab INNER JOIN T01Transaction_Header ON T01RefNo=T05RefNo WHERE T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' GROUP BY T05RefNo"
        '                'TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                'If isValidDataset(TX) Then
        '                '    _Scrap = TX.Tables(0).Rows(0)("T05Weight")
        '                'End If
        '                'Sql = "SELECT SUM(T05Weight) AS T05Weight FROM T05Scrab INNER JOIN T01Transaction_Header ON T01RefNo=T05RefNo WHERE T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' GROUP BY T05RefNo"
        '                'T05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                'If isValidDataset(T05) Then
        '                Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
        '                R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                If isValidDataset(R05) Then
        '                    nvcFieldList = "update R05Report set R05Scrap=R05Scrap +" & T01.Tables(0).Rows(i)("T05Weight") & " where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
        '                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '                Else
        '                    nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
        '                                                     " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','" & T01.Tables(0).Rows(i)("T05Weight") & "','0','0','" & netCard & "')"
        '                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '                End If

        '                'End If
        '                i = i + 1
        '            Next
        '            '---------------------------------------------------------------------------------
        '            i = 0
        '            '--------------------------------------------------------
        '            'nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status<>'I' AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
        '            'T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
        '            'For Each DTRow1 As DataRow In T01.Tables(0).Rows
        '            '    Sql = "select SUM(T04Weight) as T04Weight from T01Transaction_Header inner join T04Cutoff on T04RefNo=T01RefNo where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' group by T01OrderNo "
        '            '    T05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            '    If isValidDataset(T05) Then

        '            '        Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
        '            '        R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            '        If isValidDataset(R05) Then
        '            '            nvcFieldList = "update R05Report set R05Scrap=R05Scrap +" & T05.Tables(0).Rows(0)("T04Weight") & " where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
        '            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '            '        Else
        '            '            nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
        '            '                                             " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','" & T05.Tables(0).Rows(0)("T04Weight") & "','0','0','" & netCard & "')"
        '            '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '            '        End If
        '            '    End If
        '            '    'End If
        '            '    i = i + 1
        '            'Next
        '            '--------------------------------------------------------------------------------------
        '            'YARN SCRAP

        '            nvcFieldList = "select M03MCNo,T01OrderNo,SUM(T05Weight) AS T05Weight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo INNER JOIN T05Scrab ON T01RefNo=T05RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status<>'I' AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' AND T05Department='Knitting' group by M03MCNo,T01OrderNo"
        '            T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
        '            For Each DTRow1 As DataRow In T01.Tables(0).Rows
        '                '_Scrap = 0
        '                'Sql = "SELECT SUM(T05Weight) AS T05Weight FROM T05Scrab INNER JOIN T01Transaction_Header ON T01RefNo=T05RefNo WHERE T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' GROUP BY T05RefNo"
        '                'TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                'If isValidDataset(TX) Then
        '                '    _Scrap = TX.Tables(0).Rows(0)("T05Weight")
        '                'End If
        '                'Sql = "SELECT SUM(T05Weight) AS T05Weight FROM T05Scrab INNER JOIN T01Transaction_Header ON T01RefNo=T05RefNo WHERE T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' GROUP BY T05RefNo"
        '                'T05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                'If isValidDataset(T05) Then
        '                Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
        '                R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '                If isValidDataset(R05) Then
        '                    nvcFieldList = "update R05Report set R05Yarn=R05Yarn +" & T01.Tables(0).Rows(i)("T05Weight") & " where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
        '                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '                Else
        '                    nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
        '                                                     " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','0','0','" & T01.Tables(0).Rows(i)("T05Weight") & "','" & netCard & "')"
        '                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '                End If

        '                'End If
        '                i = i + 1
        '            Next

        '            ''QUARANTINE
        '            'nvcFieldList = "select M03MCNo,T01OrderNo,sum(T0Rollweight) as T0Rollweight from T01Transaction_Header  inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('Q') AND M03MCNo BETWEEN '" & txtM1.Text & "' AND '" & txtM2.Text & "' group by M03MCNo,T01OrderNo"
        '            'T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList)
        '            'i = 0
        '            'For Each DTRow1 As DataRow In T01.Tables(0).Rows
        '            '    _Quarantine = 0

        '            '    Sql = "select sum(T0Rollweight) as T0Rollweight from T01Transaction_Header where T01OrderNo='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and T01Status in ('RP','QR','Q') group by T01OrderNo"
        '            '    TX = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            '    If isValidDataset(TX) Then
        '            '        _Quarantine = TX.Tables(0).Rows(0)("T0Rollweight")
        '            '    End If


        '            '    Sql = "select * from R05Report where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05Id='" & netCard & "'"
        '            '    R05 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            '    If isValidDataset(R05) Then
        '            '        nvcFieldList = "update R05Report set R05Quarantine='" & _Quarantine & "' where R05Order='" & T01.Tables(0).Rows(i)("T01OrderNo") & "' and R05ID='" & netCard & "'"
        '            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '            '    Else
        '            '        nvcFieldList = "Insert Into R05Report(R05Order,R05MC,R05Usable,R05Scrap,R05Quarantine,R05Yarn,R05ID)" & _
        '            '                                         " values('" & T01.Tables(0).Rows(i)("T01OrderNo") & "', '" & T01.Tables(0).Rows(i)("M03MCNo") & "','0','0','" & _Quarantine & "','0','" & netCard & "')"
        '            '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '            '    End If



        '            '    i = i + 1
        '            'Next

        '            '  MsgBox("Records Update sucessfully", MsgBoxStyle.Information, "Information .....")
        '            transaction.Commit()
        '        Else
        '            MsgBox("Please enter the To Machine", MsgBoxStyle.Information, "Information ....")
        '            Exit Sub

        '        End If
        '    Else
        '        MsgBox("Please enter the form Machine", MsgBoxStyle.Information, "Information ....")
        '        Exit Sub
        '    End If
        '    Sql = ""

        '    MsgBox("Report Genarating Successfully", MsgBoxStyle.Information, "Report Genarating ........")
        '    ' transaction.Commit()
        '    A = ConfigurationManager.AppSettings("ReportPath") + "\KPShift.rpt"
        '    B.Load(A.ToString)
        '   B.SetDatabaseLogon("sa", "tommya")
        '    B.SetParameterValue("To", _ToTime)
        '    B.SetParameterValue("From", _FromTime)
        '    B.SetParameterValue("M/C", txtM1.Text & " - " & txtM2.Text)
        '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
        '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
        '    frmReport.CrystalReportViewer1.DisplayToolbar = True
        '    frmReport.CrystalReportViewer1.SelectionFormula = "{R05Report.R05id}='" & netCard & "'"
        '    frmReport.Refresh()
        '    ' frmReport.CrystalReportViewer1.PrintReport()
        '    ' B.PrintToPrinter(1, True, 0, 0)
        '    frmReport.MdiParent = MDIMain
        '    frmReport.Show()

        'Catch returnMessage As Exception
        '    If returnMessage.Message <> Nothing Then
        '        MessageBox.Show(returnMessage.Message)
        '    End If
        'End Try

        Call Create_Report1()
        Call Create_Report()
    End Sub

    Function Create_Report1()
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

        Dim _FromTime As String
        Dim _ToTime As String
        Dim nvcFieldList As String

        Try
            Sql = "delete from R06Report where R06ws='" & netCard & "' "
            ExecuteNonQueryText(connection, transaction, Sql)


            _FromTime = txtDate.Text & " " & txtTime1.Text
            _ToTime = txtTo.Text & " " & txtToTime.Text

            If Trim(txtM1.Text) <> "" And Trim(txtM2.Text) <> "" Then

            Else
                MsgBox("Please enter the machine no", MsgBoxStyle.Information, "Information ....")
                Exit Function
            End If

            Dim M03 As DataSet
            Dim i As Integer
            Dim X1 As Integer
            Dim _Topyarn As Double
            Dim _Totqty As Double
            Dim _TotScrap As Double
            Dim M02 As DataSet
            Dim X As Integer
            Dim _QTY As Double
            Dim _SCRAP As Double
            Dim _YARN As Double
            Dim M01 As DataSet

            Sql = "select M03MCNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo BETWEEN '" & Trim(txtM1.Text) & "' AND '" & Trim(txtM2.Text) & "' group by M03MCNo"
            M03 = DBEngin.ExecuteDataset(connection, transaction, Sql)

            i = 0
            X1 = 7
            _Topyarn = 0
            _Totqty = 0
            _TotScrap = 0

            _Topyarn = 0
            _Totqty = 0
            _TotScrap = 0

            For Each DTRow1 As DataRow In M03.Tables(0).Rows


                'worksheet.Cells(X1, 1) = M03.Tables(0).Rows(i)("M03MCNo")
                'range1 = worksheet.Cells(X1, 1)
                'range1.Interior.Color = RGB(255, 192, 255)
                'USABLE QTY
                Sql = "select M03MCNo,M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' group by M03Quality,M03Description,M03OrderNo,M03MCNo"
                M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                X = 0
                'USABLE QTY
                For Each DTRow2 As DataRow In M02.Tables(0).Rows

                    _QTY = 0
                    _SCRAP = 0
                    _YARN = 0
                    Sql = "select M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' AND M03OrderNo='" & M02.Tables(0).Rows(X)("M03OrderNo") & "' group by M03Quality,M03Description,M03OrderNo"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M01) Then
                        _QTY = Val(M01.Tables(0).Rows(0)("T0Rollweight"))
                        _Totqty = _Totqty + _QTY
                    End If

                    'SCRAP
                    Sql = "select M03OrderNo,M03Description,sum(T05Weight) as T05Weight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo INNER JOIN T05Scrab ON T05RefNo=T01RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' AND M03OrderNo='" & M02.Tables(0).Rows(X)("M03OrderNo") & "' group by M03Quality,M03Description,M03OrderNo"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M01) Then
                        _SCRAP = Val(M01.Tables(0).Rows(0)("T05Weight"))
                        _TotScrap = _TotScrap + _SCRAP
                    End If

                    'YARN SCRAP
                    'Sql = "select M03OrderNo,M03Description,sum(T05Weight) as T05Weight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo INNER JOIN T05Scrab ON T05RefNo=T01RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' AND M03OrderNo='" & M02.Tables(0).Rows(X)("M03OrderNo") & "' AND T05Department='Knitting' group by M03Quality,M03Description,M03OrderNo"
                    'M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    'If isValidDataset(M01) Then
                    '    _YARN = Val(M01.Tables(0).Rows(0)("T05Weight"))
                    '    _Topyarn = _Topyarn + _YARN
                    'End If

                    'QUARANTINE

                    Sql = "select M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Status ='Q' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' AND M03OrderNo='" & M02.Tables(0).Rows(X)("M03OrderNo") & "' group by M03Quality,M03Description,M03OrderNo"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                    If isValidDataset(M01) Then
                        _YARN = Val(M01.Tables(0).Rows(0)("T0Rollweight"))
                        _Topyarn = _Topyarn + _YARN
                    End If

                    'worksheet.Cells(X1, 2) = M02.Tables(0).Rows(X)("M03Material")
                    'worksheet.Cells(X1, 3) = M02.Tables(0).Rows(X)("M03Description")
                    'worksheet.Cells(X1, 4) = _QTY
                    'worksheet.Cells(X1, 5) = _SCRAP
                    'worksheet.Cells(X1, 6) = _YARN

                    nvcFieldList = "Insert Into R06Report(R06No,R06MC,R06Quality,R06Discription,R06TotalQty,R06Scrap,R06Qurantine,R06WS)" & _
                                                                 " values('1', '" & M03.Tables(0).Rows(i)("M03MCNo") & "','" & M02.Tables(0).Rows(X)("M03Material") & "','" & M02.Tables(0).Rows(X)("M03Description") & "','" & VB6.Format(_QTY, "#.00") & "','" & VB6.Format(_SCRAP, "#.00") & "','" & VB6.Format(_YARN, "#.00") & "','" & netCard & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    X = X + 1
                    X1 = X1 + 1
                Next




                i = i + 1
                X1 = X1 + 1
            Next


            nvcFieldList = "Insert Into R06Report(R06No,R06Dis2,R06TotalQty,R06Scrap,R06Qurantine,R06WS)" & _
                                                               " values('2','Total','" & VB6.Format(_Totqty, "#.00") & "','" & VB6.Format(_TotScrap, "#.00") & "','" & VB6.Format(_Topyarn, "#.00") & "','" & netCard & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            'QURANTINE
            nvcFieldList = "Insert Into R06Report(R06No,R06Dis,R06WS)" & _
                                                               " values('3','Quarantine - Active','" & netCard & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)


            i = 0
            _Totqty = 0
            Sql = "select M03MCNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in('Q') AND M03MCNo BETWEEN '" & Trim(txtM1.Text) & "' AND '" & Trim(txtM2.Text) & "' group by M03MCNo"
            M03 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            For Each DTRow2 As DataRow In M03.Tables(0).Rows
                'worksheet.Cells(X1, 1) = M03.Tables(0).Rows(i)("M03MCNo")
                'range1 = worksheet.Cells(X1, 1)
                'range1.Interior.Color = RGB(255, 255, 128)

                Sql = "select T01Status,T01RollNo,M03MCNo,M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('Q') AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' group by M03Quality,M03Description,M03OrderNo,M03MCNo,T01RollNo,T01Status"
                M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                X = 0
                'USABLE QTY
                For Each DTRow3 As DataRow In M02.Tables(0).Rows
                    _QTY = 0
                    _QTY = M02.Tables(0).Rows(X)("T0Rollweight")
                    _Totqty = _Totqty + _QTY
                    'worksheet.Cells(X1, 2) = M02.Tables(0).Rows(X)("M03Material")
                    'worksheet.Cells(X1, 3) = M02.Tables(0).Rows(X)("M03Description")
                    'worksheet.Cells(X1, 4) = M02.Tables(0).Rows(X)("M03OrderNo")
                    'worksheet.Cells(X1, 5) = M02.Tables(0).Rows(X)("T01RollNo")
                    'worksheet.Cells(X1, 6) = _QTY
                    'worksheet.Cells(X1, 7) = M02.Tables(0).Rows(X)("T01Status")

                    nvcFieldList = "Insert Into R06Report(R06No,R06MC,R06Quality,R06Discription,R06TotalQty,R06WS)" & _
                                                                " values('3', '" & M03.Tables(0).Rows(i)("M03MCNo") & "','" & M02.Tables(0).Rows(X)("M03Material") & "','" & M02.Tables(0).Rows(X)("M03Description") & "','" & VB6.Format(_QTY, "#.00") & "','" & netCard & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)

                    X = X + 1
                    X1 = X1 + 1
                Next

                i = i + 1
                X1 = X1 + 1
            Next

            nvcFieldList = "Insert Into R06Report(R06No,R06Dis2,R06TotalQty,R06WS)" & _
                                                              " values('4','Total','" & VB6.Format(_Totqty, "#.00") & "','" & netCard & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            'QURANTINE pass
            nvcFieldList = "Insert Into R06Report(R06No,R06Dis,R06WS)" & _
                                                               " values('5','Quarantine - Pass Roll','" & netCard & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)


            i = 0
            _Totqty = 0
            Sql = "select M03MCNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in('QP') AND M03MCNo BETWEEN '" & Trim(txtM1.Text) & "' AND '" & Trim(txtM2.Text) & "' group by M03MCNo"
            M03 = DBEngin.ExecuteDataset(connection, transaction, Sql)
            For Each DTRow2 As DataRow In M03.Tables(0).Rows
                'worksheet.Cells(X1, 1) = M03.Tables(0).Rows(i)("M03MCNo")
                'range1 = worksheet.Cells(X1, 1)
                'range1.Interior.Color = RGB(255, 255, 128)

                Sql = "select T01Status,T01RollNo,M03MCNo,M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('QP') AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' group by M03Quality,M03Description,M03OrderNo,M03MCNo,T01RollNo,T01Status"
                M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
                X = 0
                'USABLE QTY
                For Each DTRow3 As DataRow In M02.Tables(0).Rows
                    _QTY = 0
                    _QTY = M02.Tables(0).Rows(X)("T0Rollweight")
                    _Totqty = _Totqty + _QTY
                    'worksheet.Cells(X1, 2) = M02.Tables(0).Rows(X)("M03Material")
                    'worksheet.Cells(X1, 3) = M02.Tables(0).Rows(X)("M03Description")
                    'worksheet.Cells(X1, 4) = M02.Tables(0).Rows(X)("M03OrderNo")
                    'worksheet.Cells(X1, 5) = M02.Tables(0).Rows(X)("T01RollNo")
                    'worksheet.Cells(X1, 6) = _QTY
                    'worksheet.Cells(X1, 7) = M02.Tables(0).Rows(X)("T01Status")

                    nvcFieldList = "Insert Into R06Report(R06No,R06MC,R06Quality,R06Discription,R06TotalQty,R06WS)" & _
                                                              " values('5', '" & M03.Tables(0).Rows(i)("M03MCNo") & "','" & M02.Tables(0).Rows(X)("M03Material") & "','" & M02.Tables(0).Rows(X)("M03Description") & "','" & VB6.Format(_QTY, "#.00") & "','" & netCard & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    X = X + 1
                    X1 = X1 + 1
                Next

                i = i + 1
                X1 = X1 + 1
            Next


            nvcFieldList = "Insert Into R06Report(R06No,R06Dis2,R06TotalQty,R06WS)" & _
                                                              " values('6','Total','" & VB6.Format(_Totqty, "#.00") & "','" & netCard & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList)

            MsgBox("Report Genarating Sucessfully", MsgBoxStyle.Information, "TJL.........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            A = ConfigurationManager.AppSettings("ReportPath") + "\KnittingPro.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            B.SetParameterValue("To", _ToTime)
            B.SetParameterValue("From", _FromTime)
            B.SetParameterValue("M/C", txtM1.Text & " - " & txtM2.Text)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{R06Report.R06ws}='" & netCard & "'"
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.CrystalReportViewer1.DisplayGroupTree = False
            frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Function
    Function Create_Report()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet


        Dim FileName As String
        exc.Visible = True
        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        Dim range1 As Range
        Dim _NETTOTAL As Integer
        Dim _FromTime As String
        Dim _ToTime As String
        Dim _Total3mtr As Double
        Dim X As Integer
        Dim M02 As DataSet
        Dim _Total As Double
        Dim _CutoffReason As String
        Dim M03 As DataSet
        Dim _Quality As String
        Dim X1 As Integer
        Dim _QTY As Double
        Dim _SCRAP As Double
        Dim _YARN As Double

        Dim _Totqty As Double
        Dim _TotScrap As Double
        Dim _Topyarn As Double

        Dim _Totqty1 As Double
        Dim _TotScrap1 As Double
        Dim _Topyarn1 As Double

        Try

            _FromTime = txtDate.Text & " " & txtTime1.Text
            _ToTime = txtTo.Text & " " & txtToTime.Text

            If Trim(txtM1.Text) <> "" And Trim(txtM2.Text) <> "" Then

            Else
                MsgBox("Please enter the machine no", MsgBoxStyle.Information, "Information ....")
                Exit Function
            End If


            worksheet.Name = "Knitting Production Report"
            worksheet.Cells(2, 3) = "Knitting Production Report"
            worksheet.Rows(2).Font.Bold = True
            worksheet.Rows(2).Font.size = 26

            worksheet.Range("A2:J2").MergeCells = True
            worksheet.Range("A2:J2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet.Cells(4, 1) = "Knitting Production Report on "
            range1 = worksheet.Cells(4, 1)
            range1.Interior.Color = RGB(192, 192, 255)
            worksheet.Cells(4, 2) = _FromTime & "  To " & _ToTime
            worksheet.Rows(4).Font.Bold = True
            worksheet.Rows(4).Font.size = 10

            '  worksheet.Rows(6).rowheight = 20.25

            worksheet.Rows(6).Font.Bold = True
            worksheet.Rows(6).Font.size = 10
            worksheet.Cells(6, 1) = "Machine No"
            range1 = worksheet.Cells(6, 1)
            range1.Interior.Color = RGB(255, 245, 55)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(6, 2) = "Quality"
            range1 = worksheet.Cells(6, 2)
            range1.Interior.Color = RGB(255, 245, 55)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(6, 3) = "Description"
            range1 = worksheet.Cells(6, 3)
            range1.Interior.Color = RGB(255, 245, 55)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(6, 4) = "Total Qty"
            range1 = worksheet.Cells(6, 4)
            range1.Interior.Color = RGB(255, 245, 55)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(6, 5) = "Scrap"
            range1 = worksheet.Cells(6, 5)
            range1.Interior.Color = RGB(255, 245, 55)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(6, 6) = "Quarantine"
            range1 = worksheet.Cells(6, 6)
            range1.Interior.Color = RGB(255, 245, 55)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Columns(1).columnwidth = 35
            worksheet.Columns(2).columnwidth = 20
            worksheet.Columns(3).columnwidth = 45
            worksheet.Columns(4).columnwidth = 10
            worksheet.Columns(5).columnwidth = 10
            worksheet.Columns(6).columnwidth = 10

            Sql = "select M03MCNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo BETWEEN '" & Trim(txtM1.Text) & "' AND '" & Trim(txtM2.Text) & "' group by M03MCNo"
            M03 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            i = 0
            X1 = 7
            _Topyarn = 0
            _Totqty = 0
            _TotScrap = 0

            For Each DTRow1 As DataRow In M03.Tables(0).Rows
                _Topyarn = 0
                _Totqty = 0
                _TotScrap = 0

                worksheet.Cells(X1, 1) = M03.Tables(0).Rows(i)("M03MCNo")
                range1 = worksheet.Cells(X1, 1)
                range1.Interior.Color = RGB(255, 192, 255)
                'USABLE QTY
                Sql = "select M03MCNo,M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' group by M03Quality,M03Description,M03OrderNo,M03MCNo"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                X = 0
                'USABLE QTY
                For Each DTRow2 As DataRow In M02.Tables(0).Rows

                    _QTY = 0
                    _SCRAP = 0
                    _YARN = 0
                    Sql = "select M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' AND M03OrderNo='" & M02.Tables(0).Rows(X)("M03OrderNo") & "' group by M03Quality,M03Description,M03OrderNo"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M01) Then
                        _QTY = Val(M01.Tables(0).Rows(0)("T0Rollweight"))
                        _Totqty = _Totqty + _QTY
                    End If

                    'SCRAP
                    Sql = "select M03OrderNo,M03Description,sum(T05Weight) as T05Weight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo INNER JOIN T05Scrab ON T05RefNo=T01RefNo where  T01Status <>'I' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' AND M03OrderNo='" & M02.Tables(0).Rows(X)("M03OrderNo") & "' group by M03Quality,M03Description,M03OrderNo"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M01) Then
                        _SCRAP = Val(M01.Tables(0).Rows(0)("T05Weight"))
                        _TotScrap = _TotScrap + _SCRAP
                    End If

                    'YARN SCRAP
                    'Sql = "select M03OrderNo,M03Description,sum(T05Weight) as T05Weight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo INNER JOIN T05Scrab ON T05RefNo=T01RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status <>'I' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' AND M03OrderNo='" & M02.Tables(0).Rows(X)("M03OrderNo") & "' AND T05Department='Knitting' group by M03Quality,M03Description,M03OrderNo"
                    'M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    'If isValidDataset(M01) Then
                    '    _YARN = Val(M01.Tables(0).Rows(0)("T05Weight"))
                    '    _Topyarn = _Topyarn + _YARN
                    'End If

                    'QUARANTINE

                    Sql = "select M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where  T01Status ='Q' AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' AND M03OrderNo='" & M02.Tables(0).Rows(X)("M03OrderNo") & "' group by M03Quality,M03Description,M03OrderNo"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M01) Then
                        _YARN = Val(M01.Tables(0).Rows(0)("T0Rollweight"))
                        _Topyarn = _Topyarn + _YARN
                    End If

                    worksheet.Cells(X1, 2) = M02.Tables(0).Rows(X)("M03Material")
                    worksheet.Cells(X1, 3) = M02.Tables(0).Rows(X)("M03Description")
                    worksheet.Cells(X1, 4) = _QTY
                    worksheet.Cells(X1, 5) = _SCRAP
                    worksheet.Cells(X1, 6) = _YARN
                    X = X + 1
                    '  X1 = X1 + 1
                Next
                'worksheet.Cells(X1, 4) = _Totqty
                'worksheet.Cells(X1, 5) = _TotScrap
                'worksheet.Cells(X1, 6) = _Topyarn
                'worksheet.Rows(X1).Font.Bold = True
                'worksheet.Rows(X1).Font.size = 10
                'range1 = worksheet.Cells(X1, 4)
                'range1.Interior.Color = RGB(255, 192, 128)
                'range1 = worksheet.Cells(X1, 5)
                'range1.Interior.Color = RGB(255, 192, 128)
                'range1 = worksheet.Cells(X1, 6)
                'range1.Interior.Color = RGB(255, 192, 128)
                ' range1.Borders.LineStyle = XlLineStyle.xlContinuous
                _Totqty1 = _Totqty1 + _Totqty
                _TotScrap1 = +_TotScrap1 + _TotScrap
                _Topyarn1 = _Topyarn1 + _Topyarn

                i = i + 1
                X1 = X1 + 1
            Next

            worksheet.Cells(X1, 3) = "Total"
            worksheet.Cells(X1, 4) = _Totqty1
            worksheet.Cells(X1, 5) = _TotScrap1
            worksheet.Cells(X1, 6) = _Topyarn1
            worksheet.Rows(X1).Font.Bold = True
            worksheet.Rows(X1).Font.size = 10

            range1 = worksheet.Cells(X1, 3)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 4)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 5)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 6)
            range1.Interior.Color = RGB(255, 192, 128)

            X1 = X1 + 3

            worksheet.Cells(X1, 1) = "Quarantine - Active"
            range1 = worksheet.Cells(X1, 1)
            range1.Interior.Color = RGB(0, 192, 0)
            ' worksheet.Cells(4, 2) = _FromTime & "  To " & _ToTime
            worksheet.Rows(X1).Font.Bold = True
            worksheet.Rows(X1).Font.size = 10

            X1 = X1 + 2
            worksheet.Rows(X1).Font.Bold = True
            worksheet.Rows(X1).Font.size = 10
            worksheet.Cells(X1, 1) = "Machine No"
            range1 = worksheet.Cells(X1, 1)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 2) = "Quality"
            range1 = worksheet.Cells(X1, 2)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 3) = "Description"
            range1 = worksheet.Cells(X1, 3)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 4) = "Order No"
            range1 = worksheet.Cells(X1, 4)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 5) = "Roll No"
            range1 = worksheet.Cells(X1, 5)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 6) = "Qurantine"
            range1 = worksheet.Cells(X1, 6)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 7) = "Status"
            range1 = worksheet.Cells(X1, 7)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Columns(1).columnwidth = 35
            worksheet.Columns(2).columnwidth = 20
            worksheet.Columns(3).columnwidth = 45
            worksheet.Columns(4).columnwidth = 10
            worksheet.Columns(5).columnwidth = 10
            worksheet.Columns(6).columnwidth = 10

            X1 = X1 + 1
            i = 0
            _Totqty = 0
            Sql = "select M03MCNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in('Q') AND M03MCNo BETWEEN '" & Trim(txtM1.Text) & "' AND '" & Trim(txtM2.Text) & "' group by M03MCNo"
            M03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow2 As DataRow In M03.Tables(0).Rows
                worksheet.Cells(X1, 1) = M03.Tables(0).Rows(i)("M03MCNo")
                range1 = worksheet.Cells(X1, 1)
                range1.Interior.Color = RGB(255, 255, 128)

                Sql = "select T01Status,T01RollNo,M03MCNo,M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('Q') AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' group by M03Quality,M03Description,M03OrderNo,M03MCNo,T01RollNo,T01Status"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                X = 0
                'USABLE QTY
                For Each DTRow3 As DataRow In M02.Tables(0).Rows
                    _QTY = 0
                    _QTY = M02.Tables(0).Rows(X)("T0Rollweight")
                    _Totqty = _Totqty + _QTY
                    worksheet.Cells(X1, 2) = M02.Tables(0).Rows(X)("M03Material")
                    worksheet.Cells(X1, 3) = M02.Tables(0).Rows(X)("M03Description")
                    worksheet.Cells(X1, 4) = M02.Tables(0).Rows(X)("M03OrderNo")
                    worksheet.Cells(X1, 5) = M02.Tables(0).Rows(X)("T01RollNo")
                    worksheet.Cells(X1, 6) = _QTY
                    worksheet.Cells(X1, 7) = M02.Tables(0).Rows(X)("T01Status")
                    X = X + 1
                    X1 = X1 + 1
                Next

                i = i + 1
                X1 = X1 + 1
            Next
            worksheet.Cells(X1, 3) = "Total"
            ' worksheet.Cells(X1, 4) = _Totqty1
            'worksheet.Cells(X1, 5) = _TotScrap1
            worksheet.Cells(X1, 6) = _Totqty
            worksheet.Rows(X1).Font.Bold = True
            worksheet.Rows(X1).Font.size = 10

            range1 = worksheet.Cells(X1, 3)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 4)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 5)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 6)
            range1.Interior.Color = RGB(255, 192, 128)

            X1 = X1 + 3

            'QURANTINE PASS
            '
            worksheet.Cells(X1, 1) = "Quarantine - Pass Roll"
            range1 = worksheet.Cells(X1, 1)
            range1.Interior.Color = RGB(0, 192, 0)
            ' worksheet.Cells(4, 2) = _FromTime & "  To " & _ToTime
            worksheet.Rows(X1).Font.Bold = True
            worksheet.Rows(X1).Font.size = 10

            X1 = X1 + 2
            worksheet.Rows(X1).Font.Bold = True
            worksheet.Rows(X1).Font.size = 10
            worksheet.Cells(X1, 1) = "Machine No"
            range1 = worksheet.Cells(X1, 1)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 2) = "Quality"
            range1 = worksheet.Cells(X1, 2)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 3) = "Description"
            range1 = worksheet.Cells(X1, 3)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 4) = "Order No"
            range1 = worksheet.Cells(X1, 4)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 5) = "Roll No"
            range1 = worksheet.Cells(X1, 5)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 6) = "Qurantine"
            range1 = worksheet.Cells(X1, 6)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(X1, 7) = "Status"
            range1 = worksheet.Cells(X1, 7)
            range1.Interior.Color = RGB(192, 192, 255)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Columns(1).columnwidth = 35
            worksheet.Columns(2).columnwidth = 20
            worksheet.Columns(3).columnwidth = 45
            worksheet.Columns(4).columnwidth = 10
            worksheet.Columns(5).columnwidth = 10
            worksheet.Columns(6).columnwidth = 10

            X1 = X1 + 1
            i = 0
            _Totqty = 0
            Sql = "select M03MCNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in('QP') AND M03MCNo BETWEEN '" & Trim(txtM1.Text) & "' AND '" & Trim(txtM2.Text) & "' group by M03MCNo"
            M03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow2 As DataRow In M03.Tables(0).Rows
                worksheet.Cells(X1, 1) = M03.Tables(0).Rows(i)("M03MCNo")
                range1 = worksheet.Cells(X1, 1)
                range1.Interior.Color = RGB(255, 255, 128)

                Sql = "select T01Status,T01RollNo,M03MCNo,M03OrderNo,M03Quality as [M03Material],M03Description,sum(T0Rollweight) as T0Rollweight,M03OrderNo from T01Transaction_Header inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('QP') AND M03MCNo='" & M03.Tables(0).Rows(i)("M03MCNo") & "' group by M03Quality,M03Description,M03OrderNo,M03MCNo,T01RollNo,T01Status"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                X = 0
                'USABLE QTY
                For Each DTRow3 As DataRow In M02.Tables(0).Rows
                    _QTY = 0
                    _QTY = M02.Tables(0).Rows(X)("T0Rollweight")
                    _Totqty = _Totqty + _QTY
                    worksheet.Cells(X1, 2) = M02.Tables(0).Rows(X)("M03Material")
                    worksheet.Cells(X1, 3) = M02.Tables(0).Rows(X)("M03Description")
                    worksheet.Cells(X1, 4) = M02.Tables(0).Rows(X)("M03OrderNo")
                    worksheet.Cells(X1, 5) = M02.Tables(0).Rows(X)("T01RollNo")
                    worksheet.Cells(X1, 6) = _QTY
                    worksheet.Cells(X1, 7) = M02.Tables(0).Rows(X)("T01Status")
                    X = X + 1
                    X1 = X1 + 1
                Next

                i = i + 1
                X1 = X1 + 1
            Next
            worksheet.Cells(X1, 3) = "Total"
            ' worksheet.Cells(X1, 4) = _Totqty1
            'worksheet.Cells(X1, 5) = _TotScrap1
            worksheet.Cells(X1, 6) = _Totqty
            worksheet.Rows(X1).Font.Bold = True
            worksheet.Rows(X1).Font.size = 10

            range1 = worksheet.Cells(X1, 3)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 4)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 5)
            range1.Interior.Color = RGB(255, 192, 128)
            range1 = worksheet.Cells(X1, 6)
            range1.Interior.Color = RGB(255, 192, 128)


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub frmKnittingProduction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDate.Text = Today
        txtTo.Text = Today
    End Sub
End Class