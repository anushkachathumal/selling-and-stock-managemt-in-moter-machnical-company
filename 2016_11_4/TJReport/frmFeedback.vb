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
Public Class frmFeedback
    Dim Clicked As String
    Dim exc As New Application

    Dim workbooks As Workbooks = exc.Workbooks
    Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    Dim sheets As Sheets = Workbook.Worksheets
    Dim worksheet As _Worksheet = CType(Sheets.Item(1), _Worksheet)
    ' Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

    Function Create_ExelSheet()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet

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

        Try
            _FromTime = txtDate.Text & " " & txtTime1.Text
            _ToTime = txtTo.Text & " " & txtToTime.Text

            worksheet.Name = "Feed Back Report"
            worksheet.Cells(2, 3) = "Feed Back Report"
            worksheet.Rows(2).Font.Bold = True
            worksheet.Rows(2).Font.size = 26

            worksheet.Range("A2:J2").MergeCells = True
            worksheet.Range("A2:J2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet.Cells(4, 1) = "Daily Down Time Report on "
            range1 = worksheet.Cells(4, 1)
            range1.Interior.Color = RGB(192, 192, 255)
            worksheet.Cells(4, 2) = _FromTime & "  To " & _ToTime
            worksheet.Rows(4).Font.Bold = True
            worksheet.Rows(4).Font.size = 10

            '  worksheet.Rows(6).rowheight = 20.25

            worksheet.Rows(6).Font.Bold = True
            worksheet.Rows(6).Font.size = 10



            worksheet.Cells(6, 1) = "Total Cutoff Weight"
            range1 = worksheet.Cells(6, 1)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 2)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 3)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 4)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 5)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 6)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 7)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 8)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 9)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 10)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 11)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 12)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 13)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 14)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 15)
            range1.Interior.Color = RGB(255, 245, 55)
            range1 = worksheet.Cells(6, 16)
            range1.Interior.Color = RGB(255, 245, 55)

            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Columns(1).columnwidth = 35

            'LESS THAN 3MTR CUTOFF WEIGHT
            _FromTime = txtDate.Text & " " & txtTime1.Text
            _ToTime = txtDate.Text & " " & txtToTime.Text
            If chk1.Checked = True Then
                SQL = "select sum(T04Weight) as T04Weight from T04Cutoff inner join T01Transaction_Header on T04RefNo=T01RefNo where T01time between '" & _FromTime & "' and '" & _ToTime & "' group by T01RefNo"
            Else
                SQL = "select sum(T04Weight) as T04Weight from T04Cutoff inner join T01Transaction_Header on T04RefNo=T01RefNo where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T01RefNo"
            End If
            M02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            X = 0
            _Total3mtr = 0
            For Each DTRow2 As DataRow In M02.Tables(0).Rows
                _Total3mtr = _Total3mtr + Val(M02.Tables(0).Rows(X)("T04Weight"))
                X = X + 1

            Next
            _Total = _Total + _Total3mtr



            'CUTOFF WEIGHT
            If chk1.Checked = True Then
                SQL = "select sum(T05Weight) as T05Weight from T05Scrab inner join T01Transaction_Header on T05RefNo=T01RefNo where T01time between '" & _FromTime & "' and '" & _ToTime & "' group by T05RefNo"
            Else
                SQL = "select sum(T05Weight) as T05Weight from T05Scrab inner join T01Transaction_Header on T05RefNo=T01RefNo where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T05RefNo"
            End If

            M02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            X = 0
            _Total3mtr = 0
            For Each DTRow2 As DataRow In M02.Tables(0).Rows
                _Total3mtr = _Total3mtr + Val(M02.Tables(0).Rows(X)("T05Weight"))
                X = X + 1

            Next
            _Total = _Total + _Total3mtr

            worksheet.Cells(6, 2) = _Total


            worksheet.Cells(8, 4) = "Less than 3m Cutoff Weight"
            worksheet.Columns(4).columnwidth = 22
            worksheet.Cells(7, 5) = _Total3mtr
            worksheet.Cells(7, 4) = "Cutoff Weight"
            worksheet.Columns(4).columnwidth = 22
            worksheet.Cells(8, 5) = _Total - _Total3mtr

            worksheet.Cells(9, 1) = "Top 3 Cutoff Weight"
            range1 = worksheet.Cells(9, 1)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 2)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 3)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 4)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 5)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 6)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 7)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 8)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 9)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 10)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 11)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 12)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 13)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 14)
            range1.Interior.Color = RGB(255, 128, 192)
            range1 = worksheet.Cells(9, 15)
            range1.Interior.Color = RGB(255, 128, 192)

            worksheet.Rows(9).Font.Bold = True

            Dim Z As Integer
            Z = 10
            X = 0
            _Total3mtr = 0
            If chk1.Checked = True Then
                SQL = "select count(T05Fualcode),T05Fualcode,max(M02Name) as M02Name,sum(T05Weight) as T05Weight from T05Scrab inner join T01Transaction_Header on T05RefNo=T01RefNo inner join M02Fault on T05Fualcode=M02Code where T01Time between '" & _FromTime & "' and '" & _ToTime & "' group by T05Fualcode order by count(T05Fualcode) desc"
            Else
                SQL = "select count(T05Fualcode),T05Fualcode,max(M02Name) as M02Name,sum(T05Weight) as T05Weight from T05Scrab inner join T01Transaction_Header on T05RefNo=T01RefNo inner join M02Fault on T05Fualcode=M02Code where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T05Fualcode order by count(T05Fualcode) desc"

            End If
            M02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow2 As DataRow In M02.Tables(0).Rows
                'nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
                '                                         " values('3','" & M02.Tables(0).Rows(X)("M02Name") & "','" & Microsoft.VisualBasic.Format(M02.Tables(0).Rows(X)("T05Weight"), "#.00") & "','" & netCard & "')"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList)

                worksheet.Cells(Z, 4) = M02.Tables(0).Rows(X)("M02Name")
                worksheet.Cells(Z, 5) = Microsoft.VisualBasic.Format(M02.Tables(0).Rows(X)("T05Weight"), "#.00")

                Z = Z + 1
                X = X + 1
                If X = 3 Then
                    Exit For
                End If
            Next

            worksheet.Cells(Z, 1) = "Total No of Quarantine roll"
            worksheet.Rows(9).Font.Bold = True

            If chk1.Checked = True Then
                SQL = " select * from T01Transaction_Header where T01time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('Q')"
            Else
                SQL = " select * from T01Transaction_Header where T01date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status in ('Q')"

            End If
            M02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Total = 0
            _Total = M02.Tables(0).Rows.Count
            worksheet.Cells(Z, 2) = _Total
            worksheet.Rows(Z).Font.Bold = True

            For i = 1 To 15
                range1 = worksheet.Cells(Z, i)
                range1.Interior.Color = RGB(106, 106, 255)
            Next

            Z = Z + 1

            If chk1.Checked = True Then
                SQL = "select count(T07F_Code),T07F_Code,max(M02Name) as M02Name from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code where T01time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status='Q' group by T07F_Code order by count(T07F_Code) desc"
            Else
                SQL = "select count(T07F_Code),T07F_Code,max(M02Name) as M02Name from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status='Q' group by T07F_Code order by count(T07F_Code) desc"

            End If
            M02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            X = 0
            _CutoffReason = ""
            For Each DTRow2 As DataRow In M02.Tables(0).Rows

                If X = 0 Then
                    _CutoffReason = M02.Tables(0).Rows(X)("M02Name")
                Else
                    _CutoffReason = _CutoffReason & "," & M02.Tables(0).Rows(X)("M02Name")
                End If

                If chk1.Checked = True Then
                    SQL = "select count(M03MCNo) as MC,M03MCNo from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status='Q' and T07F_Code='" & M02.Tables(0).Rows(X)("T07F_Code") & "' group by M03MCNo"
                Else
                    SQL = "select count(M03MCNo) as MC,M03MCNo from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status='Q' and T07F_Code='" & M02.Tables(0).Rows(X)("T07F_Code") & "' group by M03MCNo"
                End If

                M03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                i = 0
                _Quality = ""
                For Each DTRow3 As DataRow In M03.Tables(0).Rows
                    If i = 0 Then
                        _Quality = Trim(M03.Tables(0).Rows(i)("M03MCNo")) & "|" & Trim(M03.Tables(0).Rows(i)("MC"))
                    Else
                        _Quality = _Quality & "," & Trim(M03.Tables(0).Rows(i)("M03MCNo")) & "|" & Trim(M03.Tables(0).Rows(i)("MC"))
                    End If
                    i = i + 1
                Next
                X = X + 1
                If X = 3 Then
                    Exit For
                End If
            Next

            worksheet.Cells(Z, 4) = _CutoffReason
            worksheet.Cells(Z, 5) = _Quality

            Z = Z + 1

            If chk1.Checked = True Then
                SQL = " select * from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('RP')"
            Else
                SQL = " select * from T01Transaction_Header where T01date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status in ('RP')"

            End If
            M02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            _Total = 0
            _Total = M02.Tables(0).Rows.Count

            worksheet.Cells(Z, 1) = "Total No of Reprocess roll"
            worksheet.Cells(Z, 2) = _Total
            worksheet.Rows(Z).Font.Bold = True
            i = 1
            For i = i To 15
                range1 = worksheet.Cells(Z, i)
                range1.Interior.Color = RGB(132, 255, 193)
            Next

            '---------------------------------------------------
            Z = Z + 2

            worksheet.Cells(Z, 1) = "Knitting M/C Stop"
            worksheet.Rows(Z).Font.Bold = True

            i = 1
            For i = 1 To 15
                range1 = worksheet.Cells(Z, 1)
                range1.Interior.Color = RGB(255, 172, 102)
            Next

            Z = Z + 1
            worksheet.Cells(Z, 2) = "M/C"
            range1 = worksheet.Cells(Z, 2)
            range1.Interior.Color = RGB(108, 217, 0)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(Z, 3) = "Quality No"
            range1 = worksheet.Cells(Z, 3)
            range1.Interior.Color = RGB(108, 217, 0)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(Z, 4) = "Stock Code"
            range1 = worksheet.Cells(Z, 4)
            range1.Interior.Color = RGB(108, 217, 0)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(Z, 5) = "M/C Stop Time"
            range1 = worksheet.Cells(Z, 5)
            range1.Interior.Color = RGB(108, 217, 0)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Cells(Z, 6) = "Reason"
            range1 = worksheet.Cells(Z, 6)
            range1.Interior.Color = RGB(108, 217, 0)
            range1.Borders.LineStyle = XlLineStyle.xlContinuous
            worksheet.Rows(Z).Font.Bold = True

            'SQL = "select M02Name,M03Quality,M03MCNo,M03Yarnstock,T01Time from T01Transaction_Header inner join T07Q_Reason on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code inner join M03Knittingorder on M03OrderNo=T01OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' group by M03Quality,M03MCNo"
            SQL = "select M03Quality,M03MCNo from T01Transaction_Header inner join T07Q_Reason on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code inner join M03Knittingorder on M03OrderNo=T01OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' group by M03Quality,M03MCNo"
            M02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            Z = Z + 1
            X = 0
            For Each DTRow2 As DataRow In M02.Tables(0).Rows
                SQL = "select M02Name,M03Quality,M03MCNo,M03Yarnstock,T01Time from T01Transaction_Header inner join T07Q_Reason on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code inner join M03Knittingorder on M03OrderNo=T01OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and M03Quality='" & M02.Tables(0).Rows(i)("M03Quality") & "' and M03MCNo='" & M02.Tables(0).Rows(i)("M03MCNo") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                X = 0
                For Each DTRow3 As DataRow In M03.Tables(0).Rows
                    If X = 0 Then
                        worksheet.Cells(Z, 2) = M02.Tables(0).Rows(i)("M03MCNo")
                        worksheet.Cells(Z, 3) = M02.Tables(0).Rows(i)("M03Quality")
                        worksheet.Cells(Z, 4) = M03.Tables(0).Rows(X)("M03Yarnstock")
                        worksheet.Cells(Z, 5) = Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(M03.Tables(0).Rows(X)("T01Time"), 11), 9)
                        worksheet.Cells(Z, 6) = M03.Tables(0).Rows(X)("M02Name")
                    Else
                        worksheet.Cells(Z, 5) = Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(M03.Tables(0).Rows(X)("T01Time"), 11), 9)
                        worksheet.Cells(Z, 6) = M03.Tables(0).Rows(X)("M02Name")
                    End If
                    Z = Z + 1
                    X = X + 1
                Next
                i = i + 1
            Next
            ' worksheet.Rows(Z).Font.Bold = True
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
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
        'Dim _Cutofwt As Double
        'Dim _Reasone As String
        'Dim Sql As String
        'Dim B As New ReportDocument
        'Dim A As String
        'Dim StrFromDate As String
        'Dim StrToDate As String

        'Dim ncQryType As String
        'Dim nvcFieldList As String
        'Dim nvcWhereClause As String
        'Dim nvcVccode As String
        'Dim i As Integer

        'Dim con = New SqlConnection()
        'con = DBEngin.GetConnection()
        'Dim recGRNheader As DataSet
        'Dim recStockBalance As DataSet
        '' Dim A As String
        ''  Dim B As New ReportDocument

        'Dim connection As SqlClient.SqlConnection
        'Dim transaction As SqlClient.SqlTransaction
        'Dim transactionCreated As Boolean
        'Dim connectionCreated As Boolean

        'connection = DBEngin.GetConnection(True)
        'connectionCreated = True
        'transaction = connection.BeginTransaction()
        'transactionCreated = True

        'Dim _Total3mtr As Double
        'Dim X As Integer
        'Dim M02 As DataSet
        'Dim _Total As Double
        'Dim _CutoffReason As String
        'Dim M03 As DataSet
        'Dim _Quality As String
        'Dim _FromTime As String
        'Dim _ToTime As String


        '_FromTime = txtDate.Text & " " & txtTime1.Text
        '_ToTime = txtTo.Text & " " & txtToTime.Text

        'Try

        Call Create_ExelSheet()

        '    nvcFieldList = "delete from R01Report where R01WorkStation='" & netCard & "'"
        '    ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '    If chk1.Checked = False Then

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation)" & _
        '                                                            " values('1', 'Total Cutoff weight','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        '------------------------------------------------------------------------------------
        '        'LESS THAN 3MTR CUTOFF WEIGHT
        '        Sql = "select sum(T04Weight) as T04Weight from T04Cutoff inner join T01Transaction_Header on T04RefNo=T01RefNo where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T01RefNo"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        X = 0
        '        _Total3mtr = 0
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows
        '            _Total3mtr = X + Val(M02.Tables(0).Rows(X)("T04Weight"))
        '            X = X + 1

        '        Next
        '        _Total = _Total + _Total3mtr

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                            " values('2', 'Less than 3m cutoff weight','" & Microsoft.VisualBasic.Format(_Total3mtr, "#.00") & "','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '        '------------------------------------------------------------------------------------
        '        'CUTOFF WEIGHT
        '        Sql = "select sum(T05Weight) as T05Weight from T05Scrab inner join T01Transaction_Header on T05RefNo=T01RefNo where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T05RefNo"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        X = 0
        '        _Total3mtr = 0
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows
        '            _Total3mtr = X + Val(M02.Tables(0).Rows(X)("T05Weight"))
        '            X = X + 1

        '        Next
        '        _Total = _Total + _Total3mtr


        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                            " values('2', 'Cutoff weight','" & Microsoft.VisualBasic.Format(_Total3mtr, "#.00") & "','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        'nvcFieldList = "Insert Into R01Report(R01Ref,R01RollNo,R01WorkStation)" & _
        '        '                                                    " values('2','" & Microsoft.VisualBasic.Format(_Total, "#.00") & "','" & netCard & "')"
        '        'ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        nvcFieldList = "Update R01Report set R01RollNo='" & Microsoft.VisualBasic.Format(_Total, "#.00") & "' where R01Ref='1' and R01WorkStation='" & netCard & "' "
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        '--------------------------------------------------------------
        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation)" & _
        '                                                           " values('3', 'Top 3 Cutoff weight','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        X = 0
        '        _Total3mtr = 0
        '        Sql = "select count(T05Fualcode),T05Fualcode,max(M02Name) as M02Name,sum(T05Weight) as T05Weight from T05Scrab inner join T01Transaction_Header on T05RefNo=T01RefNo inner join M02Fault on T05Fualcode=M02Code where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' group by T05Fualcode order by count(T05Fualcode) desc"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows
        '            nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                     " values('3','" & M02.Tables(0).Rows(X)("M02Name") & "','" & Microsoft.VisualBasic.Format(M02.Tables(0).Rows(X)("T05Weight"), "#.00") & "','" & netCard & "')"
        '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '            X = X + 1
        '            If X = 3 Then
        '                Exit For
        '            End If
        '        Next

        '        '--------------------------------------------
        '        'Total Quarantine Roll
        '        Sql = " select * from T01Transaction_Header where T01date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status in ('Q','RP')"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        _Total = 0
        '        _Total = M02.Tables(0).Rows.Count

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation,R01RollNo)" & _
        '                                                          " values('4', 'Total No of Quarantine roll','" & netCard & "','" & _Total & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        'Total Reprocess Roll
        '        Sql = " select * from T01Transaction_Header where T01date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status in ('RP')"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        _Total = 0
        '        _Total = M02.Tables(0).Rows.Count

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation,R01RollNo)" & _
        '                                                          " values('4', 'No of Reprocess roll','" & netCard & "','" & _Total & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        Sql = "select count(T07F_Code),T07F_Code,max(M02Name) as M02Name from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status='RP' group by T07F_Code order by count(T07F_Code) desc"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        X = 0
        '        _CutoffReason = ""
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows

        '            If X = 0 Then
        '                _CutoffReason = M02.Tables(0).Rows(X)("M02Name")
        '            Else
        '                _CutoffReason = _CutoffReason & "," & M02.Tables(0).Rows(X)("M02Name")
        '            End If

        '            Sql = "select T01OrderNo,M03MCNo from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status='RP' and T07F_Code='" & M02.Tables(0).Rows(X)("T07F_Code") & "'"
        '            M03 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            i = 0
        '            _Quality = ""
        '            For Each DTRow3 As DataRow In M03.Tables(0).Rows
        '                If i = 0 Then
        '                    _Quality = M03.Tables(0).Rows(i)("T01OrderNo") & "|" & M03.Tables(0).Rows(i)("M03MCNo")
        '                Else
        '                    _Quality = _Quality & "," & M03.Tables(0).Rows(i)("T01OrderNo") & "|" & M03.Tables(0).Rows(i)("M03MCNo")
        '                End If
        '                i = i + 1
        '            Next
        '            X = X + 1
        '            If X = 3 Then
        '                Exit For
        '            End If
        '        Next

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                     " values('4','" & _CutoffReason & "','" & _Quality & "','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '        '-----------------------------------------------------------------------------------------
        '        'Total Quarantine Roll
        '        Sql = " select * from T01Transaction_Header where T01date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status in ('Q')"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        _Total = 0
        '        _Total = M02.Tables(0).Rows.Count

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation,R01RollNo)" & _
        '                                                          " values('5', 'No of Quarantine roll','" & netCard & "','" & _Total & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        Sql = "select count(T07F_Code),T07F_Code,max(M02Name) as M02Name from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status='Q' group by T07F_Code order by count(T07F_Code) desc"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        X = 0
        '        _CutoffReason = ""
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows

        '            If X = 0 Then
        '                _CutoffReason = M02.Tables(0).Rows(X)("M02Name")
        '            Else
        '                _CutoffReason = _CutoffReason & "," & M02.Tables(0).Rows(X)("M02Name")
        '            End If

        '            Sql = "select T01OrderNo,M03MCNo from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Date between '" & txtDate.Text & "' and '" & txtTo.Text & "' and T01Status='Q' and T07F_Code='" & M02.Tables(0).Rows(X)("T07F_Code") & "'"
        '            M03 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            i = 0
        '            _Quality = ""
        '            For Each DTRow3 As DataRow In M03.Tables(0).Rows
        '                If i = 0 Then
        '                    _Quality = M03.Tables(0).Rows(i)("T01OrderNo") & "|" & M03.Tables(0).Rows(i)("M03MCNo")
        '                Else
        '                    _Quality = _Quality & "," & M03.Tables(0).Rows(i)("T01OrderNo") & "|" & M03.Tables(0).Rows(i)("M03MCNo")
        '                End If
        '                i = i + 1
        '            Next
        '            X = X + 1
        '            If X = 3 Then
        '                Exit For
        '            End If
        '        Next

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                     " values('5','" & _CutoffReason & "','" & _Quality & "','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '    Else

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation)" & _
        '                                                            " values('1', 'Total Cutoff weight','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        '------------------------------------------------------------------------------------
        '        'LESS THAN 3MTR CUTOFF WEIGHT
        '        Sql = "select sum(T04Weight) as T04Weight from T04Cutoff inner join T01Transaction_Header on T04RefNo=T01RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' group by T01RefNo"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        X = 0
        '        _Total3mtr = 0
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows
        '            _Total3mtr = X + Val(M02.Tables(0).Rows(X)("T04Weight"))
        '            X = X + 1

        '        Next
        '        _Total = _Total + _Total3mtr

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                            " values('2', 'Less than 3m cutoff weight','" & Microsoft.VisualBasic.Format(_Total3mtr, "#.00") & "','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '        '------------------------------------------------------------------------------------
        '        'CUTOFF WEIGHT
        '        Sql = "select sum(T05Weight) as T05Weight from T05Scrab inner join T01Transaction_Header on T05RefNo=T01RefNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' group by T05RefNo"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        X = 0
        '        _Total3mtr = 0
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows
        '            _Total3mtr = X + Val(M02.Tables(0).Rows(X)("T05Weight"))
        '            X = X + 1

        '        Next
        '        _Total = _Total + _Total3mtr


        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                            " values('2', 'Cutoff weight','" & Microsoft.VisualBasic.Format(_Total3mtr, "#.00") & "','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        'nvcFieldList = "Insert Into R01Report(R01Ref,R01RollNo,R01WorkStation)" & _
        '        '                                                    " values('2','" & Microsoft.VisualBasic.Format(_Total, "#.00") & "','" & netCard & "')"
        '        'ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        nvcFieldList = "Update R01Report set R01RollNo='" & Microsoft.VisualBasic.Format(_Total, "#.00") & "' where R01Ref='1' and R01WorkStation='" & netCard & "' "
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        '--------------------------------------------------------------
        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation)" & _
        '                                                           " values('3', 'Top 3 Cutoff weight','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        X = 0
        '        _Total3mtr = 0
        '        Sql = "select count(T05Fualcode),T05Fualcode,max(M02Name) as M02Name,sum(T05Weight) as T05Weight from T05Scrab inner join T01Transaction_Header on T05RefNo=T01RefNo inner join M02Fault on T05Fualcode=M02Code where T01time between '" & _FromTime & "' and '" & _ToTime & "' group by T05Fualcode order by count(T05Fualcode) desc"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows
        '            nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                     " values('3','" & M02.Tables(0).Rows(X)("M02Name") & "','" & Microsoft.VisualBasic.Format(M02.Tables(0).Rows(X)("T05Weight"), "#.00") & "','" & netCard & "')"
        '            ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '            X = X + 1
        '            If X = 3 Then
        '                Exit For
        '            End If
        '        Next

        '        '--------------------------------------------
        '        'Total Quarantine Roll
        '        Sql = " select * from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('Q','RP')"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        _Total = 0
        '        _Total = M02.Tables(0).Rows.Count

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation,R01RollNo)" & _
        '                                                          " values('4', 'Total No of Quarantine roll','" & netCard & "','" & _Total & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        'Total Reprocess Roll
        '        Sql = " select * from T01Transaction_Header where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('RP')"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        _Total = 0
        '        _Total = M02.Tables(0).Rows.Count

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation,R01RollNo)" & _
        '                                                          " values('4', 'No of Reprocess roll','" & netCard & "','" & _Total & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        Sql = "select count(T07F_Code),T07F_Code,max(M02Name) as M02Name from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status='RP' group by T07F_Code order by count(T07F_Code) desc"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        X = 0
        '        _CutoffReason = ""
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows

        '            If X = 0 Then
        '                _CutoffReason = M02.Tables(0).Rows(X)("M02Name")
        '            Else
        '                _CutoffReason = _CutoffReason & "," & M02.Tables(0).Rows(X)("M02Name")
        '            End If

        '            Sql = "select T01OrderNo,M03MCNo from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01Time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status='RP' and T07F_Code='" & M02.Tables(0).Rows(X)("T07F_Code") & "'"
        '            M03 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            i = 0
        '            _Quality = ""
        '            For Each DTRow3 As DataRow In M03.Tables(0).Rows
        '                If i = 0 Then
        '                    _Quality = M03.Tables(0).Rows(i)("T01OrderNo") & "|" & M03.Tables(0).Rows(i)("M03MCNo")
        '                Else
        '                    _Quality = _Quality & "," & M03.Tables(0).Rows(i)("T01OrderNo") & "|" & M03.Tables(0).Rows(i)("M03MCNo")
        '                End If
        '                i = i + 1
        '            Next
        '            X = X + 1
        '            If X = 3 Then
        '                Exit For
        '            End If
        '        Next

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                     " values('4','" & _CutoffReason & "','" & _Quality & "','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '        '-----------------------------------------------------------------------------------------
        '        'Total Quarantine Roll
        '        Sql = " select * from T01Transaction_Header where T01time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status in ('Q')"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        _Total = 0
        '        _Total = M02.Tables(0).Rows.Count

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01OrderNo,R01WorkStation,R01RollNo)" & _
        '                                                          " values('5', 'No of Quarantine roll','" & netCard & "','" & _Total & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)

        '        Sql = "select count(T07F_Code),T07F_Code,max(M02Name) as M02Name from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M02Fault on T07F_Code=M02Code where T01time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status='Q' group by T07F_Code order by count(T07F_Code) desc"
        '        M02 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '        X = 0
        '        _CutoffReason = ""
        '        For Each DTRow2 As DataRow In M02.Tables(0).Rows

        '            If X = 0 Then
        '                _CutoffReason = M02.Tables(0).Rows(X)("M02Name")
        '            Else
        '                _CutoffReason = _CutoffReason & "," & M02.Tables(0).Rows(X)("M02Name")
        '            End If

        '            Sql = "select T01OrderNo,M03MCNo from T07Q_Reason inner join T01Transaction_Header on T07RefNo=T01RefNo inner join M03Knittingorder on T01OrderNo=M03OrderNo where T01time between '" & _FromTime & "' and '" & _ToTime & "' and T01Status='Q' and T07F_Code='" & M02.Tables(0).Rows(X)("T07F_Code") & "'"
        '            M03 = DBEngin.ExecuteDataset(connection, transaction, Sql)
        '            i = 0
        '            _Quality = ""
        '            For Each DTRow3 As DataRow In M03.Tables(0).Rows
        '                If i = 0 Then
        '                    _Quality = M03.Tables(0).Rows(i)("T01OrderNo") & "|" & M03.Tables(0).Rows(i)("M03MCNo")
        '                Else
        '                    _Quality = _Quality & "," & M03.Tables(0).Rows(i)("T01OrderNo") & "|" & M03.Tables(0).Rows(i)("M03MCNo")
        '                End If
        '                i = i + 1
        '            Next
        '            X = X + 1
        '            If X = 3 Then
        '                Exit For
        '            End If
        '        Next

        '        nvcFieldList = "Insert Into R01Report(R01Ref,R01Reason,T01R_Reason,R01WorkStation)" & _
        '                                                     " values('5','" & _CutoffReason & "','" & _Quality & "','" & netCard & "')"
        '        ExecuteNonQueryText(connection, transaction, nvcFieldList)
        '    End If


        '    MsgBox("Report  generated successfully", MsgBoxStyle.Information, "Textued Jersey ............")
        '    transaction.Commit()

        '    If chk1.Checked = True Then
        '        A = ConfigurationManager.AppSettings("ReportPath") + "\Feedback.rpt"
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

        '    Else
        '        A = ConfigurationManager.AppSettings("ReportPath") + "\Feedback.rpt"
        '        B.Load(A.ToString)
        '       B.SetDatabaseLogon("sa", "tommya")
        '        B.SetParameterValue("From", txtDate.Text)
        '        B.SetParameterValue("To", txtTo.Text)
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

        'Catch returnMessage As Exception
        '    If returnMessage.Message <> Nothing Then
        '        MessageBox.Show(returnMessage.Message)
        '    End If
        'End Try
    End Sub

    Private Sub frmFeedback_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class