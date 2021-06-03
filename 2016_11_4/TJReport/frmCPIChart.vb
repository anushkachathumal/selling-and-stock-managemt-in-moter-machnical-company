Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Drawing.Color
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Configuration
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Spire.XlS

Public Class frmCPIChart
    Dim Clicked As String
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter
    Dim exc As New Application

    Dim workbooks As Workbooks = exc.Workbooks
    Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    Dim sheets As Sheets = Workbook.Worksheets
    Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
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

    Private Sub frmCPIChart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
            txtDate.Text = Today
            txtTo.Text = Today

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet
        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        exc.Visible = True
        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        Dim range1 As Range
        Dim _NETTOTAL As Integer
        Dim T04 As DataSet
        Dim n_per As Double
        Dim Y As Integer
        Dim _cOUNT As Integer
        Dim _Row As Integer

        '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
        If cboQuality.Text <> "" Then
        Else
            MsgBox("Please enter the quality", MsgBoxStyle.Information, "Textued Jersey ........")
            Exit Sub
        End If
        Workbooks.Application.Sheets.Add()
        Dim sheets1 As Sheets = Workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        worksheet1.Name = "CPI Chart_" & Month(txtDate.Text) & "." & Microsoft.VisualBasic.Day(txtDate.Text) & "." & Year(txtDate.Text)

        worksheet1.Rows("2:1").rowheight = 30

        worksheet1.Rows(2).Font.size = 10
        worksheet1.Rows(2).Font.bold = True
        worksheet1.Cells(2, 1) = " Order No "
        worksheet1.Cells(2, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 1).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 1).Orientation = 90
        worksheet1.Columns("A").ColumnWidth = 15

        worksheet1.Cells(2, 2) = "  Roll No"
        worksheet1.Cells(2, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 2).VerticalAlignment = XlVAlign.xlVAlignCenter
        'worksheet1.Cells(2, 2).Orientation = 90

        worksheet1.Columns("B").ColumnWidth = 15

        worksheet1.Cells(2, 3) = "   Material"
        worksheet1.Cells(2, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 3).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 3).Orientation = 90
        worksheet1.Columns("C").ColumnWidth = 15

        worksheet1.Cells(2, 4) = "   Quality"
        worksheet1.Cells(2, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 4).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 4).Orientation = 90

        worksheet1.Columns("D").ColumnWidth = 15
        worksheet1.Cells(2, 5) = "   M/C No"
        worksheet1.Cells(2, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 5).VerticalAlignment = XlVAlign.xlVAlignCenter
        '  worksheet1.Cells(2, 5).Orientation = 90

        worksheet1.Columns("E").ColumnWidth = 15
        worksheet1.Cells(2, 6) = "CPI 1"
        worksheet1.Cells(2, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 6).VerticalAlignment = XlVAlign.xlVAlignCenter
        '  worksheet1.Cells(2, 6).Orientation = 90

        worksheet1.Columns("F").ColumnWidth = 10
        worksheet1.Cells(2, 7) = "CPI 2"
        worksheet1.Cells(2, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 7).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 7).Orientation = 90

        worksheet1.Columns("G").ColumnWidth = 10
        worksheet1.Cells(2, 8) = "CPI 3"
        worksheet1.Cells(2, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 8).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 8).Orientation = 90

        worksheet1.Columns("H").ColumnWidth = 10
        worksheet1.Cells(2, 9) = "CPI 4"
        worksheet1.Cells(2, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 9).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 9).Orientation = 90

        worksheet1.Columns("I").ColumnWidth = 10
        worksheet1.Cells(2, 10) = "CPI 5"
        worksheet1.Cells(2, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 10).VerticalAlignment = XlVAlign.xlVAlignCenter
        '  worksheet1.Cells(2, 10).Orientation = 90

        worksheet1.Columns("J").ColumnWidth = 10
        worksheet1.Cells(2, 11) = "CPI GAP"
        worksheet1.Cells(2, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 11).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 11).Orientation = 90

        worksheet1.Columns("K").ColumnWidth = 10
        worksheet1.Cells(2, 12) = "MAX"
        worksheet1.Cells(2, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 12).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 12).Orientation = 90

        worksheet1.Columns("L").ColumnWidth = 10
        worksheet1.Cells(2, 13) = "MIN"
        worksheet1.Cells(2, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 13).VerticalAlignment = XlVAlign.xlVAlignCenter
        '  worksheet1.Cells(2, 13).Orientation = 90

        worksheet1.Columns("M").ColumnWidth = 10
        worksheet1.Cells(2, 14) = "Max target"
        worksheet1.Cells(2, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 14).VerticalAlignment = XlVAlign.xlVAlignCenter
        '  worksheet1.Cells(2, 14).Orientation = 90

        worksheet1.Columns("N").ColumnWidth = 10
        worksheet1.Cells(2, 15) = "Min target"
        worksheet1.Cells(2, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 15).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 15).Orientation = 90

        worksheet1.Columns("O").ColumnWidth = 10

        worksheet1.Cells(2, 16) = "EPF No"
        worksheet1.Cells(2, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 16).VerticalAlignment = XlVAlign.xlVAlignCenter
        ' worksheet1.Cells(2, 15).Orientation = 90

        worksheet1.Columns("p").ColumnWidth = 10
        '--------------------------------------------------------------
        worksheet1.Range("A2:a2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("b2:b2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("c2:c2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("d2:d2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("e2:e2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("f2:f2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("g2:g2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("h2:h2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("i2:i2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("j2:j2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("k2:k2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("l2:l2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("m2:m2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("n2:n2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("o2:o2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("p2:p2").Interior.Color = RGB(141, 180, 227)


        worksheet1.Range("A2", "a2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b2", "b2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c2", "c2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A2", "a2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A2", "a2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b2", "b2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b2", "b2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c2", "c2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c2", "c2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d2", "d2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d2", "d2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d2", "d2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e2", "e2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e2", "e2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e2", "e2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f2", "f2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f2", "f2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f2", "f2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g2", "g2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g2", "g2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g2", "g2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h2", "h2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h2", "h2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h2", "h2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i2", "i2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i2", "i2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i2", "i2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j2", "j2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j2", "j2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j2", "j2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k2", "k2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k2", "k2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k2", "k2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l2", "l2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l2", "l2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l2", "l2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m2", "m2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m2", "m2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m2", "m2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n2", "n2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n2", "n2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n2", "n2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o2", "o2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o2", "o2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("o2", "o2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p2", "p2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p2", "p2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("p2", "p2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        '======================================================================================================
        Dim X As Integer
        Dim _From As Date
        Dim _To As Date
        Dim _Min As Integer
        Dim _Max As Integer

        _From = txtDate.Text & " " & "7:30 AM"
        _To = txtTo.Text & " " & "7:30 PM"
        SQL = "select M03OrderNo,M03Quality,M03Material,T01RollNo,M03MCNo,T01RefNo,T01InsEPF from M03Knittingorder inner join T01Transaction_Header on T01OrderNo=M03OrderNo where M03Quality='" & Trim(cboQuality.Text) & "' and T01time between '" & _From & "' and '" & _To & "'"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        i = 0
        _Row = 3
        For Each DTRow1 As DataRow In T01.Tables(0).Rows
            worksheet1.Rows(_Row).Font.size = 10
            worksheet1.Cells(_Row, 1) = T01.Tables(0).Rows(i)("M03OrderNo")
            worksheet1.Cells(_Row, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(_Row, 2) = T01.Tables(0).Rows(i)("T01RollNo")
            worksheet1.Cells(_Row, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(_Row, 3) = T01.Tables(0).Rows(i)("M03Material")
            worksheet1.Cells(_Row, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(_Row, 4) = T01.Tables(0).Rows(i)("M03Quality")
            worksheet1.Cells(_Row, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(_Row, 5) = T01.Tables(0).Rows(i)("M03MCNo")
            worksheet1.Cells(_Row, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter

            SQL = "select * from T03CPI_Reading where T03RefNo=" & T01.Tables(0).Rows(i)("T01RefNo") & ""
            T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            X = 0
            _Max = 0
            _Min = 0
            If isValidDataset(T03) Then
                If Not IsDBNull(Trim(T03.Tables(0).Rows(X)("T03CPIV2"))) Then
                    worksheet1.Cells(_Row, 6) = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV1")))
                    worksheet1.Cells(_Row, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If
                If Not IsDBNull(Trim(T03.Tables(0).Rows(X)("T03CPIV2"))) Then
                    worksheet1.Cells(_Row, 7) = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV2")))
                    worksheet1.Cells(_Row, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If Not IsDBNull(Trim(T03.Tables(0).Rows(X)("T03CPIV3"))) Then
                    worksheet1.Cells(_Row, 8) = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV3")))
                    worksheet1.Cells(_Row, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                If Not IsDBNull(Trim(T03.Tables(0).Rows(X)("T03CPIV4"))) Then
                    worksheet1.Cells(_Row, 9) = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV4")))
                    worksheet1.Cells(_Row, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                If Not IsDBNull(Trim(T03.Tables(0).Rows(X)("T03CPIV5"))) Then
                    worksheet1.Cells(_Row, 10) = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV5")))
                    worksheet1.Cells(_Row, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                If CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV2"))) >= Val(Trim(T03.Tables(0).Rows(X)("T03CPIV1"))) Then
                    _Max = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV2")))
                Else
                    _Max = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV1")))
                End If

                If _Max < CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV3"))) Then
                    _Max = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV3")))
                End If

                If _Max < CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV4"))) Then
                    _Max = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV4")))
                End If

                If _Max < CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV5"))) Then
                    _Max = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV5")))
                End If

                If Val(Trim(T03.Tables(0).Rows(X)("T03CPIV1"))) = 0 Then
                    _Min = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV2")))
                Else
                    If CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV2"))) >= Val(Trim(T03.Tables(0).Rows(X)("T03CPIV1"))) Then
                        _Min = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV1")))
                    Else
                        _Min = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV2")))
                    End If
                End If


                If _Min > CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV3"))) And CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV3"))) <> 0 Then
                    _Min = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV3")))
                End If

                If _Min > CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV4"))) And CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV4"))) <> 0 Then
                    _Min = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV4")))
                End If

                If _Min > CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV5"))) And CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV5"))) <> 0 Then
                    _Min = CInt(Trim(T03.Tables(0).Rows(X)("T03CPIV5")))
                End If
            End If

            worksheet1.Cells(_Row, 11) = _Max - _Min
            worksheet1.Cells(_Row, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(_Row, 12) = _Max
            worksheet1.Cells(_Row, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(_Row, 13) = _Min
            worksheet1.Cells(_Row, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

            SQL = "select * from M11Quality_CPI where M11Quality='" & Trim(cboQuality.Text) & "'"
            T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T04) Then
                SQL = "select * from M12Knitting_MC where M12MCNo='" & T01.Tables(0).Rows(i)("M03MCNo") & "'"
                T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T03) Then
                    If Trim(T03.Tables(0).Rows(0)("M12Type")) = "ORIZIO" Then
                        worksheet1.Cells(_Row, 14) = Val(T04.Tables(0).Rows(0)("M11CPIValue")) + 1
                        worksheet1.Cells(_Row, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(_Row, 15) = Val(T04.Tables(0).Rows(0)("M11CPIValue")) - 1
                        worksheet1.Cells(_Row, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(_Row, 14) = Val(T04.Tables(0).Rows(0)("M11SAN")) + 1
                        worksheet1.Cells(_Row, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet1.Cells(_Row, 15) = Val(T04.Tables(0).Rows(0)("M11SAN")) - 1
                        worksheet1.Cells(_Row, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If
                End If
            End If

            worksheet1.Cells(_Row, 16) = T01.Tables(0).Rows(i)("T01InsEPF")
            worksheet1.Cells(_Row, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Range("A" & _Row, "a" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("A" & _Row, "a" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & _Row, "b" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & _Row, "b" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & _Row, "c" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & _Row, "c" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & _Row, "d" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & _Row, "d" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & _Row, "e" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & _Row, "e" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & _Row, "f" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & _Row, "f" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & _Row, "g" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & _Row, "g" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & _Row, "h" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & _Row, "h" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & _Row, "i" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & _Row, "i" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & _Row, "j" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & _Row, "j" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & _Row, "k" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & _Row, "k" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & _Row, "l" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & _Row, "l" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & _Row, "m" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & _Row, "m" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & _Row, "n" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & _Row, "n" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & _Row, "o" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & _Row, "o" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & _Row, "p" & _Row).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & _Row, "p" & _Row).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

            _Row = _Row + 1
            i = i + 1
        Next


        Dim chartPage As Microsoft.Office.Interop.Excel.Chart
        Dim xlCharts As Microsoft.Office.Interop.Excel.ChartObjects
        Dim myChart As Microsoft.Office.Interop.Excel.ChartObject
        Dim chartRange As Microsoft.Office.Interop.Excel.Range
        Dim chartRange1 As Microsoft.Office.Interop.Excel.Range
        Dim chartRange2 As Microsoft.Office.Interop.Excel.Range


        Dim t_SerCol As Microsoft.Office.Interop.Excel.SeriesCollection
        Dim t_Series As Microsoft.Office.Interop.Excel.Series
        Dim z1 As Integer
        Dim sh As Worksheet
        xlCharts = worksheet1.ChartObjects

        z1 = _Row * 12.5
        z1 = z1 + 40
        myChart = xlCharts.Add(10, z1, 605, 200)
        chartPage = myChart.Chart
        chartRange = worksheet1.Range("E3", "E" & (_Row - 1))
        chartRange1 = worksheet1.Range("L3", "L" & (_Row - 1))
        'chartRange = worksheet1.Range("H8", "K" & (X - 1))
        'chartRange = worksheet1.Range("H8:K39", "A9:A39")
        ' chartPage.SetSourceData(Source:=chartRange)
        t_SerCol = chartPage.SeriesCollection
        t_Series = t_SerCol.NewSeries
        With t_Series
            .Name = "Max"
            t_Series.XValues = chartRange '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
            t_Series.Values = chartRange1 '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

        End With
       

        t_Series.Border.Color = RGB(6, 13, 150)
        chartPage.SeriesCollection(1).Interior.Color = RGB(6, 13, 150)
        chartRange = worksheet1.Range("M3", "M" & (_Row - 1))
        '   chartRange1 = worksheet16.Range("A7", "A" & (x - 1))
        'chartRange = worksheet1.Range("H8", "K" & (X - 1))
        'chartRange = worksheet1.Range("H8:K39", "A9:A39")
        ' chartPage.SetSourceData(Source:=chartRange)
        t_SerCol = chartPage.SeriesCollection
        t_Series = t_SerCol.NewSeries
        With t_Series
            .Name = "Min"
            ' t_Series.XValues = chartRange '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

        End With
        t_Series.Border.Color = RGB(255, 0, 0)
        chartPage.Refresh()
   
        ' chartPage.SeriesCollection(2).Interior.Color = RGB(255, 0, 0)

        't_Series.Border.Color = RGB(0, 176, 80)
        chartPage.SeriesCollection(2).Interior.Color = RGB(0, 176, 80)
        chartRange = worksheet1.Range("N3", "N" & (_Row - 1))
        '   chartRange1 = worksheet16.Range("A7", "A" & (x - 1))
        'chartRange = worksheet1.Range("H8", "K" & (X - 1))
        'chartRange = worksheet1.Range("H8:K39", "A9:A39")
        ' chartPage.SetSourceData(Source:=chartRange)
        t_SerCol = chartPage.SeriesCollection
        t_Series = t_SerCol.NewSeries
        With t_Series
            .Name = "Max Target"
            '  t_Series.XValues = chartRange '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

        End With


        t_Series.Border.Color = RGB(0, 176, 80)
        chartPage.Refresh()
        ' chartPage.SeriesCollection(2).Interior.Color = RGB(255, 0, 0)
        chartPage.SeriesCollection(3).Interior.Color = RGB(255, 255, 0)
        chartRange = worksheet1.Range("o3", "o" & (_Row - 1))
        '   chartRange1 = worksheet16.Range("A7", "A" & (x - 1))
        'chartRange = worksheet1.Range("H8", "K" & (X - 1))
        'chartRange = worksheet1.Range("H8:K39", "A9:A39")
        ' chartPage.SetSourceData(Source:=chartRange)
        t_SerCol = chartPage.SeriesCollection
        t_Series = t_SerCol.NewSeries
        With t_Series
            .Name = "Min Target"
            '  t_Series.XValues = chartRange '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

        End With


        t_Series.Border.Color = RGB(255, 255, 0)
        chartPage.Refresh()
        chartPage.Refresh()

        chartPage.Refresh()
        chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlLine
        chartPage.chartStyle = 42
        chartPage.HasTitle = True
        chartPage.ChartTitle.Text = ("CPI Chart -" & cboQuality.Text)

        DBEngin.CloseConnection(con)
        con.ConnectionString = ""
    End Sub
End Class