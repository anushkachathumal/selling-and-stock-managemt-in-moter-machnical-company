Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.Quarrys
Imports System.IO.File
Imports System.IO.StreamWriter
Imports System.Net.Mail
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Spire.XlS
Public Class frmOrderAnalysis
    Dim strLine As String
    Dim strLineflu As String
    Dim strDash As String
    Dim StrDisCode As String
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter
    Dim exc As New Application
    Dim Clicked As String
    Dim workbooks As Workbooks = exc.Workbooks
    Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    Dim sheets As Sheets = Workbook.Worksheets
    Dim worksheet1 As _Worksheet = CType(Sheets.Item(1), _Worksheet)

    Function Load_SAPNo()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select SAPCode as [SAP Code] from Alarm group by SAPCode"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            'With cboFrom
            '    .DataSource = T01
            '    .Rows.Band.Columns(0).Width = 125
            'End With

            'With cboTo
            '    .DataSource = T01
            '    .Rows.Band.Columns(0).Width = 125
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        OPR1.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False

        'cboFrom.ToggleDropdown()

        cmdSave.Enabled = True
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False

        cmdAdd.Focus()
    End Sub

    Private Sub frmOrderAnalysis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_SAPNo()
        txtFromDate.Text = Today
        txtTodate.Text = Today

    End Sub

    Function Order_Analysis()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
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
        Dim _1stWeek As Double
        Dim _2ndWeek As Double
        Dim _3rdweek As Double
        Dim _DD As Double
        Dim _CWeek As Double
        Dim _RL As Double
        Dim _RQ As Double

        '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
        ' workbooks.Application.Sheets.Add()
        Dim sheets1 As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        worksheet1.Name = "DCA Order Analysis Report"
        ' & Month(Today) & "." & Microsoft.VisualBasic.Day(Today) & "." & Year(Today)
        Dim _WeekNo As Integer

        worksheet1.Cells(1, 2) = "Textured Jersey Lanka Pvt Ltd"
        worksheet1.Cells(2, 2) = "DCA Order Analysis Report"
        worksheet1.Cells(3, 2) = "Report Date : " & Month(Today) & "." & Microsoft.VisualBasic.Day(Today) & "." & Year(Today)
        worksheet1.Cells(4, 2) = "Report Time : " & Hour(VserverTime) & ":" & Minute(VserverTime) & ":" & Second(VserverTime)

        worksheet1.Columns("A").ColumnWidth = 12
        worksheet1.Columns("B").ColumnWidth = 40
        worksheet1.Range("A1:B1").Interior.Color = RGB(191, 191, 191)
        worksheet1.Rows(1).Font.size = 10
        worksheet1.Rows(1).Font.bold = True

        worksheet1.Range("A2:B2").Interior.Color = RGB(191, 191, 191)
        worksheet1.Rows(2).Font.size = 10
        worksheet1.Rows(2).Font.bold = True

        worksheet1.Range("A3:B3").Interior.Color = RGB(191, 191, 191)
        worksheet1.Rows(3).Font.size = 10
        worksheet1.Rows(3).Font.bold = True

        worksheet1.Range("A4:B4").Interior.Color = RGB(191, 191, 191)
        worksheet1.Rows(4).Font.size = 10
        worksheet1.Rows(4).Font.bold = True


        worksheet1.Cells(1, 16) = "D"
        worksheet1.Cells(1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(1, 17) = "57"
        worksheet1.Cells(1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(1, 18) = "(70~80)"
        worksheet1.Cells(1, 18).HorizontalAlignment = XlHAlign.xlHAlignLeft

        worksheet1.Cells(1, 16) = "D"
        worksheet1.Cells(1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(2, 16) = "A"
        worksheet1.Cells(2, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 17) = "47"
        worksheet1.Cells(2, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(2, 18) = "(60~70)"
        worksheet1.Cells(2, 18).HorizontalAlignment = XlHAlign.xlHAlignLeft

        worksheet1.Cells(3, 16) = "C"
        worksheet1.Cells(3, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(3, 17) = "22"
        worksheet1.Cells(3, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter


        worksheet1.Cells(2, 24) = "Purchasing control"
        worksheet1.Cells(2, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(3, 24) = "Reference TOD"
        worksheet1.Cells(3, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("w2:y2").Interior.Color = RGB(255, 255, 255)
        worksheet1.Range("w3:y3").Interior.Color = RGB(255, 255, 255)

        worksheet1.Columns("A").ColumnWidth = 8
        worksheet1.Columns("b").ColumnWidth = 27
        worksheet1.Columns("c").ColumnWidth = 5
        worksheet1.Columns("d").ColumnWidth = 5
        worksheet1.Columns("e").ColumnWidth = 5

        worksheet1.Columns("f").ColumnWidth = 8
        worksheet1.Columns("g").ColumnWidth = 8
        worksheet1.Columns("h").ColumnWidth = 8
        worksheet1.Columns("i").ColumnWidth = 8
        worksheet1.Columns("j").ColumnWidth = 8
        worksheet1.Columns("k").ColumnWidth = 8
        worksheet1.Columns("l").ColumnWidth = 8
        worksheet1.Columns("m").ColumnWidth = 8
        worksheet1.Columns("n").ColumnWidth = 8
        worksheet1.Columns("o").ColumnWidth = 8
        worksheet1.Columns("p").ColumnWidth = 8
        worksheet1.Columns("q").ColumnWidth = 10
        worksheet1.Columns("r").ColumnWidth = 10
        worksheet1.Columns("s").ColumnWidth = 10
        worksheet1.Columns("t").ColumnWidth = 10
        worksheet1.Columns("u").ColumnWidth = 10
        worksheet1.Columns("v").ColumnWidth = 10
        worksheet1.Columns("w").ColumnWidth = 10
        worksheet1.Columns("x").ColumnWidth = 10
        worksheet1.Columns("y").ColumnWidth = 10
        worksheet1.Columns("z").ColumnWidth = 60
     
        '-----------------------------------------------------------------------
        worksheet1.Rows("6:6").rowheight = 80
        '-----------------------------------------------------------------------
        worksheet1.Cells(6, 1) = "  MATERIAL"
        worksheet1.Cells(6, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 1).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 1).Orientation = 90
        worksheet1.Rows(6).Font.size = 9

        worksheet1.Cells(6, 2) = "  DESCRIPTION"
        worksheet1.Cells(6, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 2).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 2).Orientation = 90

        worksheet1.Cells(6, 3) = "   PkSiz"
        worksheet1.Cells(6, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 3).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 3).Orientation = 90

        worksheet1.Cells(6, 4) = "  Catg"
        worksheet1.Cells(6, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 4).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 4).Orientation = 90

        worksheet1.Cells(6, 5) = " MR"
        worksheet1.Cells(6, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 5).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 5).Orientation = 90

        worksheet1.Cells(6, 6) = " Alarm Days"
        worksheet1.Cells(6, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(6, 6).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 6).Orientation = 90

        worksheet1.Cells(6, 7) = " End stock"
        worksheet1.Cells(6, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 7).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 7).Orientation = 90

        worksheet1.Cells(6, 8) = " Pending Pos"
        worksheet1.Cells(6, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 8).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 8).Orientation = 90

        worksheet1.Cells(6, 9) = " N3"
        worksheet1.Cells(6, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 9).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 9).Orientation = 90
        Dim _month As Integer
        _month = Month(Today)

        _month = _month - 1
        If _month = 0 Then
            _month = 12
        End If
        worksheet1.Cells(6, 10) = " Consumption-" & MonthName(_month)
        worksheet1.Cells(6, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 10).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 10).Orientation = 90
        Dim _Lastweek As Integer
        Dim I1 As Integer
        I1 = 0
        Dim _NRow As Integer
        _NRow = 11
        _WeekNo = DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text))
        SQL = "select M12Year,M12Week from M12FDP where M12Week>'" & _WeekNo & "' and M12Year>='" & Year(txtFromDate.Text) & "' group by M12Year,M12Week order by M12Year,M12Week"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        For Each DTRow3 As DataRow In T01.Tables(0).Rows
            If I1 = 2 Or I1 = 4 Or I1 = 8 Then
                worksheet1.Cells(6, _NRow) = " Req up-to week " & _Lastweek
                worksheet1.Cells(6, _NRow).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(6, _NRow).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet1.Cells(6, _NRow).Orientation = 90
                _NRow = _NRow + 1
            End If
            _Lastweek = T01.Tables(0).Rows(I1)("m12week")
            ' X = X + 1
            I1 = I1 + 1
        Next

        

        'worksheet1.Cells(6, 11) = " Req up-to week" & _Lastweek
        'worksheet1.Cells(6, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
        'worksheet1.Cells(6, 11).VerticalAlignment = XlVAlign.xlVAlignCenter
        'worksheet1.Cells(6, 11).Orientation = 90
        'worksheet1.Cells(6, 12) = " Req up-to week"
        'worksheet1.Cells(6, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
        'worksheet1.Cells(6, 12).VerticalAlignment = XlVAlign.xlVAlignCenter
        'worksheet1.Cells(6, 12).Orientation = 90
        'worksheet1.Cells(6, 13) = " Req up-to week"
        'worksheet1.Cells(6, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
        'worksheet1.Cells(6, 13).VerticalAlignment = XlVAlign.xlVAlignCenter
        'worksheet1.Cells(6, 13).Orientation = 90w

        worksheet1.Cells(6, 14) = " Req up-to week " & _WeekNo
        worksheet1.Cells(6, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 14).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 14).Orientation = 90
        worksheet1.Cells(6, 15) = " Speed order "
        worksheet1.Cells(6, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 15).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 15).Orientation = 90
        worksheet1.Cells(6, 16) = " TOD N3(End stock)"
        worksheet1.Cells(6, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 16).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 16).Orientation = 90
        worksheet1.Cells(6, 17) = "TOD N3(Pending POs)"
        worksheet1.Cells(6, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 17).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 17).Orientation = 90
        worksheet1.Cells(6, 18) = "TOD " & MonthName(_month) & " (End stock)"
        worksheet1.Cells(6, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 18).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 18).Orientation = 90
        worksheet1.Cells(6, 19) = "TOD Apr (Pending POs)"
        worksheet1.Cells(6, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(6, 19).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 19).Orientation = 90
        worksheet1.Cells(6, 20) = " RL  Re-Ord. Level"
        worksheet1.Cells(6, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 20).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 20).Orientation = 90
        worksheet1.Cells(6, 21) = " RQ"
        worksheet1.Cells(6, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 21).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 21).Orientation = 90
        worksheet1.Cells(6, 22) = "New Order " & Today
        worksheet1.Cells(6, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 22).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 22).Orientation = 90
        worksheet1.Cells(6, 23) = "TOD-new order "
        worksheet1.Cells(6, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 23).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 23).Orientation = 90
        worksheet1.Cells(6, 24) = "TOD(Endstock)+TOD(PO)+TOD(neworder) "
        worksheet1.Cells(6, 24).WrapText = True
        worksheet1.Cells(6, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 24).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 24).Orientation = 90
        worksheet1.Cells(6, 25) = "End stock+pending pos < RL"
        worksheet1.Cells(6, 25).WrapText = True
        worksheet1.Cells(6, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 25).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 25).Orientation = 90

        worksheet1.Cells(6, 26) = "Comment"
        worksheet1.Cells(6, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 26).VerticalAlignment = XlVAlign.xlVAlignCenter

        '-------------------------------------------------------------------------
        worksheet1.Range("A6", "a6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b6", "b6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c6", "c6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d6", "d6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e6", "e6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f6", "f6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u6", "u6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v6", "v6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w6", "w6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("x6", "x6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y6", "y6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("z6", "z6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("z6", "z6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("z6", "z6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y6", "y6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y6", "y6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("x6", "x6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("x6", "x6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w6", "w6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w6", "w6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v6", "v6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v6", "v6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u6", "u6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u6", "u6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f6", "f6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f6", "f6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e6", "e6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e6", "e6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d6", "d6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d6", "d6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c6", "c6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c6", "c6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b6", "b6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b6", "b6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a6", "a6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a6", "a6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        '-----------------------------------------------------------------------------------------------------
        worksheet1.Range("A6:a6").Interior.Color = RGB(219, 238, 243)
        worksheet1.Range("b6:b6").Interior.Color = RGB(219, 238, 243)
        worksheet1.Range("c6:c6").Interior.Color = RGB(219, 238, 243)
        worksheet1.Range("d6:d6").Interior.Color = RGB(219, 238, 243)
        worksheet1.Range("e6:e6").Interior.Color = RGB(219, 238, 243)
        worksheet1.Range("f6:f6").Interior.Color = RGB(219, 238, 243)
        worksheet1.Range("g6:g6").Interior.Color = RGB(255, 153, 51)
        worksheet1.Range("h6:h6").Interior.Color = RGB(255, 204, 0)
        worksheet1.Range("i6:i6").Interior.Color = RGB(219, 238, 243)
        worksheet1.Range("j6:j6").Interior.Color = RGB(219, 238, 243)
        worksheet1.Range("k6:k6").Interior.Color = RGB(242, 221, 220)
        worksheet1.Range("l6:l6").Interior.Color = RGB(242, 221, 220)
        worksheet1.Range("m6:m6").Interior.Color = RGB(242, 221, 220)
        worksheet1.Range("n6:n6").Interior.Color = RGB(242, 221, 220)
        worksheet1.Range("o6:o6").Interior.Color = RGB(255, 255, 102)
        worksheet1.Range("p6:p6").Interior.Color = RGB(255, 153, 51)
        worksheet1.Range("q6:q6").Interior.Color = RGB(255, 204, 0)
        worksheet1.Range("r6:r6").Interior.Color = RGB(255, 153, 51)
        worksheet1.Range("s6:s6").Interior.Color = RGB(255, 204, 0)
        worksheet1.Range("t6:t6").Interior.Color = RGB(0, 204, 255)
        worksheet1.Range("u6:u6").Interior.Color = RGB(51, 204, 51)
        worksheet1.Range("v6:v6").Interior.Color = RGB(255, 255, 102)
        worksheet1.Range("w6:w6").Interior.Color = RGB(234, 241, 221)
        worksheet1.Range("x6:x6").Interior.Color = RGB(247, 150, 70)
        worksheet1.Range("y6:y6").Interior.Color = RGB(255, 255, 102)
        worksheet1.Range("z6:z6").Interior.Color = RGB(255, 255, 102)
        '-------------------------------------------------------------------------------------------
        i = 0
        SQL = "select * from Alarm  where rundate='" & txtFromDate.Text & "'"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        Y = 7
        For Each DTRow4 As DataRow In T01.Tables(0).Rows
            SQL = "select * from M13Material_Category where M13To='" & T01.Tables(0).Rows(i)("SAPCode") & "'"
            T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T03) Then

            Else
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = T01.Tables(0).Rows(i)("SAPCode")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 2) = T01.Tables(0).Rows(i)("Dis")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft
                SQL = "select * from M11MRS where M11SAPCode='" & T01.Tables(0).Rows(i)("SAPCode") & "'"
                T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T04) Then
                    worksheet1.Cells(Y, 3) = T04.Tables(0).Rows(0)("M11PckSize")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 5) = T04.Tables(0).Rows(0)("M11MR")
                    worksheet1.Cells(Y, 5).WrapText = True
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 5).VerticalAlignment = XlVAlign.xlVAlignCenter

                End If
            
                worksheet1.Cells(Y, 6) = T01.Tables(0).Rows(i)("Category")
                worksheet1.Cells(Y, 6).WrapText = True
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet1.Cells(Y, 6) = T01.Tables(0).Rows(i)("L_14Day")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 6)
                range1.NumberFormat = "0"
                '----------------------------------------------------------------------
                Dim _EndStock As Integer
                Dim _Pending_PO As Double
                Dim _N3 As Double
             

                Dim X As Integer
                _EndStock = 0
                _Pending_PO = 0
                _N3 = 0
                _RL = 0
                _RQ = 0

                _EndStock = T01.Tables(0).Rows(i)("EndStock")
                _Pending_PO = T01.Tables(0).Rows(i)("Pending_PO")
                _N3 = T01.Tables(0).Rows(i)("n3")
                _RL = T01.Tables(0).Rows(i)("RL")
                _RQ = T01.Tables(0).Rows(i)("RQ")

                SQL = "select sum(EndStock) as EndStock,sum(Pending_PO) as Pending_PO,sum(N3) as N3,sum(RL) as RL from Alarm inner join M13Material_Category on M13To=SAPCode where M13from='" & T01.Tables(0).Rows(i)("SAPCode") & "' group by M13from"
                T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                X = 0
                For Each DTRow1 As DataRow In T04.Tables(0).Rows
                    _EndStock = _EndStock + Val(T04.Tables(0).Rows(X)("EndStock"))
                    _Pending_PO = _Pending_PO + Val(T04.Tables(0).Rows(X)("Pending_PO"))
                    _N3 = _N3 + Val(T04.Tables(0).Rows(X)("n3"))
                    _RL = _RL + Val(T04.Tables(0).Rows(X)("RL"))
                    _RQ = _RQ + Val(T04.Tables(0).Rows(X)("RQ"))

                    X = X + 1
                Next
                'END STOCK
                worksheet1.Cells(Y, 7) = _EndStock
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(Y, 7)
                range1.NumberFormat = "0"
                '---------------------------------------------------
                'PENDING POS
                worksheet1.Cells(Y, 8) = _Pending_PO
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(Y, 8)
                range1.NumberFormat = "0"
                '----------------------------------------------------
                'N3
                worksheet1.Cells(Y, 9) = _N3
                worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(Y, 9)
                range1.NumberFormat = "0"
                '----------------------------------------------------
                'Last month Consumpshion
                Dim _LastMonthConsump As Double

                _LastMonthConsump = 0
                If Month(Today) = 1 Then
                    SQL = "select * from M11MRS where M11year='" & Year(Today) - 1 & "' and m11month='" & _month & "' and m11SAPCode='" & T01.Tables(0).Rows(i)("SAPCode") & "'"
                Else
                    SQL = "select * from M11MRS where M11year='" & Year(Today) & "' and m11month='" & _month & "' and m11SAPCode='" & T01.Tables(0).Rows(i)("SAPCode") & "'"
                End If
                T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T04) Then
                    _LastMonthConsump = T04.Tables(0).Rows(0)("M11Value")
                    'worksheet1.Cells(Y, 10) = T04.Tables(0).Rows(0)("M11Value")
                    'worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
                    'range1 = worksheet1.Cells(Y, 10)
                    'range1.NumberFormat = "0"
                End If
                '-------------------------------------------------------------------------------
                SQL = "select sum(M11Value) as M11Value  from M11MRS inner join M13Material_Category on M13To=M11SAPCode where M13from='" & T01.Tables(0).Rows(i)("SAPCode") & "' group by M13from"
                T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                X = 0
                For Each DTRow1 As DataRow In T04.Tables(0).Rows
                    _LastMonthConsump = _LastMonthConsump + Val(T04.Tables(0).Rows(X)("M11Value"))
                  

                    X = X + 1
                Next
                worksheet1.Cells(Y, 10) = _LastMonthConsump
                worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(Y, 10)
                range1.NumberFormat = "0"
                '---------------------------------------------------------------------------------
             

                _1stWeek = 0
                _2ndWeek = 0
                _3rdweek = 0
                _CWeek = 0
                SQL = "select * from M12FDP where M12sapcode='" & T01.Tables(0).Rows(i)("SAPCode") & "' and M12Week>='" & _WeekNo & "' and M12Year>='" & Year(Today) & "' order by m12sapcode,m12year,m12week"
                T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                X = 0
                _DD = 0
                For Each DTRow1 As DataRow In T04.Tables(0).Rows

                    If X = 2 Or X = 4 Or X = 8 Then
                        If X = 2 Then
                            _1stWeek = _DD
                            _DD = 0
                        ElseIf X = 4 Then
                            _2ndWeek = _DD
                            _DD = 0
                        ElseIf X = 8 Then
                            _3rdweek = _DD
                            _DD = 0
                        End If
                    End If

                    If T04.Tables(0).Rows(X)("M12Qty") > 0 Then
                        _DD = _DD + T04.Tables(0).Rows(X)("M12Qty")
                        If X = 0 Then
                            _CWeek = T04.Tables(0).Rows(X)("M12Qty")
                        End If
                    Else

                    End If



                    X = X + 1

                Next
                '------------------------------------------------------------------------------------
                SQL = "select sum(M12Qty) from M12FDP inner join M13Material_Category on M12sapcode=M13To where M13from='" & T01.Tables(0).Rows(i)("SAPCode") & "' and M12Week>='" & _WeekNo & "' and M12Year>='" & Year(Today) & "' group by M13from"
                T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                X = 0
                _DD = 0
                For Each DTRow1 As DataRow In T04.Tables(0).Rows

                    If X = 2 Or X = 4 Or X = 8 Then
                        If X = 2 Then
                            _1stWeek = _1stWeek + _DD
                            _DD = 0
                        ElseIf X = 4 Then
                            _2ndWeek = _2ndWeek + _DD
                            _DD = 0
                        ElseIf X = 8 Then
                            _3rdweek = _3rdweek + _DD
                            _DD = 0
                        End If
                    End If

                    If T04.Tables(0).Rows(X)("M12Qty") > 0 Then
                        _DD = _DD + T04.Tables(0).Rows(X)("M12Qty")
                        If X = 0 Then
                            _CWeek = _CWeek + T04.Tables(0).Rows(X)("M12Qty")
                        End If
                    Else

                    End If



                    X = X + 1

                Next
                End If
            '--------------------------------------------------------------------------------------------------------------------
            worksheet1.Cells(Y, 11) = _1stWeek
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 11)
            range1.NumberFormat = "0"

            worksheet1.Cells(Y, 12) = _2ndWeek + _1stWeek
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"

            worksheet1.Cells(Y, 13) = _3rdweek + _2ndWeek + _1stWeek
            worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 13)
            range1.NumberFormat = "0"

            worksheet1.Cells(Y, 14) = _CWeek
            worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 14)
            range1.NumberFormat = "0"
            '---------------------------------------------------------------------

            worksheet1.Range("p" & (Y)).Formula = "=G" & Y & "/i" & Y & "*30"
            worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 16)
            range1.NumberFormat = "0"

            worksheet1.Range("q" & (Y)).Formula = "=h" & Y & "/i" & Y & "*30"
            worksheet1.Cells(Y, 17).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 17)
            range1.NumberFormat = "0"


            worksheet1.Range("R" & (Y)).Formula = "=g" & Y & "/j" & Y & "*21"
            worksheet1.Cells(Y, 18).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 18)
            range1.NumberFormat = "0"

            worksheet1.Range("s" & (Y)).Formula = "=h" & Y & "/j" & Y & "*21"
            worksheet1.Cells(Y, 19).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 19)
            range1.NumberFormat = "0"
         
            worksheet1.Cells(Y, 20) = _RL
            worksheet1.Cells(Y, 20).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 20)
            range1.NumberFormat = "0"

            worksheet1.Cells(Y, 21) = _RQ
            worksheet1.Cells(Y, 21).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 21)
            range1.NumberFormat = "0"

            worksheet1.Range("w" & (Y)).Formula = "=v" & Y & "/j" & Y & "*21"
            worksheet1.Cells(Y, 23).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 23)
            range1.NumberFormat = "0"

            worksheet1.Range("x" & (Y)).Formula = "=S" & Y & "+r" & Y & "+w" & Y & ""
            worksheet1.Cells(Y, 24).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 24)
            range1.NumberFormat = "0"

            worksheet1.Range("y" & (Y)).Formula = "=g" & Y & "+h" & Y & "-t" & Y & ""
            worksheet1.Cells(Y, 24).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(Y, 24)
            range1.NumberFormat = "0"


            worksheet1.Range("z" & Y, "z" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("z" & Y, "z" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("y" & Y, "y" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("y" & Y, "y" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("x" & Y, "x" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("x" & Y, "x" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("w" & Y, "w" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("w" & Y, "w" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("v" & Y, "v" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("v" & Y, "v" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("u" & Y, "u" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("u" & Y, "u" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("t" & Y, "t" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("t" & Y, "t" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("s" & Y, "s" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("s" & Y, "s" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous


           
            worksheet1.Range("g" & Y & ":g" & Y).Interior.Color = RGB(255, 153, 51)
            worksheet1.Range("h" & Y & ":h" & Y).Interior.Color = RGB(255, 204, 0)
            worksheet1.Range("o" & Y & ":o" & Y).Interior.Color = RGB(255, 255, 102)
            worksheet1.Range("p" & Y & ":p" & Y).Interior.Color = RGB(255, 153, 51)
            worksheet1.Range("q" & Y & ":q" & Y).Interior.Color = RGB(255, 204, 0)
            worksheet1.Range("r" & Y & ":r" & Y).Interior.Color = RGB(255, 153, 51)
            worksheet1.Range("s" & Y & ":s" & Y).Interior.Color = RGB(255, 204, 0)
            worksheet1.Range("t" & Y & ":t" & Y).Interior.Color = RGB(0, 204, 255)
            worksheet1.Range("u" & Y & ":u" & Y).Interior.Color = RGB(51, 204, 51)
           
                Y = Y + 1

                i = i + 1
        Next

    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Order_Analysis()
    End Sub
End Class