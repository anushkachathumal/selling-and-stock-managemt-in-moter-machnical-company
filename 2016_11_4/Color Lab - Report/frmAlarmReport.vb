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
Public Class frmAlarmReport
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
            With cboFrom
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 125
            End With

            With cboTo
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 125
            End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmAlarmReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_SAPNo()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        OPR1.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False

        cboFrom.ToggleDropdown()

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

    Function AlarmReport()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet
        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        'Dim exc As New Application
        'Dim workbooks As Workbooks = exc.Workbooks
        'Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sheets1 As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        If exc.Visible = True Then
            exc.Visible = False
            exc.Visible = True
        Else
            ' exc.Visible = False
            exc.Visible = True
        End If

        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        Dim range1 As Range
        Dim _NETTOTAL As Integer
        Dim T04 As DataSet
        Dim n_per As Double
        Dim Y As Integer
        Dim _cOUNT As Integer

        '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
        ' workbooks.Application.Sheets.Ahedd()
        Try
            ' workbooks.Application.Sheets.Add()

            'Dim sheets1 As Sheets = workbook.Worksheets
            ' Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet1.Name = "DCA ALARM REPORT_" & Month(Today) & "." & Microsoft.VisualBasic.Day(Today) & "." & Year(Today)


            worksheet1.Cells(1, 2) = "Textured Jersey Lanka Pvt Ltd"
            worksheet1.Cells(2, 2) = "Alarm Report"
            worksheet1.Cells(3, 2) = "Report Date : " & Month(Today) & "." & Microsoft.VisualBasic.Day(Today) & "." & Year(Today)
            worksheet1.Cells(4, 2) = "Report Time : " & Hour(VserverTime) & ":" & Minute(VserverTime) & ":" & Second(VserverTime)

            worksheet1.Columns("A").ColumnWidth = 12
            worksheet1.Columns("B").ColumnWidth = 40
            worksheet1.Range("A1:B1").Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(1).Font.size = 10

            worksheet1.Range("A2:B2").Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(2).Font.size = 10
            worksheet1.Range("A3:B3").Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(3).Font.size = 10
            worksheet1.Range("A4:B4").Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(4).Font.size = 10


            worksheet1.Cells(5, 1) = "SAP-Code"
            worksheet1.Cells(5, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Rows(5).Font.size = 10
            worksheet1.Columns("A").ColumnWidth = 10

            worksheet1.Cells(5, 2) = "Description"
            worksheet1.Cells(5, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Columns("B").ColumnWidth = 40

            worksheet1.Cells(5, 3) = "Category"
            worksheet1.Cells(5, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Columns("C").ColumnWidth = 10

            worksheet1.Cells(5, 4) = "SS"
            worksheet1.Cells(5, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Columns("D").ColumnWidth = 10

            worksheet1.Cells(5, 5) = "End Stock"
            worksheet1.Cells(5, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Columns("E").ColumnWidth = 10
            worksheet1.Columns("F").ColumnWidth = 10

            worksheet1.Cells(5, 7) = "AVERAGE CONSUMPTION"
            worksheet1.Range(worksheet1.Cells(5, 7), worksheet1.Cells(5, 12)).Merge()
            worksheet1.Cells(5, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Columns("G").ColumnWidth = 10
            worksheet1.Columns("I").ColumnWidth = 10
            worksheet1.Columns("J").ColumnWidth = 10
            worksheet1.Columns("K").ColumnWidth = 10
            worksheet1.Columns("L").ColumnWidth = 10
            worksheet1.Columns("M").ColumnWidth = 10

            worksheet1.Cells(5, 13) = "COMPARISION BETWEEN CQ-MRS & BELOW (%)"
            worksheet1.Range(worksheet1.Cells(5, 13), worksheet1.Cells(5, 16)).Merge()
            worksheet1.Cells(5, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Columns("M").ColumnWidth = 10
            worksheet1.Columns("N").ColumnWidth = 10
            worksheet1.Columns("O").ColumnWidth = 10
            worksheet1.Columns("P").ColumnWidth = 10
            worksheet1.Columns("Q").ColumnWidth = 10

            worksheet1.Cells(5, 17) = "COMPARISION BETWEEN CQ-MRS & BELOW (KG)"
            worksheet1.Range(worksheet1.Cells(5, 17), worksheet1.Cells(5, 20)).Merge()
            worksheet1.Cells(5, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Columns("R").ColumnWidth = 10
            worksheet1.Columns("S").ColumnWidth = 10
            worksheet1.Columns("T").ColumnWidth = 10

            worksheet1.Columns("U").ColumnWidth = 10
            worksheet1.Cells(5, 21) = " Days"
            worksheet1.Cells(5, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("V").ColumnWidth = 10
            worksheet1.Cells(5, 22) = " RL"
            worksheet1.Cells(5, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("W").ColumnWidth = 10
            worksheet1.Cells(5, 23) = " RQ"
            worksheet1.Cells(5, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Columns("X").ColumnWidth = 10
            worksheet1.Cells(5, 24) = " LT-Days"
            worksheet1.Cells(5, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("Y").ColumnWidth = 10
            worksheet1.Cells(5, 25) = " Pindin PO"
            worksheet1.Cells(5, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("Z").ColumnWidth = 10
            worksheet1.Cells(5, 26) = " item #"
            worksheet1.Cells(5, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("AA").ColumnWidth = 10
            worksheet1.Cells(5, 27) = "  Po#"
            worksheet1.Cells(5, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("Ab").ColumnWidth = 10
            worksheet1.Cells(5, 28) = "  Po issue dat"
            worksheet1.Cells(5, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("Ac").ColumnWidth = 10
            worksheet1.Cells(5, 29) = "  Po del date"
            worksheet1.Cells(5, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("AD").ColumnWidth = 10
            worksheet1.Cells(5, 30) = "  ETD"
            worksheet1.Cells(5, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Columns("AE").ColumnWidth = 10
            worksheet1.Cells(5, 31) = "  ETA"
            worksheet1.Cells(5, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Rows(6).Font.size = 10
            worksheet1.Cells(6, 7) = "  AVG 12"
            worksheet1.Cells(6, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(6, 8) = "    M6"
            worksheet1.Cells(6, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 9) = "   N3"
            worksheet1.Cells(6, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 10) = "    CQ-MRS"
            worksheet1.Cells(6, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 11) = "    Last Month"
            worksheet1.Cells(6, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 12) = " L14D Con"
            worksheet1.Cells(6, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 13) = " Last 30 days"
            worksheet1.Cells(6, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 14) = "SD"
            worksheet1.Cells(6, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 15) = "  2ND L WEEK"
            worksheet1.Cells(6, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 16) = "  LAST WEEK"
            worksheet1.Cells(6, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 17) = "   Last 30 days"
            worksheet1.Cells(6, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 18) = " SD"
            worksheet1.Cells(6, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 19) = " 2ND L WEEK"
            worksheet1.Cells(6, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(6, 20) = " LAST WEEK"
            worksheet1.Cells(6, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
            '-----------------------------------------------------------------------------\
            'COLOUR
            worksheet1.Range("A5:a5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("b5:b5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("c5:c5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("d5:d5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("e5:e5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("f5:f5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("g5:g5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("m5:m5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("q5:q5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("u5:u5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("v5:v5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("w5:w5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("x5:x5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("y5:y5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("z5:z5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("aA5:aa5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("ab5:ab5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("ac5:ac5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("ad5:ad5").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("ae5:ae5").Interior.Color = RGB(191, 191, 191)


            worksheet1.Range("A6:a6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("b6:b6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("c6:c6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("d6:d6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("e6:e6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("f6:f6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("g6:g6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("h6:h6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("i6:i6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("j6:j6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("k6:k6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("l6:l6").Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("m6:m6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("n6:n6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("o6:o6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("p6:p6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("q6:q6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("r6:r6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("s6:s6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("t6:t6").Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("u6:u6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("v6:v6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("w6:w6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("x6:x6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("y6:y6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("z6:z6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("aA6:aa6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("ab6:ab6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("ac6:ac6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("ad6:ad6").Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("ae6:ae6").Interior.Color = RGB(191, 191, 191)


            worksheet1.Range("A5", "a5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b5", "b5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c5", "c5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d5", "d5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e5", "e5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f5", "f5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g5", "g5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h5", "h5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i5", "i5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j5", "j5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k5", "k5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l5", "l5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m5", "m5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n5", "n5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o5", "o5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p5", "p5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q5", "q5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r5", "r5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("s5", "s5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("t5", "t5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("u5", "u5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("v5", "v5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("w5", "w5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("x5", "x5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("y5", "y5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("z5", "z5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("aa5", "aa5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ab5", "ab5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ac5", "ac5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ad5", "ad5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ae5", "ae5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("ae5", "ae5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ae6", "ae6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ae6", "ae6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ae6", "ae6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ae5", "ae5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("ad6", "ad6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ad6", "ad6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ad5", "ad5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("ac6", "ac6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ac6", "ac6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ac5", "ac5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("ab6", "ab6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ab6", "ab6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("ab5", "ab5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("aa6", "aa6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("aa6", "aa6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("aa5", "aa5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("z6", "z6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("z6", "z6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("z5", "z5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("y6", "y6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("y6", "y6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("y5", "y5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("x6", "x6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("x6", "x6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("x5", "x5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("w6", "w6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("w6", "w6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("w5", "w5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("v6", "v6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("v6", "v6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("v5", "v5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("u6", "u6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("u6", "u6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("u5", "u5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q5", "q5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m5", "m5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i5", "i5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g5", "g5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f6", "f6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f6", "f6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f5", "f5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e6", "e6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e6", "e6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e5", "e5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d6", "d6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d6", "d6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d5", "d5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("c6", "c6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c6", "c6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c5", "c5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b6", "b6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b6", "b6").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b5", "b5").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("a6", "a6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous


            i = 0
            Y = 7
            SQL = "select * from Alarm where SAPCode between '" & cboFrom.Text & "' and '" & cboTo.Text & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 8
                If T01.Tables(0).Rows(i)("SAPCode") = "600250" Then
                    '     MsgBox("")
                End If
                worksheet1.Cells(Y, 1) = T01.Tables(0).Rows(i)("SAPCode")
                worksheet1.Cells(Y, 2) = T01.Tables(0).Rows(i)("Dis")
                SQL = "select * from M11MRS where M11SAPCode='" & T01.Tables(0).Rows(i)("SAPCode") & "'"
                T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T03) Then
                    worksheet1.Cells(Y, 3) = T03.Tables(0).Rows(0)("M11Category")
                    worksheet1.Cells(Y, 6) = T03.Tables(0).Rows(0)("M11Purchase")
                    range1 = worksheet1.Cells(Y, 6)
                    range1.NumberFormat = "0"
                End If
                worksheet1.Cells(Y, 4) = T01.Tables(0).Rows(i)("SS")
                worksheet1.Cells(Y, 5) = T01.Tables(0).Rows(i)("EndStock")
                range1 = worksheet1.Cells(Y, 5)
                range1.NumberFormat = "0"

                'worksheet1.Cells(Y, 6) = T01.Tables(0).Rows(i)("L14Day_Con")
                'range1 = worksheet1.Cells(Y, 6)
                'range1.NumberFormat = "0"

                worksheet1.Cells(Y, 7) = T01.Tables(0).Rows(i)("Avg")
                range1 = worksheet1.Cells(Y, 7)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 8) = T01.Tables(0).Rows(i)("M6")
                range1 = worksheet1.Cells(Y, 8)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 9) = T01.Tables(0).Rows(i)("N3")
                range1 = worksheet1.Cells(Y, 9)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 10) = T01.Tables(0).Rows(i)("CQ_MRS")
                range1 = worksheet1.Cells(Y, 10)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 11) = T01.Tables(0).Rows(i)("Last_MntQty")
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = T01.Tables(0).Rows(i)("L14Day_Con")
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = T01.Tables(0).Rows(i)("Last_30Day")
                range1 = worksheet1.Cells(Y, 13)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 14) = T01.Tables(0).Rows(i)("SD")
                range1 = worksheet1.Cells(Y, 14)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 15) = T01.Tables(0).Rows(i)("L_Week2nd")
                range1 = worksheet1.Cells(Y, 15)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 16) = T01.Tables(0).Rows(i)("L_week2")
                range1 = worksheet1.Cells(Y, 16)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 17) = T01.Tables(0).Rows(i)("L_30Day2")
                range1 = worksheet1.Cells(Y, 17)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 18) = T01.Tables(0).Rows(i)("SD2")
                range1 = worksheet1.Cells(Y, 18)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 19) = T01.Tables(0).Rows(i)("L_Week2nd2")
                range1 = worksheet1.Cells(Y, 19)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 20) = T01.Tables(0).Rows(i)("L_Week3")
                range1 = worksheet1.Cells(Y, 20)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 21) = T01.Tables(0).Rows(i)("L_14Day")
                range1 = worksheet1.Cells(Y, 21)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 22) = T01.Tables(0).Rows(i)("RL")
                range1 = worksheet1.Cells(Y, 22)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 23) = T01.Tables(0).Rows(i)("RQ")
                range1 = worksheet1.Cells(Y, 23)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 24) = T01.Tables(0).Rows(i)("LT_day")
                range1 = worksheet1.Cells(Y, 24)
                range1.NumberFormat = "0"
                worksheet1.Cells(Y, 25) = T01.Tables(0).Rows(i)("Pending_PO")
                worksheet1.Cells(Y, 26) = T01.Tables(0).Rows(i)("ItemNo")
                worksheet1.Cells(Y, 27) = T01.Tables(0).Rows(i)("PO")
                If T01.Tables(0).Rows(i)("PODate") = "1/1/1900" Then
                Else
                    worksheet1.Cells(Y, 28) = T01.Tables(0).Rows(i)("PODate")
                End If
                If T01.Tables(0).Rows(i)("PODelDate") = "1/1/1900" Then
                Else
                    worksheet1.Cells(Y, 29) = T01.Tables(0).Rows(i)("PODelDate")
                End If

                If T01.Tables(0).Rows(i)("ETD") = "1/1/1900" Then
                Else
                    worksheet1.Cells(Y, 30) = T01.Tables(0).Rows(i)("ETD")
                End If

                If T01.Tables(0).Rows(i)("ETA") = "1/1/1900" Then
                Else
                    worksheet1.Cells(Y, 31) = T01.Tables(0).Rows(i)("ETA")
                End If

                worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("s" & Y, "s" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("s" & Y, "s" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("t" & Y, "t" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("t" & Y, "t" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("u" & Y, "u" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("u" & Y, "u" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("v" & Y, "v" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("v" & Y, "v" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("w" & Y, "w" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("w" & Y, "w" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("x" & Y, "x" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("x" & Y, "x" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("y" & Y, "y" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("y" & Y, "y" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("z" & Y, "z" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("z" & Y, "z" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Aa" & Y, "Aa" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Aa" & Y, "Aa" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Ab" & Y, "Ab" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Ab" & Y, "Ab" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Ac" & Y, "Ac" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Ac" & Y, "Ac" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Ad" & Y, "Ad" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Ad" & Y, "Ad" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                worksheet1.Range("Ae" & Y, "Ae" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("Ae" & Y, "Ae" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash









                Y = Y + 1
                i = i + 1
            Next


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call AlarmReport()
        Call DCA()
    End Sub
    Function DCA()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim T04 As DataSet

        Dim dsUser As DataSet
        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        ' exc.Visible = True
        Dim i As Integer
        Dim _GrandTotal As Integer
        Dim _STGrand As String
        Dim range1 As Range
        Dim _NETTOTAL As Integer
        ' Dim T04 As DataSet
        Dim n_per As Double
        Dim Y As Integer
        Dim _cOUNT As Integer
        Dim x As Integer
        '' Dim exc As New Application
        '' Dim workbooks As Workbooks = exc.Workbooks
        '' Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        ''Dim sheets As Sheets = workbook.Worksheets
        ' Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
        '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
        workbooks.Application.Sheets.Add()
        Dim sheets1 As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        Try
            ' workbooks.Application.Sheets.Add()
            'Dim sheets1 As Sheets = workbook.Worksheets
            ' Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet1.Name = "DCA ALARM REPORT_" & Microsoft.VisualBasic.Day(Today) & "." & Month(Today) & ".Days >30"

            worksheet1.Rows(1).Font.size = 10
            worksheet1.Rows(2).Font.size = 10
            worksheet1.Rows(3).Font.size = 10
            worksheet1.Rows(4).Font.size = 10
            worksheet1.Rows(4).Font.bold = True
            worksheet1.Cells(1, 2) = "Textured Jersey Lanka Pvt Ltd"
            worksheet1.Cells(2, 2) = "DCA ALARM SYSTEM"
            worksheet1.Cells(3, 2) = "Report Date : " & Month(Today) & "." & Microsoft.VisualBasic.Day(Today) & "." & Year(Today)
            worksheet1.Cells(4, 2) = "Purchase PTL  " ' & Hour(VserverTime) & ":" & Minute(VserverTime) & ":" & Second(VserverTime)

            worksheet1.Columns("A").ColumnWidth = 12
            worksheet1.Columns("B").ColumnWidth = 40
            worksheet1.Columns("C").ColumnWidth = 10
            worksheet1.Columns("D").ColumnWidth = 10
            worksheet1.Columns("E").ColumnWidth = 10
            worksheet1.Columns("F").ColumnWidth = 10
            worksheet1.Columns("G").ColumnWidth = 10
            worksheet1.Columns("I").ColumnWidth = 10
            worksheet1.Columns("J").ColumnWidth = 10
            worksheet1.Columns("K").ColumnWidth = 10
            worksheet1.Columns("L").ColumnWidth = 10
            worksheet1.Columns("M").ColumnWidth = 10
            worksheet1.Columns("M").ColumnWidth = 10
            worksheet1.Columns("N").ColumnWidth = 10
            worksheet1.Columns("O").ColumnWidth = 10
            worksheet1.Columns("P").ColumnWidth = 10
            worksheet1.Columns("Q").ColumnWidth = 10
            worksheet1.Columns("R").ColumnWidth = 10
            worksheet1.Columns("S").ColumnWidth = 10
            worksheet1.Columns("T").ColumnWidth = 10

            worksheet1.Columns("U").ColumnWidth = 10

            worksheet1.Columns("V").ColumnWidth = 10
            worksheet1.Columns("W").ColumnWidth = 10
            worksheet1.Columns("X").ColumnWidth = 10
            worksheet1.Columns("Y").ColumnWidth = 10
            worksheet1.Columns("Z").ColumnWidth = 10
            worksheet1.Columns("AA").ColumnWidth = 10
            worksheet1.Columns("Ab").ColumnWidth = 10
            worksheet1.Columns("Ac").ColumnWidth = 10
            worksheet1.Columns("AD").ColumnWidth = 10
            worksheet1.Columns("AE").ColumnWidth = 10


            worksheet1.Range("A1:B1").Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(1).Font.size = 10
            worksheet1.Range("A2:B2").Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(2).Font.size = 10
            worksheet1.Range("A3:B3").Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(3).Font.size = 10
            worksheet1.Range("A4:B4").Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(4).Font.size = 10

            Y = 5
            worksheet1.Cells(Y, 1) = "SAP-Code"
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Rows(Y).Font.size = 10

            worksheet1.Cells(Y, 2) = "Description"
            worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 3) = "Category"
            worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 4) = "SS"
            worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 5) = "End Stock"
            worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 6) = "Purchase"
            worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 7) = "AVERAGE CONSUMPTION"
            'worksheet1.Range(worksheet1.Cells(Y, 7), worksheet1.Cells(5, 12)).Merge()
            worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 8) = " Days"
            worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 9) = " RL"
            worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 10) = " RQ"
            worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 11) = " LT-Days"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 12) = " Pindin PO"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 13) = " Req No"
            worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 14) = "  Po#"
            worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 15) = "  Po issue dat"
            worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 16) = "  Po del date"
            worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 17) = "  ETD"
            worksheet1.Cells(Y, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 18) = "  ETA"
            worksheet1.Cells(Y, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter

            Y = Y + 1
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Cells(Y, 6) = "  From"
            worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 7) = "    L14D Con"
            worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Range("A" & Y - 1 & ":A" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("A" & Y & ":A" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("b" & Y - 1 & ":b" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("b" & Y & ":b" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("c" & Y - 1 & ":c" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("c" & Y & ":c" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("d" & Y - 1 & ":d" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("d" & Y & ":d" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("e" & Y - 1 & ":e" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("e" & Y & ":e" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("f" & Y - 1 & ":f" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("f" & Y & ":f" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("g" & Y - 1 & ":g" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("g" & Y & ":g" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("h" & Y - 1 & ":h" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("h" & Y & ":h" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("i" & Y - 1 & ":i" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("i" & Y & ":i" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("j" & Y - 1 & ":j" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("j" & Y & ":j" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("k" & Y - 1 & ":k" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("k" & Y & ":k" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("l" & Y - 1 & ":l" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("l" & Y & ":l" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("m" & Y - 1 & ":m" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("m" & Y & ":m" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("n" & Y - 1 & ":n" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("n" & Y & ":n" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("o" & Y - 1 & ":o" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("o" & Y & ":o" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("p" & Y - 1 & ":p" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("p" & Y & ":p" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("q" & Y - 1 & ":q" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("q" & Y & ":q" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("r" & Y - 1 & ":r" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("r" & Y & ":r" & Y).Interior.Color = RGB(191, 191, 191)


            worksheet1.Range("A" & Y - 1, "A" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("A" & Y - 1, "A" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("b" & Y - 1, "b" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y - 1, "b" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("c" & Y - 1, "c" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y - 1, "c" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("d" & Y - 1, "d" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y - 1, "d" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("e" & Y - 1, "e" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y - 1, "e" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("f" & Y - 1, "f" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y - 1, "f" & Y - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y - 1, "f" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("g" & Y - 1, "g" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y - 1, "g" & Y - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y - 1, "g" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("h" & Y - 1, "h" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y - 1, "h" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("i" & Y - 1, "i" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y - 1, "i" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("j" & Y - 1, "j" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y - 1, "j" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("k" & Y - 1, "k" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y - 1, "k" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("l" & Y - 1, "l" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y - 1, "l" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("m" & Y - 1, "m" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y - 1, "m" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("n" & Y - 1, "n" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y - 1, "n" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("o" & Y - 1, "o" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y - 1, "o" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("p" & Y - 1, "p" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y - 1, "p" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("q" & Y - 1, "q" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y - 1, "q" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("r" & Y - 1, "r" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y - 1, "r" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            '==========================================================================================================================
            'PTL
            Y = Y + 1
            i = 0
            SQL = "select max(M11Purchase) as M11Purchase,M11Sapcode from M11MRS where M11Sapcode between '" & cboFrom.Text & "' and '" & cboTo.Text & "' and M11Purchase='PTL' group by M11Sapcode"
            T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow3 As DataRow In T03.Tables(0).Rows
                ' SQL = "select Last_MntQty,SAPCode,Dis,Category,SS,EndStock,M11Purchase,L_14Day,RL,RQ,LT_day,Pending_PO,PODate,PODelDate,ItemNo,Po,ETA,ETD from Alarm inner join M11MRS on SAPCode=M11SAPCode where L_14Day<=30  and SAPCode ='" & T03.Tables(0).Rows(i)("M11Sapcode") & "'"
                x = 0
                SQL = "select * from Alarm where L_14Day<=30  and SAPCode ='" & T03.Tables(0).Rows(i)("M11Sapcode") & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                For Each DTRow2 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(Y).Font.size = 8
                    worksheet1.Cells(Y, 1) = T01.Tables(0).Rows(x)("SAPCode")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = T01.Tables(0).Rows(x)("Dis")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    SQL = "select * from M11MRS where M11SAPCode='" & T01.Tables(0).Rows(x)("SAPCode") & "'"
                    T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(T03) Then


                        worksheet1.Cells(Y, 3) = T04.Tables(0).Rows(0)("M11Category")
                        worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If
                    worksheet1.Cells(Y, 4) = T01.Tables(0).Rows(x)("SS")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 4)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 5) = T01.Tables(0).Rows(x)("EndStock")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 5)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 6) = T03.Tables(0).Rows(i)("M11Purchase")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 7) = T01.Tables(0).Rows(x)("Last_MntQty")
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 7)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 8) = T01.Tables(0).Rows(x)("L_14Day")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 8)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 9) = T01.Tables(0).Rows(x)("RL")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 9)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 10) = T01.Tables(0).Rows(x)("RQ")
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 10)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 11) = T01.Tables(0).Rows(x)("LT_day")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = T01.Tables(0).Rows(x)("Pending_PO")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 13) = T01.Tables(0).Rows(x)("ItemNo")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 14) = T01.Tables(0).Rows(x)("PO")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    If T01.Tables(0).Rows(x)("PODate") = "1/1/1900" Then
                    Else
                        worksheet1.Cells(Y, 15) = T01.Tables(0).Rows(x)("PODate")
                        worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If

                    If T01.Tables(0).Rows(x)("PODelDate") = "1/1/1900" Then
                    Else
                        worksheet1.Cells(Y, 16) = T01.Tables(0).Rows(x)("PODelDate")
                        worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If


                    If T01.Tables(0).Rows(x)("ETD") = "1/1/1900" Then
                    Else
                        worksheet1.Cells(Y, 17) = T01.Tables(0).Rows(x)("ETD")
                        worksheet1.Cells(Y, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If

                    If T01.Tables(0).Rows(x)("ETA") = "1/1/1900" Then
                    Else
                        worksheet1.Cells(Y, 18) = T01.Tables(0).Rows(x)("ETA")
                        worksheet1.Cells(Y, 18).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If

                    worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("F" & Y, "F" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("G" & Y, "G" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash

                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                    Y = Y + 1
                    x = x + 1
                Next
                ' Y = Y + 1

                i = i + 1

            Next

            Y = Y + 2
            worksheet1.Cells(Y, 2) = "Textured Jersey Lanka Pvt Ltd"
            Y = Y + 1
            worksheet1.Cells(Y, 2) = "DCA ALARM SYSTEM"
            Y = Y + 1
            worksheet1.Cells(Y, 2) = "Report Date : " & Month(Today) & "." & Microsoft.VisualBasic.Day(Today) & "." & Year(Today)
            Y = Y + 1
            worksheet1.Cells(Y, 2) = "Purchase Local  " ' & Hour(VserverTime) & ":" & Minute(VserverTime) & ":" & Second(VserverTime)

            worksheet1.Range("A" & Y - 3 & ":B" & Y - 3).Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(Y - 3).Font.size = 10
            worksheet1.Range("A" & Y - 2 & ":B" & Y - 2).Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(Y - 2).Font.size = 10
            worksheet1.Range("A" & Y - 1 & ":B" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(Y - 1).Font.size = 10
            worksheet1.Range("A" & Y & ":B" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Rows(Y).Font.size = 10

            Y = Y + 1
            worksheet1.Cells(Y, 1) = "SAP-Code"
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Rows(Y).Font.size = 10

            worksheet1.Cells(Y, 2) = "Description"
            worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 3) = "Category"
            worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 4) = "SS"
            worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 5) = "End Stock"
            worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 6) = "Purchase"
            worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 7) = "AVERAGE CONSUMPTION"
            'worksheet1.Range(worksheet1.Cells(Y, 7), worksheet1.Cells(5, 12)).Merge()
            worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 8) = " Days"
            worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 9) = " RL"
            worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 10) = " RQ"
            worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 11) = " LT-Days"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 12) = " Pindin PO"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 13) = " Req No"
            worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet1.Cells(Y, 14) = "  Po#"
            worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 15) = "  Po issue dat"
            worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 16) = "  Po del date"
            worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 17) = "  ETD"
            worksheet1.Cells(Y, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 18) = "  ETA"
            worksheet1.Cells(Y, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter

            Y = Y + 1
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Cells(Y, 6) = "  From"
            worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(Y, 7) = "    L14D Con"
            worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Range("A" & Y - 1 & ":A" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("A" & Y & ":A" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("b" & Y - 1 & ":b" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("b" & Y & ":b" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("c" & Y - 1 & ":c" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("c" & Y & ":c" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("d" & Y - 1 & ":d" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("d" & Y & ":d" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("e" & Y - 1 & ":e" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("e" & Y & ":e" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("f" & Y - 1 & ":f" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("f" & Y & ":f" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("g" & Y - 1 & ":g" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("g" & Y & ":g" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("h" & Y - 1 & ":h" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("h" & Y & ":h" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("i" & Y - 1 & ":i" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("i" & Y & ":i" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("j" & Y - 1 & ":j" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("j" & Y & ":j" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("k" & Y - 1 & ":k" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("k" & Y & ":k" & Y).Interior.Color = RGB(191, 191, 191)

            worksheet1.Range("l" & Y - 1 & ":l" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("l" & Y & ":l" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("m" & Y - 1 & ":m" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("m" & Y & ":m" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("n" & Y - 1 & ":n" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("n" & Y & ":n" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("o" & Y - 1 & ":o" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("o" & Y & ":o" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("p" & Y - 1 & ":p" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("p" & Y & ":p" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("q" & Y - 1 & ":q" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("q" & Y & ":q" & Y).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("r" & Y - 1 & ":r" & Y - 1).Interior.Color = RGB(191, 191, 191)
            worksheet1.Range("r" & Y & ":r" & Y).Interior.Color = RGB(191, 191, 191)


            worksheet1.Range("A" & Y - 1, "A" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("A" & Y - 1, "A" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("b" & Y - 1, "b" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y - 1, "b" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("c" & Y - 1, "c" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y - 1, "c" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("d" & Y - 1, "d" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y - 1, "d" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("e" & Y - 1, "e" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y - 1, "e" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("f" & Y - 1, "f" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y - 1, "f" & Y - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y - 1, "f" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("g" & Y - 1, "g" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y - 1, "g" & Y - 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y - 1, "g" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("h" & Y - 1, "h" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y - 1, "h" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("i" & Y - 1, "i" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y - 1, "i" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("j" & Y - 1, "j" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y - 1, "j" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("k" & Y - 1, "k" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y - 1, "k" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("l" & Y - 1, "l" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y - 1, "l" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("m" & Y - 1, "m" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y - 1, "m" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("n" & Y - 1, "n" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y - 1, "n" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("o" & Y - 1, "o" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y - 1, "o" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("p" & Y - 1, "p" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y - 1, "p" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("q" & Y - 1, "q" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y - 1, "q" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet1.Range("r" & Y - 1, "r" & Y - 1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y - 1, "r" & Y - 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            '==========================================================================================================================
            'PTL
            Y = Y + 1
            i = 0
            SQL = "select max(M11Purchase) as M11Purchase,M11Sapcode from M11MRS where M11Sapcode between '" & cboFrom.Text & "' and '" & cboTo.Text & "' and M11Purchase in ('Local','Local-Import') group by M11Sapcode"
            T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow3 As DataRow In T03.Tables(0).Rows
                ' SQL = "select Last_MntQty,SAPCode,Dis,Category,SS,EndStock,M11Purchase,L_14Day,RL,RQ,LT_day,Pending_PO,PODate,PODelDate,ItemNo,Po,ETA,ETD from Alarm inner join M11MRS on SAPCode=M11SAPCode where L_14Day<=30  and SAPCode ='" & T03.Tables(0).Rows(i)("M11Sapcode") & "'"
                'T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)

                SQL = "select * from Alarm where L_14Day<=30  and SAPCode ='" & T03.Tables(0).Rows(i)("M11Sapcode") & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                x = 0
                For Each DTRow2 As DataRow In T01.Tables(0).Rows
                    worksheet1.Rows(Y).Font.size = 8
                    worksheet1.Cells(Y, 1) = T01.Tables(0).Rows(x)("SAPCode")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = T01.Tables(0).Rows(x)("Dis")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    SQL = "select * from M11MRS where M11SAPCode='" & T01.Tables(0).Rows(x)("SAPCode") & "'"
                    T04 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(T03) Then


                        worksheet1.Cells(Y, 3) = T04.Tables(0).Rows(0)("M11Category")
                        worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If

                    worksheet1.Cells(Y, 4) = T01.Tables(0).Rows(x)("SS")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 4)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 5) = T01.Tables(0).Rows(x)("EndStock")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 5)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 6) = T03.Tables(0).Rows(i)("M11Purchase")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 7) = T01.Tables(0).Rows(x)("Last_MntQty")
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 7)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 8) = T01.Tables(0).Rows(x)("L_14Day")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 8)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 9) = T01.Tables(0).Rows(x)("RL")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 9)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 10) = T01.Tables(0).Rows(x)("RQ")
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 10)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 11) = T01.Tables(0).Rows(x)("LT_day")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = T01.Tables(0).Rows(x)("Pending_PO")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 13) = T01.Tables(0).Rows(x)("ItemNo")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 14) = T01.Tables(0).Rows(x)("PO")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    If T01.Tables(0).Rows(x)("PODate") = "1/1/1900" Then
                    Else
                        worksheet1.Cells(Y, 15) = T01.Tables(0).Rows(x)("PODate")
                        worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If

                    If T01.Tables(0).Rows(x)("PODelDate") = "1/1/1900" Then
                    Else
                        worksheet1.Cells(Y, 16) = T01.Tables(0).Rows(x)("PODelDate")

                        worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If


                    If T01.Tables(0).Rows(x)("ETD") = "1/1/1900" Then
                    Else
                        worksheet1.Cells(Y, 17) = T01.Tables(0).Rows(x)("ETD")
                        worksheet1.Cells(Y, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If

                    If T01.Tables(0).Rows(x)("ETA") = "1/1/1900" Then
                    Else
                        worksheet1.Cells(Y, 18) = T01.Tables(0).Rows(x)("ETA")
                        worksheet1.Cells(Y, 18).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If

                    worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("A" & Y, "A" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("F" & Y, "F" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("G" & Y, "G" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash

                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("q" & Y, "q" & Y).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("r" & Y, "r" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash

                    x = x + 1

                    Y = Y + 1
                Next
                i = i + 1

            Next
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
End Class