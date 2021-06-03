
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
Public Class MRSReport
    Dim strLine As String
    Dim strLineflu As String
    Dim strDash As String
    Dim StrDisCode As String
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter
    Dim exc As New Application

    Dim workbooks As Workbooks = exc.Workbooks
    Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    Dim sheets As Sheets = Workbook.Worksheets
    Dim worksheet1 As _Worksheet = CType(Sheets.Item(1), _Worksheet)

    Private Sub MRSReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
        'Call Daily_Boliout()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        OPR1.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False

        txtFromDate.Text = Today
        txtTodate.Text = Today

        cmdSave.Enabled = True

    End Sub

    Function ptlMRS_Report()
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

        '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
        workbooks.Application.Sheets.Add()
        Dim sheets1 As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        worksheet1.Name = "MRS_" & Month(txtFromDate.Text) & "." & Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Year(txtFromDate.Text)

        n_Date = CDate(txtFromDate.Text).AddDays(-365)
        ' n_Date = n_Date & " " & "7:30AM"
        N_Date1 = CDate(txtTodate.Text).AddDays(+1)
        ' N_Date1 = N_Date1 & " " & "7:30AM"

        worksheet1.Cells(1, 2) = "Textured Jersey Lanka Pvt Ltd"
        worksheet1.Cells(2, 2) = "MRS Report"
        worksheet1.Cells(3, 2) = "Report Date : " & Month(txtFromDate.Text) & "." & Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        worksheet1.Cells(4, 2) = "Report Time : " & Hour(VserverTime) & ":" & Minute(VserverTime) & ":" & Second(VserverTime)

        worksheet1.Columns("A").ColumnWidth = 12
        worksheet1.Columns("B").ColumnWidth = 28

        worksheet1.Range("A1:B1").Interior.Color = RGB(141, 180, 227)
        worksheet1.Rows(1).Font.size = 10

        worksheet1.Range("A2:B2").Interior.Color = RGB(141, 180, 227)
        worksheet1.Rows(2).Font.size = 10
        worksheet1.Range("A3:B3").Interior.Color = RGB(141, 180, 227)
        worksheet1.Rows(3).Font.size = 10
        worksheet1.Range("A4:B4").Interior.Color = RGB(141, 180, 227)
        worksheet1.Rows(4).Font.size = 10
        '-----------------------------------------------------------------------------------------
        worksheet1.Cells(5, 1) = "SAP-Code"
        worksheet1.Rows(5).Font.size = 10
        ' worksheet1.Range("A5", "a5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        'worksheet1.Range("A5", "a5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        'worksheet1.Range("A5", "a5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Cells(5, 2) = " Description"
        worksheet1.Columns("B").ColumnWidth = 40
        worksheet1.Cells(5, 3) = "  Category"
        worksheet1.Columns("c").ColumnWidth = 8
        worksheet1.Cells(5, 4) = "   PkSize"
        worksheet1.Columns("d").ColumnWidth = 6
        worksheet1.Cells(5, 5) = "MR"
        worksheet1.Columns("E").ColumnWidth = 6

        i = 6
        Dim y_1 As Integer
        Dim strDis As String
        Dim MonthCount As Integer
        Dim YearCount As Integer

        MonthCount = Month(n_Date)
        YearCount = Year(n_Date)
        For y_1 = 1 To 12
            strDis = ""
            If MonthCount = 1 Then
                strDis = "Jan"
            ElseIf MonthCount = 2 Then
                strDis = "Feb"
            ElseIf MonthCount = 3 Then
                strDis = "Mar"
            ElseIf MonthCount = 4 Then
                strDis = "Apr"
            ElseIf MonthCount = 5 Then
                strDis = "May"
            ElseIf MonthCount = 6 Then
                strDis = "Jun"
            ElseIf MonthCount = 7 Then
                strDis = "Jul"
            ElseIf MonthCount = 8 Then
                strDis = "Aug"
            ElseIf MonthCount = 9 Then
                strDis = "Sep"
            ElseIf MonthCount = 10 Then
                strDis = "Oct"
            ElseIf MonthCount = 11 Then
                strDis = "Nov"
            ElseIf MonthCount = 12 Then
                strDis = "Dec"
            End If

            ' worksheet1.Cells(5, i) = MonthName(MonthCount, True) & "-" & Year(n_Date)
            strDis = strDis & " - " & YearCount.ToString
            worksheet1.Cells(5, i) = strDis
            worksheet1.Cells(5, i).HorizontalAlignment = XlHAlign.xlHAlignCenter
            If Microsoft.VisualBasic.Left(strDis, 3) = "Dec" Then
                YearCount = YearCount + 1
                MonthCount = 1
            Else
                MonthCount = MonthCount + 1
            End If

            i = i + 1

        Next
        worksheet1.Cells(5, i) = MonthName(Month(n_Date)) & "-" & YearCount
        worksheet1.Cells(5, i).HorizontalAlignment = XlHAlign.xlHAlignCenter
        i = i + 1
        worksheet1.Cells(5, i) = "SD"
        worksheet1.Columns("S").ColumnWidth = 6
        i = i + 1
        worksheet1.Cells(5, i) = " 12 MAvg"
        worksheet1.Columns("T").ColumnWidth = 8
        i = i + 1
        worksheet1.Cells(5, i) = "  SD%"
        worksheet1.Columns("U").ColumnWidth = 6
        i = i + 1
        worksheet1.Cells(5, i) = "   Hgt 6 MAvg"
        worksheet1.Columns("V").ColumnWidth = 8
        i = i + 1
        worksheet1.Cells(5, i) = "Hgt 3 MAvg"
        worksheet1.Columns("w").ColumnWidth = 8
        i = i + 1
        worksheet1.Cells(5, i) = "n3"
        worksheet1.Columns("x").ColumnWidth = 8
        i = i + 1
        worksheet1.Cells(5, i) = "CQ-MRS"
        worksheet1.Columns("y").ColumnWidth = 8
        i = i + 1
        worksheet1.Cells(5, i) = "Purchase"
        worksheet1.Columns("z").ColumnWidth = 15

        i = i + 1
        worksheet1.Cells(5, i) = "RQLT"
        worksheet1.Columns("aa").ColumnWidth = 6
        i = i + 1
        worksheet1.Cells(5, i) = "RLLT"
        worksheet1.Columns("ab").ColumnWidth = 6
        i = i + 1
        worksheet1.Cells(5, i) = "SS"
        worksheet1.Columns("ac").ColumnWidth = 10
        i = i + 1
        worksheet1.Cells(5, i) = "RL"
        worksheet1.Columns("ad").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "RQ"
        worksheet1.Columns("ae").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "ML"
        worksheet1.Columns("af").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "RQ"
        worksheet1.Columns("ag").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "wh stock"
        worksheet1.Columns("ah").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "End Stock"
        worksheet1.Columns("ai").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "TOD-CQ"
        worksheet1.Columns("aj").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "TOD n3"
        worksheet1.Columns("ak").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "TOt Req"
        worksheet1.Columns("al").ColumnWidth = 10

        i = i + 1
        worksheet1.Cells(5, i) = "Po qty"
        worksheet1.Columns("am").ColumnWidth = 10

        worksheet1.Range("A5:a5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Rows(5).Font.size = 10

        worksheet1.Range("b5:b5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("c5:c5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("d5:d5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("e5:e5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("f5:f5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("g5:g5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("h5:h5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("i5:i5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("j5:j5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("k5:k5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("l5:l5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("m5:m5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("n5:n5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("o5:o5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("p5:p5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("q5:q5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("r5:r5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("s5:s5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("t5:t5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("u5:u5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("v5:v5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("w5:w5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("x5:x5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("y5:y5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("z5:z5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Aa5:aa5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ab5:ab5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ac5:ac5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ad5:ad5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ae5:ae5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Af5:af5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ag5:ag5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ah5:ah5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ai5:ai5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Aj5:aj5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ak5:ak5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Al5:al5").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Am5:am5").Interior.Color = RGB(141, 180, 227)



        worksheet1.Range("A5", "a5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A5", "a5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A5", "a5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b5", "b5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b5", "b5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b5", "b5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("c5", "c5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c5", "c5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c5", "c5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d5", "d5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d5", "d5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d5", "d5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e5", "e5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e5", "e5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e5", "e5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("f5", "f5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f5", "f5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f5", "f5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("g5", "g5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g5", "g5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g5", "g5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("h5", "h5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h5", "h5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h5", "h5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("i5", "i5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i5", "i5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i5", "i5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("j5", "j5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j5", "j5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j5", "j5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("k5", "k5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k5", "k5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k5", "k5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("l5", "l5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l5", "l5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l5", "l5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("m5", "m5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m5", "m5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m5", "m5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("n5", "n5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n5", "n5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n5", "n5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("o5", "o5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o5", "o5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o5", "o5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("p5", "p5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p5", "p5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p5", "p5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("q5", "q5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("q5", "q5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("q5", "q5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("r5", "r5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r5", "r5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r5", "r5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("s5", "s5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("s5", "s5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("s5", "s5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("t5", "t5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t5", "t5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t5", "t5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("u5", "u5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u5", "u5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u5", "u5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("v5", "v5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v5", "v5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v5", "v5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("w5", "w5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w5", "w5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w5", "w5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("x5", "x5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("x5", "x5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("x5", "x5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("y5", "y5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y5", "y5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y5", "y5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("z5", "z5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("z5", "z5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("z5", "z5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous



        worksheet1.Range("Aa5", "aa5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Aa5", "aa5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Aa5", "aa5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ab5", "ab5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ab5", "ab5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ab5", "ab5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ac5", "ac5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ac5", "ac5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ac5", "ac5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ad5", "ad5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ad5", "ad5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ad5", "ad5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ae5", "ae5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ae5", "ae5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ae5", "ae5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Af5", "af5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Af5", "af5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Af5", "af5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ag5", "ag5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ag5", "ag5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ag5", "ag5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ah5", "ah5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ah5", "ah5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ah5", "ah5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ai5", "ai5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ai5", "ai5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ai5", "ai5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Aj5", "aj5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Aj5", "aj5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Aj5", "aj5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ak5", "ak5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ak5", "ak5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ak5", "ak5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Al5", "al5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Al5", "al5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Al5", "al5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Am5", "am5").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Am5", "am5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Am5", "am5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Rows(6).Font.size = 8
        worksheet1.Cells(6, 26) = "From"
        worksheet1.Cells(6, 29) = " Safty Stock"
        worksheet1.Cells(6, 30) = " Re-Ord.Level"
        worksheet1.Cells(6, 31) = " Re-Ord. Qty"
        worksheet1.Cells(6, 32) = " Max Stock"
        worksheet1.Cells(6, 34) = " (wh + wip)"

        worksheet1.Range("A6:a6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("b6:b6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("c6:c6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("d6:d6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("e6:e6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("f6:f6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("g6:g6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("h6:h6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("i6:i6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("j6:j6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("k6:k6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("l6:l6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("m6:m6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("n6:n6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("o6:o6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("p6:p6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("q6:q6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("r6:r6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("s6:s6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("t6:t6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("u6:u6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("v6:v6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("w6:w6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("x6:x6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("y6:y6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("z6:z6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Aa6:aa6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ab6:ab6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ac6:ac6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ad6:ad6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ae6:ae6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Af6:af6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ag6:ag6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ah6:ah6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ai6:ai6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Aj6:aj6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Ak6:ak6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Al6:al6").Interior.Color = RGB(141, 180, 227)
        worksheet1.Range("Am6:am6").Interior.Color = RGB(141, 180, 227)



        worksheet1.Range("A6", "a6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A6", "a6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("A6", "a6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b6", "b6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b6", "b6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b6", "b6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("c6", "c6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c6", "c6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c6", "c6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d6", "d6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d6", "d6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d6", "d6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e6", "e6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e6", "e6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e6", "e6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("f6", "f6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f6", "f6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("f6", "f6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g6", "g6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h6", "h6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i6", "i6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("j6", "j6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("k6", "k6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("l6", "l6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("m6", "m6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("n6", "n6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("o6", "o6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("p6", "p6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("q6", "q6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("r6", "r6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("s6", "s6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("t6", "t6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("u6", "u6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u6", "u6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("u6", "u6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("v6", "v6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v6", "v6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("v6", "v6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("w6", "w6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w6", "w6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("w6", "w6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("x6", "x6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("x6", "x6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("x6", "x6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("y6", "y6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y6", "y6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("y6", "y6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("z6", "z6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("z6", "z6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("z6", "z6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous



        worksheet1.Range("Aa6", "aa6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Aa6", "aa6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Aa6", "aa6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ab6", "ab6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ab6", "ab6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ab6", "ab6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ac6", "ac6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ac6", "ac6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ac6", "ac6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ad6", "ad6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ad6", "ad6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ad6", "ad6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ae6", "ae6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ae6", "ae6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ae6", "ae6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Af6", "af6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Af6", "af6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Af6", "af6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ag6", "ag6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ag6", "ag6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ag6", "ag6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ah6", "ah6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ah6", "ah6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ah6", "ah6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ai6", "ai6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ai6", "ai6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ai6", "ai6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Aj6", "aj6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Aj6", "aj6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Aj6", "aj6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Ak6", "ak6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ak6", "ak6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Ak6", "ak6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Al6", "al6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Al6", "al6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Al6", "al6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("Am6", "am6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Am6", "am6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("Am6", "am6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        n_Date = CDate(txtFromDate.Text).AddDays(-365)
        ' n_Date = n_Date & " " & "7:30AM"
        n_Date = Month(n_Date) & "/1/" & Year(n_Date)
        N_Date1 = Month(txtTodate.Text) & "/1/" & Year(txtTodate.Text)
        ' N_Date1 = CDate(txtTodate.Text).AddDays(+1)
        'N_Date1 = N_Date1 & " " & "7:30AM"
        Dim X As Integer
        Dim Y1 As Integer
        Dim nColum As Integer
        X = 7
        nColum = 0
        SQL = "select m11SAPCode from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "' group by m11SAPCode"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        i = 0
        For Each DTRow4 As DataRow In T01.Tables(0).Rows
            Y1 = 0
            nColum = 6
            worksheet1.Rows(X).Font.size = 8
            SQL = "select * from  M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "' and m11SAPCode='" & T01.Tables(0).Rows(i)("m11SAPCode") & "' order by M11year,m11month"
            T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            For Each DTRow3 As DataRow In T03.Tables(0).Rows
                If Y1 = 0 Then
                    worksheet1.Cells(X, 1) = T03.Tables(0).Rows(Y1)("M11SAPCode")
                    worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(X, 2) = T03.Tables(0).Rows(Y1)("M11Dis")
                    worksheet1.Cells(X, 3) = T03.Tables(0).Rows(Y1)("M11Category")
                    'worksheet1.Cells(X, 4) = T03.Tables(0).Rows(Y1)("M11PRC")
                    'worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
                    'range1 = worksheet1.Cells(X, 4)
                    'range1.NumberFormat = "0"
                    worksheet1.Cells(X, 4) = T03.Tables(0).Rows(Y1)("M11PckSize")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X, 4)
                    range1.NumberFormat = "0"
                    worksheet1.Cells(X, 5) = T03.Tables(0).Rows(Y1)("M11MR")
                    worksheet1.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 19) = T03.Tables(0).Rows(Y1)("M11SD")
                    worksheet1.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 19)
                    range1.NumberFormat = "0"
                    worksheet1.Cells(X, 20) = T03.Tables(0).Rows(Y1)("M11MAvg")
                    worksheet1.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 20)
                    range1.NumberFormat = "0"
                    worksheet1.Cells(X, 21) = T03.Tables(0).Rows(Y1)("M11NSD")
                    worksheet1.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 21)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(X, 22) = T03.Tables(0).Rows(Y1)("M11HGT6")
                    worksheet1.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 22)
                    range1.NumberFormat = "0"


                    worksheet1.Cells(X, 23) = T03.Tables(0).Rows(Y1)("M11HGT3")
                    worksheet1.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 23)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(X, 24) = T03.Tables(0).Rows(Y1)("M11N3")
                    worksheet1.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 24)
                    range1.NumberFormat = "0"
                    worksheet1.Cells(X, 25) = T03.Tables(0).Rows(Y1)("M11CQ")
                    worksheet1.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 25)
                    range1.NumberFormat = "0"
                    worksheet1.Cells(X, 26) = T03.Tables(0).Rows(Y1)("M11Purchase")
                    worksheet1.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(X, 27) = T03.Tables(0).Rows(Y1)("M11RQLT")
                    worksheet1.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 28) = T03.Tables(0).Rows(Y1)("M11RLLT")
                    worksheet1.Cells(X, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 29) = T03.Tables(0).Rows(Y1)("M11SS")
                    worksheet1.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 30) = T03.Tables(0).Rows(Y1)("M11RL")
                    worksheet1.Cells(X, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 31) = T03.Tables(0).Rows(Y1)("M11RQ")
                    worksheet1.Cells(X, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 32) = T03.Tables(0).Rows(Y1)("M11ML")
                    worksheet1.Cells(X, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    If T03.Tables(0).Rows(Y1)("M11WHStock") = 0 Then
                        worksheet1.Cells(X, 33) = "-"
                        worksheet1.Cells(X, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else

                        worksheet1.Cells(X, 33) = T03.Tables(0).Rows(Y1)("M11WHStock")
                        worksheet1.Cells(X, 33).HorizontalAlignment = XlHAlign.xlHAlignRight
                    End If
                    If T03.Tables(0).Rows(Y1)("M11EndStock") = 0 Then
                        worksheet1.Cells(X, 34) = "-"
                        worksheet1.Cells(X, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                    Else
                        worksheet1.Cells(X, 34) = T03.Tables(0).Rows(Y1)("M11EndStock")
                        worksheet1.Cells(X, 34).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet1.Cells(X, 34)
                        range1.NumberFormat = "0"
                    End If

                    worksheet1.Cells(X, 35) = T03.Tables(0).Rows(Y1)("M11TOD_CQ")
                    worksheet1.Cells(X, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, 36) = T03.Tables(0).Rows(Y1)("M11TOD_N3")
                    worksheet1.Cells(X, 36).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    range1 = worksheet1.Cells(X, 36)
                    range1.NumberFormat = "0"
                    worksheet1.Cells(X, 37) = T03.Tables(0).Rows(Y1)("M11Tot_Req")
                    worksheet1.Cells(X, 37).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 37)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(X, 38) = T03.Tables(0).Rows(Y1)("M11PO_QTY")
                    worksheet1.Cells(X, 38).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, 38)
                    range1.NumberFormat = "0"
                End If
              
                worksheet1.Cells(X, nColum) = T03.Tables(0).Rows(Y1)("M11Value")
                worksheet1.Cells(X, nColum).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, nColum)
                range1.NumberFormat = "0"
                nColum = nColum + 1
                Y1 = Y1 + 1
            Next

            '  worksheet1.Range("A6", "a6").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("A" & X, "a" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("A" & X, "a" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("b" & X, "b" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & X, "b" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("c" & X, "c" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & X, "c" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("d" & X, "d" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & X, "d" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("e" & X, "e" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & X, "e" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("f" & X, "f" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & X, "f" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("g" & X, "g" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & X, "g" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("h" & X, "h" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & X, "h" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("i" & X, "i" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & X, "i" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("j" & X, "j" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & X, "j" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("k" & X, "k" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & X, "k" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("l" & X, "l" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & X, "l" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("m" & X, "m" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & X, "m" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("n" & X, "n" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & X, "n" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("o" & X, "o" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & X, "o" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("p" & X, "p" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & X, "p" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("q" & X, "q" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("q" & X, "q" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("r" & X, "r" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("r" & X, "r" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("s" & X, "s" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("s" & X, "s" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("t" & X, "t" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("t" & X, "t" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("u" & X, "u" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("u" & X, "u" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("v" & X, "v" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("v" & X, "v" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("w" & X, "w" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("w" & X, "w" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("x" & X, "x" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("x" & X, "x" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("y" & X, "y" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("y" & X, "y" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("z" & X, "z" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("z" & X, "z" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash



            worksheet1.Range("Aa" & X, "aa" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Aa" & X, "aa" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Ab" & X, "ab" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Ab" & X, "ab" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Ac" & X, "ac" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Ac" & X, "ac" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Ad" & X, "ad" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Ad" & X, "ad" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Ae" & X, "ae" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Ae" & X, "ae" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Af" & X, "af" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Af" & X, "af" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Ag" & X, "ag" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Ag" & X, "ag" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Ah" & X, "ah" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Ah" & X, "ah" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Ai" & X, "ai" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Ai" & X, "ai" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Aj" & X, "aj" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Aj" & X, "aj" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Ak" & X, "ak" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Ak" & X, "ak" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Al" & X, "al" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Al" & X, "al" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            worksheet1.Range("Am" & X, "am" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("Am" & X, "am" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            X = X + 1
            i = i + 1
        Next
        '----------------------------------------------------------------------------------------------------------------
        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("F" & (X)).Formula = "=SUM(F7:F" & X - 1 & ")"
        worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 6)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("G" & (X)).Formula = "=SUM(G7:G" & X - 1 & ")"
        worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 7)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("h" & (X)).Formula = "=SUM(h7:h" & X - 1 & ")"
        worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 8)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("i" & (X)).Formula = "=SUM(i7:i" & X - 1 & ")"
        worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 9)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("j" & (X)).Formula = "=SUM(j7:j" & X - 1 & ")"
        worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 10)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("k" & (X)).Formula = "=SUM(k7:k" & X - 1 & ")"
        worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 11)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("l" & (X)).Formula = "=SUM(l7:l" & X - 1 & ")"
        worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 12)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("m" & (X)).Formula = "=SUM(m7:m" & X - 1 & ")"
        worksheet1.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 13)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("n" & (X)).Formula = "=SUM(n7:n" & X - 1 & ")"
        worksheet1.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 14)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("o" & (X)).Formula = "=SUM(o7:o" & X - 1 & ")"
        worksheet1.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 15)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("p" & (X)).Formula = "=SUM(p7:p" & X - 1 & ")"
        worksheet1.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 16)
        range1.NumberFormat = "0"

        worksheet1.Rows(X).Font.size = 8
        worksheet1.Range("Q" & (X)).Formula = "=SUM(Q7:Q" & X - 1 & ")"
        worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(X, 17)
        range1.NumberFormat = "0"


    End Function

    Function TOD_Report()
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

        '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
        ' workbooks.Application.Sheets.Add()
        Dim sheets1 As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        worksheet1.Name = "TOD_" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        n_Date = CDate(txtFromDate.Text).AddDays(-365)
        ' n_Date = n_Date & " " & "7:30AM"
        N_Date1 = CDate(txtTodate.Text).AddDays(+1)
        ' N_Date1 = N_Date1 & " " & "7:30AM"
        worksheet1.Rows(2).Font.size = 11
        worksheet1.Rows(2).Font.Bold = True
        worksheet1.Columns("A").ColumnWidth = 40
        worksheet1.Cells(2, 1) = "TOD(Overall)" & Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        worksheet1.Cells(2, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("B").ColumnWidth = 10
        worksheet1.Cells(2, 2) = "D"
        worksheet1.Cells(2, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("C").ColumnWidth = 10
        worksheet1.Cells(2, 3) = "C"
        worksheet1.Cells(2, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("D").ColumnWidth = 10
        worksheet1.Cells(2, 4) = "A"
        worksheet1.Cells(2, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Range("a2,A2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b2,b2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c2,c2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d2,d2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a2,a2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b2,b2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c2,c2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d2,d2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a2,a2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b2,b2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c2,c2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d2,d2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b2:b2").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("c2:c2").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("d2:d2").Interior.Color = RGB(197, 190, 151)

        'A
        SQL = "select sum(M11endstock) as M11endstock ,sum(M11N3) as M11N3 from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then
            worksheet1.Rows(3).Font.size = 11
            worksheet1.Cells(3, 1) = "End Stock"
            worksheet1.Cells(3, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(3, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(3, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(3, 2)
            range1.NumberFormat = "0"

            worksheet1.Rows(4).Font.size = 11
            worksheet1.Cells(4, 1) = "N3"
            worksheet1.Cells(4, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(4, 2) = T01.Tables(0).Rows(0)("M11N3")
            worksheet1.Cells(4, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(4, 2)
            range1.NumberFormat = "0"
        End If
        Dim firstDayLastMonth As DateTime
        Dim lastDayLastMonth As DateTime
        Dim thisMonth As Date
        Dim _MonthNo As Integer

        thisMonth = txtFromDate.Text
        firstDayLastMonth = thisMonth.AddMonths(-1)
        firstDayLastMonth = Month(firstDayLastMonth) & "/1/" & Year(firstDayLastMonth)
        _MonthNo = Month(firstDayLastMonth)

        SQL = "select sum(M11Value) as M11Value from M11MRS where M11Date ='" & firstDayLastMonth & "'  and left(M11Category,1)='D' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then
            worksheet1.Rows(5).Font.size = 11
            worksheet1.Cells(5, 1) = "LMQ(" & MonthName(_MonthNo) & ")"
            worksheet1.Cells(5, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(5, 2) = T01.Tables(0).Rows(0)("M11Value")
            worksheet1.Cells(5, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(5, 2)
            range1.NumberFormat = "0"
        End If

        worksheet1.Rows(6).Font.size = 11
        worksheet1.Rows(6).Font.bold = True
        worksheet1.Cells(6, 2) = "D"
        worksheet1.Cells(6, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 3) = "C"
        worksheet1.Cells(6, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 4) = "A"
        worksheet1.Cells(6, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Range("b6:b6").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("c6:c6").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("d6:d6").Interior.Color = RGB(197, 190, 151)
        worksheet1.Rows(7).Font.size = 11
        worksheet1.Cells(7, 1) = "TOT - N3"
        worksheet1.Cells(7, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("b7").Formula = "=(B3/B4)*30"
        worksheet1.Cells(7, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(7, 2)
        range1.NumberFormat = "0"
        '---------------------------------------------------------------------
        worksheet1.Rows(8).Font.size = 11
        worksheet1.Cells(8, 1) = "TOT - LMQ (" & MonthName(_MonthNo) & " )"
        worksheet1.Cells(8, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Range("b8").Formula = "=(B3/B5)*30"
        worksheet1.Cells(8, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(8, 2)
        range1.NumberFormat = "0"
        '---------------------------------------------------------------------
        'C
        '
        '
        '
        SQL = "select sum(M11endstock) as M11endstock ,sum(M11N3) as M11N3 from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then
           
            worksheet1.Cells(3, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(3, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(3, 3)
            range1.NumberFormat = "0"

            worksheet1.Cells(4, 3) = T01.Tables(0).Rows(0)("M11N3")
            worksheet1.Cells(4, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(4, 3)
            range1.NumberFormat = "0"
        End If
       

        thisMonth = txtFromDate.Text
        firstDayLastMonth = thisMonth.AddMonths(-1)
        firstDayLastMonth = Month(firstDayLastMonth) & "/1/" & Year(firstDayLastMonth)
        _MonthNo = Month(firstDayLastMonth)

        SQL = "select sum(M11Value) as M11Value from M11MRS where M11Date ='" & firstDayLastMonth & "'  and left(M11Category,1)='C' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then
       
            worksheet1.Cells(5, 3) = T01.Tables(0).Rows(0)("M11Value")
            worksheet1.Cells(5, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(5, 3)
            range1.NumberFormat = "0"
        End If

        worksheet1.Range("C7").Formula = "=(c3/c4)*30"
        worksheet1.Cells(7, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(7, 3)
        range1.NumberFormat = "0"
        '---------------------------------------------------------------------

        worksheet1.Range("C8").Formula = "=(c3/c5)*30"
        worksheet1.Cells(8, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(8, 3)
        range1.NumberFormat = "0"
        '---------------------------------------------------------------------
        'A
        SQL = "select sum(M11endstock) as M11endstock ,sum(M11N3) as M11N3 from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then

            worksheet1.Cells(3, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(3, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(3, 4)
            range1.NumberFormat = "0"

            worksheet1.Cells(4, 4) = T01.Tables(0).Rows(0)("M11N3")
            worksheet1.Cells(4, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(4, 4)
            range1.NumberFormat = "0"
        End If


        thisMonth = txtFromDate.Text
        firstDayLastMonth = thisMonth.AddMonths(-1)
        firstDayLastMonth = Month(firstDayLastMonth) & "/1/" & Year(firstDayLastMonth)
        _MonthNo = Month(firstDayLastMonth)

        SQL = "select sum(M11Value) as M11Value from M11MRS where M11Date ='" & firstDayLastMonth & "'  and left(M11Category,1)='A' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then

            worksheet1.Cells(5, 4) = T01.Tables(0).Rows(0)("M11Value")
            worksheet1.Cells(5, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(5, 4)
            range1.NumberFormat = "0"
        End If

        worksheet1.Range("D7").Formula = "=(D3/D4)*30"
        worksheet1.Cells(7, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(7, 4)
        range1.NumberFormat = "0"
        '---------------------------------------------------------------------

        worksheet1.Range("D8").Formula = "=(D3/D5)*30"
        worksheet1.Cells(8, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        range1 = worksheet1.Cells(8, 4)
        range1.NumberFormat = "0"

        worksheet1.Range("d8:d8").Interior.Color = RGB(255, 255, 0)
        worksheet1.Range("c8:c8").Interior.Color = RGB(255, 255, 0)
        worksheet1.Range("b8:b8").Interior.Color = RGB(255, 255, 0)
        worksheet1.Range("a8:a8").Interior.Color = RGB(255, 255, 0)
        worksheet1.Range("a7:a7").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("b7:b7").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("c7:c7").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("d7:d7").Interior.Color = RGB(252, 213, 180)

        worksheet1.Range("a3,A3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b3,b3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c3,c3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d3,d3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a3,a3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b3,b3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c3,c3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d3,d3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a4,A4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b4,b4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c4,c4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d4,d4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a4,a4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b4,b4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c4,c4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d4,d4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a5,A5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b5,b5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c5,c5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d5,d5").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a5,a5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b5,b5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c5,c5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d5,d5").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a6,A6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b6,b6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c6,c6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d6,d6").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a6,a6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b6,b6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c6,c6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d6,d6").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a7,A7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b7,b7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c7,c7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d7,d7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a7,a7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b7,b7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c7,c7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d7,d7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a8,A8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b8,b8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c8,c8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d8,d8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a8,a8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b8,b8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c8,c8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d8,d8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


        '====================================================================================================
        worksheet1.Rows(10).Font.size = 11
        worksheet1.Rows(10).Font.Bold = True
        worksheet1.Columns("A").ColumnWidth = 40
        worksheet1.Cells(10, 1) = "Detailed TOD Data - " & Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        worksheet1.Cells(10, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("B").ColumnWidth = 10
        worksheet1.Cells(10, 2) = "D"
        worksheet1.Cells(10, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("C").ColumnWidth = 10
        worksheet1.Cells(10, 3) = "C"
        worksheet1.Cells(10, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("D").ColumnWidth = 10
        worksheet1.Cells(10, 4) = "A"
        worksheet1.Cells(10, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Range("a10,A10").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b10,b10").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c10,c10").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d10,d10").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a10,a10").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b10,b10").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c10,c10").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d10,d10").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a10,a10").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b10,b10").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c10,c10").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d10,d10").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b10:b10").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("c10:c10").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("d10:d10").Interior.Color = RGB(197, 190, 151)

        worksheet1.Rows(11).Font.size = 11
        worksheet1.Cells(11, 1) = "Endstock-PTL"

        'A
        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase='PTL' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then

 
            worksheet1.Cells(11, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(11, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(11, 2)
            range1.NumberFormat = "0"
        End If
        '----------------------------------------------------------------------------------------
        'C
        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase='PTL' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(11, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(11, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(11, 3)
            range1.NumberFormat = "0"
        End If
        '---------------------------------------------------------------------------------------
        'A
        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase='PTL' group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(11, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(11, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(11, 4)
            range1.NumberFormat = "0"
        End If
        '--------------------------------------------------------------------------
        worksheet1.Rows(12).Font.size = 11
        worksheet1.Cells(12, 1) = "Endstock-Local(Import)"
        'D
        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(12, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(12, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(12, 2)
            range1.NumberFormat = "0"
        End If
        'C

        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then
            worksheet1.Cells(12, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(12, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(12, 3)
            range1.NumberFormat = "0"
        End If
        'A
        'D
        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(12, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(12, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(12, 4)
            range1.NumberFormat = "0"
        End If

        worksheet1.Rows(13).Font.size = 11
        worksheet1.Cells(13, 1) = "Endstock-Local(ex-stock)"
        'D
        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(13, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(13, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(13, 2)
            range1.NumberFormat = "0"
        End If

        'C
        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(13, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(13, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(13, 3)
            range1.NumberFormat = "0"
        End If

        'A
        SQL = "select sum(M11endstock) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(13, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(13, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(13, 4)
            range1.NumberFormat = "0"
        End If
        '-----------------------------------------------------------------------------------
        worksheet1.Range("a11,a11").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b11,b11").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c11,c11").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d11,d11").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a11,a11").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b11,b11").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c11,c11").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d11,d11").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a12,a12").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b12,b12").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c12,c12").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d12,d12").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a12,a12").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b12,b12").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c12,c12").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d12,d12").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a13,a13").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b13,b13").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c13,c13").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d13,d13").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a13,a13").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b13,b13").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c13,c13").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d13,d13").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b11:b11").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("c11:c11").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("d11:d11").Interior.Color = RGB(252, 213, 180)

        worksheet1.Range("b12:b12").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("c12:c12").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("d12:d12").Interior.Color = RGB(182, 221, 232)

        worksheet1.Range("b13:b13").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("c13:c13").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("d13:d13").Interior.Color = RGB(215, 228, 188)


        worksheet1.Rows(15).Font.size = 11
        worksheet1.Rows(15).Font.Bold = True
        worksheet1.Columns("A").ColumnWidth = 40
        worksheet1.Cells(15, 1) = "% Stock Distribution" '& Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        worksheet1.Cells(15, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("B").ColumnWidth = 10
        worksheet1.Cells(15, 2) = "D"
        worksheet1.Cells(15, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("C").ColumnWidth = 10
        worksheet1.Cells(15, 3) = "C"
        worksheet1.Cells(15, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("D").ColumnWidth = 10
        worksheet1.Cells(15, 4) = "A"
        worksheet1.Cells(15, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Range("a15,A15").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b15,b15").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c15,c15").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d15,d15").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a15,a15").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b15,b15").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c15,c15").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d15,d15").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a15,a15").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b15,b15").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c15,c15").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d15,d15").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b15:b15").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("c15:c15").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("d15:d15").Interior.Color = RGB(197, 190, 151)
        '-------------------------------------------------------------------------------------------------------
        worksheet1.Rows(16).Font.size = 11
        worksheet1.Cells(16, 1) = "Endstock-PTL"
        worksheet1.Range("B16").Formula = "=(B11/B3)"
        worksheet1.Cells(16, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(16, 2)
        range1.NumberFormat = "0%"

        worksheet1.Range("c16").Formula = "=(c11/c3)"
        worksheet1.Cells(16, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(16, 3)
        range1.NumberFormat = "0%"

        worksheet1.Range("D16").Formula = "=(D11/D3)"
        worksheet1.Cells(16, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(16, 4)
        range1.NumberFormat = "0%"

        worksheet1.Rows(17).Font.size = 11
        worksheet1.Cells(17, 1) = "Endstock-Local(Import)"
        worksheet1.Range("B17").Formula = "=(B12/B3)"
        worksheet1.Cells(17, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(17, 2)
        range1.NumberFormat = "0%"

        worksheet1.Range("c17").Formula = "=(c12/c3)"
        worksheet1.Cells(17, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(17, 3)
        range1.NumberFormat = "0%"

        worksheet1.Range("D17").Formula = "=(D12/D3)"
        worksheet1.Cells(17, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(17, 4)
        range1.NumberFormat = "0%"
        '---------------------------------------------------------------------------------
        worksheet1.Rows(18).Font.size = 11
        worksheet1.Cells(18, 1) = "Endstock-Local(ex-stock)"
        worksheet1.Range("B18").Formula = "=(B13/B3)"
        worksheet1.Cells(18, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(18, 2)
        range1.NumberFormat = "0%"

        worksheet1.Range("c18").Formula = "=(c13/c3)"
        worksheet1.Cells(18, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(18, 3)
        range1.NumberFormat = "0%"

        worksheet1.Range("D18").Formula = "=(D13/D3)"
        worksheet1.Cells(18, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(18, 4)
        range1.NumberFormat = "0%"
        '-------------------------------------------------------------------------
        worksheet1.Range("a16,a16").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b16,b16").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c16,c16").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d16,d16").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a16,a16").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b16,b16").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c16,c16").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d16,d16").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a17,a17").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b17,b17").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c17,c17").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d17,d17").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a17,a17").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b17,b17").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c17,c17").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d17,d17").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a18,a18").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b18,b18").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c18,c18").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d18,d18").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a18,a18").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b18,b18").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c18,c18").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d18,d18").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b16:b16").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("c16:c16").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("d16:d16").Interior.Color = RGB(252, 213, 180)

        worksheet1.Range("b17:b17").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("c17:c17").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("d17:d17").Interior.Color = RGB(182, 221, 232)

        worksheet1.Range("b18:b18").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("c18:c18").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("d18:d18").Interior.Color = RGB(215, 228, 188)
        '--------------------------------------------------------------------------
        worksheet1.Rows(20).Font.size = 11
        worksheet1.Cells(20, 1) = "N3 - PTL"
        'D
        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase in ('PTL') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(20, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(20, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(20, 2)
            range1.NumberFormat = "0"
        End If

        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase in ('PTL') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(20, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(20, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(20, 3)
            range1.NumberFormat = "0"
        End If

        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase in ('PTL') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(20, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(20, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(20, 4)
            range1.NumberFormat = "0"
        End If
        '---------------------------------------------------------------------------------------------------
        worksheet1.Rows(21).Font.size = 11
        worksheet1.Cells(21, 1) = "N3-Local(Import)"
        'D
        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(21, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(21, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(21, 2)
            range1.NumberFormat = "0"
        End If

        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(21, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(21, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(21, 3)
            range1.NumberFormat = "0"
        End If

        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(21, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(21, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(21, 4)
            range1.NumberFormat = "0"
        End If
        '------------------------------------------------------------------------------------------------
        worksheet1.Rows(22).Font.size = 11
        worksheet1.Cells(22, 1) = "N3-(ex-stock)"
        'D
        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(22, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(22, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(22, 2)
            range1.NumberFormat = "0"
        End If

        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(22, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(22, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(22, 3)
            range1.NumberFormat = "0"
        End If

        SQL = "select sum(M11N3) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(22, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(22, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(22, 4)
            range1.NumberFormat = "0"
        End If
        '====================================================================================
        worksheet1.Range("a19,a19").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b19,b19").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c19,c19").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d19,d19").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a19,a19").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b19,b19").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c19,c19").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d19,d19").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a20,a20").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b20,b20").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c20,c20").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d20,d20").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a20,a20").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b20,b20").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c20,c20").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d20,d20").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a21,a21").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b21,b21").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c21,c21").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d21,d21").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a21,a21").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b21,b21").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c21,c21").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d21,d21").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a22,a22").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b22,b22").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c22,c22").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d22,d22").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a22,a22").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b22,b22").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c22,c22").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d22,d22").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b20:b20").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("c20:c20").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("d20:d20").Interior.Color = RGB(252, 213, 180)

        worksheet1.Range("b21:b21").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("c21:c21").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("d21:d21").Interior.Color = RGB(182, 221, 232)

        worksheet1.Range("b22:b22").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("c22:c22").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("d22:d22").Interior.Color = RGB(215, 228, 188)
        '-------------------------------------------------------------------
        worksheet1.Rows(24).Font.size = 11
        worksheet1.Cells(24, 1) = "LMQ (" & MonthName(_MonthNo) & " ) -PTL"
        'D
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase in ('PTL') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(24, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(24, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(24, 2)
            range1.NumberFormat = "0"
        End If
        '---------------------------------------------------------------------
        'C
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase in ('PTL') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(24, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(24, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(24, 3)
            range1.NumberFormat = "0"
        End If
        '----------------------------------------------------------------------
        'A
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase in ('PTL') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(24, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(24, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(24, 4)
            range1.NumberFormat = "0"
        End If
        '----------------------------------------------------------------------
        worksheet1.Rows(25).Font.size = 11
        worksheet1.Cells(25, 1) = "LMQ (" & MonthName(_MonthNo) & " ) -Local(Import)"
        'D
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(25, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(25, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(25, 2)
            range1.NumberFormat = "0"
        End If
        '---------------------------------------------------------------------
        'C
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(25, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(25, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(25, 3)
            range1.NumberFormat = "0"
        End If
        '----------------------------------------------------------------------
        'A
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase in ('Local-Imports') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(25, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(25, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(25, 4)
            range1.NumberFormat = "0"
        End If
        '----------------------------------------------------------------------
        worksheet1.Rows(26).Font.size = 11
        worksheet1.Cells(26, 1) = "LMQ (" & MonthName(_MonthNo) & " ) - (ex-stock)"
        'D
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='D' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(26, 2) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(26, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(26, 2)
            range1.NumberFormat = "0"
        End If
        '---------------------------------------------------------------------
        'C
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='C' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(26, 3) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(26, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(26, 3)
            range1.NumberFormat = "0"
        End If
        '----------------------------------------------------------------------
        'A
        SQL = "select sum(M11Value) as M11endstock  from M11MRS where M11Date between '" & n_Date & "' and '" & N_Date1 & "'  and left(M11Category,1)='A' and M11Purchase in ('Local','Local-Huntsman') group by left(M11Category,1)"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


            worksheet1.Cells(26, 4) = T01.Tables(0).Rows(0)("M11endstock")
            worksheet1.Cells(26, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(26, 4)
            range1.NumberFormat = "0"
        End If
        '----------------------------------------------------------------------
        worksheet1.Range("a23,a23").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b23,b23").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c23,c23").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d23,d23").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a23,a23").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b23,b23").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c23,c23").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d23,d23").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a24,a24").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b24,b24").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c24,c24").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d24,d24").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a24,a24").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b24,b24").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c24,c24").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d24,d24").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a25,a25").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b25,b25").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c25,c25").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d25,d25").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a25,a25").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b25,b25").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c25,c25").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d25,d25").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a26,a26").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b26,b26").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c26,c26").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d26,d26").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a26,a26").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b26,b26").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c26,c26").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d26,d26").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b24:b24").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("c24:c24").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("d24:d24").Interior.Color = RGB(252, 213, 180)

        worksheet1.Range("b25:b25").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("c25:c25").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("d25:d25").Interior.Color = RGB(182, 221, 232)

        worksheet1.Range("b26:b26").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("c26:c26").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("d26:d26").Interior.Color = RGB(215, 228, 188)
        '================================================================================
        worksheet1.Rows(28).Font.size = 11
        worksheet1.Rows(28).Font.Bold = True
        worksheet1.Columns("A").ColumnWidth = 40
        'worksheet1.Cells(10, 1) = "Detailed TOD Data - " & Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        'worksheet1.Cells(10, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("B").ColumnWidth = 10
        worksheet1.Cells(28, 2) = "D"
        worksheet1.Cells(28, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("C").ColumnWidth = 10
        worksheet1.Cells(28, 3) = "C"
        worksheet1.Cells(28, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("D").ColumnWidth = 10
        worksheet1.Cells(28, 4) = "A"
        worksheet1.Cells(28, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Range("a28,A28").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b28,b28").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c28,c28").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d28,d28").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a28,a28").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b28,b28").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c28,c28").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d28,d28").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a28,a28").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b28,b28").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c28,c28").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d28,d28").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b28:b28").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("c28:c28").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("d28:d28").Interior.Color = RGB(197, 190, 151)

       
        worksheet1.Cells(29, 1) = "TOD(N3)-PTL " '& Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        worksheet1.Range("B29").Formula = "=(B11/B20)*30"
        worksheet1.Cells(29, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(29, 2)
        range1.NumberFormat = "0"

        worksheet1.Range("C29").Formula = "=(c11/c20)*30"
        worksheet1.Cells(29, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(29, 3)
        range1.NumberFormat = "0"

        worksheet1.Range("D29").Formula = "=(d11/d20)*30"
        worksheet1.Cells(29, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(29, 4)
        range1.NumberFormat = "0"

        worksheet1.Cells(30, 1) = "TOD(n3)-Local(Import)" '& Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        worksheet1.Range("B30").Formula = "=(B12/B21)*30"
        worksheet1.Cells(30, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(30, 2)
        range1.NumberFormat = "0"

        worksheet1.Range("C30").Formula = "=(c12/c21)*30"
        worksheet1.Cells(30, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(30, 3)
        range1.NumberFormat = "0"

        worksheet1.Range("D30").Formula = "=(d12/d21)*30"
        worksheet1.Cells(30, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(30, 4)
        range1.NumberFormat = "0"

        worksheet1.Cells(31, 1) = "TOD(n3)-Local(ex-stock)" '& Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        worksheet1.Range("B31").Formula = "=(B13/B22)*30"
        worksheet1.Cells(31, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(31, 2)
        range1.NumberFormat = "0"

        worksheet1.Range("C31").Formula = "=(c13/c22)*30"
        worksheet1.Cells(31, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(31, 3)
        range1.NumberFormat = "0"

        worksheet1.Range("D31").Formula = "=(d13/d22)*30"
        worksheet1.Cells(31, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(31, 4)
        range1.NumberFormat = "0"

        '---------------------------------------------------------------------
     

        worksheet1.Range("a28,a28").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b28,b28").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c28,c28").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d28,d28").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a28,a28").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b28,b28").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c28,c28").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d28,d28").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a29,a29").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b29,b29").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c29,c29").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d29,d29").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a29,a29").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b29,b29").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c29,c29").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d29,d29").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a30,a30").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b30,b30").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c30,c30").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d30,d30").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a30,a30").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b30,b30").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c30,c30").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d30,d30").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a31,a31").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b31,b31").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c31,c31").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d31,d31").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a31,a31").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b31,b31").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c31,c31").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d31,d31").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b29:b29").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("c29:c29").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("d29:d29").Interior.Color = RGB(252, 213, 180)

        worksheet1.Range("b30:b30").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("c30:c30").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("d30:d30").Interior.Color = RGB(182, 221, 232)

        worksheet1.Range("b31:b31").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("c31:c31").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("d31:d31").Interior.Color = RGB(215, 228, 188)
        '----------------------------------------------------------------------
        worksheet1.Rows(32).Font.size = 11
        worksheet1.Rows(32).Font.Bold = True
        worksheet1.Columns("A").ColumnWidth = 40
        'worksheet1.Cells(10, 1) = "Detailed TOD Data - " & Microsoft.VisualBasic.Day(txtFromDate.Text) & "." & Month(txtFromDate.Text) & "." & Year(txtFromDate.Text)
        'worksheet1.Cells(10, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("B").ColumnWidth = 10
        worksheet1.Cells(32, 2) = "D"
        worksheet1.Cells(32, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("C").ColumnWidth = 10
        worksheet1.Cells(32, 3) = "C"
        worksheet1.Cells(32, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Columns("D").ColumnWidth = 10
        worksheet1.Cells(32, 4) = "A"
        worksheet1.Cells(32, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Range("a32,A32").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b32,b32").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c32,c32").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d32,d32").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a32,a32").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b32,b32").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c32,c32").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d32,d32").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a32,a32").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b32,b32").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c32,c32").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d32,d32").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("b32:b32").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("c32:c32").Interior.Color = RGB(197, 190, 151)
        worksheet1.Range("d32:d32").Interior.Color = RGB(197, 190, 151)
        '-----------------------------------------------------------------------------------------------------
        worksheet1.Rows(33).Font.size = 11
        worksheet1.Cells(33, 1) = "TOD (" & MonthName(_MonthNo) & " ) - PTL"
        worksheet1.Range("B33").Formula = "=(B11/B24)*30"
        worksheet1.Cells(33, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(33, 2)
        range1.NumberFormat = "0"

        worksheet1.Range("C33").Formula = "=(c11/c24)*30"
        worksheet1.Cells(33, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(33, 3)
        range1.NumberFormat = "0"

        worksheet1.Range("D33").Formula = "=(d11/d24)*30"
        worksheet1.Cells(33, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(33, 4)
        range1.NumberFormat = "0"
        '-------------------------------------------------------------------------------------------------------
        worksheet1.Rows(34).Font.size = 11
        worksheet1.Cells(34, 1) = "TOD (" & MonthName(_MonthNo) & " ) - Local(Import)"
        worksheet1.Range("B34").Formula = "=(B12/B25)*30"
        worksheet1.Cells(34, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(34, 2)
        range1.NumberFormat = "0"

        worksheet1.Range("C34").Formula = "=(c12/c25)*30"
        worksheet1.Cells(34, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(34, 3)
        range1.NumberFormat = "0"

        worksheet1.Range("D34").Formula = "=(d12/d25)*30"
        worksheet1.Cells(34, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(34, 4)
        range1.NumberFormat = "0"
        '------------------------------------------------------------------------------------------
        worksheet1.Rows(35).Font.size = 11
        worksheet1.Cells(35, 1) = "TOD (" & MonthName(_MonthNo) & " ) - Local(ex-stock)"
        worksheet1.Range("B35").Formula = "=(B13/B26)*30"
        worksheet1.Cells(35, 2).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(35, 2)
        range1.NumberFormat = "0"

        worksheet1.Range("C35").Formula = "=(c13/c26)*30"
        worksheet1.Cells(35, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(35, 3)
        range1.NumberFormat = "0"

        worksheet1.Range("D35").Formula = "=(d13/d26)*30"
        worksheet1.Cells(35, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
        range1 = worksheet1.Cells(35, 4)
        range1.NumberFormat = "0"
        '------------------------------------------------------------------------------------
        worksheet1.Range("a33,a33").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b33,b33").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c33,c33").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d33,d33").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a33,a33").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b33,b33").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c33,c33").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d33,d33").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a34,a34").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b34,b34").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c34,c34").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d34,d34").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a34,a34").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b34,b34").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c34,c34").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d34,d34").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("a35,a35").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b35,b35").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c35,c35").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d35,d35").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("a35,a35").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b35,b35").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c35,c35").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d35,d35").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

      

        worksheet1.Range("b33:b33").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("c33:c33").Interior.Color = RGB(252, 213, 180)
        worksheet1.Range("d33:d33").Interior.Color = RGB(252, 213, 180)

        worksheet1.Range("b34:b34").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("c34:c34").Interior.Color = RGB(182, 221, 232)
        worksheet1.Range("d34:d34").Interior.Color = RGB(182, 221, 232)

        worksheet1.Range("b35:b35").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("c35:c35").Interior.Color = RGB(215, 228, 188)
        worksheet1.Range("d35:d35").Interior.Color = RGB(215, 228, 188)

    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call TOD_Report()
        Call ptlMRS_Report()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click

    End Sub
End Class
