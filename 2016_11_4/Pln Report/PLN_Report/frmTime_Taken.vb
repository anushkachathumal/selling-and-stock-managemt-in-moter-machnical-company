
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Imports Microsoft.Office.Interop.Excel
Public Class frmTime_Taken
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Const MAX_SERIALS = 156000

    

    Private Sub frmTime_Taken_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today
        chkDye.Checked = True
    End Sub

    Function Create_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet
        Dim tblDye As DataSet
        Dim vcWhere As String

        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        Dim _FirstChr As Integer
        Dim _Possible_Date As Date
        Dim _Last As Integer
        Dim _Total_NoFail As Integer
        Dim X As Integer
        Dim exc As New Application
        Try
            Dim range1 As Range
            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
            Dim _ActualBooked As Double
            Dim _weekNo As Integer
            Dim _FromWeek As Date
            Dim _Toweek As Date

            Dim vcWharer As String
            Dim vcWharer1 As String

            exc.Visible = True

            'SQL = "select * from M20Week where M20Dis='" & Trim(cboWeek.Text) & "'"
            'T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            'If isValidDataset(T01) Then
            '    _weekNo = T01.Tables(0).Rows(0)("M20Code")
            'Else
            '    MsgBox("Please select the week", MsgBoxStyle.Exclamation, "Technova .......")
            '    Exit Function
            'End If
            Dim currentCulture As System.Globalization.CultureInfo
            currentCulture = System.Globalization.CultureInfo.CurrentCulture
            _weekNo = currentCulture.Calendar.GetWeekOfYear(txtFromDate.Text, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)
            '-------------------------------------------------------------------------------------
            worksheet1.Columns("A").ColumnWidth = 20
            worksheet1.Columns("B").ColumnWidth = 20
            worksheet1.Columns("C").ColumnWidth = 10
            worksheet1.Columns("D").ColumnWidth = 10
            worksheet1.Columns("E").ColumnWidth = 10
            worksheet1.Columns("F").ColumnWidth = 10
            worksheet1.Columns("G").ColumnWidth = 10
            worksheet1.Columns("H").ColumnWidth = 10
            worksheet1.Columns("I").ColumnWidth = 10
            worksheet1.Columns("J").ColumnWidth = 10
            worksheet1.Columns("K").ColumnWidth = 10
            worksheet1.Columns("L").ColumnWidth = 10
            worksheet1.Columns("M").ColumnWidth = 10
            worksheet1.Columns("N").ColumnWidth = 10
            worksheet1.Columns("O").ColumnWidth = 10
            worksheet1.Columns("P").ColumnWidth = 10
            worksheet1.Columns("Q").ColumnWidth = 10

            worksheet1.Rows(3).Font.size = 15
            worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(3).Font.name = "Times New Roman"
            worksheet1.Rows(3).Font.BOLD = True
            If chkDye.Checked = True Then
                worksheet1.Cells(3, 1) = " Time taken from dyeing to FG Stock"
            ElseIf chkPrep.Checked = True Then
                worksheet1.Cells(3, 1) = " Time taken from Preparation to FG Stock"
            End If
            worksheet1.Cells(3, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("A3:Q3").MergeCells = True
            worksheet1.Range("A3:Q3").VerticalAlignment = XlVAlign.xlVAlignCenter


            X = 5
            worksheet1.Rows(X).Font.size = 10
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Rows(X).Font.BOLD = True
            worksheet1.Cells(X, 1) = " Stock moved Time Period"
            Dim _FROM As Date
            Dim _TO As Date

            'Dim startDate As New DateTime(txtYear.Text, 1, 1)
            'Dim weekDate As DateTime = DateAdd(DateInterval.WeekOfYear, _weekNo - 1, startDate)

            'Dim thisCulture = Globalization.CultureInfo.CurrentCulture
            'Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(weekDate)
            'Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

            'If dayName = "Monday" Then
            '    '_FromWeek = CDate(weekDate).AddDays(-1)
            'ElseIf dayName = "Tuesday" Then
            '    _FromWeek = CDate(weekDate).AddDays(-1)
            'ElseIf dayName = "Wednesday" Then
            '    _FromWeek = CDate(weekDate).AddDays(-2)
            'ElseIf dayName = "Thursday" Then
            '    _FromWeek = CDate(weekDate).AddDays(-3)
            'ElseIf dayName = "Friday" Then
            '    _FromWeek = CDate(weekDate).AddDays(-4)
            'ElseIf dayName = "Saturday" Then
            '    _FromWeek = CDate(weekDate).AddDays(-5)
            'End If

            '_Toweek = _FromWeek.AddDays(+6)
            worksheet1.Cells(X, 2) = txtFromDate.Text & " to " & txtTodate.Text
            X = X + 1

            worksheet1.Rows(X).Font.size = 10
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Rows(X).Font.BOLD = True
            worksheet1.Cells(X, 1) = "Week #"
            worksheet1.Cells(X, 2) = _weekNo
            worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft

            Dim _chr As Integer
            Dim i As Integer
            Dim Y As Integer
            Dim z As Integer
            z = 1
            _chr = 97
            X = 5
            For Y = 1 To 2
                _chr = 97
                z = 1
                For i = 1 To 2
                    worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                    ' worksheet1.Cells(X, z).WrapText = True
                    z = z + 1
                    _chr = _chr + 1
                Next
                X = X + 1
            Next

            X = X + 3
            worksheet1.Rows(X).Font.size = 10
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Rows(X).Font.BOLD = True
            worksheet1.Cells(X, 3) = "0-4 Days"
            worksheet1.Range("c10:d10").MergeCells = True
            worksheet1.Range("c10:d10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(X, 5) = "5-6 Days"
            worksheet1.Range("e10:F10").MergeCells = True
            worksheet1.Range("e10:F10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(X, 7) = "7-9 Days"
            worksheet1.Range("G10:H10").MergeCells = True
            worksheet1.Range("G10:H10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(X, 9) = "10-11 Days"
            worksheet1.Range("I10:J10").MergeCells = True
            worksheet1.Range("I10:J10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(X, 11) = "12-13 Days"
            worksheet1.Range("K10:L10").MergeCells = True
            worksheet1.Range("K10:L10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(X, 13) = "14-20 Days"
            worksheet1.Range("M10:N10").MergeCells = True
            worksheet1.Range("M10:N10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(X, 15) = "Over 20 Days"
            worksheet1.Range("O10:P10").MergeCells = True
            worksheet1.Range("O10:P10").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Cells(X, 17) = "Grand Total (KG)"
            worksheet1.Range("Q10:Q11").MergeCells = True
            worksheet1.Range("Q10:Q11").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(X, 17).WrapText = True

            _chr = 99
            For i = 1 To 15
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(217, 217, 217)
                z = z + 1
                _chr = _chr + 1
            Next

            X = X + 1
            worksheet1.Rows(X).Font.size = 8
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Rows(X).Font.BOLD = True

            z = 3

            _chr = 99
            For Y = 1 To 7
                worksheet1.Cells(X, z) = "Qty (Kg)"
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).MergeCells = True
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet1.Cells(X, z).HorizontalAlignment = XlHAlign.xlHAlignCenter
                z = z + 1
                _chr = _chr + 1
                worksheet1.Cells(X, z) = "%"
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).MergeCells = True
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet1.Cells(X, z).HorizontalAlignment = XlHAlign.xlHAlignCenter
                z = z + 1
                _chr = _chr + 1
            Next

            _chr = 99
            For i = 1 To 15
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(217, 217, 217)
                z = z + 1
                _chr = _chr + 1
            Next
            worksheet1.Rows(X).rowheight = 18
            X = X + 1
            worksheet1.Rows(X).rowheight = 15
            worksheet1.Cells(X, 2) = "Marl"
            worksheet1.Rows(X).Font.size = 8
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Cells(X, 2).Font.BOLD = True
            worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Cells(X, 2).WrapText = True

            _chr = 98
            For i = 1 To 16
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                ' worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(217, 217, 217)
                z = z + 1
                _chr = _chr + 1
            Next


            X = X + 1
            worksheet1.Rows(X).rowheight = 15
            worksheet1.Cells(X, 2) = "White"
            worksheet1.Rows(X).Font.size = 8
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Cells(X, 2).Font.BOLD = True
            worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Cells(X, 2).WrapText = True

            _chr = 98
            For i = 1 To 16
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                ' worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(217, 217, 217)
                z = z + 1
                _chr = _chr + 1
            Next

            X = X + 1
            worksheet1.Rows(X).rowheight = 15
            worksheet1.Cells(X, 2) = "Shade"
            worksheet1.Rows(X).Font.size = 8
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Cells(X, 2).Font.BOLD = True
            worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Cells(X, 2).WrapText = True

            _chr = 98
            For i = 1 To 16
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                ' worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(217, 217, 217)
                z = z + 1
                _chr = _chr + 1
            Next


            X = X + 1
            worksheet1.Rows(X).rowheight = 15
            worksheet1.Cells(X, 2) = "Yarn Dye"
            worksheet1.Rows(X).Font.size = 8
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Cells(X, 2).Font.BOLD = True
            worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Cells(X, 2).WrapText = True

            _chr = 98
            For i = 1 To 16
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                ' worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(217, 217, 217)
                z = z + 1
                _chr = _chr + 1
            Next

            X = X + 1
            worksheet1.Rows(X).rowheight = 15
            worksheet1.Cells(X, 2) = "Total"
            worksheet1.Rows(X).Font.size = 8
            ' worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Cells(X, 2).Font.BOLD = True
            worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Cells(X, 2).WrapText = True

            _chr = 98
            For i = 1 To 16
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_chr) & X, Chr(_chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                ' worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(217, 217, 217)
                z = z + 1
                _chr = _chr + 1
            Next

            X = 12
            _chr = 99
            For i = 1 To 5
                _chr = 99
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(146, 208, 80)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(146, 208, 80)
                X = X + 1
            Next
            X = 12
            _chr = 101
            For i = 1 To 5
                _chr = 101
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(255, 192, 0)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(255, 192, 0)
                X = X + 1
            Next
            X = 12
            _chr = 103
            For i = 1 To 5
                _chr = 103
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(218, 150, 148)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(218, 150, 148)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(218, 150, 148)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(218, 150, 148)
                X = X + 1
            Next

            X = 12
            _chr = 107
            For i = 1 To 5
                _chr = 107
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(255, 0, 0)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(255, 0, 0)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(255, 0, 0)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(255, 0, 0)
                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(255, 0, 0)

                _chr = _chr + 1
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(255, 0, 0)

                X = X + 1
            Next


            X = 12
            _chr = 113
            For i = 1 To 5
                _chr = 113
                worksheet1.Range(Chr(_chr) & X & ":" & Chr(_chr) & X).Interior.Color = RGB(217, 217, 217)


                X = X + 1
            Next
            '--------------------------------------------------------------------------------------------
            Dim _MarlQTY As Double

            _MarlQTY = 0
            'MARL
            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Dye_Date)-1 between '0' and '4' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Prep_Date)-1 between '0' and '4' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 12
            worksheet1.Cells(X, 3) = _MarlQTY
            worksheet1.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 3)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Dye_Date)-1 between '5' and '6' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Prep_Date)-1 between '5' and '6' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 12
            worksheet1.Cells(X, 5) = _MarlQTY
            worksheet1.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 5)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Dye_Date)-1 between '7' and '9' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Prep_Date)-1 between '7' and '9' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 12
            worksheet1.Cells(X, 7) = _MarlQTY
            worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 7)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Dye_Date)-1 between '10' and '11' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Prep_Date)-1 between '10' and '11' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 12
            worksheet1.Cells(X, 9) = _MarlQTY
            worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 9)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Dye_Date)-1 between '12' and '13' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Prep_Date)-1 between '12' and '13' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 12
            worksheet1.Cells(X, 11) = _MarlQTY
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 11)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Dye_Date)-1 between '14' and '20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Prep_Date)-1 between '14' and '20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 12
            worksheet1.Cells(X, 13) = _MarlQTY
            worksheet1.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 13)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Dye_Date)-1 >='20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Product_Type='Marl' and day(M19Posting_Date-M19Prep_Date)-1 >='20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 12
            worksheet1.Cells(X, 15) = _MarlQTY
            worksheet1.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 15)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 17) = "=C12+E12+G12+I12+K12+M12+O12"
            worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 17)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 4) = "=C12/Q12"
            worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 4)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 6) = "=E12/Q12"
            worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 6)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 8) = "=G12/Q12"
            worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 8)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 10) = "=I12/Q12"
            worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 10)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 12) = "=K12/Q12"
            worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 12)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 14) = "=M12/Q12"
            worksheet1.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 14)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 16) = "=O12/Q12"
            worksheet1.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 16)
            range1.NumberFormat = "0%"
            '==========================================================================================
            'WHITE

            _MarlQTY = 0

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Dye_Date)-1 between '0' and '4' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Prep_Date)-1 between '0' and '4' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 13
            worksheet1.Cells(X, 3) = _MarlQTY
            worksheet1.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 3)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Dye_Date)-1 between '5' and '6' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Prep_Date)-1 between '5' and '6' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 13
            worksheet1.Cells(X, 5) = _MarlQTY
            worksheet1.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 5)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Dye_Date)-1 between '7' and '9' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Prep_Date)-1 between '7' and '9' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 13
            worksheet1.Cells(X, 7) = _MarlQTY
            worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 7)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Dye_Date)-1 between '10' and '11' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Prep_Date)-1 between '10' and '11' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 9) = _MarlQTY
            worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 9)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Dye_Date)-1 between '12' and '13' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Prep_Date)-1 between '12' and '13' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            worksheet1.Cells(X, 11) = _MarlQTY
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 11)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Dye_Date)-1 between '14' and '20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Prep_Date)-1 between '14' and '20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 13) = _MarlQTY
            worksheet1.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 13)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Dye_Date)-1 >='20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='White' and day(M19Posting_Date-M19Prep_Date)-1 >='20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 15) = _MarlQTY
            worksheet1.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 15)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 17) = "=C13+E13+G13+I13+K13+M13+O13"
            worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 17)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 4) = "=C13/Q13"
            worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 4)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 6) = "=E13/Q13"
            worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 6)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 8) = "=G13/Q13"
            worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 8)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 10) = "=I13/Q13"
            worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 10)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 12) = "=K13/Q13"
            worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 12)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 14) = "=M13/Q13"
            worksheet1.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 14)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 16) = "=O13/Q13"
            worksheet1.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 16)
            range1.NumberFormat = "0%"
            '=========================================================================
            'SHADE
            _MarlQTY = 0

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Dye_Date)-1 between '0' and '4' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Prep_Date)-1 between '0' and '4' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 14
            worksheet1.Cells(X, 3) = _MarlQTY
            worksheet1.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 3)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Dye_Date)-1 between '5' and '6' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Prep_Date)-1 between '5' and '6' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 14
            worksheet1.Cells(X, 5) = _MarlQTY
            worksheet1.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 5)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Dye_Date)-1 between '7' and '9' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Prep_Date)-1 between '7' and '9' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 14
            worksheet1.Cells(X, 7) = _MarlQTY
            worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 7)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Dye_Date)-1 between '10' and '11' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Prep_Date)-1 between '10' and '11' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 9) = _MarlQTY
            worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 9)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Dye_Date)-1 between '12' and '13' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Prep_Date)-1 between '12' and '13' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            worksheet1.Cells(X, 11) = _MarlQTY
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 11)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Dye_Date)-1 between '14' and '20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Prep_Date)-1 between '14' and '20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 13) = _MarlQTY
            worksheet1.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 13)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Dye_Date)-1 >='20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type<>'White' AND M16Product_Type='Solid' and day(M19Posting_Date-M19Prep_Date)-1 >='20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 15) = _MarlQTY
            worksheet1.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 15)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 17) = "=C14+E14+G14+I14+K14+M14+O14"
            worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 17)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 4) = "=C14/Q14"
            worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 4)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 6) = "=E14/Q14"
            worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 6)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 8) = "=G14/Q14"
            worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 8)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 10) = "=I14/Q14"
            worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 10)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 12) = "=K14/Q14"
            worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 12)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 14) = "=M14/Q14"
            worksheet1.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 14)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 16) = "=O14/Q14"
            worksheet1.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 16)
            range1.NumberFormat = "0%"
            '===================================================================
            'yarn dye
            _MarlQTY = 0

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Dye_Date)-1 between '0' and '4' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Prep_Date)-1 between '0' and '4' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 15
            worksheet1.Cells(X, 3) = _MarlQTY
            worksheet1.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 3)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Dye_Date)-1 between '5' and '6' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Prep_Date)-1 between '5' and '6' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 15
            worksheet1.Cells(X, 5) = _MarlQTY
            worksheet1.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 5)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Dye_Date)-1 between '7' and '9' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Prep_Date)-1 between '7' and '9' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            X = 15
            worksheet1.Cells(X, 7) = _MarlQTY
            worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 7)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Dye_Date)-1 between '10' and '11' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='Yarn Dyes' and day(M19Posting_Date-M19Prep_Date)-1 between '10' and '11' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 9) = _MarlQTY
            worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 9)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Dye_Date)-1 between '12' and '13' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Prep_Date)-1 between '12' and '13' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next
            worksheet1.Cells(X, 11) = _MarlQTY
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 11)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Dye_Date)-1 between '14' and '20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Prep_Date)-1 between '14' and '20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 13) = _MarlQTY
            worksheet1.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 13)
            range1.NumberFormat = "0.0"

            If chkDye.Checked = True Then
                vcWharer1 = "year(m19dye_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Dye_Date)-1 >='20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            ElseIf chkPrep.Checked = True Then
                vcWharer1 = "year(M19Prep_Date)<>'1900' and M16Shade_Type='Yarn Dyes'  and day(M19Posting_Date-M19Prep_Date)-1 >='20' AND M19Posting_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM19Taken_quarry", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer1))
            i = 0
            _MarlQTY = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _MarlQTY = _MarlQTY + T01.Tables(0).Rows(i)("M19Qty")
                i = i + 1
            Next

            worksheet1.Cells(X, 15) = _MarlQTY
            worksheet1.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 15)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 17) = "=C15+E15+G15+I15+K15+M15+O15"
            worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 17)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 4) = "=C15/Q15"
            worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 4)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 6) = "=E15/Q15"
            worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 6)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 8) = "=G15/Q15"
            worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 8)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 10) = "=I15/Q15"
            worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 10)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 12) = "=K15/Q15"
            worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 12)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 14) = "=M15/Q15"
            worksheet1.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 14)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 16) = "=O15/Q15"
            worksheet1.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 16)
            range1.NumberFormat = "0%"

            X = 16
            worksheet1.Cells(X, 3) = "=sum(C12:C15)"
            worksheet1.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 3)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 5) = "=sum(E12:E15)"
            worksheet1.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 5)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 7) = "=sum(G12:G15)"
            worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 7)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 9) = "=sum(I12:I15)"
            worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 9)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 11) = "=sum(K12:K15)"
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 11)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 13) = "=sum(M12:M15)"
            worksheet1.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 13)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 15) = "=sum(O12:O15)"
            worksheet1.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 15)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 17) = "=sum(Q12:Q15)"
            worksheet1.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 17)
            range1.NumberFormat = "0.0"

            worksheet1.Cells(X, 4) = "=C16/Q16"
            worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 4)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 6) = "=E16/Q16"
            worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 6)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 8) = "=G16/Q16"
            worksheet1.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 8)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 10) = "=I16/Q16"
            worksheet1.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 10)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 12) = "=K16/Q16"
            worksheet1.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 12)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 14) = "=M16/Q16"
            worksheet1.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 14)
            range1.NumberFormat = "0%"

            worksheet1.Cells(X, 16) = "=O16/Q16"
            worksheet1.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(X, 16)
            range1.NumberFormat = "0%"


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try


    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        txtFromDate.Text = Today
        txtTodate.Text = Today

    End Sub

    Function Upload_Textfile()
        Dim strFileName As String
        'strFileName = "\\Tjlapp04\grginspec_dload$\TJL CAT.txt"
        ' strFileName = "E:\TJL_MILAN\SAP_DOWNLOADS\Sales Forcust\time tkn.txt"
        'Dim strFileName As String
        strFileName = ConfigurationManager.AppSettings("FilePath") + "\time tkn.txt"
        Dim CurrGameWinningSerials(0 To MAX_SERIALS) As Long
        Dim fileHndl As Long
        Dim lLineNo As Long

        Dim strOrder, strLineItem, _
      strMerchant, strDis As String
        Dim strDep As String
        Dim strSpec As Double

        Dim strKg As Double
        Dim strMtr As Double
        ' Dim strFileName As String '= _
        Dim strDate As String
        ' Dim strDate As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim M03Knittingorder As DataSet
        Dim ncQryType As String
        Dim nvcVccode As String
        Dim linesList As New List(Of String)(IO.File.ReadAllLines(strFileName))
        Dim _RefNo As Integer
        Dim P01Parameter As DataSet
        Dim _Value As Double
        Dim M06Cls As DataSet
        Dim str30Class As String
        Dim strMaterial As String

        Dim nvcFieldList As String
        Dim nvcWhereClause As String
        Dim nvcVcLine As String
        Dim nvcVcDate As String
        Dim nvcFrom As String
        Dim nvcTo As String
        Dim nvcQtype As String
        Dim strFrom As Integer
        Dim strTo As String
        Dim strMovement As String
        Dim strSorder As String
        Dim strLine_Item As String
        Dim strOrdr_No As String
        Dim strOrder_type As String
        Dim strQty As Double
        Dim strDye_Date As String
        Dim strPrefDate As String
        Dim strFreshBatch As String
        Dim vcWharer As String
        Dim strWeek As Integer

        Dim characterToRemove As String


        Try

            'nvcFieldList1 = "delete from M09ZPL_ORDER "
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            fileHndl = FreeFile()


            ' strFileName = Dir(strFileName)

            'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object strValidSerialFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileHndl, strFileName, OpenMode.Input)
            lLineNo = 0
            Dim strRow As String

            Do Until EOF(fileHndl)


                '  Line Input #fileHndl, strRow
                'UPGRADE_WARNING: Couldn't resolve default property of object fileHndl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strRow = LineInput(fileHndl)

                'If lLineNo = 18 Then
                '    MsgBox("")
                'End If
                If Trim(strRow) <> "" Then
                    ncQryType = "LST"
                    If InStr(1, strRow, vbTab) > 0 Then
                        If (Trim(Split(strRow, vbTab)(0))) <> "" Then
                            '  CurrGameWinningSerials(lLineNo) = Trim(Split(strRow, vbTab)(0))
                            strFrom = (Trim(Split(strRow, vbTab)(0)))
                            strTo = (Trim(Split(strRow, vbTab)(1)))
                            strMovement = CInt(Trim(Split(strRow, vbTab)(3)))
                            '    strMaterial = Microsoft.VisualBasic.Left(strMaterial, 2) & "-" & Microsoft.VisualBasic.Right(strMaterial, 5)
                            strSorder = (Trim(Split(strRow, vbTab)(4)))
                            strLine_Item = (Trim(Split(strRow, vbTab)(5)))
                            strMaterial = (Trim(Split(strRow, vbTab)(6)))

                            characterToRemove = "-"

                            strMaterial = (Replace(strMaterial, characterToRemove, ""))

                            strDate = (Trim(Split(strRow, vbTab)(7)))
                            'strDate = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 7), 2)
                            'strDate = strDate & "/" & Microsoft.VisualBasic.Right(Trim(Split(strRow, vbTab)(7)), 2)
                            'strDate = strDate & "/" & Microsoft.VisualBasic.Left(Trim(Split(strRow, vbTab)(7)), 4)

                            strOrdr_No = (Trim(Split(strRow, vbTab)(8)))
                            strOrder_type = (Trim(Split(strRow, vbTab)(9)))
                            strQty = 0
                            strQty = (Trim(Split(strRow, vbTab)(10)))
                            strDye_Date = ""
                            If (Trim(Split(strRow, vbTab)(12))) <> "" Then

                                strDye_Date = (Trim(Split(strRow, vbTab)(12)))
                                'strDye_Date = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDye_Date, 6), 2)
                                'strDye_Date = strDye_Date & "/" & Microsoft.VisualBasic.Right(Trim(Split(strRow, vbTab)(12)), 2)
                                'strDye_Date = strDye_Date & "/" & Microsoft.VisualBasic.Left(Trim(Split(strRow, vbTab)(12)), 4)


                            End If

                            strPrefDate = ""
                            If (Trim(Split(strRow, vbTab)(13))) <> "" Then

                                strPrefDate = (Trim(Split(strRow, vbTab)(13)))
                                'strPrefDate = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strPrefDate, 6), 2)
                                'strPrefDate = strPrefDate & "/" & Microsoft.VisualBasic.Right(Trim(Split(strRow, vbTab)(13)), 2)
                                'strPrefDate = strPrefDate & "/" & Microsoft.VisualBasic.Left(Trim(Split(strRow, vbTab)(13)), 4)


                            End If
                    

                            Dim currentCulture As System.Globalization.CultureInfo
                            currentCulture = System.Globalization.CultureInfo.CurrentCulture
                            strWeek = currentCulture.Calendar.GetWeekOfYear(strDate, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

                            strFreshBatch = (Trim(Split(strRow, vbTab)(14)))
                            vcWharer = "M19From=" & strFrom & "  and M19Posting_Date='" & strDate & "' and M19Order_No='" & strOrdr_No & "' and M19Week='" & strWeek & " and M19Year=" & Year(strDate) & " and M19Qty='" & strQty & "' "

                            ' M06Cls = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetM19Time_Taken", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate), New SqlParameter("@vcBATCH", strBatch))
                            M06Cls = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetM19Time_Taken", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause", vcWharer))
                            If isValidDataset(M06Cls) Then
                                'For Each DTRow As DataRow In M06Cls.Tables(0).Rows
                                '    ' nvcWhereClause = "M09Sales_Oredr='" & strOrder & "' AND M09Line_Item='" & strLineItem & "' AND M09Del_Date='" & strDate & "' AND M09BatchNo='" & nvcBatch & "'"
                                '    ncQryType = "UPD"
                                '    nvcFieldList1 = "M09Qty_KG='" & (Trim(Split(strRow, vbTab)(6))) & "',M09Qty_Mtr='" & (Trim(Split(strRow, vbTab)(13))) & "'"
                                '    up_GetSetM09ZPL_ORDER(ncQryType, nvcFieldList1, nvcWhereClause, nvcVcLine, nvcVcDate, nvcBatch, nvcVccode, connection, transaction)
                                '    ' ExecuteNonQueryText(connection, transaction, "up_GetSetM06Delivary_Qty", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcCode", strOrder), New SqlParameter("@vcLine", strLineItem), New SqlParameter("@vcDate", strDate))
                                'Next
                            Else

                                ncQryType = "ADD"
                                nvcFieldList1 = "(M19From," & "M19To," & "M19Move_Type," & "M19Sales_Order," & "M19Line_Items," & "M19Material," & "M19Posting_Date," & "M19Order_No," & "M19Order_Type," & "M19Qty," & "M19Dye_Date," & "M19Prep_Date," & "M19Fresh_Batch," & "M19Week," & "M19Year) " & "values('" & Trim(strFrom) & "','" & strTo & "','" & strMovement & "','" & strSorder & "','" & strLine_Item & "','" & strMaterial & "','" & strDate & "','" & strOrdr_No & "','" & strOrder_type & "','" & strQty & "','" & strDye_Date & "','" & strPrefDate & "','" & strFreshBatch & "','" & strWeek & "','" & Year(strDate) & "')"
                                up_GetSetup_GetSetM19Time_Taken(ncQryType, nvcFieldList1, nvcWhereClause, connection, transaction)
                            End If
                            '---------------------------------------------------------------------------------------------

                            linesList.RemoveAt(0)
                            ''  MsgBox(linesList.ToArray().ToString)
                            'IO.File.WriteAllLines(strFileName, linesList.ToArray())

                            strMtr = 0
                            strKg = 0

                        Else
                            '  Err.Raise(vbObjectError + 18001, "GenerateInstantFile(str,str,str)", "Invalid Record At Line " & CStr(lLineNo))
                        End If
                    End If
                End If

                lLineNo = lLineNo + 1

            Loop
            MsgBox("Upload successfully", MsgBoxStyle.Information, "Technova ...")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            ' transaction.Rollback()
            ' MsgBox("M09ZPL_ORDER ")
            connection.Close()
            FileClose()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                MsgBox(lLineNo & "time tkn.txt")
                connection.Close()
                FileClose()
            End If
        End Try
    End Function

    Private Sub chkUpload_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUpload.CheckedChanged
        If chkUpload.Checked = True Then
            Call Upload_Textfile()
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Create_File()
    End Sub

    Private Sub chkDye_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDye.CheckedChanged
        If chkDye.Checked = True Then
            chkPrep.Checked = False
        End If
    End Sub

    Private Sub chkPrep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPrep.CheckedChanged
        If chkPrep.Checked = True Then
            chkDye.Checked = False
        End If
    End Sub
End Class