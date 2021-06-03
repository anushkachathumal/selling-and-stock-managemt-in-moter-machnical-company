
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
Imports System.Object
Imports Microsoft.Office.Tools.Excel
Imports System.Drawing
Imports System.Globalization.CultureInfo.CurrentCulture
Imports System.DateTime
Imports System.DayOfWeek
Public Class frmOrder_Placement
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable
    Dim _Customer As String
    Dim _Department As String
    Dim _Merchant As String
    Dim _Status As String

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmOrder_Placement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today

        chkIN.Checked = True
        chkOCI.Checked = True
        chkPTL.Checked = True
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        txtFromDate.Text = Today
        txtTodate.Text = Today

        chkIN.Checked = True
        chkOCI.Checked = True
        chkPTL.Checked = True
    End Sub

    Function Create_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet
        Dim _1stQualityROW As Integer
        Dim _PreShade As String

        Dim vcWhere As String
        Dim _FirstRow As Integer
        Dim _SHADE As String
        Dim _FromDate As Date
        Dim _ToDate As Date
        '  Dim M02 As DataSet
        Dim cargoWeights(5) As Double
        Dim _20sd(5) As String
        Dim _t As Integer

        Try
            Dim exc As New Application

            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
            Dim range1 As Range
            Dim _Chr As String
            Dim i As Integer
            Dim X As Integer

            exc.Visible = True

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 11
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 25
            worksheet2.Columns("B").ColumnWidth = 13
            worksheet2.Columns("C").ColumnWidth = 13
            worksheet2.Columns("D").ColumnWidth = 13
            worksheet2.Columns("E").ColumnWidth = 13
            worksheet2.Columns("F").ColumnWidth = 13
            worksheet2.Columns("G").ColumnWidth = 13
            worksheet2.Columns("H").ColumnWidth = 13


            worksheet2.Rows(2).Font.size = 10
            worksheet2.Rows(2).rowheight = 35
            worksheet2.Rows(2).Font.name = "Times New Roman"
            worksheet2.Rows(2).Font.BOLD = True

            worksheet2.Cells(2, 2) = "Unit Head "
            worksheet2.Cells(2, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("B2:B2").MergeCells = True
            worksheet2.Range("B2:B2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(2, 3) = "Quantity(m) "
            worksheet2.Cells(2, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("C2:C2").MergeCells = True
            worksheet2.Range("C2:C2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(2, 4) = "Quantity(Kg) "
            worksheet2.Cells(2, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("D2:D2").MergeCells = True
            worksheet2.Range("D2:D2").VerticalAlignment = XlVAlign.xlVAlignCenter
            X = 2
            _Chr = 98
            For i = 1 To 3


                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)

                _Chr = _Chr + 1

            Next
            '===================================================================================================
            X = 3
            _FromDate = txtFromDate.Text
            Dim thisCulture = Globalization.CultureInfo.CurrentCulture
            Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_FromDate)
            Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

            If dayName = "Sunday" Then
                _FromDate = CDate(_FromDate).AddDays(-6)
            ElseIf dayName = "Tuesday" Then
                _FromDate = CDate(_FromDate).AddDays(-1)
            ElseIf dayName = "Wednesday" Then
                _FromDate = CDate(_FromDate).AddDays(-2)
            ElseIf dayName = "Thursday" Then
                _FromDate = CDate(_FromDate).AddDays(-3)
            ElseIf dayName = "Friday" Then
                _FromDate = CDate(_FromDate).AddDays(-4)
            ElseIf dayName = "Saturday" Then
                _FromDate = CDate(_FromDate).AddDays(-5)
            End If

            Dim z As Integer

            _ToDate = _FromDate.AddDays(+6)

            If chkIN.Checked = True And chkOCI.Checked = True Then
                vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "'"

            ElseIf chkIN.Checked = True Then
                vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "' and M17Location<>'4004'"
            ElseIf chkOCI.Checked = True Then
                vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "' and M17Location='4004'"
            End If

            M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "UNT"), New SqlParameter("@vcWhereClause1", vcWhere))
            i = 0
            For Each DTRow5 As DataRow In M03.Tables(0).Rows
                worksheet2.Cells(X, 2) = Trim(M03.Tables(0).Rows(i)("M14Name"))
                worksheet2.Cells(X, 3) = Trim(M03.Tables(0).Rows(i)("M01SO_Qty"))
                range1 = worksheet2.Cells(X, 3)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(X, 4) = Trim(M03.Tables(0).Rows(i)("qtykg"))
                range1 = worksheet2.Cells(X, 4)
                range1.NumberFormat = "0.00"



                _Chr = 98
                For z = 1 To 3


                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    '  worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)

                    _Chr = _Chr + 1

                Next
                X = X + 1
                i = i + 1
            Next
            worksheet2.Rows(X).Font.size = 12
            worksheet2.Rows(X).Font.Bold = True
            worksheet2.Cells(X, 2) = "Total"
            worksheet2.Range("C" & (X)).Formula = "=SUM(C3:C" & X - 1 & ")"
            range1 = worksheet2.Cells(X, 3)
            range1.NumberFormat = "0.00"

            worksheet2.Range("D" & (X)).Formula = "=SUM(D3:D" & X - 1 & ")"
            range1 = worksheet2.Cells(X, 4)
            range1.NumberFormat = "0.00"

            _Chr = 98
            For z = 1 To 3


                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)

                

                _Chr = _Chr + 1

            Next

          


            Dim chartPage As Microsoft.Office.Interop.Excel.Chart
            Dim xlCharts As Microsoft.Office.Interop.Excel.ChartObjects
            Dim myChart As Microsoft.Office.Interop.Excel.ChartObject
            Dim chartRange As Microsoft.Office.Interop.Excel.Range
            Dim chartRange1 As Microsoft.Office.Interop.Excel.Range
            Dim chartRange2 As Microsoft.Office.Interop.Excel.Range


            Dim t_SerCol As Microsoft.Office.Interop.Excel.SeriesCollection
            Dim t_Series As Microsoft.Office.Interop.Excel.Series
            Dim X_ChartH As Integer

            Dim rh As Double

            rh = (X - 5) * 12.75
            rh = rh + 25.5
            rh = rh + (15 * 5)
            rh = rh + 33.75
            xlCharts = worksheet2.ChartObjects
            X_ChartH = (X + 18) * 10
            myChart = xlCharts.Add(7, rh, 905, 300)

            chartPage = myChart.Chart
            'chartPage.ChartStyle = "Style 6"
            ' chartRange = worksheet1.Range(worksheet1.Cells("A8", "E" & (X - 1)), worksheet1.Cells("A9", "A" & (X - 1)))


            ' chartPage.ChartType = XlChartType.xl3DColumnClustered
            '  chartPage.ChartStyle = 6
            chartRange = worksheet2.Range("B2", "D" & (X - 1))
            chartPage.SetSourceData(Source:=chartRange)
            '  chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xl3DColumnStacked100
            chartPage.ChartType = XlChartType.xlColumnClustered
            'chartPage.fill.color = Color.Black

            worksheet2.ChartObjects.select()
            '  workbook.ribbon.tabs(0).select()
            'xlCharts.ChartStyle = 4


            ' workbook.ActiveChart.ChartStyle = 8
            chartPage.HasTitle = True
            'Dim X1 As Integer, Y1 As Integer
            'X1 = 628
            'Y1 = 146
            'SetCursorPos(X1, Y1)
            'Dim Layout As Integer
            'Dim ChartType As Object

            'ChartType = XlChartType.xl3DColumnClustered
            chartPage.ApplyLayout(5)
            ' chartPage.ChartStyle = 8
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '         msoElementChartTitleCenteredOverlay)
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '                   msoElementPrimaryCategoryAxisTitleHorizontal)
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '                   msoElementPrimaryValueAxisTitleRotated)
            chartPage.ChartTitle.Text = ("Confirmed orders received qty during last week – Unit wise")

            '=======================================
            Dim _WeekNo As Integer

            X = 37
            _FromDate = _ToDate
            dayOfWeek = thisCulture.Calendar.GetDayOfWeek(_FromDate)
            dayName = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

            'If dayName = "Sunday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-6)
            'ElseIf dayName = "Tuesday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-1)
            'ElseIf dayName = "Wednesday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-2)
            'ElseIf dayName = "Thursday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-3)
            'ElseIf dayName = "Friday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-4)
            'ElseIf dayName = "Saturday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-5)
            'End If
            '_FromDate = _FromDate.AddDays(+6)
            _FromDate = _FromDate.AddDays(-42)
            _WeekNo = weekNumber(_FromDate)
            worksheet2.Rows(X).Font.size = 10
            worksheet2.Rows(X).rowheight = 35
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            Dim Y As Integer
            Y = 3
            For z = 1 To 6
                If _WeekNo = 54 Then
                    _WeekNo = 1
                End If

                worksheet2.Cells(X, Y) = "Week " & _WeekNo
                worksheet2.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _WeekNo = _WeekNo + 1
                Y = Y + 1
            Next

          
            _Chr = 98
            For z = 1 To 7


                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                _Chr = _Chr + 1

            Next
            X = X + 1
            worksheet2.Rows(X).Font.size = 10
            ' worksheet2.Rows(X).rowheight = 35
            worksheet2.Rows(X).Font.name = "Times New Roman"
            Y = 2
            worksheet2.Cells(X, Y) = "Quantity (m)"
            ' Dim thisCulture = Globalization.CultureInfo.CurrentCulture
            'dayOfWeek = thisCulture.Calendar.GetDayOfWeek(_FromDate)
            'dayName = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)
            Y = Y + 1
            If dayName = "Sunday" Then
                _FromDate = CDate(_FromDate).AddDays(+1)
            ElseIf dayName = "Tuesday" Then
                _FromDate = CDate(_FromDate).AddDays(-1)
            ElseIf dayName = "Wednesday" Then
                _FromDate = CDate(_FromDate).AddDays(-2)
            ElseIf dayName = "Thursday" Then
                _FromDate = CDate(_FromDate).AddDays(-3)
            ElseIf dayName = "Friday" Then
                _FromDate = CDate(_FromDate).AddDays(-4)
            ElseIf dayName = "Saturday" Then
                _FromDate = CDate(_FromDate).AddDays(-5)
            End If
            For z = 1 To 6
                _ToDate = _FromDate.AddDays(+6)
                If chkIN.Checked = True And chkOCI.Checked = True Then
                    vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "'"

                ElseIf chkIN.Checked = True Then
                    vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "' and M17Location<>'4004'"
                ElseIf chkOCI.Checked = True Then
                    vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "' and M17Location='4004'"
                End If

                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "OPL"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M03) Then
                    If Not DBNull.Value.Equals(M03.Tables(0).Rows(0)("M01SO_Qty")) Then

                        worksheet2.Cells(X, Y) = Trim(M03.Tables(0).Rows(0)("M01SO_Qty"))
                        range1 = worksheet2.Cells(X, Y)
                        range1.NumberFormat = "0.00"
                    End If
                End If
                Y = Y + 1
                _FromDate = _ToDate.AddDays(+1)
            Next
            _Chr = 98
            For z = 1 To 7


                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                'worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                'worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                'worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                _Chr = _Chr + 1

            Next


            rh = (X) * 12.75
            rh = rh + 25.5
            rh = rh + (15 * 5)
            rh = rh + 33.75
            xlCharts = worksheet2.ChartObjects
            X_ChartH = (X + 18) * 10
            myChart = xlCharts.Add(7, rh, 905, 300)

            chartPage = myChart.Chart
            'chartPage.ChartStyle = "Style 6"
            ' chartRange = worksheet1.Range(worksheet1.Cells("A8", "E" & (X - 1)), worksheet1.Cells("A9", "A" & (X - 1)))


            ' chartPage.ChartType = XlChartType.xl3DColumnClustered
            '  chartPage.ChartStyle = 6
            chartRange = worksheet2.Range("B36", "H" & (X))
            chartPage.SetSourceData(Source:=chartRange)
            '  chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xl3DColumnStacked100
            chartPage.ChartType = XlChartType.xlLineMarkersStacked
            'chartPage.fill.color = Color.Black

            worksheet2.ChartObjects.select()
            '  workbook.ribbon.tabs(0).select()
            'xlCharts.ChartStyle = 4


            ' workbook.ActiveChart.ChartStyle = 8
            chartPage.HasTitle = True
            'Dim X1 As Integer, Y1 As Integer
            'X1 = 628
            'Y1 = 146
            'SetCursorPos(X1, Y1)
            'Dim Layout As Integer
            'Dim ChartType As Object

            'ChartType = XlChartType.xl3DColumnClustered
            chartPage.ApplyLayout(5)
            '  chartPage.ChartStyle = 8
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '         msoElementChartTitleCenteredOverlay)
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '                   msoElementPrimaryCategoryAxisTitleHorizontal)
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '                   msoElementPrimaryValueAxisTitleRotated)
            chartPage.ChartTitle.Text = ("Last six weeks total orders received")
            '--------------------------------------------------------------------
            X = 61
            _FromDate = _ToDate
            dayOfWeek = thisCulture.Calendar.GetDayOfWeek(_FromDate)
            dayName = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

            'If dayName = "Sunday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-6)
            'ElseIf dayName = "Tuesday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-1)
            'ElseIf dayName = "Wednesday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-2)
            'ElseIf dayName = "Thursday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-3)
            'ElseIf dayName = "Friday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-4)
            'ElseIf dayName = "Saturday" Then
            '    _FromDate = CDate(_FromDate).AddDays(-5)
            'End If
            '_FromDate = _FromDate.AddDays(+6)
            _FromDate = _FromDate.AddDays(-42)
            _WeekNo = weekNumber(_FromDate)
            worksheet2.Rows(X).Font.size = 10
            worksheet2.Rows(X).rowheight = 35
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            '  Dim Y As Integer
            Y = 3
            For z = 1 To 6
                If _WeekNo = 54 Then
                    _WeekNo = 1
                End If

                worksheet2.Cells(X, Y) = "Week " & _WeekNo
                worksheet2.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _WeekNo = _WeekNo + 1
                Y = Y + 1
            Next


            _Chr = 98
            For z = 1 To 7


                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                _Chr = _Chr + 1

            Next
            '---------------------------------------------------------------------------------------------------------

            Dim X1 As Integer

            If dayName = "Sunday" Then
                _FromDate = CDate(_FromDate).AddDays(+1)
            ElseIf dayName = "Tuesday" Then
                _FromDate = CDate(_FromDate).AddDays(-1)
            ElseIf dayName = "Wednesday" Then
                _FromDate = CDate(_FromDate).AddDays(-2)
            ElseIf dayName = "Thursday" Then
                _FromDate = CDate(_FromDate).AddDays(-3)
            ElseIf dayName = "Friday" Then
                _FromDate = CDate(_FromDate).AddDays(-4)
            ElseIf dayName = "Saturday" Then
                _FromDate = CDate(_FromDate).AddDays(-5)
            End If
            Dim y1 As Integer
            Y = 2
            X = 62

            For z = 1 To 6
                Y = Y + 1
                X = 62
                _ToDate = _FromDate.AddDays(+6)
                If chkIN.Checked = True And chkOCI.Checked = True Then
                    vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "'"

                ElseIf chkIN.Checked = True Then
                    vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "' and M17Location<>'4004'"
                ElseIf chkOCI.Checked = True Then
                    vcWhere = "M01SO_Date between '" & _FromDate & "' and '" & _ToDate & "' and M17Location='4004'"
                End If

                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "UNT"), New SqlParameter("@vcWhereClause1", vcWhere))
                ' X = X + 1
                ' Y = 2
                X1 = 0
                For Each DTRow5 As DataRow In M03.Tables(0).Rows
                    ' X = 62
                    ' Y = 2

                    worksheet2.Rows(X).Font.size = 10
                    ' worksheet2.Rows(X).rowheight = 35
                    worksheet2.Rows(X).Font.name = "Times New Roman"
                    worksheet2.Cells(X, 2) = Trim(M03.Tables(0).Rows(X1)("M14Name"))
                    '  Y = Y + 1
                    If Not DBNull.Value.Equals(M03.Tables(0).Rows(X1)("M01SO_Qty")) Then

                        worksheet2.Cells(X, Y) = Trim(M03.Tables(0).Rows(X1)("M01SO_Qty"))
                        range1 = worksheet2.Cells(X, Y)
                        range1.NumberFormat = "0.00"
                    End If

                    _Chr = 98
                    For y1 = 1 To 7


                        worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        'worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        'worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                        'worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                        _Chr = _Chr + 1

                    Next

                    X = X + 1
                    X1 = X1 + 1
                Next
                X = 62

                _FromDate = _ToDate.AddDays(+1)
            Next


            rh = (X) * 12.75
            rh = rh + 25.5
            rh = rh + (15 * 5)
            rh = rh + 33.75
            xlCharts = worksheet2.ChartObjects
            X_ChartH = (X + 28) * 10
            myChart = xlCharts.Add(7, rh, 905, 300)

            chartPage = myChart.Chart
            'chartPage.ChartStyle = "Style 6"
            ' chartRange = worksheet1.Range(worksheet1.Cells("A8", "E" & (X - 1)), worksheet1.Cells("A9", "A" & (X - 1)))


            ' chartPage.ChartType = XlChartType.xl3DColumnClustered
            '  chartPage.ChartStyle = 6
            chartRange = worksheet2.Range("B61", "H65")
            chartPage.SetSourceData(Source:=chartRange)
            '  chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xl3DColumnStacked100
            chartPage.ChartType = XlChartType.xlLineMarkersStacked
            'chartPage.fill.color = Color.Black

            worksheet2.ChartObjects.select()
            '  workbook.ribbon.tabs(0).select()
            'xlCharts.ChartStyle = 4


            ' workbook.ActiveChart.ChartStyle = 8
            chartPage.HasTitle = True
            'Dim X1 As Integer, Y1 As Integer
            'X1 = 628
            'Y1 = 146
            'SetCursorPos(X1, Y1)
            'Dim Layout As Integer
            'Dim ChartType As Object

            'ChartType = XlChartType.xl3DColumnClustered
            chartPage.ApplyLayout(5)
            ' chartPage.ChartStyle = 8
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '         msoElementChartTitleCenteredOverlay)
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '                   msoElementPrimaryCategoryAxisTitleHorizontal)
            'chartPage.SetElement(Microsoft.Office.Core.MsoChartElementType. _
            '                   msoElementPrimaryValueAxisTitleRotated)
            chartPage.ChartTitle.Text = ("Unit Wise")
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Public Function weekNumber(ByVal d As Date) As Integer
        weekNumber = DatePart(DateInterval.WeekOfYear, d, FirstDayOfWeek.Monday, FirstWeekOfYear.System)

    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Create_File()
    End Sub
End Class