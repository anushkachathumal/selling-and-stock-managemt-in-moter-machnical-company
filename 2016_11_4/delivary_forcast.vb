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


Module delivary_forcast
    Function Create_ReportDFCT(ByVal STRMONTH As Date, ByVal _To As Date)
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet
        Dim tblDye As DataSet
        Dim _PromDelQty(4) As Double
        Dim _PromDelValue(4) As Double

        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        Dim _FirstChr As Integer
        Dim _Possible_Date As Date
        Dim _Last As Integer

        Try
            Dim exc As New Application

            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

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
            Dim _Fail_Batch As Integer

            _Fail_Batch = 0
            '  Try
            '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
            ' workbooks.Application.Sheets.Add()
            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 11
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 20
            worksheet2.Columns("B").ColumnWidth = 10
            worksheet2.Columns("C").ColumnWidth = 10
            worksheet2.Columns("D").ColumnWidth = 10
            worksheet2.Columns("E").ColumnWidth = 10
            worksheet2.Columns("F").ColumnWidth = 10
            worksheet2.Columns("G").ColumnWidth = 10
            worksheet2.Columns("H").ColumnWidth = 10
            worksheet2.Columns("I").ColumnWidth = 10
            worksheet2.Columns("L").ColumnWidth = 10
            worksheet2.Columns("M").ColumnWidth = 10
            worksheet2.Columns("N").ColumnWidth = 10
            worksheet2.Columns("O").ColumnWidth = 8
            worksheet2.Columns("P").ColumnWidth = 8
            worksheet2.Columns("Q").ColumnWidth = 8
            worksheet2.Columns("R").ColumnWidth = 8
            worksheet2.Columns("S").ColumnWidth = 8
            worksheet2.Columns("T").ColumnWidth = 10
            worksheet2.Columns("U").ColumnWidth = 10
            worksheet2.Columns("V").ColumnWidth = 10
            worksheet2.Columns("W").ColumnWidth = 10
            worksheet2.Columns("X").ColumnWidth = 10
            worksheet2.Columns("Y").ColumnWidth = 10
            worksheet2.Columns("Z").ColumnWidth = 10
            worksheet2.Columns("AA").ColumnWidth = 8
            worksheet2.Columns("AB").ColumnWidth = 8
            worksheet2.Columns("AC").ColumnWidth = 8
            worksheet2.Columns("AD").ColumnWidth = 8
            worksheet2.Columns("AE").ColumnWidth = 8
            worksheet2.Columns("AF").ColumnWidth = 8
            worksheet2.Columns("AG").ColumnWidth = 14
            worksheet2.Columns("AH").ColumnWidth = 12
            worksheet2.Columns("AI").ColumnWidth = 12
            worksheet2.Columns("AJ").ColumnWidth = 12
            worksheet2.Columns("AK").ColumnWidth = 12
            worksheet2.Columns("AL").ColumnWidth = 8
            worksheet2.Columns("AM").ColumnWidth = 8
            worksheet2.Columns("AN").ColumnWidth = 8

            worksheet2.Cells(1, 1) = "Forecasted sales for " & MonthName(Month(STRMONTH))
            worksheet2.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Range("A1:M1").Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A2:M2").Interior.Color = RGB(197, 217, 241)
            worksheet2.Rows(1).Font.size = 13
            worksheet2.Rows(1).rowheight = 55
            worksheet2.Rows(1).Font.name = "Times New Roman"
            worksheet2.Rows(1).Font.BOLD = True

            worksheet2.Range("A1:M1").MergeCells = True
            worksheet2.Range("A1:M1").VerticalAlignment = XlVAlign.xlVAlignCenter

            Dim _Chr As Integer

            Dim X As Integer

            X = 3
            worksheet2.Rows(X).rowheight = 18
            worksheet2.Cells(X, 1) = "Business Unit"
            worksheet2.Rows(X).Font.size = 9
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True
            worksheet2.Range("A3:A4").MergeCells = True
            worksheet2.Range("A3:A4").VerticalAlignment = XlVAlign.xlVAlignCenter

       
            worksheet2.Cells(X, 2) = "TOTAL"
            worksheet2.Range("B3:E3").MergeCells = True
            worksheet2.Range("B3:E3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 6) = "IN HOUSE"
            worksheet2.Range("F3:I3").MergeCells = True
            worksheet2.Range("F3:I3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 10) = "OUT SOURCE"
            worksheet2.Range("J3:M3").MergeCells = True
            worksheet2.Range("J3:M3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

           

            worksheet2.Cells(1, 14) = "Delivered"
            worksheet2.Cells(1, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Range("N1:Q1").Interior.Color = RGB(242, 220, 219)
            worksheet2.Range("N2:Q2").Interior.Color = RGB(242, 220, 219)
            worksheet2.Range("N1:Q1").MergeCells = True
            worksheet2.Range("N1:Q1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 14) = "IN HOUSE"
            worksheet2.Range("N3:O3").MergeCells = True
            worksheet2.Range("N3:O3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 16) = "OUT SOURCE"
            worksheet2.Range("P3:Q3").MergeCells = True
            worksheet2.Range("P3:Q3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(1, 18) = "Returns"
            worksheet2.Cells(1, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("R1:S1").Interior.Color = RGB(242, 220, 219)
            worksheet2.Range("R2:S2").Interior.Color = RGB(242, 220, 219)
            worksheet2.Range("R1:S1").MergeCells = True
            worksheet2.Range("R1:S1").VerticalAlignment = XlVAlign.xlVAlignCenter

          

            ' X = X + 1
            worksheet2.Rows(X + 1).Font.size = 9
            worksheet2.Rows(X + 1).Font.name = "Times New Roman"
            worksheet2.Rows(X + 1).Font.BOLD = True

            worksheet2.Cells(X + 1, 2) = "ASP(M)"
            worksheet2.Cells(X + 1, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A4:A4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 3) = "ASP(kg)"
            worksheet2.Cells(X + 1, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("b4:b4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 4) = "Qty(m)"
            worksheet2.Cells(X + 1, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("c3:c3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 5) = "Value($)"
            worksheet2.Cells(X + 1, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("d3:d3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("E3:E3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 6) = "ASP(M)"
            worksheet2.Cells(X + 1, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("f4:f4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 7) = "ASP(kg)"
            worksheet2.Cells(X + 1, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("G4:G4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 8) = "Qty(m)"
            worksheet2.Cells(X + 1, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("H4:H4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 9) = "Value($)"
            worksheet2.Cells(X + 1, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("I4:I4").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("I4:I4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 10) = "ASP(M)"
            worksheet2.Cells(X + 1, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("J4:J4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 11) = "ASP(kg)"
            worksheet2.Cells(X + 1, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("K4:K4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 12) = "Qty(m)"
            worksheet2.Cells(X + 1, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("L4:L4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 13) = "Value($)"
            worksheet2.Cells(X + 1, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("M4:M4").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("M4:M4").VerticalAlignment = XlVAlign.xlVAlignCenter
            _Chr = 97
            For i = 1 To 13
                worksheet2.Range(Chr(_Chr) & "3", Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "3", Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "3", Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & X + 1 & ":" & Chr(_Chr) & X - 1).Interior.Color = RGB(197, 217, 241)
                _Chr = _Chr + 1
            Next

          
            worksheet2.Cells(X + 1, 14) = "Qty(m)"
            worksheet2.Cells(X + 1, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("N4:N4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 15) = "Value($)"
            worksheet2.Cells(X + 1, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("O3:O3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 16) = "Qty(m)"
            worksheet2.Cells(X + 1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("P4:P4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 17) = "Value($)"
            worksheet2.Cells(X + 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("Q3:Q3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 18) = "Qty(m)"
            worksheet2.Cells(X + 1, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("R4:R4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X + 1, 19) = "Value($)"
            worksheet2.Cells(X + 1, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("S3:S3").VerticalAlignment = XlVAlign.xlVAlignCenter

            '_Chr = 102
            'For i = 6 To 7
            '    worksheet2.Range(Chr(_Chr) & "2", Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            '    _Chr = _Chr + 1
            'Next
            _Chr = 110
            For i = 14 To 19
                worksheet2.Range(Chr(_Chr) & "3", Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "3", Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "3", Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(242, 220, 219)
                worksheet2.Range(Chr(_Chr) & X + 1 & ":" & Chr(_Chr) & X + 1).Interior.Color = RGB(242, 220, 219)

                _Chr = _Chr + 1
            Next
            ' _Chr = _Chr - 1
            worksheet2.Range("M1,m1").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("M2,m2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet2.Range("Q1,Q1").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("Q2,Q2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet2.Range("S1,S1").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("S2,S2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet2.Range("R3:R4").MergeCells = True
            worksheet2.Range("R3:R4").VerticalAlignment = XlVAlign.xlVAlignCenter
            ' worksheet2.Cells(3, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Range("S3:S4").MergeCells = True
            worksheet2.Range("S3:S4").VerticalAlignment = XlVAlign.xlVAlignCenter
            '===========================================================================================
            worksheet2.Cells(1, 20) = "Balance To be delivered"
            worksheet2.Cells(1, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Range("T1:Z1").Interior.Color = RGB(216, 228, 188)
            worksheet2.Range("T2:Z2").Interior.Color = RGB(216, 228, 188)
            worksheet2.Range("T1:Z1").MergeCells = True
            worksheet2.Range("T1:Z1").VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 104

            X = 2

            worksheet2.Rows(X).Font.size = 9
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            worksheet2.Cells(X, 20) = MonthName(Month(STRMONTH) - 1)
            worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ' worksheet2.Range(Chr(_Chr) & "2:" & Chr(_Chr) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("T2:U2").MergeCells = True
            worksheet2.Range("T2:U2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("T2:U2").Interior.Color = RGB(216, 228, 188)

            _Chr = 106
            worksheet2.Cells(X, 22) = MonthName(Month(STRMONTH))
            worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ' worksheet2.Range(Chr(_Chr) & "2:" & Chr(_Chr) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("V2:W2").MergeCells = True
            worksheet2.Range("V2:W2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("V2:W2").Interior.Color = RGB(216, 228, 188)

            _Chr = 108
            worksheet2.Cells(X, 24) = MonthName(Month(STRMONTH) + 1)
            worksheet2.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ' worksheet2.Range(Chr(_Chr) & "2:" & Chr(_Chr) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("X2:Z2").MergeCells = True
            worksheet2.Range("X2:Z2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("X2:Z2").Interior.Color = RGB(216, 228, 188)
           

            _Chr = 116
            For i = 20 To 26
                worksheet2.Range(Chr(_Chr) & "2", Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "2", Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "2", Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                '  worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(242, 220, 219)
                _Chr = _Chr + 1
            Next
            _Chr = _Chr - 1
            worksheet2.Range("z1,z1").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("z2,z2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("z3,z3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("z4,z4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            ' worksheet2.Rows(X).rowheight = 14

            X = 3
            worksheet2.Cells(X, 20) = "Qty(m)"
            worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("t3:t3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 21) = "Value($)"
            worksheet2.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("u3:u3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 22) = "Qty(m)"
            worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("v3:v3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 23) = "Value($)"
            worksheet2.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("w3:w3").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 24) = "1st Week of " & MonthName(Month(STRMONTH) + 1)
            worksheet2.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("x3:x3").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 25) = "2nd Week of " & MonthName(Month(STRMONTH) + 1)
            worksheet2.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("y3:y3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 26) = "Value$ "
            worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("z3:z3").VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 116
            For i = 20 To 26
                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & "3", Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(216, 228, 188)
                worksheet2.Range(Chr(_Chr) & X + 1 & ":" & Chr(_Chr) & X + 1).Interior.Color = RGB(216, 228, 188)

                worksheet2.Range(Chr(_Chr) & "3:" & Chr(_Chr) & "4").MergeCells = True
                worksheet2.Range(Chr(_Chr) & "3:" & Chr(_Chr) & "4").VerticalAlignment = XlVAlign.xlVAlignCenter

                _Chr = _Chr + 1
            Next
            '_Chr = _Chr - 1
            'worksheet2.Range(Chr(_Chr) & "2", Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            '----------------------------------------------------------------------------------------------------------
            '----------------------------------------------------------------------------------------------------------
            worksheet2.Cells(1, 27) = "Promised to deliver"
            worksheet2.Cells(1, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Range("AA1:AF1").Interior.Color = RGB(183, 222, 232)
            worksheet2.Range("AA2:AF2").Interior.Color = RGB(183, 222, 232)
            worksheet2.Range("AA1:AF1").MergeCells = True
            worksheet2.Range("AA1:AF1").VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97

            X = 2
            worksheet2.Cells(X, 27) = "Deliver from Stock" 'MonthName(Month(STRMONTH) - 1)
            worksheet2.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & Chr(_Chr) & "2:A" & Chr(_Chr) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("AA2:AB2").MergeCells = True
            worksheet2.Range("AA2:AB2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("AA2:AB2").Interior.Color = RGB(183, 222, 232)

            _Chr = _Chr + 1
            worksheet2.Cells(X, 29) = "Deliver from WIP" ' MonthName(Month(STRMONTH))
            worksheet2.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & Chr(_Chr) & "2:A" & Chr(_Chr) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("AC2:AD2").MergeCells = True
            worksheet2.Range("AC2:AD2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("AC2:AD2").Interior.Color = RGB(183, 222, 232)

            _Chr = _Chr + 1
            worksheet2.Cells(X, 31) = "ToTal Delivery" 'MonthName(Month(STRMONTH) + 1)
            worksheet2.Cells(X, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & Chr(_Chr) & "2:A" & Chr(_Chr) & "2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("AE2:AF2").MergeCells = True
            worksheet2.Range("AE2:AF2").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Range("AE2:AF2").Interior.Color = RGB(183, 222, 232)
            worksheet2.Rows(X).Font.size = 9
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            _Chr = 97
            For i = 123 To 128
                worksheet2.Range("A" & Chr(_Chr) & "2", "A" & Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "2", "A" & Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "2", "A" & Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                '  worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(242, 220, 219)
                _Chr = _Chr + 1
            Next
            _Chr = _Chr - 1
            worksheet2.Range("A" & Chr(_Chr) & "1", "A" & Chr(_Chr) & "1").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Rows(X).rowheight = 14

            X = 3
            worksheet2.Cells(X, 27) = "Qty(m)"
            worksheet2.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AA3:AA3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 28) = "Value($)"
            worksheet2.Cells(X, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AB3:AB3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 29) = "Qty(m)"
            worksheet2.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AC3:AC3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 30) = "Value($)"
            worksheet2.Cells(X, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AD3:AD3").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 31) = "Qty(m)"
            worksheet2.Cells(X, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AE3:AE3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 32) = "Value($)"
            worksheet2.Cells(X, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AF3:AF3").VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For i = 123 To 128
                worksheet2.Range("A" & Chr(_Chr) & "3", "A" & Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "3", "A" & Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "3", "A" & Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "3", "A" & Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

                worksheet2.Range("A" & Chr(_Chr) & "4", "A" & Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "4", "A" & Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "4", "A" & Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "4", "A" & Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

                worksheet2.Range("A" & Chr(_Chr) & X & ":" & "A" & Chr(_Chr) & X).Interior.Color = RGB(183, 222, 232)


                worksheet2.Range("A" & Chr(_Chr) & "3:" & "A" & Chr(_Chr) & "4").MergeCells = True
                worksheet2.Range("A" & Chr(_Chr) & "3:" & "A" & Chr(_Chr) & "4").VerticalAlignment = XlVAlign.xlVAlignCenter

                _Chr = _Chr + 1
            Next
            '_Chr = _Chr - 1
            'worksheet2.Range(Chr(_Chr) & "2", Chr(_Chr) & "2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            '===================================================================================================================
            worksheet2.Cells(1, 33) = "Total Possible Qty to be delivered"
            worksheet2.Cells(1, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Range("AG1:AG1").Interior.Color = RGB(183, 222, 232)
            worksheet2.Range("AG2:AG2").Interior.Color = RGB(183, 222, 232)
            worksheet2.Range("AG1:AG4").MergeCells = True
            worksheet2.Range("AG1:AG1").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(1, 33).WrapText = True

            worksheet2.Range("AG3", "AG3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AG2", "AG2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AG1", "AG1").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AG3", "AG3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AG4", "AG4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AG4", "AG4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            '--------------------------------------------------------------------------------------------------------
            worksheet2.Cells(1, 34) = "FG Stocks"
            worksheet2.Cells(1, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Range("AH1:AH1").Interior.Color = RGB(252, 213, 180)
            worksheet2.Range("AH2:AK2").Interior.Color = RGB(252, 213, 180)
            worksheet2.Range("AH1:AK1").MergeCells = True
            worksheet2.Range("AH2:AK2").MergeCells = True
            worksheet2.Range("AH1:AH1").VerticalAlignment = XlVAlign.xlVAlignCenter

            X = 3
            If Month(STRMONTH) = 1 Then
                worksheet2.Cells(X, 34) = "January"
            Else
                worksheet2.Cells(X, 34) = MonthName(Month(STRMONTH) - 1)
            End If
            worksheet2.Cells(X, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AH3:AH3").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 35) = MonthName(Month(STRMONTH))
            worksheet2.Cells(X, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AI3:AI3").VerticalAlignment = XlVAlign.xlVAlignCenter

            If Month(STRMONTH) = 12 Then
                worksheet2.Cells(X, 36) = "1st Wk of January"
            Else
                worksheet2.Cells(X, 36) = "1st Wk of " & MonthName(Month(STRMONTH) + 1)
            End If
            worksheet2.Cells(X, 36).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AJ3:AJ3").VerticalAlignment = XlVAlign.xlVAlignCenter

            If Month(STRMONTH) = 12 Then
                worksheet2.Cells(X, 37) = "2nd Wk of January"
            Else
                worksheet2.Cells(X, 37) = "2nd Wk of " & MonthName(Month(STRMONTH) + 1)
            End If
            worksheet2.Cells(X, 37).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AJ3:AJ3").VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 104
            For i = 34 To 37
                worksheet2.Range("A" & Chr(_Chr) & "3", "A" & Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "3", "A" & Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "3", "A" & Chr(_Chr) & "3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                worksheet2.Range("A" & Chr(_Chr) & "4", "A" & Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "4", "A" & Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range("A" & Chr(_Chr) & "4", "A" & Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


                worksheet2.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).Interior.Color = RGB(252, 213, 180)

                worksheet2.Range("A" & Chr(_Chr) & "3:" & "A" & Chr(_Chr) & "4").MergeCells = True
                worksheet2.Range("A" & Chr(_Chr) & "3:" & "A" & Chr(_Chr) & "4").VerticalAlignment = XlVAlign.xlVAlignCenter

                _Chr = _Chr + 1
            Next
            _Chr = _Chr - 1
            worksheet2.Range("AK2", "AK2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            '=========================================================================================
            worksheet2.Cells(1, 38) = "WIP"
            worksheet2.Cells(1, 38).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Range("al1:AL1").Interior.Color = RGB(221, 217, 196)
            worksheet2.Range("AL2:AN2").Interior.Color = RGB(221, 217, 196)
            worksheet2.Range("AL3:AN3").Interior.Color = RGB(221, 217, 196)
            worksheet2.Range("AL1:AN1").MergeCells = True
            worksheet2.Range("AL1:AL1").VerticalAlignment = XlVAlign.xlVAlignCenter

            '----------------------------------------------------------------------------------------
            X = 3

            worksheet2.Cells(X, 38) = "Exam"

            worksheet2.Cells(X, 38).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AL3:AL3").VerticalAlignment = XlVAlign.xlVAlignCenter


            'worksheet2.Cells(X, 26) = MonthName(Month(STRMONTH))
            'worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet2.Range("z3:z3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 39) = "Finishing"

            worksheet2.Cells(X, 39).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AM3:AM3").VerticalAlignment = XlVAlign.xlVAlignCenter
            'worksheet2.Cells(X, 27) = MonthName(Month(STRMONTH))
            'worksheet2.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet2.Range("AA3:AA3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 40) = "Dyeing"

            worksheet2.Cells(X, 40).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AN3:AN3").VerticalAlignment = XlVAlign.xlVAlignCenter
            'worksheet2.Cells(X, 28) = MonthName(Month(STRMONTH))
            'worksheet2.Cells(X, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet2.Range("AB3:AB3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Range("AL3", "AL3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AL3", "AL3").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AL3", "AL3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AL3", "AL3").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AL3", "AL3").Interior.Color = RGB(221, 217, 196)

            worksheet2.Range("AL2", "AL2").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AL1", "AL1").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

            worksheet2.Range("AM3", "AM3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AM3", "AM3").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AM3", "AM3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AM3", "AM3").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AM3", "AM3").Interior.Color = RGB(221, 217, 196)

            worksheet2.Range("AN3", "AN3").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AN3", "AN3").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AN3", "AN3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AN3", "AN3").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AN3", "AN3").Interior.Color = RGB(221, 217, 196)


            'worksheet2.Range("AB2", "AB2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            'worksheet2.Range("AB2", "AB2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            'worksheet2.Range("AB2", "AB2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            ' worksheet2.Range("AB2", "AB2").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous


            worksheet2.Range("AN1", "AN1").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AN2", "AN2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AN3", "AN3").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AN4", "AN4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet2.Range("AL4", "AL4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AM4", "AM4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet2.Range("AN4", "AN4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            '  worksheet2.Range("AB1", "AB1").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            '==============================================================================================================

            worksheet2.Range("AL3:AL4").MergeCells = True
            worksheet2.Range("AL3:AL4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Range("AM3:AM4").MergeCells = True
            worksheet2.Range("AM3:AM4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Range("AN3:AN4").MergeCells = True
            worksheet2.Range("AN3:AN4").VerticalAlignment = XlVAlign.xlVAlignCenter
            X = 5
            worksheet2.Rows(X).rowheight = 18
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            worksheet2.Cells(X, 1) = "M&S"
            worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet2.Range("A" & X & ":A" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For i = 1 To 40
                If i >= 27 Then
                    If i = 27 Then
                        _Chr = 97
                    End If
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                Else
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                End If
            Next
            X = 6
            worksheet2.Rows(X).rowheight = 18
            _Chr = 97
            For i = 1 To 40
                If i >= 27 Then
                    If i = 27 Then
                        _Chr = 97
                    End If
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                Else
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                End If
            Next

            X = 7
            worksheet2.Rows(X).rowheight = 18
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            worksheet2.Cells(X, 1) = "LMTD-BRANDS"
            worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet2.Range("A" & X & ":A" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For i = 1 To 40
                If i >= 27 Then
                    If i = 27 Then
                        _Chr = 97
                    End If
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Interior.Color = RGB(132, 151, 176)
                    _Chr = _Chr + 1
                Else
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(132, 151, 176)
                    _Chr = _Chr + 1
                End If
            Next
            X = 8
            worksheet2.Rows(X).rowheight = 18
            _Chr = 97

            For i = 1 To 40
                If i >= 27 Then
                    If i = 27 Then
                        _Chr = 97
                    End If
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    '  worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                Else
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                End If
            Next

            X = 9
            worksheet2.Rows(X).rowheight = 18
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            worksheet2.Cells(X, 1) = "INTIMISIMI/ TEZENEIS"
            worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet2.Range("A" & X & ":A" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For i = 1 To 40
                If i >= 27 Then
                    If i = 27 Then
                        _Chr = 97
                    End If
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Interior.Color = RGB(189, 215, 238)
                    _Chr = _Chr + 1
                Else
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(189, 215, 238)
                    _Chr = _Chr + 1
                End If
            Next

            X = 10
            worksheet2.Rows(X).rowheight = 18
            _Chr = 97
            For i = 1 To 40
                If i >= 27 Then
                    If i = 27 Then
                        _Chr = 97
                    End If
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                Else
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                End If
            Next


            X = 11
            worksheet2.Rows(X).rowheight = 18
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            worksheet2.Cells(X, 1) = "EMG- BRANDS"
            worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet2.Range("A" & X & ":A" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For i = 1 To 40
                If i >= 27 Then
                    If i = 27 Then
                        _Chr = 97
                    End If
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Interior.Color = RGB(248, 203, 173)
                    _Chr = _Chr + 1
                Else
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(248, 203, 173)
                    _Chr = _Chr + 1
                End If
            Next

            X = 12
            worksheet2.Rows(X).rowheight = 18
            _Chr = 97
            For i = 1 To 40
                If i >= 27 Then
                    If i = 27 Then
                        _Chr = 97
                    End If
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    '  worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                Else
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(155, 187, 89)
                    _Chr = _Chr + 1
                End If
            Next
            '-----------------------------------------------------------------
            'DELIVARY QTY
            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from M06Delivary_Qty inner  join M13Biz_Unit on M13Merchant=M06Merchant where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='1'  group by M13Department"
            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from View_Delivary  where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='1' and M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                worksheet2.Cells(5, 14) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                worksheet2.Cells(5, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("n5:n5").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(5, 14)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(5, 15) = dsUser.Tables(0).Rows(0)("TotValue")
                worksheet2.Cells(5, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("o5:o5").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(5, 15)
                range1.NumberFormat = "0.00"

            End If

            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from View_Delivary  where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='1' and M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                worksheet2.Cells(5, 16) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                worksheet2.Cells(5, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("p5:p5").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(5, 16)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(5, 17) = dsUser.Tables(0).Rows(0)("TotValue")
                worksheet2.Cells(5, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("q5:q5").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(5, 17)
                range1.NumberFormat = "0.00"

            End If

            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from View_Delivary  where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='2' and M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                worksheet2.Cells(7, 14) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                worksheet2.Cells(7, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("n7:n7").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(7, 14)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(7, 15) = dsUser.Tables(0).Rows(0)("TotValue")
                worksheet2.Cells(7, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("o7:o7").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(7, 15)
                range1.NumberFormat = "0.00"

            End If

            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from View_Delivary  where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='2' and M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                worksheet2.Cells(7, 16) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                worksheet2.Cells(7, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("p7:p7").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(7, 16)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(7, 17) = dsUser.Tables(0).Rows(0)("TotValue")
                worksheet2.Cells(7, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("q7:q7").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(7, 17)
                range1.NumberFormat = "0.00"

            End If


            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from View_Delivary  where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='3' and M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                worksheet2.Cells(9, 14) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                worksheet2.Cells(9, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("n9:n9").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(9, 14)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(9, 15) = dsUser.Tables(0).Rows(0)("TotValue")
                worksheet2.Cells(9, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("o9:o9").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(9, 15)
                range1.NumberFormat = "0.00"

            End If

            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from View_Delivary  where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='3' and M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                worksheet2.Cells(9, 16) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                worksheet2.Cells(9, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("p9:p9").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(9, 16)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(9, 17) = dsUser.Tables(0).Rows(0)("TotValue")
                worksheet2.Cells(9, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("q9:q9").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(9, 17)
                range1.NumberFormat = "0.00"

            End If

            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from View_Delivary  where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='4' and M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                worksheet2.Cells(11, 14) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                worksheet2.Cells(11, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("n11:n11").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(11, 14)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(11, 15) = dsUser.Tables(0).Rows(0)("TotValue")
                worksheet2.Cells(11, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("o11:o11").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(11, 15)
                range1.NumberFormat = "0.00"

            End If

            SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr,sum(M06D_Qty_Mtr*M06Unit_Mtr) as TotValue from View_Delivary  where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='4' and M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                worksheet2.Cells(11, 16) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")
                worksheet2.Cells(11, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("p11:p11").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(11, 16)
                range1.NumberFormat = "0.00"

                worksheet2.Cells(11, 17) = dsUser.Tables(0).Rows(0)("TotValue")
                worksheet2.Cells(11, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Range("q11:q5").VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(11, 17)
                range1.NumberFormat = "0.00"

            End If


            X = 13
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.bold = True
            worksheet2.Range("N" & (X)).Formula = "=SUM(N5:N11)"
            worksheet2.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 14)
            range1.NumberFormat = "0.00"
            worksheet2.Range("N" & X, "N" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("O" & (X)).Formula = "=SUM(O5:O11)"
            worksheet2.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 15)
            range1.NumberFormat = "0.00"
            worksheet2.Range("O" & X, "O" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("P" & (X)).Formula = "=SUM(P5:P11)"
            worksheet2.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 16)
            range1.NumberFormat = "0.00"
            worksheet2.Range("P" & X, "P" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("Q" & (X)).Formula = "=SUM(Q5:Q11)"
            worksheet2.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 17)
            range1.NumberFormat = "0.00"
            worksheet2.Range("Q" & X, "Q" & X).Interior.Color = RGB(141, 180, 227)

            '-------------------------------------------------------------------------------
            'TO BE DELIVERY
            Dim _TotalStock(4) As Double
            _TotalStock(0) = 0
            _TotalStock(1) = 0
            _TotalStock(2) = 0
            _TotalStock(3) = 0


            Dim L_DateofMonth As Integer
            Dim _From As Date
            Dim EndDate As Date
            Dim _ToDate As Date

            _From = CDate(Today)
            _From = Month(_From) & "/1/" & Year(_From)
            If Month(_From) = 1 Then
                _From = "12/1/" & Year(_From) - 1
            Else
                _From = Month(_From) - 1 & "/1/" & Year(_From)
            End If
            EndDate = _From.AddDays(DateTime.DaysInMonth(_From.Year, _From.Month) - 1)
            L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)
            L_DateofMonth = L_DateofMonth - 1
            _ToDate = _From.AddDays(+L_DateofMonth)
            Dim _Del As String
            X = 5
            SQL = "select * from M14Retailer order by M14Code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            Y = 0
            _Chr = 104
            Dim z As Integer
            Dim _String As String
            ' _TotalStock = 0


            z = 6
            For Each DTRow3 As DataRow In T01.Tables(0).Rows

                SQL = "SELECT SUM(M08Qty_Mtr) AS M08Qty_Mtr FROM M08Stock INNER JOIN M07TobeDelivered ON M07Sales_Order=M08Sales_Order AND M07Line_Item=M08Line_Item inner join M13Biz_Unit on M13Merchant=M08Merchant  where M08Location in ('2060','2059') and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' GROUP BY M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    _TotalStock(Y) = dsUser.Tables(0).Rows(0)("M08Qty_Mtr")
                End If

                SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr,sum(M07Qty_Mtr*M07Unit_PriceMTR) as TotalValue from M07TobeDelivered  inner join M13Biz_Unit on M13Merchant=M07Merchant where  M07Date <= '" & _ToDate & "' and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then


                    '  _BTQty01 = _BTQty01 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")

                    worksheet2.Cells(X, 20) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("T" & X & ":T" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 20)
                    range1.NumberFormat = "0.00"

                    _Chr = _Chr + 1
                    worksheet2.Cells(X, 21) = dsUser.Tables(0).Rows(0)("TotalValue")
                    worksheet2.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("U" & X & ":U" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 21)
                    range1.NumberFormat = "0.00"

                    If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") < _TotalStock(Y) Then
                        _TotalStock(Y) = _TotalStock(Y) - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        worksheet2.Cells(X, 34) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        worksheet2.Cells(X, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AH" & X & ":AH" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 34)
                        range1.NumberFormat = "0.00"

                    ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") >= _TotalStock(Y) Then
                        worksheet2.Cells(X, 34) = _TotalStock(Y)
                        worksheet2.Cells(X, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AH" & X & ":AH" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 34)
                        range1.NumberFormat = "0.00"

                        _TotalStock(Y) = 0
                    End If
                    _Chr = _Chr + 1
                Else
                    _Chr = _Chr + 2
                End If

                X = X + 2

                Y = Y + 1
            Next

            worksheet2.Range("T" & (X)).Formula = "=SUM(T5:T11)"
            worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 20)
            range1.NumberFormat = "0.00"
            worksheet2.Range("T" & X, "T" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("U" & (X)).Formula = "=SUM(U5:U11)"
            worksheet2.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 21)
            range1.NumberFormat = "0.00"
            worksheet2.Range("U" & X, "U" & X).Interior.Color = RGB(141, 180, 227)


            _From = CDate(Today)
            _From = Month(_From) & "/1/" & Year(_From)
            'If Month(_From) = 1 Then
            '    _From = "12/1/" & Year(_From) - 1
            'Else
            '    _From = Month(_From) - 1 & "/1/" & Year(_From)
            'End If
            EndDate = _From.AddDays(DateTime.DaysInMonth(_From.Year, _From.Month) - 1)
            L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)
            L_DateofMonth = L_DateofMonth - 1
            _ToDate = _From.AddDays(+L_DateofMonth)

            X = 5
            SQL = "select * from M14Retailer order by M14Code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            Y = 0
            _Chr = 106

            z = 6
            For Each DTRow3 As DataRow In T01.Tables(0).Rows

                'SQL = "SELECT SUM(M08Qty_Mtr) AS M08Qty_Mtr FROM M08Stock INNER JOIN M07TobeDelivered ON M07Sales_Order=M08Sales_Order AND M07Line_Item=M08Line_Item inner join M13Biz_Unit on M13Merchant=M08Merchant  where M08Location in ('2060','2059') and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' GROUP BY M13Department"
                'dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                'If isValidDataset(dsUser) Then
                '    _TotalStock = dsUser.Tables(0).Rows(0)("M08Qty_Mtr")
                'End If


                SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr,sum(M07Qty_Mtr*M07Unit_PriceMTR) as TotalValue from M07TobeDelivered  inner join M13Biz_Unit on M13Merchant=M07Merchant where  M07Date between '" & _From & "' and '" & _To & "' and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then


                    '  _BTQty01 = _BTQty01 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")

                    worksheet2.Cells(X, 22) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("V" & X & ":V" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 22)
                    range1.NumberFormat = "0.00"

                    _Chr = _Chr + 1
                    worksheet2.Cells(X, 23) = dsUser.Tables(0).Rows(0)("TotalValue")
                    worksheet2.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("W" & X & ":W" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 23)
                    range1.NumberFormat = "0.00"

                    If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") < _TotalStock(Y) Then
                        _TotalStock(Y) = _TotalStock(Y) - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        worksheet2.Cells(X, 35) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        worksheet2.Cells(X, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AI" & X & ":AI" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 35)
                        range1.NumberFormat = "0.00"

                    ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") >= _TotalStock(Y) Then
                        worksheet2.Cells(X, 35) = _TotalStock(Y)
                        worksheet2.Cells(X, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AI" & X & ":AI" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 35)
                        range1.NumberFormat = "0.00"

                        _TotalStock(Y) = 0
                    End If

                    _Chr = _Chr + 1
                Else
                    _Chr = _Chr + 2
                End If

                X = X + 2

                Y = Y + 1
            Next

            worksheet2.Range("V" & (X)).Formula = "=SUM(V5:V11)"
            worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 22)
            range1.NumberFormat = "0.00"
            worksheet2.Range("V" & X, "V" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("W" & (X)).Formula = "=SUM(W5:W11)"
            worksheet2.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 23)
            range1.NumberFormat = "0.00"
            worksheet2.Range("W" & X, "W" & X).Interior.Color = RGB(141, 180, 227)

            _From = Today
            If Month(_From) = 12 Then
                _From = "1/1/" & Year(_From) + 1
            Else
                _From = Month(Today) + 1 & "/1/" & Year(Today)
            End If

            EndDate = _From.AddDays(DateTime.DaysInMonth(_From.Year, _From.Month) - 1)

            Dim _wkFIRST As Date
            Dim _wkTo As Date

            Dim _wkFIRST1 As Date
            Dim _wkTo1 As Date
            Dim _AppVALUE As Double

            _wkFIRST = Month(_From) & "/1/" & Year(_From)

            Dim thisCulture = Globalization.CultureInfo.CurrentCulture
            Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_wkFIRST)
            Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)

            If dayName = "Sunday" Then
                _wkTo = CDate(_wkFIRST).AddDays(+6)
            ElseIf dayName = "Tuesday" Then
                _wkTo = CDate(_wkFIRST).AddDays(+5)
            ElseIf dayName = "Wednesday" Then
                _wkTo = CDate(_wkFIRST).AddDays(+4)
            ElseIf dayName = "Thuesday" Then
                _wkTo = CDate(_wkFIRST).AddDays(+3)
            ElseIf dayName = "Friday" Then
                _wkTo = CDate(_wkFIRST).AddDays(+2)
            ElseIf dayName = "Saturday" Then
                _wkTo = CDate(_wkFIRST).AddDays(+1)
            ElseIf dayName = "Monday" Then
                _wkTo = CDate(_wkFIRST).AddDays(+6)
            End If
            '1ST WEEK

            X = 5
            SQL = "select * from M14Retailer order by M14Code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            Y = 0
            _Chr = 108

            z = 6
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _AppVALUE = 0
                SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr,sum(M07Qty_Mtr*M07Unit_PriceMTR) as TotalValue from M07TobeDelivered  inner join M13Biz_Unit on M13Merchant=M07Merchant where  M07Date BETWEEN  '" & _wkFIRST & "' AND '" & _wkTo & "' and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then


                    '  _BTQty01 = _BTQty01 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")

                    worksheet2.Cells(X, 24) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    worksheet2.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("X" & X & ":X" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 24)
                    range1.NumberFormat = "0.00"

                    _AppVALUE = dsUser.Tables(0).Rows(0)("TotalValue")

                    If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") < _TotalStock(Y) Then
                        _TotalStock(Y) = _TotalStock(Y) - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        worksheet2.Cells(X, 36) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        worksheet2.Cells(X, 36).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AJ" & X & ":AJ" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 36)
                        range1.NumberFormat = "0.00"

                    ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") >= _TotalStock(Y) Then
                        worksheet2.Cells(X, 36) = _TotalStock(Y)
                        worksheet2.Cells(X, 36).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AJ" & X & ":AJ" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 36)
                        range1.NumberFormat = "0.00"

                        _TotalStock(Y) = 0
                    End If
                    _Chr = _Chr + 1


                Else
                    _Chr = _Chr + 2
                End If


                _wkFIRST1 = _wkTo.AddDays(+1)
                _wkTo1 = _wkFIRST1.AddDays(+6)


                SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr,sum(M07Qty_Mtr*M07Unit_PriceMTR) as TotalValue from M07TobeDelivered  inner join M13Biz_Unit on M13Merchant=M07Merchant where  M07Date BETWEEN  '" & _wkFIRST1 & "' AND '" & _wkTo1 & "' and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then


                    '  _BTQty01 = _BTQty01 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")

                    worksheet2.Cells(X, 25) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    worksheet2.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("Y" & X & ":Y" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 25)
                    range1.NumberFormat = "0"
                    _AppVALUE = _AppVALUE + dsUser.Tables(0).Rows(0)("TotalValue")

                    If dsUser.Tables(0).Rows(0)("M07Qty_Mtr") < _TotalStock(Y) Then
                        _TotalStock(Y) = _TotalStock(Y) - dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        worksheet2.Cells(X, 37) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                        worksheet2.Cells(X, 37).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AK" & X & ":AK" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 37)
                        range1.NumberFormat = "0"

                    ElseIf dsUser.Tables(0).Rows(0)("M07Qty_Mtr") >= _TotalStock(Y) Then
                        worksheet2.Cells(X, 37) = _TotalStock(Y)
                        worksheet2.Cells(X, 37).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AK" & X & ":AK" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 37)
                        range1.NumberFormat = "0.00"

                        _TotalStock(Y) = 0
                    End If
                    _Chr = _Chr + 1


                Else
                    _Chr = _Chr + 2
                End If

                worksheet2.Cells(X, 26) = _AppVALUE
                worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
                _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                worksheet2.Range("Z" & X & ":Z" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet2.Cells(X, 26)
                range1.NumberFormat = "0"


                X = X + 2

                Y = Y + 1

            Next

            worksheet2.Range("X" & (X)).Formula = "=SUM(X5:X11)"
            worksheet2.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 24)
            range1.NumberFormat = "0"
            worksheet2.Range("X" & X, "X" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("Y" & (X)).Formula = "=SUM(Y5:Y11)"
            worksheet2.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 25)
            range1.NumberFormat = "0"
            worksheet2.Range("Y" & X, "Y" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("Z" & (X)).Formula = "=SUM(Z5:Z11)"
            worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 26)
            range1.NumberFormat = "0"
            worksheet2.Range("Z" & X, "Z" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("AH" & (X)).Formula = "=SUM(AH5:AH11)"
            worksheet2.Cells(X, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 34)
            range1.NumberFormat = "0"
            worksheet2.Range("AH" & X, "AH" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("AI" & (X)).Formula = "=SUM(AI5:AI11)"
            worksheet2.Cells(X, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 35)
            range1.NumberFormat = "0"
            worksheet2.Range("AI" & X, "AI" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("AJ" & (X)).Formula = "=SUM(AJ5:AJ11)"
            worksheet2.Cells(X, 36).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 36)
            range1.NumberFormat = "0"
            worksheet2.Range("AJ" & X, "AJ" & X).Interior.Color = RGB(141, 180, 227)


            worksheet2.Range("AK" & (X)).Formula = "=SUM(AK5:AK11)"
            worksheet2.Cells(X, 37).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 37)
            range1.NumberFormat = "0"
            worksheet2.Range("AK" & X, "AK" & X).Interior.Color = RGB(141, 180, 227)


            '-------------------------------------------------------------------------------------------
            'Promised to deliver

            SQL = "SELECT * FROM T08DelayComment ORDER BY T08DATE DESC"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                _From = T01.Tables(0).Rows(0)("T08DATE")
                _ToDate = Month(_From) & "/1/" & Year(_From)
            End If
            X = 5
            SQL = "select * from M14Retailer order by M14Code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            Y = 0
            _Chr = 113

            z = 6
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant  where d_Date between '" & _ToDate & "' and '" & frmDelivary_Forcust.txtTodate.Text & "' and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' group by M08Sales_Order,M08Line_Item"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    worksheet2.Cells(X, 27) = dsUser.Tables(0).Rows(0)("Qty")
                    worksheet2.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("AA" & X & ":AA" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 27)
                    range1.NumberFormat = "0"
                End If

                SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant  where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then


                    '  _BTQty01 = _BTQty01 + dsUser.Tables(0).Rows(0)("M07Qty_Mtr")

                    _PromDelQty(Y) = dsUser.Tables(0).Rows(0)("KGQty")
                    _PromDelValue(Y) = dsUser.Tables(0).Rows(0)("KGValue")
                    worksheet2.Cells(X, 29) = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
                    worksheet2.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("AC" & X & ":AC" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 29)
                    range1.NumberFormat = "0"

                    worksheet2.Cells(X, 30) = dsUser.Tables(0).Rows(0)("TotalValue1")
                    worksheet2.Cells(X, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("AD" & X & ":AD" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 30)
                    range1.NumberFormat = "0"

                    worksheet2.Cells(X, 31) = dsUser.Tables(0).Rows(0)("T08Pro_Qty2")
                    worksheet2.Cells(X, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("AE" & X & ":AE" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 31)
                    range1.NumberFormat = "0"

                    worksheet2.Cells(X, 32) = dsUser.Tables(0).Rows(0)("TotalValue2")
                    worksheet2.Cells(X, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("AF" & X & ":AF" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 32)
                    range1.NumberFormat = "0"
                    _Chr = _Chr + 1


                Else
                    _Chr = _Chr + 2
                End If
                X = X + 2
                Y = Y + 1
            Next


            worksheet2.Range("AC" & (X)).Formula = "=SUM(AC5:AC11)"
            worksheet2.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 29)
            range1.NumberFormat = "0.00"
            worksheet2.Range("AC" & X, "AC" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("AD" & (X)).Formula = "=SUM(AD5:AD11)"
            worksheet2.Cells(X, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 30)
            range1.NumberFormat = "0.00"
            worksheet2.Range("AD" & X, "AD" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("AE" & (X)).Formula = "=SUM(AE5:AE11)"
            worksheet2.Cells(X, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 31)
            range1.NumberFormat = "0.00"
            worksheet2.Range("AE" & X, "AE" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("AF" & (X)).Formula = "=SUM(AF5:AF11)"
            worksheet2.Cells(X, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 32)
            range1.NumberFormat = "0.00"
            worksheet2.Range("AF" & X, "AF" & X).Interior.Color = RGB(141, 180, 227)
            '---------------------------------------------------------------------------------
            'WIP
            X = 5
            SQL = "select * from M14Retailer order by M14Code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            Y = 0
            _Chr = 113

            z = 6
            For Each DTRow3 As DataRow In T01.Tables(0).Rows

                SQL = "select sum(M09Qty_Mtr) as M09Qty_Mtr from M09ZPL_ORDER inner join M07TobeDelivered ON M07Material=M09Meterial INNER JOIN M13Biz_Unit ON M13Merchant=M09Merchant  where M09Oredr_Type='Exam' and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    worksheet2.Cells(X, 38) = dsUser.Tables(0).Rows(0)("M09Qty_Mtr")
                    worksheet2.Cells(X, 38).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("AL" & X & ":AL" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 38)
                    range1.NumberFormat = "0.00"
                End If

                SQL = "select sum(M09Qty_Mtr) as M09Qty_Mtr from M09ZPL_ORDER inner join M07TobeDelivered ON M07Material=M09Meterial INNER JOIN M13Biz_Unit ON M13Merchant=M09Merchant  where M09Oredr_Type='Finishing' and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    worksheet2.Cells(X, 39) = dsUser.Tables(0).Rows(0)("M09Qty_Mtr")
                    worksheet2.Cells(X, 39).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                    worksheet2.Range("AM" & X & ":AM" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    range1 = worksheet2.Cells(X, 39)
                    range1.NumberFormat = "0.00"
                End If

                If frmDelivary_Forcust.chkCus.Checked = True Then
                    SQL = "select sum(M09Qty_Mtr) as M09Qty_Mtr from M09ZPL_ORDER inner join FR_Update on M09BatchNo=Batch_No inner join M07TobeDelivered ON M07Material=M09Meterial INNER JOIN M13Biz_Unit ON M13Merchant=M09Merchant where M09Oredr_Type='Dyeing' and M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' and Dye_Pln_Date<= '" & frmDelivary_Forcust.txtDyeDate.Text & "' group by  M13Department"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        worksheet2.Cells(X, 40) = dsUser.Tables(0).Rows(0)("M09Qty_Mtr")
                        worksheet2.Cells(X, 40).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _String = ChrW(_Chr) & X & ":" & ChrW(_Chr) & X
                        worksheet2.Range("AN" & X & ":AN" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet2.Cells(X, 40)
                        range1.NumberFormat = "0.00"
                    End If

                End If


                X = X + 2
                Y = Y + 1
            Next

            worksheet2.Range("AL" & (X)).Formula = "=SUM(AL5:AL11)"
            worksheet2.Cells(X, 38).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 38)
            range1.NumberFormat = "0.00"
            worksheet2.Range("AL" & X, "AL" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("AM" & (X)).Formula = "=SUM(AM5:AM11)"
            worksheet2.Cells(X, 39).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 39)
            range1.NumberFormat = "0.00"
            worksheet2.Range("AM" & X, "AM" & X).Interior.Color = RGB(141, 180, 227)

            worksheet2.Range("AN" & (X)).Formula = "=SUM(AN5:AN11)"
            worksheet2.Cells(X, 40).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 40)
            range1.NumberFormat = "0.00"
            worksheet2.Range("AN" & X, "AN" & X).Interior.Color = RGB(141, 180, 227)

            '----------------------------------------------------------------------------------
            'TOTAL POSSIBLE QTY
            X = 5
            worksheet2.Range("AG" & (X)).Formula = "=SUM(AH5:AN5)"
            worksheet2.Cells(X, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 33)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("AG" & (X)).Formula = "=SUM(AH7:AN7)"
            worksheet2.Cells(X, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 33)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("AG" & (X)).Formula = "=SUM(AH9:AN9)"
            worksheet2.Cells(X, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 33)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("AG" & (X)).Formula = "=SUM(AH11:AN11)"
            worksheet2.Cells(X, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 33)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("AG" & (X)).Formula = "=SUM(AG5:AG11)"
            worksheet2.Cells(X, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 33)
            range1.NumberFormat = "0.00"
            worksheet2.Range("AG" & X, "AG" & X).Interior.Color = RGB(141, 180, 227)
            '----------------------------------------------------------------------------------
            'QTY
            X = 5
            worksheet2.Range("D" & (X)).Formula = "=(H5+L5)-R5"
            worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 4)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("D" & (X)).Formula = "=(H7+L7)-R7"
            worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 4)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("D" & (X)).Formula = "=(H9+L9)-R9"
            worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 4)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("D" & (X)).Formula = "=(H11+L11)-R11"
            worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 4)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("D" & (X)).Formula = "=(D5+D7+D9+D11)-R13"
            worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 4)
            range1.NumberFormat = "0.00"
            worksheet2.Range("D" & X, "D" & X).Interior.Color = RGB(141, 180, 227)


            'IN HOUSE
            Dim _PROMISE_TO_dEL As Double
            Dim _STOCKFG As Double

            _STOCKFG = 0

            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='1' AND M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
            End If

            SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant INNER JOIN M17Material_Master ON M17Code= Material where d_Date between '" & _ToDate & "' and '" & _From & "' and M13Department='1' AND M17Location='4004' group by M08Sales_Order,M08Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _STOCKFG = dsUser.Tables(0).Rows(0)("Qty")
            End If
            X = 5
            worksheet2.Range("H" & (X)).Formula = "=N5+" & _STOCKFG & "+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 8)
            range1.NumberFormat = "0.00"

            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='2' AND M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
            End If
            _STOCKFG = 0
            SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant INNER JOIN M17Material_Master ON M17Code= Material where d_Date between '" & _ToDate & "' and '" & _From & "' and M13Department='2' AND M17Location='4004' group by M08Sales_Order,M08Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _STOCKFG = dsUser.Tables(0).Rows(0)("Qty")
            End If
            worksheet2.Range("H" & (X)).Formula = "=N7+" & _STOCKFG & "+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 8)
            range1.NumberFormat = "0.00"


            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='3' AND M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
            End If

            _STOCKFG = 0
            SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant INNER JOIN M17Material_Master ON M17Code= Material where d_Date between '" & _ToDate & "' and '" & _From & "' and M13Department='3' AND M17Location='4004' group by M08Sales_Order,M08Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _STOCKFG = dsUser.Tables(0).Rows(0)("Qty")
            End If
            worksheet2.Range("H" & (X)).Formula = "=N9+" & _STOCKFG & "+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 8)
            range1.NumberFormat = "0.00"


            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='4' AND M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
            End If

            _STOCKFG = 0
            SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant INNER JOIN M17Material_Master ON M17Code= Material where d_Date between '" & _ToDate & "' and '" & _From & "' and M13Department='4' AND M17Location='4004' group by M08Sales_Order,M08Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _STOCKFG = dsUser.Tables(0).Rows(0)("Qty")
            End If

            worksheet2.Range("H" & (X)).Formula = "=N11+" & _STOCKFG & "+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 8)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("H" & (X)).Formula = "=SUM(H5:H11)"
            worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 8)
            range1.NumberFormat = "0.00"
            worksheet2.Range("H" & X, "H" & X).Interior.Color = RGB(141, 180, 227)
            '================================================================================
            'OUT SOURCE
            X = 5
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='1' AND M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
            End If

            _STOCKFG = 0
            SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant INNER JOIN M17Material_Master ON M17Code= Material where d_Date between '" & _ToDate & "' and '" & _From & "' and M13Department='1' AND M17Location<>'4004' group by M08Sales_Order,M08Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _STOCKFG = dsUser.Tables(0).Rows(0)("Qty")
            End If

            worksheet2.Range("L" & (X)).Formula = "=P5+" & _STOCKFG & "+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 12)
            range1.NumberFormat = "0.00"

            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='2' AND M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
            End If

            _STOCKFG = 0
            SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant INNER JOIN M17Material_Master ON M17Code= Material where d_Date between '" & _ToDate & "' and '" & _From & "' and M13Department='2' AND M17Location<>'4004' group by M08Sales_Order,M08Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _STOCKFG = dsUser.Tables(0).Rows(0)("Qty")
            End If

            worksheet2.Range("L" & (X)).Formula = "=P7+" & _STOCKFG & "+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 12)
            range1.NumberFormat = "0.00"


            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='3' AND M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
            End If

            _STOCKFG = 0

            SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant INNER JOIN M17Material_Master ON M17Code= Material where d_Date between '" & _ToDate & "' and '" & _From & "' and M13Department='3' AND M17Location<>'4004' group by M08Sales_Order,M08Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _STOCKFG = dsUser.Tables(0).Rows(0)("Qty")
            End If

            worksheet2.Range("L" & (X)).Formula = "=P9+" & _STOCKFG & "+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 12)
            range1.NumberFormat = "0.00"


            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='4' AND M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("T08Pro_Qty1")
            End If

            _STOCKFG = 0
            SQL = "select sum(Qty) as Qty from View_FG_StockComment inner join M13Biz_Unit on M13Merchant=M08Merchant INNER JOIN M17Material_Master ON M17Code= Material where d_Date between '" & _ToDate & "' and '" & _From & "' and M13Department='4' AND M17Location<>'4004' group by M08Sales_Order,M08Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _STOCKFG = dsUser.Tables(0).Rows(0)("Qty")
            End If

            worksheet2.Range("L" & (X)).Formula = "=P11+" & _STOCKFG & "+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 12)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("L" & (X)).Formula = "=SUM(L5:L11)"
            worksheet2.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 12)
            range1.NumberFormat = "0.00"
            worksheet2.Range("L" & X, "L" & X).Interior.Color = RGB(141, 180, 227)
            '================================================================================
            'VALUE
            X = 5
            worksheet2.Range("E" & (X)).Formula = "=(I5+M5)-S5"
            worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 5)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("E" & (X)).Formula = "=(I7+M7)-S7"
            worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 5)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("E" & (X)).Formula = "=(I9+M9)-S9"
            worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 5)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("E" & (X)).Formula = "=(I11+M11)-S11"
            worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 5)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("E" & (X)).Formula = "=(E5+E7+E9+E11)-S13"
            worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 5)
            range1.NumberFormat = "0.00"
            worksheet2.Range("E" & X, "E" & X).Interior.Color = RGB(141, 180, 227)


            'IN HOUSE

            X = 5
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='1' AND M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("TotalValue1")
            End If
            worksheet2.Range("I" & (X)).Formula = "=O5+AB5+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 9)
            range1.NumberFormat = "0.00"

            X = X + 2

            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='2' AND M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("TotalValue1")
            End If
            worksheet2.Range("I" & (X)).Formula = "=O7+AB7+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 9)
            range1.NumberFormat = "0.00"


            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='3' AND M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("TotalValue1")
            End If
            worksheet2.Range("I" & (X)).Formula = "=O9+AB9+" & _PROMISE_TO_dEL
            worksheet2.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 9)
            range1.NumberFormat = "0.00"


            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='4' AND M17Location='4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("TotalValue1")
            End If
            worksheet2.Range("I" & (X)).Formula = "=O11+AB11+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 9)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("I" & (X)).Formula = "=SUM(I5:I11)"
            worksheet2.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 9)
            range1.NumberFormat = "0.00"
            worksheet2.Range("I" & X, "I" & X).Interior.Color = RGB(141, 180, 227)
            '------------------------------------------------------------------------------
            'OUT SOURCE
            X = 5
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='1' AND M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("TotalValue1")
            End If
            worksheet2.Range("M" & (X)).Formula = "=Q5+AB5+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 13)
            range1.NumberFormat = "0.00"

            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='2' AND M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("TotalValue1")
            End If
            worksheet2.Range("M" & (X)).Formula = "=Q7+AB7+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 13)
            range1.NumberFormat = "0.00"


            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='3' AND M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("TotalValue1")
            End If
            worksheet2.Range("M" & (X)).Formula = "=Q9+AB9+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 13)
            range1.NumberFormat = "0.00"


            X = X + 2
            _PROMISE_TO_dEL = 0
            SQL = "select sum(T08Pro_Qty1) as T08Pro_Qty1,sum(T08Pro_Qty1*T08Value) as TotalValue1,sum(T08Pro_Qty2) as T08Pro_Qty2,sum(T08Pro_Qty2*T08Value) as TotalValue2,sum(T08Pro_Qty1)/sum(T08Con_Fact) as KGQty,(sum(T08Pro_Qty1)/sum(T08Con_Fact))*sum(T08Value_KG ) as KgValue from T08DelayComment inner join M13Biz_Unit on M13Merchant=T08Merchant INNER JOIN M17Material_Master ON M17Code=T08Material where  T08Date between '" & _ToDate & "' and  '" & _From & "'  and M13Department='4' AND M17Location<>'4004' group by M13Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(dsUser) Then
                _PROMISE_TO_dEL = dsUser.Tables(0).Rows(0)("TotalValue1")
            End If
            worksheet2.Range("M" & (X)).Formula = "=Q11+AB11+" & _PROMISE_TO_dEL & ""
            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 13)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("M" & (X)).Formula = "=SUM(M5:M11)"
            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 13)
            range1.NumberFormat = "0.00"
            worksheet2.Range("M" & X, "M" & X).Interior.Color = RGB(141, 180, 227)
            '------------------------------------------------------------------------------
            'ASP(M)
            X = 5
            worksheet2.Range("B" & (X)).Formula = "=E5/D5"
            worksheet2.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 2)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("B" & (X)).Formula = "=E7/D7"
            worksheet2.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 2)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("B" & (X)).Formula = "=E9/D9"
            worksheet2.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 2)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("B" & (X)).Formula = "=E11/D11"
            worksheet2.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 2)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("B" & (X)).Formula = "=E13/D13"
            worksheet2.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 2)
            range1.NumberFormat = "0.00"
            worksheet2.Range("B" & X, "B" & X).Interior.Color = RGB(141, 180, 227)

            'IN HOUSE
            X = 5
            worksheet2.Range("F" & (X)).Formula = "=I5/H5"
            worksheet2.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 6)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("F" & (X)).Formula = "=I7/H7"
            worksheet2.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 6)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("F" & (X)).Formula = "=I9/H9"
            worksheet2.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 6)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("F" & (X)).Formula = "=I11/H11"
            worksheet2.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 6)
            range1.NumberFormat = "0.00"

            X = X + 2

            worksheet2.Range("F" & (X)).Formula = "=I13/H13"
            worksheet2.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 6)
            range1.NumberFormat = "0.00"
            worksheet2.Range("F" & X, "F" & X).Interior.Color = RGB(141, 180, 227)

            'OUT SOURCE
            X = 5
            worksheet2.Range("J" & (X)).Formula = "=M5/L5"
            worksheet2.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 10)
            range1.NumberFormat = "0.00"


            X = X + 2
            worksheet2.Range("J" & (X)).Formula = "=M7/L7"
            worksheet2.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 10)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("J" & (X)).Formula = "=M9/L9"
            worksheet2.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 10)
            range1.NumberFormat = "0.00"

            X = X + 2
            worksheet2.Range("J" & (X)).Formula = "=M11/L11"
            worksheet2.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 10)
            range1.NumberFormat = "0.00"

            X = X + 2

            worksheet2.Range("J" & (X)).Formula = "=M13/L13"
            worksheet2.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 10)
            range1.NumberFormat = "0.00"
            worksheet2.Range("J" & X, "J" & X).Interior.Color = RGB(141, 180, 227)
            '---------------------------------------------------------------------------------
            'ASP(KG) TOTAAL
            Dim _DelivaryQty(4) As Double
            Dim _DelivaryValue(4) As Double



            SQL = "select * from M14Retailer order by M14Code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            Y = 0
            _Chr = 106
            X = 5
            z = 6
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                SQL = "select sum(M06D_Qty_KG) as M06D_Qty_Mtr,sum(M06D_Qty_KG*M06Unit_Kg) as TotValue from M06Delivary_Qty inner  join M13Biz_Unit on M13Merchant=M06Merchant where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "'  group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    _DelivaryQty(Y) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")

                    _DelivaryValue(Y) = dsUser.Tables(0).Rows(0)("TotValue")


                End If

                Y = Y + 1
            Next
            Dim _Value As Double
            Dim _Qty As Double

            _Value = 0
            _Qty = 0

            X = 5
            For Y = 0 To 3
                If (_DelivaryValue(Y) + _PromDelValue(Y)) > 0 And (_DelivaryQty(Y) + _PromDelQty(Y)) > 0 Then
                    _Value = _Value + (_DelivaryValue(Y) + _PromDelValue(Y))
                    _Qty = _Qty + (_DelivaryQty(Y) + _PromDelQty(Y))

                    worksheet2.Range("C" & (X)).Formula = (_DelivaryValue(Y) + _PromDelValue(Y)) / (_DelivaryQty(Y) + _PromDelQty(Y))
                    worksheet2.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(X, 3)
                    range1.NumberFormat = "0.00"
                End If
                X = X + 2
            Next
            X = 13
            worksheet2.Range("C" & (X)).Formula = _Value / _Qty
            worksheet2.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 3)
            range1.NumberFormat = "0.00"
            worksheet2.Range("C" & X, "C" & X).Interior.Color = RGB(141, 180, 227)

            'ASP KG IN HOUSE

            SQL = "select * from M14Retailer order by M14Code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            Y = 0
            _Chr = 106
            X = 5
            z = 6
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                SQL = "select sum(M06D_Qty_KG) as M06D_Qty_Mtr,sum(M06D_Qty_KG*M06Unit_Kg) as TotValue from M06Delivary_Qty inner  join M13Biz_Unit on M13Merchant=M06Merchant INNER JOIN M17Material_Master ON M17Code=M06Material where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' AND M17Location<>'4004' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    _DelivaryQty(Y) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")

                    _DelivaryValue(Y) = dsUser.Tables(0).Rows(0)("TotValue")


                End If

                Y = Y + 1
            Next

            X = 5
            For Y = 0 To 3
                If (_DelivaryValue(Y) + _PromDelValue(Y)) > 0 And (_DelivaryQty(Y) + _PromDelQty(Y)) > 0 Then
                    worksheet2.Range("G" & (X)).Formula = (_DelivaryValue(Y) + _PromDelValue(Y)) / (_DelivaryQty(Y) + _PromDelQty(Y))
                    worksheet2.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(X, 7)
                    range1.NumberFormat = "0.00"
                End If
                X = X + 2
            Next
            X = 12
            'ASP KG OUT SOURCE
            SQL = "select * from M14Retailer order by M14Code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            Y = 0
            _Chr = 106
            X = 5
            z = 6
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                SQL = "select sum(M06D_Qty_KG) as M06D_Qty_Mtr,sum(M06D_Qty_KG*M06Unit_Kg) as TotValue from M06Delivary_Qty inner  join M13Biz_Unit on M13Merchant=M06Merchant INNER JOIN M17Material_Master ON M17Code=M06Material where M06Date between '" & STRMONTH & "' and '" & _To & "' and  M13Department='" & T01.Tables(0).Rows(Y)("M14Code") & "' AND M17Location='4004' group by M13Department"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    _DelivaryQty(Y) = dsUser.Tables(0).Rows(0)("M06D_Qty_Mtr")

                    _DelivaryValue(Y) = dsUser.Tables(0).Rows(0)("TotValue")
                Else
                    _DelivaryQty(Y) = 0
                    _DelivaryValue(Y) = 0
                End If

                Y = Y + 1
            Next

            X = 5
            For Y = 0 To 3
                If (_DelivaryValue(Y) + _PromDelValue(Y)) > 0 And (_DelivaryQty(Y) + _PromDelQty(Y)) > 0 Then
                    worksheet2.Range("K" & (X)).Formula = (_DelivaryValue(Y) + _PromDelValue(Y)) / (_DelivaryQty(Y) + _PromDelQty(Y))
                    worksheet2.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(X, 11)
                    range1.NumberFormat = "0.00"
                End If
                X = X + 2
            Next
            X = 12

            SQL = "select sum(M18Qty_Mtr) as M18Qty_Mtr,sum(M18Value) as M18Value from M18Return where M18Month='" & Month(STRMONTH) & "' and M18Year='" & Year(STRMONTH) & "' group by M18Month, M18Year"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            X = X + 1
            If isValidDataset(dsUser) Then
                worksheet2.Cells(X, 18) = dsUser.Tables(0).Rows(0)("M18Qty_Mtr")
                worksheet2.Cells(X, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X, 18)
                range1.NumberFormat = "0.00"
                worksheet2.Range("R" & X, "R" & X).Interior.Color = RGB(141, 180, 227)

                worksheet2.Cells(X, 19) = dsUser.Tables(0).Rows(0)("M18Value")
                worksheet2.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(X, 19)
                range1.NumberFormat = "0.00"
                worksheet2.Range("S" & X, "S" & X).Interior.Color = RGB(141, 180, 227)

            End If
            con.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' MsgBox(i)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.Close()
            End If
        End Try
    End Function

End Module
