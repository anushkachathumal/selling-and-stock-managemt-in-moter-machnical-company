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
Imports System.Windows
Public Class frmQualityWise
    Dim Clicked As String
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter
    'Dim exc As New Application

    'Dim workbooks As Workbooks = exc.Workbooks
    'Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    'Dim sheets As Sheets = Workbook.Worksheets
    'Dim worksheet1 As _Worksheet = CType(Sheets.Item(1), _Worksheet)

    Private Sub frmQualityWise_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDate.Text = Today
        txtTo.Text = Today

        Call Load_Quality()
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim X As Integer
        Dim range1 As Range
        Dim T02 As DataSet

        Dim exc As New Application
        Dim workbooks As Workbooks = exc.Workbooks
        Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
        exc.Visible = True

        Dim Y1 As Date
        Dim X1 As Date
        Dim i As Integer
        Dim T03 As DataSet

        Y1 = txtDate.Text & " " & txtTime1.Text
        X1 = txtTo.Text & " " & txtToTime.Text


        If cboFrom.Text <> "" Then
            If chk1.Checked = True Then

                workbooks.Application.Sheets.Add()
                Dim sheets1 As Sheets = workbook.Worksheets
                Dim worksheet11 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
                worksheet11.Name = "Quality wise Report"
                worksheet11.Cells(2, 3) = "Textured Jersey PLC"
                worksheet11.Rows(2).Font.Bold = True
                worksheet11.Rows(2).Font.size = 26
                worksheet11.Range("A2:J2").MergeCells = True
                worksheet11.Range("A2:J2").VerticalAlignment = XlVAlign.xlVAlignCenter


                worksheet11.Cells(4, 1) = "Quality wise Report "
                worksheet11.Rows(4).Font.Bold = True
                worksheet11.Rows(4).Font.size = 10

                worksheet11.Cells(5, 1) = "Quality :" & cboFrom.Text
                worksheet11.Rows(5).Font.Bold = True
                worksheet11.Rows(5).Font.size = 10

                worksheet11.Columns(1).columnwidth = 12


                worksheet11.Cells(7, 2) = "Knitting M/C No"
                worksheet11.Rows(7).Font.Bold = True
                worksheet11.Rows(7).Font.size = 10
                worksheet11.Columns(2).columnwidth = 12
                worksheet1.Cells(7, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(7, 2).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet1.Cells(7, 2).Orientation = 90


                worksheet11.Range("B7:B7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet11.Range("B7:B7").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet11.Range("B7:B7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet11.Range("B7:B7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                worksheet11.Cells(7, 3) = "No of Knitted Rolls"
                worksheet11.Columns(3).columnwidth = 12
                worksheet11.Cells(7, 3).WrapText = True
                'worksheet11.Cells(7, 3).VerticalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(7, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(7, 3).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet1.Cells(7, 3).Orientation = 90

                worksheet11.Range("b7:b7").Interior.Color = RGB(215, 228, 188)
            

                worksheet11.Range("c7:c7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet11.Range("c7:c7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet11.Range("c7:c7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                worksheet11.Range("c7:c7").Interior.Color = RGB(215, 228, 188)
                Dim Y As Integer
                SQL = "select M02Name from T01Transaction_Header inner join T02Trans_Fault on T02Ref=T01RefNo inner join M02Fault on T02FualtCode=M02Code where T01Time between '" & Y1 & "' and '" & X1 & "' and T02FualtCode not in ('101 - BB','101-B','101-TB','102','104','117','118','120','122','125','128','129','130','132','133','135','136','137','138','139','140','141','142','201','205','211','212','301') group by M02Name "
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                i = 0
                Y = 4
                For Each DTRow2 As DataRow In T01.Tables(0).Rows
                    worksheet11.Cells(7, y) = T01.Tables(0).Rows(i)("M02Name")
                    worksheet11.Cells(7, Y).WrapText = True
                    'worksheet11.Cells(7, Y).VerticalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(7, Y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(7, Y).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Cells(7, Y).Orientation = 90

                    worksheet11.Cells(7, Y).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet11.Cells(7, Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet11.Cells(7, Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet11.Cells(7, Y).Interior.Color = RGB(215, 228, 188)
                    Y = Y + 1
                    i = i + 1
                Next

                X = 8
                Dim Z As Integer
                Dim _RollCount As Integer

                SQL = "select M03MCNo,count(M03MCNo) as Mcount from T01Transaction_Header inner join M03Knittingorder on M03OrderNo=T01OrderNo where T01Time between '" & Y1 & "' and '" & X1 & "' and M03Quality='" & cboFrom.Text & "' group by M03MCNo"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                i = 0
                For Each DTRow2 As DataRow In T01.Tables(0).Rows
                    Z = 2
                    _RollCount = 0

                    worksheet11.Cells(X, Z) = T01.Tables(0).Rows(i)("M03MCNo")
                    worksheet1.Cells(X, Z).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDot
                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDot
                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDot

                    Z = Z + 1
                    _RollCount = T01.Tables(0).Rows(i)("Mcount")
                    worksheet11.Cells(X, Z) = T01.Tables(0).Rows(i)("Mcount")
                    worksheet11.Cells(X, Z).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, Z)
                    range1.NumberFormat = "0"


                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDot
                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDot
                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDot

                    SQL = "select M02Name from T01Transaction_Header inner join T02Trans_Fault on T02Ref=T01RefNo inner join M02Fault on T02FualtCode=M02Code where T01Time between '" & Y1 & "' and '" & X1 & "' and  T02FualtCode not in ('101 - BB','101-B','101-TB','102','104','117','118','120','122','125','128','129','130','132','133','135','136','137','138','139','140','141','142','201','205','211','212','301') group by M02Name  "
                    T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    Y = 0
                    Z = Z + 1
                    For Each DTRow3 As DataRow In T02.Tables(0).Rows
                        SQL = "select count(T02FualtCode) as T02FualtCode,sum(T02count) as T02count from T01Transaction_Header inner join M03Knittingorder on M03OrderNo=T01OrderNo inner join T02Trans_Fault on T01RefNo=T02Ref inner join M02Fault on M02code=T02FualtCode  where T01Time between '" & Y1 & "' and '" & X1 & "'  and M02Name='" & T02.Tables(0).Rows(Y)("M02Name") & "' and M03MCNo='" & T01.Tables(0).Rows(i)("M03MCNo") & "' and M03Quality='" & cboFrom.Text & "' group by M02Code"
                        T03 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        If isValidDataset(T03) Then
                            If IsDBNull(T03.Tables(0).Rows(0)("T02count")) Then
                                worksheet11.Cells(X, Z) = T03.Tables(0).Rows(0)("T02FualtCode") / _RollCount
                            Else
                                worksheet11.Cells(X, Z) = T03.Tables(0).Rows(0)("T02count") / _RollCount
                            End If
                            worksheet11.Cells(X, Z).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet11.Cells(X, Z)
                            range1.NumberFormat = "0"



                        Else

                        End If
                        worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDot
                        worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDot
                        worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDot
                        Z = Z + 1
                        Y = Y + 1
                    Next
               
                    X = X + 1
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
                xlCharts = worksheet11.ChartObjects
                Dim rh As Integer

                Dim _Chartlocation As Integer

                rh = (15 * 3) + (30.25 * 4)
                _Chartlocation = (X + 2) * 10

                rh = rh + _Chartlocation
                myChart = xlCharts.Add(10, rh, 905, 300)
                ' chartPage = myChart.Chart
                Dim _CH As Integer
                '_CH = 7
                'For _CH = 7 To x - 1
                chartPage = myChart.Chart
                Dim _String As String
                Dim _Chr As Char
                SQL = "select M02Name from T01Transaction_Header inner join T02Trans_Fault on T02Ref=T01RefNo inner join M02Fault on T02FualtCode=M02Code where T01Time between '" & Y1 & "' and '" & X1 & "' and  T02FualtCode not in ('101 - BB','101-B','101-TB','102','104','117','118','120','122','125','128','129','130','132','133','135','136','137','138','139','140','141','142','201','205','211','212','301') group by M02Name  "
                T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                Y = 0
                Z = 68
                For Each DTRow3 As DataRow In T02.Tables(0).Rows
                    _Chr = Convert.ToChar(Z)
                    _String = _Chr & "8," & _Chr & (X - 1)

                    If Y = 0 Then
                        chartRange = worksheet11.Range("D8", "D" & (X - 1))
                        chartRange1 = worksheet11.Range("B8", "B" & (X - 1))
                        'chartRange = worksheet1.Range("H8", "K" & (X - 1))
                        'chartRange = worksheet1.Range("H8:K39", "A9:A39")
                        ' chartPage.SetSourceData(Source:=chartRange)
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

                        End With
                    ElseIf Y = 1 Then
                        chartRange = worksheet11.Range("e8", "e" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 2 Then
                        chartRange = worksheet11.Range("f8", "f" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 3 Then
                        chartRange = worksheet11.Range("G8", "G" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 4 Then
                        chartRange = worksheet11.Range("H8", "H" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 5 Then
                        chartRange = worksheet11.Range("I8", "I" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 6 Then
                        chartRange = worksheet11.Range("J8", "J" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 7 Then
                        chartRange = worksheet11.Range("K8", "K" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 8 Then
                        chartRange = worksheet11.Range("L8", "L" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 9 Then
                        chartRange = worksheet11.Range("M8", "M" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 10 Then
                        chartRange = worksheet11.Range("N8", "N" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 11 Then
                        chartRange = worksheet11.Range("O8", "O" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 12 Then
                        chartRange = worksheet11.Range("P8", "P" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 13 Then
                        chartRange = worksheet11.Range("Q8", "Q" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 14 Then
                        chartRange = worksheet11.Range("R8", "R" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 15 Then
                        chartRange = worksheet11.Range("S8", "S" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 16 Then
                        chartRange = worksheet11.Range("T8", "T" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 17 Then
                        chartRange = worksheet11.Range("u8", "u" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 18 Then
                        chartRange = worksheet11.Range("v8", "v" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 19 Then
                        chartRange = worksheet11.Range("w8", "w" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 20 Then
                        chartRange = worksheet11.Range("x8", "x" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 21 Then
                        chartRange = worksheet11.Range("y8", "y" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    ElseIf Y = 22 Then
                        chartRange = worksheet11.Range("z8", "z" & (X - 1))
                        t_SerCol = chartPage.SeriesCollection
                        t_Series = t_SerCol.NewSeries
                        With t_Series
                            .Name = T02.Tables(0).Rows(Y)("M02name")
                            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE
                        End With
                    End If
                    Y = Y + 1
                    Z = Z + 1
                Next


                i = 0

            End If

        Else
            Dim result3 As DialogResult = MessageBox.Show("Please enter the Quality No", _
                "Information ...", _
                MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button2)
            If result3 = Forms.DialogResult.OK Then
                cboFrom.ToggleDropdown()
            End If

            Exit Sub
        End If
    End Sub

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

    Function Load_Quality()
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m03 As DataSet
        Dim Sql As String

        Try
            'Load Production Quality

            Sql = "select M03Quality as [Quality] from M03Knittingorder group by M03Quality"
            m03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboFrom
                .DataSource = m03
                .Rows.Band.Columns(0).Width = 340
            End With


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
End Class