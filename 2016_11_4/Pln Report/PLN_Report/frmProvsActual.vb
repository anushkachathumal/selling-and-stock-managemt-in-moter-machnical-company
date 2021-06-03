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
'Imports System.Drawing
'Imports Spire.XlS

Public Class frmProvsActual
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As System.Data.DataTable

    Private Sub frmProvsActual_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Month()
        txtYear.Text = Year(Today)
        Call Load_BizUnit()
        Call Load_Merchant()

    End Sub

    Function Load_Month()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M15Name as [Month Name] from M15Month ORDER by M15Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboMonth
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 245
                End With
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try

    End Function

    Function Load_Merchant()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M13Merchant as [Merchant] from M13Biz_Unit ORDER by M13Merchant"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboMerchnt
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 245
                End With
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try

    End Function


    Function Load_BizUnit()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M14Name as [Business Unit] from M14Retailer ORDER by M14Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboBunit
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 245
                End With
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try

    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        cboMonth.Text = ""
        txtYear.Text = Year(Today)
        cboBunit.Text = ""
        cboMerchnt.Text = ""
    End Sub

    Function Create_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T03 As DataSet
        Dim dsUser As DataSet
        Dim tblDye As DataSet

        Dim n_Date As Date
        Dim N_Date1 As Date
        Dim FileName As String
        Dim _FirstChr As Integer
        Dim _Possible_Date As Date
        Dim _Last As Integer
        Dim _Total_NoFail As Integer
        Dim X As Integer
        Dim exc As New Application
        ' Try
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

        Try
            exc.Visible = True

            Dim y As Integer
            Dim z As Integer
            Dim _Fromdate As Date
            Dim _ToDate As Date
            Dim L_DateofMonth As Integer

            SQL = "select * from M15Month where M15Name='" & cboMonth.Text & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                If IsNumeric(txtYear.Text) Then
                    If Microsoft.VisualBasic.Len(txtYear.Text) = 4 Then
                        _Fromdate = T01.Tables(0).Rows(0)("m15code") & "/1/" & txtYear.Text

                        Dim aDate As DateTime

                        aDate = _Fromdate
                        Dim EndDate As DateTime = aDate.AddDays(DateTime.DaysInMonth(aDate.Year, aDate.Month) - 1)
                        '  MsgBox(MonthName(T01.Tables(0).Rows(i)("T01month"), True))
                        L_DateofMonth = Microsoft.VisualBasic.Day(EndDate)

                        _ToDate = T01.Tables(0).Rows(0)("m15code") & "/" & L_DateofMonth & "/" & txtYear.Text
                    Else
                        MsgBox("Please enter the correct Year", MsgBoxStyle.Information, "Information .......")
                        Exit Function
                    End If
                Else
                    MsgBox("Please enter the correct Year", MsgBoxStyle.Information, "Information .......")
                    Exit Function
                End If
            End If


            _weekNo = L_DateofMonth / 7

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

            worksheet1.Columns("R").ColumnWidth = 20
            worksheet1.Columns("Q").ColumnWidth = 20
            worksheet1.Columns("S").ColumnWidth = 20

            worksheet1.Rows(3).Font.size = 15
            worksheet1.Rows(3).rowheight = 24
            worksheet1.Rows(3).Font.name = "Times New Roman"
            worksheet1.Rows(3).Font.BOLD = True
            worksheet1.Cells(3, 1) = Trim(cboMonth.Text) & "-" & txtYear.Text & " Prjection  vs Actual Analysis"
            worksheet1.Cells(3, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            If _weekNo = 4 Then
                worksheet1.Range("A3:x3").MergeCells = True
                worksheet1.Range("A3:x3").VerticalAlignment = XlVAlign.xlVAlignCenter
            Else
                worksheet1.Range("A3:y3").MergeCells = True
                worksheet1.Range("A3:y3").VerticalAlignment = XlVAlign.xlVAlignCenter
            End If
            'worksheet2.Rows(4).Font.size = 11
            ' worksheet1.Range("A1:D1").Interior.Color = RGB(197, 217, 241)
            worksheet1.Rows(4).Font.size = 11
            worksheet1.Rows(4).rowheight = 24
            worksheet1.Rows(4).Font.name = "Times New Roman"
            worksheet1.Rows(4).Font.BOLD = True

            worksheet1.Range("A4:F4").MergeCells = True
            worksheet1.Range("A4:F4").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(4, 1) = "Projection"
            worksheet1.Cells(4, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            X = 4

            worksheet1.Range("A" & X & ":F" & X).MergeCells = True
            worksheet1.Range("A" & X & ":F" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            If _weekNo = 4 Then
                worksheet1.Range("G" & X & ":n" & X).MergeCells = True
                worksheet1.Range("G" & X & ":n" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            Else
                worksheet1.Range("G" & X & ":P" & X).MergeCells = True
                worksheet1.Range("G" & X & ":P" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            End If
            worksheet1.Range("A" & X & ":F" & X).Interior.Color = RGB(251, 228, 213)
            If _weekNo = 4 Then
                worksheet1.Range("G" & X & ":n" & X).Interior.Color = RGB(197, 190, 151)
            Else
                worksheet1.Range("G" & X & ":P" & X).Interior.Color = RGB(197, 190, 151)
            End If
            worksheet1.Cells(4, 7) = "Actuals"
            worksheet1.Cells(4, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            Dim _Chr As Integer
            Dim I As Integer


            _Chr = 97
            If _weekNo = 4 Then
                For I = 1 To 14
                    worksheet1.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                    _Chr = _Chr + 1
                Next
            Else
                For I = 1 To 16
                    worksheet1.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & "4", Chr(_Chr) & "4").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                    _Chr = _Chr + 1
                Next
            End If
            X = X + 1
            worksheet1.Rows(X).rowheight = 18
            worksheet1.Cells(X, 1) = "BU"
            worksheet1.Rows(X).Font.size = 9
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Rows(X).Font.BOLD = True
            worksheet1.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("A" & X & ":A" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(X, 2) = "Quality"
            worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("B" & X & ":B" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(X, 3) = "Projection"
            worksheet1.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("C" & X & ":C" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(X, 4) = "Actual booked"
            worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("D" & X & ":D" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(X, 5) = "Inquiry orders Booked"
            worksheet1.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("E" & X & ":E" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(X, 6) = "Weekly projection"
            worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("F" & X & ":F" & X).VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet1.Cells(X, 7) = "Week 01"
            worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("G" & X & ":G" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("G" & X & ":H" & X).MergeCells = True
            worksheet1.Range("G" & X & ":H" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(X, 9) = "Week 02"
            worksheet1.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("I" & X & ":I" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("I" & X & ":J" & X).MergeCells = True
            worksheet1.Range("I" & X & ":J" & X).VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet1.Cells(X, 11) = "Week 03"
            worksheet1.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & X & ":K" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("K" & X & ":L" & X).MergeCells = True
            worksheet1.Range("K" & X & ":L" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(X, 13) = "Week 04"
            worksheet1.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("M" & X & ":M" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("M" & X & ":N" & X).MergeCells = True
            worksheet1.Range("M" & X & ":N" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            If _weekNo = 5 Then
                worksheet1.Cells(X, 14) = "Week 05"
                worksheet1.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("N" & X & ":N" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet1.Range("O" & X & ":P" & X).MergeCells = True
                worksheet1.Range("O" & X & ":P" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            End If


            _Chr = 103
            If _weekNo = 4 Then
                For I = 1 To 4
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(142, 170, 220)
                    worksheet1.Cells(X, I).WrapText = True
                    _Chr = _Chr + 1
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    _Chr = _Chr + 1
                Next
            Else
                For I = 1 To 5
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(142, 170, 220)
                    worksheet1.Cells(X, I).WrapText = True
                    _Chr = _Chr + 1
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    _Chr = _Chr + 1
                Next
            End If

            _Chr = 97
            For I = 1 To 6

                worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(255, 192, 0)
                worksheet1.Range(Chr(_Chr) & X + 1 & ":" & Chr(_Chr) & X + 1).Interior.Color = RGB(255, 192, 0)
                worksheet1.Cells(X, I).WrapText = True
                worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X + 1).MergeCells = True

                worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_Chr) & X + 1, Chr(_Chr) & X + 1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_Chr) & X + 1, Chr(_Chr) & X + 1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(Chr(_Chr) & X + 1, Chr(_Chr) & X + 1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                _Chr = _Chr + 1

            Next

            y = 1
            X = X + 1
            worksheet1.Rows(X).rowheight = 20
            worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Rows(X).Font.BOLD = True
            worksheet1.Rows(X).Font.size = 9
            y = 7
            _Chr = 103
            If _weekNo = 4 Then
                For I = 1 To 4
                    worksheet1.Cells(X, y) = "Actual"
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                    _Chr = _Chr + 1
                    y = y + 1

                    worksheet1.Cells(X, y) = "Variance"
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                    _Chr = _Chr + 1
                    y = y + 1

                Next
            Else
                For I = 1 To 5
                    worksheet1.Cells(X, y) = "Actual"
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                    _Chr = _Chr + 1
                    y = y + 1

                    worksheet1.Cells(X, y) = "Variance"
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                    _Chr = _Chr + 1
                    y = y + 1

                Next

            End If

            _Chr = 103
            If _weekNo = 4 Then
                For I = 1 To 8
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(142, 170, 220)
                    ' worksheet1.Cells(X, I).WrapText = True

                    _Chr = _Chr + 1
                Next
            Else
                For I = 1 To 10
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    '  worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(142, 170, 220)
                    ' worksheet1.Cells(X, I).WrapText = True
                    _Chr = _Chr + 1
                Next
            End If

            y = 1
            X = X + 1

            If cboMerchnt.Text <> "" Then
                SQL = "select left(M07Met_Dis,5) as M07Met_Dis,sum(M07Qty_Mtr) as M07Qty_Mtr,MAX(m14name)as Dis from View_DelivaryForcus inner join M13Biz_Unit on M13Merchant=M07Merchant inner join M14Retailer on M13Department=M14Code where M07date between '" & _Fromdate & "' and '" & _ToDate & "' and M07Merchant='" & Trim(cboMerchnt.Text) & "' group by left(M07Met_Dis,5)"
            ElseIf cboBunit.Text <> "" Then
                SQL = "select left(M07Met_Dis,5) as M07Met_Dis,sum(M07Qty_Mtr) as M07Qty_Mtr,MAX(m14name)as Dis from View_DelivaryForcus inner join M13Biz_Unit on M13Merchant=M07Merchant inner join M14Retailer on M13Department=M14Code where M07date between '" & _Fromdate & "' and '" & _ToDate & "' and m14name='" & Trim(cboBunit.Text) & "' group by left(M07Met_Dis,5)"
            Else
                SQL = "select left(M07Met_Dis,5) as M07Met_Dis,sum(M07Qty_Mtr) as M07Qty_Mtr,MAX(m14name)as Dis from View_DelivaryForcus inner join M13Biz_Unit on M13Merchant=M07Merchant inner join M14Retailer on M13Department=M14Code where M07date between '" & _Fromdate & "' and '" & _ToDate & "' group by left(M07Met_Dis,5)"
            End If
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            I = 0
            For Each DTRow3 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(X).Font.size = 8
                worksheet1.Rows(X).Font.name = "Times New Roman"
                worksheet1.Rows(X).rowheight = 15


                _ActualBooked = 0
                If IsNumeric(Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 1)) Then
                    y = 1
                    worksheet1.Cells(X, y) = (dsUser.Tables(0).Rows(I)("Dis"))
                    y = y + 1

                    worksheet1.Cells(X, y) = Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 5)
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    _Chr = 97
                    If _weekNo = 4 Then
                        For z = 1 To 14
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            worksheet1.Cells(X, z).WrapText = True
                            _Chr = _Chr + 1
                        Next
                    Else
                        For z = 1 To 16
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            worksheet1.Cells(X, z).WrapText = True
                            _Chr = _Chr + 1
                        Next
                    End If
                    worksheet1.Cells(X, 4) = dsUser.Tables(0).Rows(I)("M07Qty_Mtr")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X, 4)
                    range1.NumberFormat = "0.0"

                    Dim _Si As String
                    _Si = "=C" & X & "/" & L_DateofMonth & "*7"
                    worksheet1.Cells(X, 6) = _Si
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X, 6)
                    range1.NumberFormat = "0"
                    worksheet1.Range("F" & X & ":F" & X).Interior.Color = RGB(255, 192, 0)

                    y = 7
                    ' If _weekNo = 4 Then
                    z = 0
                    _Chr = 103
                    For z = 1 To _weekNo
                        _ActualBooked = 0
                        If z = 1 Then
                            _FromWeek = _Fromdate
                        End If
                        Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                        Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_FromWeek)
                        Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)
                        If dayName = "Monday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-1)
                        ElseIf dayName = "Tuesday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-2)
                        ElseIf dayName = "Wednesday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-3)
                        ElseIf dayName = "Thuesday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-4)
                        ElseIf dayName = "Friday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-5)
                        ElseIf dayName = "Saturday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-6)
                        End If

                        _Toweek = _FromWeek.AddDays(+6)

                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus where left(M07Met_Dis,5)='" & Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 5) & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' group by left(M07Met_Dis,5)"
                        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        If isValidDataset(T01) Then
                            _ActualBooked = T01.Tables(0).Rows(0)("M07Qty_Mtr")
                        End If

                        'SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr from M06Delivary_Qty where left(M06Met_Dis,5)='" & Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 5) & "' and M06Date between '" & _FromWeek & "' and '" & _Toweek & "' group by M06Material"
                        'T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        'If isValidDataset(T01) Then
                        '    _ActualBooked = _ActualBooked + T01.Tables(0).Rows(0)("M06D_Qty_Mtr")
                        'End If

                        worksheet1.Cells(X, y) = _ActualBooked
                        worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet1.Cells(X, y)
                        range1.NumberFormat = "0"
                        y = y + 1
                        Dim Si As String

                        Si = "=F" & X & "-" & Chr(_Chr) & X
                        worksheet1.Cells(X, y) = Si
                        worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet1.Cells(X, y)
                        range1.NumberFormat = "0"
                        _Chr = _Chr + 1
                        worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        _Chr = _Chr + 1

                        y = y + 1
                        _FromWeek = _Toweek.AddDays(+1)

                    Next

                    X = X + 1
                Else
                    ' worksheet1.Cells(X, y) = Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 8)
                End If
                ' worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter


                ' End If







                ' X = X + 1
                I = I + 1
                y = y + 1

            Next

            y = 1
            If cboMerchnt.Text <> "" Then
                SQL = "select left(M07Met_Dis,8) as M07Met_Dis,sum(M07Qty_Mtr) as M07Qty_Mtr,MAX(m14name)as Dis from View_DelivaryForcus inner join M13Biz_Unit on M13Merchant=M07Merchant inner join M14Retailer on M13Department=M14Code where M07date between '" & _Fromdate & "' and '" & _ToDate & "' and M13Merchant='" & Trim(cboMerchnt.Text) & "' group by left(M07Met_Dis,8)"
            ElseIf cboBunit.Text <> "" Then
                SQL = "select left(M07Met_Dis,8) as M07Met_Dis,sum(M07Qty_Mtr) as M07Qty_Mtr,MAX(m14name)as Dis from View_DelivaryForcus inner join M13Biz_Unit on M13Merchant=M07Merchant inner join M14Retailer on M13Department=M14Code where M07date between '" & _Fromdate & "' and '" & _ToDate & "' and m14name='" & Trim(cboBunit.Text) & "' group by left(M07Met_Dis,8)"
            Else
                SQL = "select left(M07Met_Dis,8) as M07Met_Dis,sum(M07Qty_Mtr) as M07Qty_Mtr,MAX(m14name)as Dis from View_DelivaryForcus inner join M13Biz_Unit on M13Merchant=M07Merchant inner join M14Retailer on M13Department=M14Code where M07date between '" & _Fromdate & "' and '" & _ToDate & "' group by left(M07Met_Dis,8)"
            End If
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            I = 0
            For Each DTRow3 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(X).Font.size = 8
                worksheet1.Rows(X).Font.name = "Times New Roman"
                worksheet1.Rows(X).rowheight = 15


                _ActualBooked = 0
                If IsNumeric(Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 1)) Then
                Else
                    y = 1
                    worksheet1.Cells(X, y) = (dsUser.Tables(0).Rows(I)("Dis"))
                    y = y + 1
                    worksheet1.Cells(X, y) = Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 8)
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    _Chr = 97
                    If _weekNo = 4 Then
                        For z = 1 To 14
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            worksheet1.Cells(X, z).WrapText = True
                            _Chr = _Chr + 1
                        Next
                    Else
                        For z = 1 To 16
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            worksheet1.Cells(X, z).WrapText = True
                            _Chr = _Chr + 1
                        Next
                    End If
                    worksheet1.Cells(X, 4) = dsUser.Tables(0).Rows(I)("M07Qty_Mtr")
                    worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X, 4)
                    range1.NumberFormat = "0.0"

                    Dim _Si As String
                    _Si = "=C" & X & "/" & L_DateofMonth & "*7"
                    worksheet1.Cells(X, 6) = _Si
                    worksheet1.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X, 6)
                    range1.NumberFormat = "0"
                    worksheet1.Range("F" & X & ":F" & X).Interior.Color = RGB(255, 192, 0)

                    y = 7
                    ' If _weekNo = 4 Then
                    z = 0
                    _Chr = 103
                    For z = 1 To _weekNo
                        _ActualBooked = 0
                        If z = 1 Then
                            _FromWeek = _Fromdate
                        End If
                        Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                        Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_FromWeek)
                        Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)
                        If dayName = "Monday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-1)
                        ElseIf dayName = "Tuesday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-2)
                        ElseIf dayName = "Wednesday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-3)
                        ElseIf dayName = "Thuesday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-4)
                        ElseIf dayName = "Friday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-5)
                        ElseIf dayName = "Saturday" Then
                            _FromWeek = CDate(_FromWeek).AddDays(-6)
                        End If

                        _Toweek = _FromWeek.AddDays(+6)

                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus where left(M07Met_Dis,8)='" & Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 8) & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' group by left(M07Met_Dis,8)"
                        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        If isValidDataset(T01) Then
                            _ActualBooked = T01.Tables(0).Rows(0)("M07Qty_Mtr")
                        End If

                        ' SQL = "select sum(M06D_Qty_Mtr) as M06D_Qty_Mtr from M06Delivary_Qty where left(M06Met_Dis,8)='" & Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 8) & "' and M06Date between '" & _FromWeek & "' and '" & _Toweek & "' group by M06Material"
                        'T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        'If isValidDataset(T01) Then
                        '    _ActualBooked = _ActualBooked + T01.Tables(0).Rows(0)("M06D_Qty_Mtr")
                        'End If

                        worksheet1.Cells(X, y) = _ActualBooked
                        worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet1.Cells(X, y)
                        range1.NumberFormat = "0"
                        y = y + 1
                        Dim Si As String

                        Si = "=F" & X & "-" & Chr(_Chr) & X
                        worksheet1.Cells(X, y) = Si
                        worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet1.Cells(X, y)
                        range1.NumberFormat = "0"
                        _Chr = _Chr + 1
                        worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        _Chr = _Chr + 1

                        y = y + 1
                        _FromWeek = _Toweek.AddDays(+1)

                    Next
                    X = X + 1
                    ' worksheet1.Cells(X, y) = Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(I)("M07Met_Dis"), 8)
                End If
                ' worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter


                ' End If








                I = I + 1
                y = y + 1

            Next
            'worksheet1.Cells(X, 7) = "AA"
            'worksheet1.Cells(X, 7).Font.Bold = True
            '=====================================================================================================================================
            X = 7
            y = 18
            ' worksheet1.Rows(X).rowheight = 18
            worksheet1.Cells(X, y) = "Business Unit"
            worksheet1.Cells(X, y).Font.size = 9
            ' worksheet1.Rows(X).Font.name = "Times New Roman"
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("R" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet1.Range("R7:S8").MergeCells = True
            worksheet1.Range("R7:S8").VerticalAlignment = XlVAlign.xlVAlignCenter

            y = y + 2
            _Chr = 116
            For I = 1 To _weekNo
                If I = 1 Then
                    worksheet1.Cells(X, y) = "Week1"
                    worksheet1.Cells(X, y).Font.size = 9
                    worksheet1.Cells(X, y).Font.BOLD = True
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Range("T" & X & ":V" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range("T" & X & ":V" & X).Interior.Color = RGB(197, 217, 241)
                    worksheet1.Range("T" & X & ":V" & X).MergeCells = True
                    worksheet1.Range("T" & X & ":V" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                ElseIf I = 2 Then
                    worksheet1.Cells(X, y) = "Week2"
                    worksheet1.Cells(X, y).Font.size = 9
                    worksheet1.Cells(X, y).Font.BOLD = True
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Range("W" & X & ":Y" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range("W" & X & ":Y" & X).Interior.Color = RGB(197, 217, 241)
                    worksheet1.Range("W" & X & ":Y" & X).MergeCells = True
                    worksheet1.Range("W" & X & ":Y" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                ElseIf I = 3 Then
                    worksheet1.Cells(X, y) = "Week3"
                    worksheet1.Cells(X, y).Font.size = 9
                    worksheet1.Cells(X, y).Font.BOLD = True
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Range("Z" & X & ":AB" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range("Z" & X & ":AB" & X).Interior.Color = RGB(197, 217, 241)
                    worksheet1.Range("Z" & X & ":AB" & X).MergeCells = True
                    worksheet1.Range("Z" & X & ":AB" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                ElseIf I = 4 Then
                    worksheet1.Cells(X, y) = "Week4"
                    worksheet1.Cells(X, y).Font.size = 9
                    worksheet1.Cells(X, y).Font.BOLD = True
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Range("AC" & X & ":AE" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range("AC" & X & ":AE" & X).Interior.Color = RGB(197, 217, 241)
                    worksheet1.Range("AC" & X & ":AE" & X).MergeCells = True
                    worksheet1.Range("AC" & X & ":AE" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                ElseIf I = 5 Then
                    worksheet1.Cells(X, y) = "Week5"
                    worksheet1.Cells(X, y).Font.size = 9
                    worksheet1.Cells(X, y).Font.BOLD = True
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Range("AF" & X & ":AH" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet1.Range("AF" & X & ":AH" & X).Interior.Color = RGB(197, 217, 241)
                    worksheet1.Range("AF" & X & ":AH" & X).MergeCells = True
                    worksheet1.Range("AF" & X & ":AH" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                End If
                y = y + 3
            Next

            _Chr = 114

            If _weekNo = 4 Then
                For z = 18 To 31
                    If z = 27 Then
                        _Chr = 97
                    End If
                    If z < 27 Then
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

                        ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        ' worksheet1.Cells(X, z).WrapText = True
                    Else
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        ' worksheet1.Cells(X, z).WrapText = True
                    End If
                    _Chr = _Chr + 1
                Next

            Else
                For z = 18 To 34
                    If z = 27 Then
                        _Chr = 97
                    End If
                    If z > 27 Then
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        ' worksheet1.Cells(X, z).WrapText = True
                    Else
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        ' worksheet1.Cells(X, z).WrapText = True
                    End If
                    _Chr = _Chr + 1
                Next
            End If

            X = X + 1
            _Chr = 116
            y = 20
            For I = 1 To _weekNo

                worksheet1.Cells(X, y) = "Projection"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet1.Range("T" & X & ":V" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                'worksheet1.Range("T" & X & ":V" & X).Interior.Color = RGB(197, 217, 241)
                If y = 27 Then
                    _Chr = 97
                End If
                If y < 26 Then
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                Else
                    worksheet1.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                End If
                y = y + 1
                _Chr = _Chr + 1

                worksheet1.Cells(X, y) = "Actual"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet1.Range("T" & X & ":V" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                'worksheet1.Range("T" & X & ":V" & X).Interior.Color = RGB(197, 217, 241)
                If y = 27 Then
                    _Chr = 97
                End If
                If y < 26 Then
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                Else
                    worksheet1.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                End If
                y = y + 1
                _Chr = _Chr + 1

                worksheet1.Cells(X, y) = "Variance"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet1.Range("T" & X & ":V" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                'worksheet1.Range("T" & X & ":V" & X).Interior.Color = RGB(197, 217, 241)
                If y = 27 Then
                    _Chr = 97
                End If
                If y < 26 Then
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                Else
                    worksheet1.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).MergeCells = True
                    worksheet1.Range("A" & Chr(_Chr) & X & ":A" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                End If
                y = y + 1
                _Chr = _Chr + 1
            Next


            _Chr = 114

            If _weekNo = 4 Then
                For z = 18 To 31
                    If z = 27 Then
                        _Chr = 97
                    End If
                    If z < 27 Then
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous

                        ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        ' worksheet1.Cells(X, z).WrapText = True
                    Else
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        ' worksheet1.Cells(X, z).WrapText = True
                    End If
                    _Chr = _Chr + 1
                Next

            Else
                For z = 18 To 34
                    If z = 27 Then
                        _Chr = 97
                    End If
                    If z > 27 Then
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        ' worksheet1.Cells(X, z).WrapText = True
                    Else
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet1.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        ' worksheet1.Cells(X, z).WrapText = True
                    End If
                    _Chr = _Chr + 1
                Next
            End If
            X = X + 1
            _FirstChr = X
            SQL = "select * from M14Retailer order by m14code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            I = 0

            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _FirstChr = X
                y = 18
                worksheet1.Cells(X, 18) = T01.Tables(0).Rows(I)("M14Name")
                worksheet1.Cells(X, 18).Font.size = 9
                worksheet1.Cells(X, 18).Font.BOLD = True
                worksheet1.Cells(X, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter

                y = y + 1
                worksheet1.Cells(X, y) = "Presetting"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Range("S" & X & ":S" & X).MergeCells = True
                worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                X = X + 1

                worksheet1.Cells(X, y) = "Compact"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Range("S" & X & ":S" & X).MergeCells = True
                worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                X = X + 1

                worksheet1.Cells(X, y) = "Total Finishing"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Range("S" & X & ":S" & X).MergeCells = True
                worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                X = X + 1

                worksheet1.Cells(X, y) = "White/Marl "
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Range("S" & X & ":S" & X).MergeCells = True
                worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                X = X + 1

                worksheet1.Cells(X, y) = "Bio Polish"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Range("S" & X & ":S" & X).MergeCells = True
                worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                X = X + 1

                worksheet1.Cells(X, y) = "Peaching"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Range("S" & X & ":S" & X).MergeCells = True
                worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                X = X + 1

                worksheet1.Cells(X, y) = "ASP(m)"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Range("S" & X & ":S" & X).MergeCells = True
                worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                X = X + 1

                worksheet1.Cells(X, y) = "ASP(Kg)"
                worksheet1.Cells(X, y).Font.size = 9
                worksheet1.Cells(X, y).Font.BOLD = True
                worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Range("S" & X & ":S" & X).MergeCells = True
                worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet1.Range("R" & _FirstChr & ":R" & X).MergeCells = True
                worksheet1.Range("R" & _FirstChr & ":R" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                X = X + 1

                _Chr = 114
                Dim M2 As Integer
                For M2 = 0 To 8
                    If _weekNo = 4 Then
                        _Chr = 114
                        For z = 18 To 31
                            If z = 27 Then
                                _Chr = 97
                            End If
                            If z < 27 Then
                                worksheet1.Range(Chr(_Chr) & _FirstChr + M2, Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(Chr(_Chr) & _FirstChr + M2, Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(Chr(_Chr) & _FirstChr + M2, Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(Chr(_Chr) & _FirstChr + M2, Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                                If M2 = 8 Then
                                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                                End If
                                ' worksheet1.Cells(X, z).WrapText = True
                            Else
                                worksheet1.Range("A" & Chr(_Chr) & _FirstChr + M2, "A" & Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range("A" & Chr(_Chr) & _FirstChr + M2, "A" & Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range("A" & Chr(_Chr) & _FirstChr + M2, "A" & Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range("A" & Chr(_Chr) & _FirstChr + M2, "A" & Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                                If M2 = 8 Then
                                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                                End If
                                ' worksheet1.Cells(X, z).WrapText = True
                            End If
                            _Chr = _Chr + 1
                        Next

                    Else
                        For z = 18 To 34
                            If z = 27 Then
                                _Chr = 97
                            End If
                            If z > 27 Then
                                worksheet1.Range(Chr(_Chr) & _FirstChr + M2, Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(Chr(_Chr) & _FirstChr + M2, Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(Chr(_Chr) & _FirstChr + M2, Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range(Chr(_Chr) & _FirstChr + M2, Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                                If M2 = 8 Then
                                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                                End If
                                ' worksheet1.Cells(X, z).WrapText = True
                            Else
                                worksheet1.Range("A" & Chr(_Chr) & _FirstChr + M2, "A" & Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range("A" & Chr(_Chr) & _FirstChr + M2, "A" & Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range("A" & Chr(_Chr) & _FirstChr + M2, "A" & Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                                worksheet1.Range("A" & Chr(_Chr) & _FirstChr + M2, "A" & Chr(_Chr) & _FirstChr + M2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                                If M2 = 8 Then
                                    worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                                End If
                                ' worksheet1.Cells(X, z).WrapText = True
                            End If
                            _Chr = _Chr + 1
                        Next
                    End If
                Next
                I = I + 1
            Next
            '  y = 18
            _FirstChr = 41
            worksheet1.Cells(X, 18) = "TOTAL"
            worksheet1.Cells(X, 18).Font.size = 9
            worksheet1.Cells(X, 18).Font.BOLD = True
            worksheet1.Cells(X, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
            y = 19
            worksheet1.Cells(X, y) = "Presetting"
            worksheet1.Cells(X, y).Font.size = 9
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Range("S" & X & ":S" & X).MergeCells = True
            worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            X = X + 1

            worksheet1.Cells(X, y) = "Compact"
            worksheet1.Cells(X, y).Font.size = 9
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Range("S" & X & ":S" & X).MergeCells = True
            worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            X = X + 1

            worksheet1.Cells(X, y) = "Total Finishing"
            worksheet1.Cells(X, y).Font.size = 9
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Range("S" & X & ":S" & X).MergeCells = True
            worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            X = X + 1

            worksheet1.Cells(X, y) = "White/Marl "
            worksheet1.Cells(X, y).Font.size = 9
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Range("S" & X & ":S" & X).MergeCells = True
            worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            X = X + 1

            worksheet1.Cells(X, y) = "Bio Polish"
            worksheet1.Cells(X, y).Font.size = 9
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Range("S" & X & ":S" & X).MergeCells = True
            worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            X = X + 1

            worksheet1.Cells(X, y) = "Peaching"
            worksheet1.Cells(X, y).Font.size = 9
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Range("S" & X & ":S" & X).MergeCells = True
            worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            X = X + 1

            worksheet1.Cells(X, y) = "ASP(m)"
            worksheet1.Cells(X, y).Font.size = 9
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Range("S" & X & ":S" & X).MergeCells = True
            worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            X = X + 1

            worksheet1.Cells(X, y) = "ASP(Kg)"
            worksheet1.Cells(X, y).Font.size = 9
            worksheet1.Cells(X, y).Font.BOLD = True
            worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignLeft
            worksheet1.Range("S" & X & ":S" & X).MergeCells = True
            worksheet1.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("R" & _FirstChr & ":R" & X).MergeCells = True
            worksheet1.Range("R" & _FirstChr & ":R" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            Dim m1 As Integer

            For m1 = 0 To 7
                If _weekNo = 4 Then
                    _Chr = 114
                    For z = 18 To 31
                        If z = 27 Then
                            _Chr = 97
                        End If
                        If z < 27 Then
                            worksheet1.Range(Chr(_Chr) & _FirstChr + m1, Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & _FirstChr + m1, Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & _FirstChr + m1, Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & _FirstChr + m1, Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                            If m1 = 8 Then
                                worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            End If
                            ' worksheet1.Cells(X, z).WrapText = True
                        Else
                            worksheet1.Range("A" & Chr(_Chr) & _FirstChr + m1, "A" & Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range("A" & Chr(_Chr) & _FirstChr + m1, "A" & Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range("A" & Chr(_Chr) & _FirstChr + m1, "A" & Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range("A" & Chr(_Chr) & _FirstChr + m1, "A" & Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                            If m1 = 8 Then
                                worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            End If
                            ' worksheet1.Cells(X, z).WrapText = True
                        End If
                        _Chr = _Chr + 1
                    Next

                Else
                    For z = 18 To 34
                        If z = 27 Then
                            _Chr = 97
                        End If
                        If z > 27 Then
                            worksheet1.Range(Chr(_Chr) & _FirstChr + m1, Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & _FirstChr + m1, Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & _FirstChr + m1, Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range(Chr(_Chr) & _FirstChr + m1, Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                            If m1 = 8 Then
                                worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            End If
                            ' worksheet1.Cells(X, z).WrapText = True
                        Else
                            worksheet1.Range("A" & Chr(_Chr) & _FirstChr + m1, "A" & Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range("A" & Chr(_Chr) & _FirstChr + m1, "A" & Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range("A" & Chr(_Chr) & _FirstChr + m1, "A" & Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            worksheet1.Range("A" & Chr(_Chr) & _FirstChr + m1, "A" & Chr(_Chr) & _FirstChr + m1).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                            If m1 = 8 Then
                                worksheet1.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                            End If
                            ' worksheet1.Cells(X, z).WrapText = True
                        End If
                        _Chr = _Chr + 1
                    Next
                End If
            Next

            '=====================================================================================================
            If cboBunit.Text <> "" Then
                SQL = "select * from M14Retailer where M14Name='" & Trim(cboBunit.Text) & "' order by m14code"
            Else
                SQL = "select * from M14Retailer order by m14code"
            End If
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            I = 0
            X = 9
            y = 21
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                z = 0
                _Chr = 103
                For z = 1 To _weekNo
                    _ActualBooked = 0
                    If z = 1 Then
                        _FromWeek = _Fromdate
                    End If
                    Dim thisCulture = Globalization.CultureInfo.CurrentCulture
                    Dim dayOfWeek As DayOfWeek = thisCulture.Calendar.GetDayOfWeek(_FromWeek)
                    Dim dayName As String = thisCulture.DateTimeFormat.GetDayName(dayOfWeek)
                    If dayName = "Monday" Then
                        _FromWeek = CDate(_FromWeek).AddDays(-1)
                    ElseIf dayName = "Tuesday" Then
                        _FromWeek = CDate(_FromWeek).AddDays(-2)
                    ElseIf dayName = "Wednesday" Then
                        _FromWeek = CDate(_FromWeek).AddDays(-3)
                    ElseIf dayName = "Thuesday" Then
                        _FromWeek = CDate(_FromWeek).AddDays(-4)
                    ElseIf dayName = "Friday" Then
                        _FromWeek = CDate(_FromWeek).AddDays(-5)
                    ElseIf dayName = "Saturday" Then
                        _FromWeek = CDate(_FromWeek).AddDays(-6)
                    End If

                    _Toweek = _FromWeek.AddDays(+6)

                    _ActualBooked = 0
                    If cboMerchnt.Text <> "" Then
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department  where M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M15Dis2='Pre set' and M13Merchant='" & Trim(cboMerchnt.Text) & "' group by M14Name"
                    ElseIf cboBunit.Text <> "" Then
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department  where M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M15Dis2='Pre set' and m14Name='" & Trim(cboBunit.Text) & "' group by M14Name"
                    Else
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department  where M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M15Dis2='Pre set' group by M14Name"
                    End If
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        _ActualBooked = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    End If


                    worksheet1.Cells(X, y) = _ActualBooked
                    worksheet1.Cells(X, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X, y)
                    range1.NumberFormat = "0"

                    _ActualBooked = 0

                    If cboMerchnt.Text <> "" Then
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where M15Final_Finish='compact' and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M07Merchant='" & Trim(cboMerchnt.Text) & "' group by M14Name"
                    ElseIf cboBunit.Text <> "" Then
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where M15Final_Finish='compact' and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M14Name='" & Trim(cboBunit.Text) & "' group by M14Name"
                    Else
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where M15Final_Finish='compact' and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' group by M14Name"
                    End If
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        _ActualBooked = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    End If

                    worksheet1.Cells(X + 1, y) = _ActualBooked
                    worksheet1.Cells(X + 1, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X + 1, y)
                    range1.NumberFormat = "0"

                    _ActualBooked = 0

                    If cboMerchnt.Text <> "" Then
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where M15Final_Finish in ('compact','Stenter') and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M07Merchant='" & Trim(cboMerchnt.Text) & "' group by M14Name"
                    ElseIf cboBunit.Text <> "" Then
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where M15Final_Finish in ('compact','Stenter') and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M14Name='" & Trim(cboBunit.Text) & "' group by M14Name"
                    Else
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where M15Final_Finish in ('compact','Stenter') and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' group by M14Name"
                    End If
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        _ActualBooked = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    End If

                    worksheet1.Cells(X + 2, y) = _ActualBooked
                    worksheet1.Cells(X + 2, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X + 2, y)
                    range1.NumberFormat = "0"

                    _ActualBooked = 0
                    Dim f As Integer
                    f = 0

                    If cboMerchnt.Text <> "" Then
                        vcWharer = "M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M16Shade_Type ='PFP' AND M16Product_Type='Solid' and M13Merchant='" & Trim(cboMerchnt.Text) & "'"
                        vcWharer1 = "M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M16Shade_Type in ('Marls','Yarn Dyes','White') and M13Merchant='" & Trim(cboMerchnt.Text) & "'"
                    Else
                        vcWharer = "M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M16Shade_Type ='PFP' AND M16Product_Type='Solid'"
                        vcWharer1 = "M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M16Shade_Type in ('Marls','Yarn Dyes','White')"
                    End If
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetWhiteMarl", New SqlParameter("@cQryType", "RCODE"), New SqlParameter("@vcWhereClause1", vcWharer1), New SqlParameter("@vcWhereClause", vcWharer))
                    For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                        _ActualBooked = _ActualBooked + dsUser.Tables(0).Rows(f)("M07Qty_Mtr")
                        f = f + 1
                    Next

                    worksheet1.Cells(X + 3, y) = _ActualBooked / 0.9
                    worksheet1.Cells(X + 3, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X + 3, y)
                    range1.NumberFormat = "0"

                    _ActualBooked = 0
                    If cboMerchnt.Text <> "" Then
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where M15Final_Finish in ('compact','Stenter') and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M15Dis1='Bio polish' and M13Merchant='" & Trim(cboMerchnt.Text) & "' group by M14Name"
                    Else
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where M15Final_Finish in ('compact','Stenter') and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M15Dis1='Bio polish' group by M14Name"
                    End If
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then
                        _ActualBooked = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    End If

                    worksheet1.Cells(X + 4, y) = _ActualBooked
                    worksheet1.Cells(X + 4, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X + 4, y)
                    range1.NumberFormat = "0"

                    _ActualBooked = 0
                    If cboMerchnt.Text <> "" Then
                        vcWharer1 = "M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M16Product_Type ='Peach' and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9') and M13Merchant='" & Trim(cboMerchnt.Text) & "'"
                    Else

                        vcWharer1 = "M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M16Product_Type ='Peach' and ISNUMERIC(M16R_Code)  in ('1','0','2','3','4','5','6','7','8','9')"
                    End If
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetWhiteMarl", New SqlParameter("@cQryType", "RCODE"), New SqlParameter("@vcWhereClause1", vcWharer1))
                    f = 0
                    For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                        _ActualBooked = _ActualBooked + dsUser.Tables(0).Rows(f)("M07Qty_Mtr")
                        f = f + 1
                    Next

                    worksheet1.Cells(X + 5, y) = _ActualBooked
                    worksheet1.Cells(X + 5, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                    range1 = worksheet1.Cells(X + 5, y)
                    range1.NumberFormat = "0"

                    _ActualBooked = 0
                    If cboMerchnt.Text <> "" Then
                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr,sum( M07Qty_Kg) as  M07Qty_Kg,sum(MtrPrice) as MtrPrice,sum(KgPrice) as KgPrice from View_DelivaryForcus  inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where  M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "' and M13Merchant='" & Trim(cboMerchnt.Text) & "' group by M14Name"
                    Else

                        SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr,sum( M07Qty_Kg) as  M07Qty_Kg,sum(MtrPrice) as MtrPrice,sum(KgPrice) as KgPrice from View_DelivaryForcus  inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department where  M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _FromWeek & "' and '" & _Toweek & "'  group by M14Name"
                    End If
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                    If isValidDataset(dsUser) Then


                        _ActualBooked = dsUser.Tables(0).Rows(0)("MtrPrice") / dsUser.Tables(0).Rows(0)("M07Qty_Mtr")


                        worksheet1.Cells(X + 6, y) = _ActualBooked
                        worksheet1.Cells(X + 6, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet1.Cells(X + 6, y)
                        range1.NumberFormat = "0.00"


                        _ActualBooked = 0
                        _ActualBooked = dsUser.Tables(0).Rows(0)("KgPrice") / dsUser.Tables(0).Rows(0)("M07Qty_Kg")


                        worksheet1.Cells(X + 7, y) = _ActualBooked
                        worksheet1.Cells(X + 7, y).HorizontalAlignment = XlHAlign.xlHAlignRight
                        range1 = worksheet1.Cells(X + 7, y)
                        range1.NumberFormat = "0.00"
                    End If



                    y = y + 3


                    'y = y + 1
                    _FromWeek = _Toweek.AddDays(+1)

                Next
                X = X + 8
                y = 21
                I = I + 1
            Next
            Dim Y1 As Integer
            y = 9
            X = 41
            _Chr = 116

            Dim t2 As Integer

            For t2 = 1 To 8
                _Chr = 116
                For m1 = 20 To 31
                    If m1 = 27 Then
                        _Chr = 97
                    End If

                    If m1 > 26 Then
                        worksheet1.Cells(X, m1) = "=" & "A" & Chr(_Chr) & y & "+A" & Chr(_Chr) & y + 8 & "+A" & Chr(_Chr) & y + 16 & "+A" & Chr(_Chr) & y + 24
                    Else
                        worksheet1.Cells(X, m1) = "=" & Chr(_Chr) & y & "+" & Chr(_Chr) & y + 8 & "+" & Chr(_Chr) & y + 16 & "+" & Chr(_Chr) & y + 24
                    End If
                    _Chr = _Chr + 1

                Next
                y = y + 1
                X = X + 1
            Next
            '------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------

            workbooks.Application.Sheets.Add()
            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet117 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet117.Name = "SJ-RIB"
            worksheet117.Columns("A").ColumnWidth = 20
            worksheet117.Columns("B").ColumnWidth = 10
            worksheet117.Columns("C").ColumnWidth = 10
            worksheet117.Columns("D").ColumnWidth = 10
            worksheet117.Columns("E").ColumnWidth = 10
            worksheet117.Columns("F").ColumnWidth = 10
            worksheet117.Columns("G").ColumnWidth = 10
            worksheet117.Columns("H").ColumnWidth = 10

            worksheet117.Columns("J").ColumnWidth = 10
            worksheet117.Columns("K").ColumnWidth = 10
            worksheet117.Columns("L").ColumnWidth = 10
            worksheet117.Columns("M").ColumnWidth = 20

            X = 2
            worksheet117.Rows(X).rowheight = 20
            worksheet117.Rows(X).Font.name = "Times New Roman"
            worksheet117.Rows(X).Font.BOLD = True
            worksheet117.Rows(X).Font.size = 9

            worksheet117.Cells(X, 1) = "Unit"
            ' worksheet1.Cells(X, 1).Font.size = 9
            'worksheet1.Cells(X, y).Font.BOLD = True
            worksheet117.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet117.Range("A" & X & ":A" & X + 1).MergeCells = True
            worksheet117.Range("A" & X & ":A" & X + 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet117.Cells(X, 2) = "Projection (m)" & cboMonth.Text
            ' worksheet1.Cells(X, 1).Font.size = 9
            'worksheet1.Cells(X, y).Font.BOLD = True
            worksheet117.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet117.Range("B" & X & ":D" & X).MergeCells = True
            worksheet117.Range("B" & X & ":D" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet117.Cells(X, 5) = "Quoted Qty(m)"
            ' worksheet1.Cells(X, 1).Font.size = 9
            'worksheet1.Cells(X, y).Font.BOLD = True
            worksheet117.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet117.Range("E" & X & ":G" & X).MergeCells = True
            worksheet117.Range("E" & X & ":G" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet117.Cells(X, 8) = "Difference"
            ' worksheet1.Cells(X, 1).Font.size = 9
            'worksheet1.Cells(X, y).Font.BOLD = True
            worksheet117.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet117.Range("H" & X & ":J" & X).MergeCells = True
            worksheet117.Range("H" & X & ":J" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet117.Cells(X, 11) = "Speed orders"
            ' worksheet1.Cells(X, 1).Font.size = 9
            'worksheet1.Cells(X, y).Font.BOLD = True
            worksheet117.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet117.Range("K" & X & ":K" & X).MergeCells = True
            worksheet117.Range("K" & X & ":K" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet117.Cells(X, 12) = "Final figure"
            ' worksheet1.Cells(X, 1).Font.size = 9
            'worksheet1.Cells(X, y).Font.BOLD = True
            worksheet117.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet117.Range("L" & X & ":L" & X).MergeCells = True
            worksheet117.Range("L" & X & ":L" & X).VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet117.Cells(X, 13) = "Projection Vs actual %"
            ' worksheet1.Cells(X, 1).Font.size = 9
            'worksheet1.Cells(X, y).Font.BOLD = True
            worksheet117.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet117.Range("M" & X & ":M" & X).MergeCells = True
            worksheet117.Range("M" & X & ":M" & X).VerticalAlignment = XlVAlign.xlVAlignCenter

            _FirstChr = X
            _Chr = 97
            For m1 = 1 To 13
                worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1
            Next

            worksheet117.Range("B2:B2").Interior.Color = RGB(197, 217, 241)
            worksheet117.Range("E2:E2").Interior.Color = RGB(0, 176, 240)
            worksheet117.Range("H2:H2").Interior.Color = RGB(255, 192, 0)

            X = X + 1
            worksheet117.Rows(X).Font.name = "Times New Roman"
            worksheet117.Rows(X).Font.BOLD = True
            worksheet117.Rows(X).Font.size = 9
            _Chr = 98
            For m1 = 2 To 10
                If m1 = 2 Or m1 = 5 Or m1 = 8 Then
                    worksheet117.Cells(X, m1) = "S.J"
                    worksheet117.Cells(X, m1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                  
                ElseIf m1 = 3 Or m1 = 6 Or m1 = 9 Then
                    worksheet117.Cells(X, m1) = "RIB"
                    worksheet117.Cells(X, m1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter


                ElseIf m1 = 4 Or m1 = 7 Or m1 = 10 Then
                    worksheet117.Cells(X, m1) = "TOTAL"
                    worksheet117.Cells(X, m1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).MergeCells = True
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    'worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(255, 192, 0)
                End If
                If m1 = 2 Or m1 = 3 Or m1 = 4 Then
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)

                ElseIf m1 = 5 Or m1 = 6 Or m1 = 7 Then
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(0, 176, 240)
                ElseIf m1 = 8 Or m1 = 9 Or m1 = 10 Then
                    worksheet117.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(255, 192, 0)

                End If

                _Chr = _Chr + 1
            Next

            _FirstChr = X
            _Chr = 97
            For m1 = 1 To 13
                worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1
            Next
            X = X + 1

            SQL = "select * from M14Retailer order by m14code"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            I = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                worksheet117.Rows(X).Font.size = 9
                ' worksheet117.Cells(X, 1).Font.BOLD = True
                worksheet117.Rows(X).rowheight = 20
                worksheet117.Cells(X, 1) = T01.Tables(0).Rows(I)("m14NAME")
                worksheet117.Cells(X, 1).Font.size = 9
                worksheet117.Cells(X, 1).Font.BOLD = True
                worksheet117.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet117.Range("A" & X & ":A" & X).MergeCells = True
                worksheet117.Range("A" & X & ":A" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                '----------------------------------------------------------------------------------------
                SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department  where M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _Fromdate & "' and '" & _ToDate & "' and m15Fabric_Type='SINGLE JERSEY' group by M14Name"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    worksheet117.Cells(X, 5) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    worksheet117.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Range("E" & X & ":E" & X).MergeCells = True
                    worksheet117.Range("E" & X & ":E" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                End If

                range1 = worksheet117.Cells(X, 5)
                range1.NumberFormat = "0.00"


                SQL = "select sum(M07Qty_Mtr) as M07Qty_Mtr from View_DelivaryForcus inner join M16Quality_RCode on M07Material=M16Material inner join M15Preset on M15Quality=M16Quality inner join M13Biz_Unit on M13Merchant= M07Merchant  inner join M14Retailer on M14Code=M13Department  where M14Name='" & T01.Tables(0).Rows(I)("M14Name") & "' and M07Date between '" & _Fromdate & "' and '" & _ToDate & "' and m15Fabric_Type IN ('RIB','INTERLOCK') group by M14Name"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(dsUser) Then
                    worksheet117.Cells(X, 6) = dsUser.Tables(0).Rows(0)("M07Qty_Mtr")
                    worksheet117.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet117.Range("F" & X & ":F" & X).MergeCells = True
                    worksheet117.Range("F" & X & ":F" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                End If
                range1 = worksheet117.Cells(X, 6)
                range1.NumberFormat = "0.00"

                worksheet117.Cells(X, 7) = "=" & "E" & X & "+F" & X
                worksheet117.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet117.Range("g" & X & ":g" & X).MergeCells = True
                worksheet117.Range("g" & X & ":g" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet117.Cells(X, 7)
                range1.NumberFormat = "0.00"

                worksheet117.Cells(X, 8) = "=" & "B" & X & "-E" & X
                worksheet117.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet117.Range("H" & X & ":H" & X).MergeCells = True
                worksheet117.Range("H" & X & ":H" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet117.Cells(X, 8)
                range1.NumberFormat = "0.00"

                worksheet117.Cells(X, 9) = "=" & "C" & X & "-F" & X
                worksheet117.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet117.Range("I" & X & ":I" & X).MergeCells = True
                worksheet117.Range("I" & X & ":I" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet117.Cells(X, 9)
                range1.NumberFormat = "0.00"

                worksheet117.Cells(X, 10) = "=" & "H" & X & "+I" & X
                worksheet117.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet117.Range("J" & X & ":J" & X).MergeCells = True
                worksheet117.Range("J" & X & ":J" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet117.Cells(X, 10)
                range1.NumberFormat = "0.00"

                worksheet117.Cells(X, 12) = "=" & "G" & X & "+K" & X
                worksheet117.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet117.Range("L" & X & ":L" & X).MergeCells = True
                worksheet117.Range("L" & X & ":L" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet117.Cells(X, 12)
                range1.NumberFormat = "0.00"

                worksheet117.Cells(X, 4) = "=" & "B" & X & "+C" & X
                worksheet117.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet117.Range("D" & X & ":D" & X).MergeCells = True
                worksheet117.Range("D" & X & ":D" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet117.Cells(X, 4)
                range1.NumberFormat = "0.00"

                worksheet117.Cells(X, 13) = "=" & "L" & X & "/D" & X
                worksheet117.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet117.Range("M" & X & ":M" & X).MergeCells = True
                worksheet117.Range("M" & X & ":M" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                range1 = worksheet117.Cells(X, 13)
                range1.NumberFormat = "0%"

                _FirstChr = X
                _Chr = 97
                For m1 = 1 To 13
                    worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet117.Range(Chr(_Chr) & _FirstChr, Chr(_Chr) & _FirstChr).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    _Chr = _Chr + 1
                Next

                X = X + 1

                I = I + 1
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
            xlCharts = worksheet117.ChartObjects

            Dim _Chartlocation As Integer

            _Chartlocation = (X + 5) * 10

            myChart = xlCharts.Add(10, _Chartlocation, 455, 300)
            chartPage = myChart.Chart

            chartRange = worksheet117.Range("A4", "A7")
            chartRange1 = worksheet117.Range("G4", "G7")
            'chartRange = worksheet1.Range("H8", "K" & (X - 1))
            'chartRange = worksheet1.Range("H8:K39", "A9:A39")
            ' chartPage.SetSourceData(Source:=chartRange)
            t_SerCol = chartPage.SeriesCollection
            t_Series = t_SerCol.NewSeries
            With t_Series
                .Name = "Quoted Qty(m)"
                t_Series.XValues = chartRange '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
                t_Series.Values = chartRange1 '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

            End With

            chartRange = worksheet117.Range("D4", "D7")
            t_SerCol = chartPage.SeriesCollection
            t_Series = t_SerCol.NewSeries
            With t_Series
                .Name = "Projection (m) "
                t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

            End With
            t_Series.Border.Color = RGB(0, 153, 0)
            ' chartPage.Refresh()
            chartPage.SeriesCollection(2).Interior.Color = RGB(0, 153, 0)
            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xl3DColumn
            ' chartPage.ChartType = Microsoft.Office.Interop.Excel.XlCharttool.
            MsgBox("Report genarated successfully", MsgBoxStyle.Information, "Information .......")


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
            ' worksheet1.Cells(4, 5) = _Fail_Batch
            'worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ' MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
            ' MsgBox(_Fail_Batch)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try

    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Create_File()
    End Sub
End Class