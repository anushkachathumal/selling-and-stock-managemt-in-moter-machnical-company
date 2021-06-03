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
Public Class frmPurchasing_Report
    Dim strLine As String
    Dim strLineflu As String
    Dim strDash As String
    Dim StrDisCode As String
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter

    Dim Clicked As String
    'Dim exc As New Application
    'Dim workbooks As Workbooks = exc.Workbooks
    'Dim workbook As _Workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    'Dim sheets As Sheets = Workbook.Worksheets
    'Dim worksheet1 As _Worksheet = CType(Sheets.Item(1), _Worksheet)


    Private Sub frmPurchasing_Report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtTodate.Text = Today
        txtFromDate.Text = Today

    End Sub


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

    Function Total_ValuePurchasing()
        Dim exc As New Application
        Dim workbooks As Workbooks = exc.Workbooks
        Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
        Dim sheets1 As Sheets = workbook.Worksheets
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

        Dim I As Integer
        Dim Y As Integer
        Dim X As Integer
        Dim range1 As Range


        worksheet1.Name = "Total Value"
        worksheet1.Columns("A").ColumnWidth = 4
        worksheet1.Columns("b").ColumnWidth = 14
        worksheet1.Columns("C").ColumnWidth = 14
        worksheet1.Columns("D").ColumnWidth = 14
        worksheet1.Columns("E").ColumnWidth = 14
        worksheet1.Columns("G").ColumnWidth = 14
        worksheet1.Columns("H").ColumnWidth = 14
        worksheet1.Columns("I").ColumnWidth = 14

        worksheet1.Cells(1, 2) = "Value of Purchasing, Consumption, End Stock & TODS "
        worksheet1.Cells(1, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        ' worksheet1.Columns("A").ColumnWidth = 12
        worksheet1.Range("B1:i1").MergeCells = True
        worksheet1.Range("B1:i1").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Range("B1:i1").Interior.Color = RGB(182, 221, 232)
        worksheet1.Rows(1).Font.size = 15
        worksheet1.Rows(1).Font.bold = True


        worksheet1.Rows(1).rowheight = 30.25

        worksheet1.Cells(2, 2) = "Value in USD"
        ' worksheet1.Columns("A").ColumnWidth = 12
        worksheet1.Cells(2, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Range("B2:E2").MergeCells = True
        worksheet1.Range("B2:E2").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Range("B2:E2").Interior.Color = RGB(255, 0, 0)

        worksheet1.Cells(2, 7) = "Value in USD"
        worksheet1.Cells(2, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
        ' worksheet1.Columns("A").ColumnWidth = 12
        worksheet1.Range("G2:I2").MergeCells = True
        worksheet1.Range("G2:I2").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Range("G2:I2").Interior.Color = RGB(255, 0, 0)


        worksheet1.Rows(2).Font.size = 15
        worksheet1.Rows(2).Font.bold = True
        worksheet1.Rows(2).rowheight = 18.25

        worksheet1.Rows(3).Font.size = 10
        worksheet1.Rows(3).Font.bold = True

        worksheet1.Cells(3, 5) = "Grand Total"
        worksheet1.Cells(3, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
        ' worksheet1.Columns("A").ColumnWidth = 12
        worksheet1.Range("B3:E3").MergeCells = True
        worksheet1.Range("B3:E3").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Range("B3:E3").Interior.Color = RGB(255, 204, 255)

        worksheet1.Cells(3, 7) = "Grand Total"
        worksheet1.Cells(3, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
        ' worksheet1.Columns("A").ColumnWidth = 12
        worksheet1.Range("G3:I3").MergeCells = True
        worksheet1.Range("G3:I3").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Range("G3:I3").Interior.Color = RGB(255, 204, 255)

        worksheet1.Rows(4).Font.size = 8
        worksheet1.Rows(4).Font.bold = True

        worksheet1.Cells(4, 2) = "Month"
        worksheet1.Cells(4, 2).WrapText = True
        worksheet1.Cells(4, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(4, 2).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(4, 2).Orientation = 0


        worksheet1.Cells(4, 3) = "Purchasing "
        worksheet1.Cells(4, 3).WrapText = True
        worksheet1.Cells(4, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(4, 3).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(4, 3).Orientation = 0


        worksheet1.Cells(4, 4) = "Consumption"
        worksheet1.Cells(4, 4).WrapText = True
        worksheet1.Cells(4, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(4, 4).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(4, 4).Orientation = 0


        worksheet1.Cells(4, 5) = "End Stock"
        worksheet1.Cells(4, 5).WrapText = True
        worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(4, 5).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(4, 5).Orientation = 0


        worksheet1.Cells(4, 7) = "Month"
        worksheet1.Cells(4, 7).WrapText = True
        worksheet1.Cells(4, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(4, 7).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(4, 7).Orientation = 0


        worksheet1.Cells(4, 8) = "Purchasing TODS"
        worksheet1.Cells(4, 8).WrapText = True
        worksheet1.Cells(4, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(4, 8).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(4, 8).Orientation = 0


        worksheet1.Cells(4, 9) = "End Stock TODS"
        worksheet1.Cells(4, 9).WrapText = True
        worksheet1.Cells(4, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(4, 9).VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(4, 9).Orientation = 0


        worksheet1.Range("B4:B4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("c4:c4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("d4:d4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("e4:e4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("g4:g4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("h4:h4").Interior.Color = RGB(255, 192, 0)
        worksheet1.Range("i4:i4").Interior.Color = RGB(255, 192, 0)

        X = 4
        worksheet1.Range("b" & X & ":b" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b" & X & ":b" & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b" & X & ":b" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("b" & X & ":b" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("c" & X & ":c" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c" & X & ":c" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("c" & X & ":c" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("d" & X & ":d" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d" & X & ":d" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("d" & X & ":d" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("e" & X & ":e" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e" & X & ":e" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("e" & X & ":e" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("g" & X & ":g" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g" & X & ":g" & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g" & X & ":g" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("g" & X & ":g" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("h" & X & ":h" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h" & X & ":h" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("h" & X & ":h" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet1.Range("i" & X & ":i" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i" & X & ":i" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        worksheet1.Range("i" & X & ":i" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous


        Dim daysInFeb As Integer = System.DateTime.DaysInMonth(Microsoft.VisualBasic.Year(txtFromDate.Text), Microsoft.VisualBasic.Month(txtFromDate.Text))
        n_Date = Month(txtFromDate.Text) & "/" & daysInFeb & "/" & Year(txtFromDate.Text)
        n_Date = CDate(txtFromDate.Text).AddDays(-365)

        X = 5
        Dim daysInFeb1 As Integer
        For I = 1 To 12
            daysInFeb1 = System.DateTime.DaysInMonth(Microsoft.VisualBasic.Year(n_Date), Microsoft.VisualBasic.Month(n_Date))
            N_Date1 = Month(n_Date) & "/" & daysInFeb1 & "/" & Year(n_Date)
            n_Date = Month(n_Date) & "/1/" & Year(n_Date)

            worksheet1.Cells(X, 2) = MonthName(Month(n_Date)) & "-" & Year(n_Date)
            worksheet1.Rows(X).Font.size = 8
            worksheet1.Rows(X).Font.bold = True
            worksheet1.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignRight


            worksheet1.Cells(X, 7) = MonthName(Month(n_Date)) & "-" & Year(n_Date)
            worksheet1.Rows(X).Font.size = 8
            worksheet1.Rows(X).Font.bold = True
            worksheet1.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight


            'PURCHASING
            SQL = "select sum(Qty) as Qty from ZDCA_PURCHASE where Posting_Date between '" & n_Date & "' and '" & N_Date1 & "' and Status='P' group by Status"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 3) = T01.Tables(0).Rows(0)("Qty")
                ' worksheet1.Rows(X).Font.size = 10
                worksheet1.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 3)
                range1.NumberFormat = "0"
            End If

            'CONSUMPTION
            SQL = "select sum(Qty) as Qty from ZDCA_PURCHASE where Posting_Date between '" & n_Date & "' and '" & N_Date1 & "' and Status='C' group by Status"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(X, 4) = T01.Tables(0).Rows(0)("Qty")
                ' worksheet1.Rows(X).Font.size = 10
                worksheet1.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet1.Cells(X, 4)
                range1.NumberFormat = "0"
            End If

            worksheet1.Range("b" & X & ":b" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & X & ":b" & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & X & ":b" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b" & X & ":b" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("c" & X & ":c" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & X & ":c" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c" & X & ":c" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("d" & X & ":d" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & X & ":d" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d" & X & ":d" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("e" & X & ":e" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & X & ":e" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e" & X & ":e" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("g" & X & ":g" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & X & ":g" & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & X & ":g" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g" & X & ":g" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("h" & X & ":h" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & X & ":h" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h" & X & ":h" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("i" & X & ":i" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & X & ":i" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i" & X & ":i" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            n_Date = CDate(n_Date).AddDays(+daysInFeb1)


            X = X + 1
        Next
    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Total_ValuePurchasing()

    End Sub
End Class