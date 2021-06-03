
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
Public Class frmDailyDye_Production
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

    Private Sub frmDailyDye_Production_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
        ' Call Daily_Boliout()

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

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        txtFromDate.Value = Today
        txtTodate.Value = Today
    End Sub

    Function Daily_Dye()
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
        Dim worksheet11 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
        worksheet11.Name = "Daily Dye Production"

        worksheet11.Columns("A").ColumnWidth = 22
        worksheet11.Columns("B").ColumnWidth = 14


        Dim daysInFeb As Integer = System.DateTime.DaysInMonth(Microsoft.VisualBasic.Year(txtFromDate.Text), Microsoft.VisualBasic.Month(txtFromDate.Text))
        n_Date = CDate(txtFromDate.Text).AddDays(-daysInFeb)
        n_Date = CDate(n_Date).AddDays(+1)
        '  n_Date = n_Date & " " & "7:30AM"
        ' N_Date1 = CDate(txtTodate.Text).AddDays(+1)
        ' N_Date1 = txtTodate.Text & " " & "7:30AM"
        Dim X As Integer
        X = 3
        Dim X_date As Date
        X_date = n_Date
        worksheet11.Range(worksheet11.Cells(2, 1), worksheet11.Cells(2, 2)).Merge()
        ' worksheet11.Cells(2, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

        For i = 1 To daysInFeb

            worksheet11.Cells(2, X) = X_date
            worksheet11.Cells(2, X).EntireColumn.NumberFormat = "dd-MMM"
            worksheet11.Cells(2, X).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet11.Rows(2).Font.Bold = True
            x_Date = CDate(x_Date).AddDays(+1)
            X = X + 1
        Next

        worksheet11.Range("A2:A2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("A2:A2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("B2:B2").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("B2:B2").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("B2:B2").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        X = 2
        ' For i = 1 To daysInFeb
        worksheet11.Range("C" & X & ":C" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("C" & X & ":C" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("C" & X & ":C" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("D" & X & ":D" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("D" & X & ":D" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("D" & X & ":D" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("E" & X & ":E" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("E" & X & ":E" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("E" & X & ":E" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("F" & X & ":F" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("F" & X & ":F" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("F" & X & ":F" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("G" & X & ":G" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("G" & X & ":G" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("G" & X & ":G" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("H" & X & ":H" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("H" & X & ":H" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("H" & X & ":H" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("I" & X & ":I" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("I" & X & ":I" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("I" & X & ":I" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("J" & X & ":J" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("J" & X & ":J" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("J" & X & ":J" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("K" & X & ":K" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("K" & X & ":K" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("K" & X & ":K" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("L" & X & ":L" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("L" & X & ":L" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("L" & X & ":L" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("M" & X & ":M" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("M" & X & ":M" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("M" & X & ":M" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("N" & X & ":N" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("N" & X & ":N" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("N" & X & ":N" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("O" & X & ":O" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("O" & X & ":O" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("O" & X & ":O" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("P" & X & ":P" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("P" & X & ":P" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("P" & X & ":P" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("Q" & X & ":Q" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("Q" & X & ":Q" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("Q" & X & ":Q" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("R" & X & ":R" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("R" & X & ":R" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("R" & X & ":R" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("S" & X & ":S" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("S" & X & ":S" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("S" & X & ":S" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("T" & X & ":T" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("T" & X & ":T" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("T" & X & ":T" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


        worksheet11.Range("U" & X & ":U" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("U" & X & ":U" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("U" & X & ":U" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("V" & X & ":V" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("V" & X & ":V" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("V" & X & ":V" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("W" & X & ":W" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("W" & X & ":W" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("W" & X & ":W" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("X" & X & ":X" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("X" & X & ":X" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("X" & X & ":X" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("Y" & X & ":Y" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("Y" & X & ":Y" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("Y" & X & ":Y" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        worksheet11.Range("Z" & X & ":Z" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("Z" & X & ":Z" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        worksheet11.Range("Z" & X & ":Z" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        If daysInFeb = 31 Then
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AF" & X & ":AF" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AF" & X & ":AF" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AF" & X & ":AF" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous


            worksheet11.Range("AG" & X & ":AG" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AG" & X & ":AG" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AG" & X & ":AG" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        ElseIf daysInFeb = 30 Then
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AF" & X & ":AF" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AF" & X & ":AF" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AF" & X & ":AF" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        ElseIf daysInFeb = 29 Then
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

        ElseIf daysInFeb = 28 Then
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AA" & X & ":AA" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ab" & X & ":Ab" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("Ac" & X & ":Ac" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AD" & X & ":AD" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet11.Range("AE" & X & ":AE" & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        End If

        SQL = "SELECT M04Week ,max(M04WeekNo) as [Ddate],count(T03TYPE) as nCount,m04year as Eyear FROM M04Lot INNER JOIN T03Machine ON M04Machine_No=T03Code " & _
           "INNER JOIN M01Dyeing_MC_Type ON T03TYPE=M01Code WHERE  T03TYPE IN ('01','02') and M04ETime between '" & N_Date1 & "' and '" & n_Date & "' and M04ProgrameType in ('N','R','S') GROUP BY m04year,M04Week order by m04year,M04Week"
        T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
        If isValidDataset(T01) Then


        End If





    End Function

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Daily_Dye()
    End Sub
End Class