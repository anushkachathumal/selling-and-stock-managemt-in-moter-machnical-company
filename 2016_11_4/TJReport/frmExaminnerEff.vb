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
Public Class frmExaminnerEff
    Dim Clicked As String
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter


    Private Sub frmExaminnerEff_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim m03 As DataSet
        Dim Sql As String

        Try
            'Load Production Quality

            Sql = "select M13Name as [Month] from M13Month"
            m03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboMonth
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
    End Sub


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' cboDep.ToggleDropdown()
        cmdEdit.Enabled = True
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        'cmdSave.Enabled = False
        cmdAdd.Focus()
    End Sub

    Private Sub txtyear_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtyear.KeyUp
        If e.KeyCode = 13 Then
            cboMonth.ToggleDropdown()

        End If
    End Sub

    Private Sub txtyear_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtyear.ValueChanged
        If IsNumeric(txtyear.Text) Then
        Else
            MsgBox("Please enter the correct year", MsgBoxStyle.Information, "Information ......")

        End If
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
        Dim _Monthh As Integer
        Dim n_date As Date
        Dim _Fromdate As String


        Try
            If Trim(txtyear.Text) <> "" And cboMonth.Text <> "" Then
                workbooks.Application.Sheets.Add()
                Dim sheets1 As Sheets = workbook.Worksheets
                Dim worksheet11 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
                worksheet11.Name = "Examinner wise efficency Report"
                worksheet11.Cells(2, 3) = "Textured Jersey PLC"
                worksheet11.Rows(2).Font.Bold = True
                worksheet11.Rows(2).Font.size = 26
                worksheet11.Range("A2:J2").MergeCells = True
                worksheet11.Range("A2:J2").VerticalAlignment = XlVAlign.xlVAlignCenter


                worksheet11.Cells(4, 1) = "Examinner wise efficency Report "
                worksheet11.Rows(4).Font.Bold = True
                worksheet11.Rows(4).Font.size = 10

                SQL = "select * from M13month where M13Name='" & cboMonth.Text & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then
                    _Monthh = T01.Tables(0).Rows(0)("M13Code")
                End If

                Dim daysInJuly As Integer = System.DateTime.DaysInMonth(txtyear.Text, _Monthh)
                Dim n_date1 As Date

                _Fromdate = _Monthh & "/" & daysInJuly & "/" & txtyear.Text

                n_date1 = CDate(_Fromdate).AddDays(-364)
                n_date = n_date1

                worksheet11.Cells(7, 2) = "Employee Name"
                worksheet11.Rows(7).Font.Bold = True
                worksheet11.Rows(7).Font.size = 10
                worksheet11.Columns(2).columnwidth = 28
                worksheet11.Cells(7, 3).WrapText = True
                worksheet11.Cells(7, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet11.Cells(7, 2).VerticalAlignment = XlVAlign.xlVAlignCenter
                ' worksheet11.Cells(7, 2).Orientation = 90
                worksheet11.Rows(7).rowheight = 20.25

                worksheet11.Range("B7:B7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet11.Range("B7:B7").Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet11.Range("B7:B7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet11.Range("B7:B7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous

                worksheet11.Range("b7:b7").Interior.Color = RGB(255, 193, 0)
                X = 7
                Dim Z As Integer
                Dim _FromTime As Date
                Dim _ToTime As Date

                Z = 3
                For i = 1 To 12
                    worksheet11.Cells(X, Z) = MonthName(Month(n_date)) & "-" & Year(n_date)
                    worksheet11.Cells(7, Z).WrapText = True
                    worksheet11.Cells(X, Z).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet11.Cells(X, Z).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet11.Columns(Z).columnwidth = 10
                    daysInJuly = System.DateTime.DaysInMonth(Year(n_date), Month(n_date))

                    n_date = CDate(n_date).AddDays(+daysInJuly)

                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    '  worksheet1.Cells(X, Z).Orientation = 90

                    worksheet11.Cells(7, Z).Interior.Color = RGB(255, 192, 0)
                    Z = Z + 1
                    '  X = X + 1
                Next

                SQL = "select T01InsEPF,max(FirstName) as FirstName from  T01Transaction_Header inner join Users on EPFNo=T01InsEPF where T01Date between '" & n_date1 & "' and '" & _Fromdate & "' and Department='01' group by T01InsEPF"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                i = 0
                X = 8
                Dim y As Integer
                Dim _TotalRoll As Integer
                Dim _TotalInspecTime As Integer
                Dim _AvgTime As Double
                Dim _TOBEIRoll As Double
                Dim n_Eff As Double

                For Each DTRow2 As DataRow In T01.Tables(0).Rows
                    worksheet11.Cells(X, 2) = T01.Tables(0).Rows(i)("FirstName")
                    worksheet11.Rows(X).Font.size = 10

                    If i = 0 Then
                        worksheet11.Cells(X, 2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet11.Cells(X, 2).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet11.Cells(X, 2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        worksheet11.Cells(X, 2).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    Else
                        worksheet11.Cells(X, 2).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlDot
                        worksheet11.Cells(X, 2).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlDot
                        worksheet11.Cells(X, 2).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDot
                        worksheet11.Cells(X, 2).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDot
                    End If
                    daysInJuly = System.DateTime.DaysInMonth(txtyear.Text, _Monthh)


                    _Fromdate = _Monthh & "/" & daysInJuly & "/" & txtyear.Text

                    n_date1 = CDate(_Fromdate).AddDays(-364)
                    n_date = n_date1

                    n_date = n_date1
                    Z = 3
                    For y = 1 To 12
                        daysInJuly = System.DateTime.DaysInMonth(Year(n_date), Month(n_date))
                        n_date1 = n_date
                        n_date = CDate(n_date).AddDays(+daysInJuly)

                        If y = 1 Then
                            _FromTime = n_date1 & " " & "07:30 AM"
                        Else
                            _FromTime = n_date1 & " " & "07:30 AM"
                        End If
                        ' _ToTime = System.DateTime.FromOADate(CDate(n_date).ToOADate + 1)
                        _ToTime = n_date
                        ' If y = 1 Then
                        _ToTime = _ToTime & " " & "07:30 AM"
                        'Else
                        '_ToTime = _ToTime
                        'End If
                        '----------------------------------------------------------------------------
                        _TotalRoll = 0
                        SQL = "select T04Emp,sum(T04TOTAL) as T04TOTAL from T04Summery where T04Time between '" & _FromTime & "' and '" & _ToTime & "' and T04Emp='" & Trim(T01.Tables(0).Rows(i)("T01InsEPF")) & "' group by T04Emp"
                        T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        If isValidDataset(T02) Then
                            _TotalRoll = T02.Tables(0).Rows(0)("T04TOTAL")
                        End If

                        _TotalInspecTime = 0
                        SQL = "select sum(T18Time) as T18Time from T18Downtime where T18User='" & Trim(T01.Tables(0).Rows(i)("T01InsEPF")) & "' and T18Timein between '" & _FromTime & "' and '" & _ToTime & "' group by T18User"
                        T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                        If isValidDataset(T02) Then
                            _TotalInspecTime = T02.Tables(0).Rows(0)("T18Time")
                        End If

                        _AvgTime = 0
                        If _TotalRoll > 0 Then
                            _AvgTime = _TotalInspecTime / _TotalRoll
                        End If
                        _TOBEIRoll = 0
                        'TO BE INSPECTED ROLL
                        If _AvgTime > 0 Then
                            _TOBEIRoll = (630 * 20) / _AvgTime
                        End If

                        n_Eff = 0
                        If _TOBEIRoll > 0 Then
                            n_Eff = _TotalRoll / _TOBEIRoll

                            ' n_Eff = n_Eff * 100
                        End If

                        worksheet11.Cells(X, Z) = n_Eff
                        worksheet11.Cells(X, Z).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        worksheet11.Cells(X, Z).VerticalAlignment = XlVAlign.xlVAlignCenter
                        range1 = worksheet11.Cells(X, Z)
                        range1.NumberFormat = "0.00%"
                        If i = 0 Then
                            worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        Else
                            worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlDot
                            worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDot
                            worksheet11.Cells(X, Z).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDot
                        End If
                        Z = Z + 1
                    Next
                    X = X + 1
                    i = i + 1
                Next
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class