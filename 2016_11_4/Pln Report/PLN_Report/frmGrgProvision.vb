
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
Public Class frmGrgProvision
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
        Dim _from As Date
        Dim _to As Date

        Dim vcWhere As String
        Dim _FirstRow As Integer

        Dim exc As New Application

        Dim workbooks As Workbooks = exc.Workbooks
        Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
        Dim range1 As Range
        Dim I As Integer
        Dim Z As Integer
        Try
            exc.Visible = True

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 11
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 15
            worksheet2.Columns("B").ColumnWidth = 10
            worksheet2.Columns("C").ColumnWidth = 20
            worksheet2.Columns("D").ColumnWidth = 10
            worksheet2.Columns("E").ColumnWidth = 10
            worksheet2.Columns("F").ColumnWidth = 10
            worksheet2.Columns("G").ColumnWidth = 15
            worksheet2.Columns("H").ColumnWidth = 12
            worksheet2.Columns("I").ColumnWidth = 12
            worksheet2.Columns("J").ColumnWidth = 12
            worksheet2.Columns("K").ColumnWidth = 15
            worksheet2.Columns("L").ColumnWidth = 12

            worksheet2.Columns("M").ColumnWidth = 8
            worksheet2.Columns("N").ColumnWidth = 8
            worksheet2.Columns("O").ColumnWidth = 8
            worksheet2.Columns("P").ColumnWidth = 8
            worksheet2.Columns("Q").ColumnWidth = 8
            worksheet2.Columns("R").ColumnWidth = 8
            worksheet2.Columns("S").ColumnWidth = 8
            worksheet2.Columns("T").ColumnWidth = 8
            worksheet2.Columns("U").ColumnWidth = 8
            worksheet2.Columns("V").ColumnWidth = 8
            worksheet2.Columns("W").ColumnWidth = 8
            worksheet2.Columns("X").ColumnWidth = 8
            worksheet2.Columns("Y").ColumnWidth = 8
            worksheet2.Columns("Z").ColumnWidth = 8
            worksheet2.Columns("AA").ColumnWidth = 10
            worksheet2.Columns("AB").ColumnWidth = 10
            worksheet2.Columns("AC").ColumnWidth = 10
            worksheet2.Columns("AD").ColumnWidth = 10
            worksheet2.Columns("AE").ColumnWidth = 10
            worksheet2.Columns("AF").ColumnWidth = 10
            worksheet2.Columns("AG").ColumnWidth = 10
            worksheet2.Columns("AH").ColumnWidth = 10
            worksheet2.Columns("AI").ColumnWidth = 10
            worksheet2.Columns("AJ").ColumnWidth = 10

            worksheet2.Cells(1, 1) = "Greige Provision vs orders analysis "
            worksheet2.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Range("A1:AL1").Interior.Color = RGB(197, 217, 241)
            ' worksheet2.Range("A2:M2").Interior.Color = RGB(197, 217, 241)
            worksheet2.Rows(1).Font.size = 13
            worksheet2.Rows(1).rowheight = 35
            worksheet2.Rows(1).Font.name = "Times New Roman"
            worksheet2.Rows(1).Font.BOLD = True
            worksheet2.Range("A1:AL1").MergeCells = True
            worksheet2.Range("A1:AL1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(2, 1) = "Base Data "
            worksheet2.Cells(2, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

            ' worksheet2.Range("A2:L1").Interior.Color = RGB(197, 217, 241)
            ' worksheet2.Range("A2:M2").Interior.Color = RGB(197, 217, 241)
            worksheet2.Rows(2).Font.size = 10
            worksheet2.Rows(2).rowheight = 20
            worksheet2.Rows(2).Font.name = "Times New Roman"
            worksheet2.Rows(2).Font.BOLD = True
            worksheet2.Range("A2:L2").MergeCells = True
            worksheet2.Range("A2:L2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(2, 13) = "Greige stock ageing "
            worksheet2.Cells(2, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("M2:Z2").MergeCells = True
            worksheet2.Range("M2:Z2").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(2, 27) = "Stock summary"
            worksheet2.Cells(2, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AA2:AD2").MergeCells = True
            worksheet2.Range("AA2:AD2").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(2, 31) = "Clearance Plan"
            worksheet2.Cells(2, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AE2:AL2").MergeCells = True
            worksheet2.Range("AE2:AL2").VerticalAlignment = XlVAlign.xlVAlignCenter



            Dim X As Integer
            Dim _Chr As Integer
            X = 2

            _Chr = 97
            For I = 1 To 38
                If I = 27 Then
                    _Chr = 97
                End If

                If I >= 27 Then
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                Else

                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                End If
                _Chr = _Chr + 1

            Next

            X = X + 1
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).rowheight = 35
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            worksheet2.Cells(X, 1) = "Quality #"
            worksheet2.Range("A3:A3").MergeCells = True
            worksheet2.Range("A3:A3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 2) = "Fabric Type "
            worksheet2.Range("B3:B3").MergeCells = True
            worksheet2.Range("B3:B3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 3) = "Preset/Non Preset "
            worksheet2.Range("C3:C3").MergeCells = True
            worksheet2.Range("C3:C3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 4) = "Width"
            worksheet2.Range("D3:D3").MergeCells = True
            worksheet2.Range("D3:D3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(X, 5) = "GSM"
            worksheet2.Range("E3:E3").MergeCells = True
            worksheet2.Range("E3:E3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(X, 6) = "C.factor"
            worksheet2.Range("F3:F3").MergeCells = True
            worksheet2.Range("F3:F3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 7) = "Shade Catogery"
            worksheet2.Range("G3:G3").MergeCells = True
            worksheet2.Range("G3:G3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 8) = "Relevant Planner"
            worksheet2.Range("H3:H3").MergeCells = True
            worksheet2.Range("H3:H3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 8)
            range1.WrapText = True

            worksheet2.Cells(X, 9) = "Relevant Merchant"
            worksheet2.Range("I3:I3").MergeCells = True
            worksheet2.Range("I3:I3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 9)
            range1.WrapText = True

            worksheet2.Cells(X, 10) = "Common Quality "
            worksheet2.Range("J3:J3").MergeCells = True
            worksheet2.Range("J3:J3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 10)
            range1.WrapText = True

            worksheet2.Cells(X, 11) = "Available greige orders "
            worksheet2.Range("K3:K3").MergeCells = True
            worksheet2.Range("K3:K3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 11)
            range1.WrapText = True

            worksheet2.Cells(X, 12) = "Special comment  "
            worksheet2.Range("L3:L3").MergeCells = True
            worksheet2.Range("L3:L3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 12)
            range1.WrapText = True

            worksheet2.Cells(X, 13) = "below 1 month (B)"
            worksheet2.Range("m3:m3").MergeCells = True
            worksheet2.Range("m3:m3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 13)
            range1.WrapText = True
            worksheet2.Range("m3:M3").Interior.Color = RGB(142, 169, 219)

            worksheet2.Cells(X, 14) = "Value ($)"
            worksheet2.Range("n3:n3").MergeCells = True
            worksheet2.Range("n3:n3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 14)
            range1.WrapText = True
            worksheet2.Range("n3:n3").Interior.Color = RGB(84, 130, 53)

            worksheet2.Cells(X, 15) = "Within1-2month ©"
            worksheet2.Range("o3:o3").MergeCells = True
            worksheet2.Range("o3:o3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 15)
            range1.WrapText = True
            worksheet2.Range("o3:o3").Interior.Color = RGB(142, 169, 219)



            worksheet2.Cells(X, 16) = "Value ($)"
            worksheet2.Range("p3:p3").MergeCells = True
            worksheet2.Range("p3:p3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 16)
            range1.WrapText = True
            worksheet2.Range("p3:p3").Interior.Color = RGB(84, 130, 53)


            worksheet2.Cells(X, 17) = "Within 2-3month (D)"
            worksheet2.Range("Q3:Q3").MergeCells = True
            worksheet2.Range("Q3:Q3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 17)
            range1.WrapText = True
            worksheet2.Range("Q3:Q3").Interior.Color = RGB(142, 169, 219)



            worksheet2.Cells(X, 18) = "Value ($)"
            worksheet2.Range("R3:R3").MergeCells = True
            worksheet2.Range("R3:R3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 18)
            range1.WrapText = True
            worksheet2.Range("R3:R3").Interior.Color = RGB(84, 130, 53)

            worksheet2.Cells(X, 19) = "Within 3-4 month (E)"
            worksheet2.Range("S3:S3").MergeCells = True
            worksheet2.Range("S3:S3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 19)
            range1.WrapText = True
            worksheet2.Range("S3:S3").Interior.Color = RGB(142, 169, 219)



            worksheet2.Cells(X, 20) = "Value ($)"
            worksheet2.Range("T3:T3").MergeCells = True
            worksheet2.Range("T3:T3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 20)
            range1.WrapText = True
            worksheet2.Range("T3:T3").Interior.Color = RGB(84, 130, 53)

            worksheet2.Cells(X, 21) = "Within 4-5 month (F)"
            worksheet2.Range("U3:U3").MergeCells = True
            worksheet2.Range("U3:U3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 21)
            range1.WrapText = True
            worksheet2.Range("U3:U3").Interior.Color = RGB(142, 169, 219)



            worksheet2.Cells(X, 22) = "Value ($)"
            worksheet2.Range("V3:V3").MergeCells = True
            worksheet2.Range("V3:V3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 22)
            range1.WrapText = True
            worksheet2.Range("V3:V3").Interior.Color = RGB(84, 130, 53)


            worksheet2.Cells(X, 23) = "Within 5-6 month (G)"
            worksheet2.Range("W3:W3").MergeCells = True
            worksheet2.Range("W3:W3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 23)
            range1.WrapText = True
            worksheet2.Range("W3:W3").Interior.Color = RGB(142, 169, 219)



            worksheet2.Cells(X, 24) = "Value ($)"
            worksheet2.Range("X3:X3").MergeCells = True
            worksheet2.Range("X3:X3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 24)
            range1.WrapText = True
            worksheet2.Range("X3:X3").Interior.Color = RGB(84, 130, 53)

            worksheet2.Cells(X, 25) = "Over six month (H)"
            worksheet2.Range("Y3:Y3").MergeCells = True
            worksheet2.Range("Y3:Y3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 25)
            range1.WrapText = True
            worksheet2.Range("y3:y3").Interior.Color = RGB(142, 169, 219)



            worksheet2.Cells(X, 26) = "Value ($)"
            worksheet2.Range("Z3:Z3").MergeCells = True
            worksheet2.Range("Z3:Z3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 26)
            range1.WrapText = True
            worksheet2.Range("Z3:Z3").Interior.Color = RGB(84, 130, 53)

            worksheet2.Cells(X, 27) = "Total stock (I)"
            worksheet2.Range("AA3:AA3").MergeCells = True
            worksheet2.Range("AA3:AA3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 27)
            range1.WrapText = True

            worksheet2.Cells(X, 28) = "Total Value ($)"
            worksheet2.Range("AB3:AB3").MergeCells = True
            worksheet2.Range("AB3:AB3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 28)
            range1.WrapText = True

            worksheet2.Cells(X, 29) = "Planned qty for Dyeing "
            worksheet2.Range("AC3:AC3").MergeCells = True
            worksheet2.Range("AC3:AC3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 29)
            range1.WrapText = True

            worksheet2.Cells(X, 30) = "Remaining qty "
            worksheet2.Range("AD3:AD3").MergeCells = True
            worksheet2.Range("AD3:AD3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 30)
            range1.WrapText = True

            _to = Today
            _from = _to.AddDays(-150)

            worksheet2.Cells(X, 31) = MonthName(Month(_to))
            worksheet2.Range("Ae3:Ae3").MergeCells = True
            worksheet2.Range("Ae3:Ae3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 31).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 31)
            range1.WrapText = True

            _from = _to.AddDays(+31)

            worksheet2.Cells(X, 32) = MonthName(Month(_from))
            worksheet2.Range("AF3:AF3").MergeCells = True
            worksheet2.Range("AF3:AF3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 32).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 32)
            range1.WrapText = True

            _from = _to.AddDays(+60)

            worksheet2.Cells(X, 33) = MonthName(Month(_from))
            worksheet2.Range("AG3:AG3").MergeCells = True
            worksheet2.Range("AG3:AG3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 33).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 33)
            range1.WrapText = True

            _from = _to.AddDays(+90)

            worksheet2.Cells(X, 34) = MonthName(Month(_from))
            worksheet2.Range("AH3:AH3").MergeCells = True
            worksheet2.Range("AH3:AH3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 34).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 34)
            range1.WrapText = True

            _from = _to.AddDays(+120)

            worksheet2.Cells(X, 35) = MonthName(Month(_from))
            worksheet2.Range("AI3:AI3").MergeCells = True
            worksheet2.Range("AI3:AI3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 35).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 35)
            range1.WrapText = True


            _from = _to.AddDays(+150)
            worksheet2.Cells(X, 36) = MonthName(Month(_from))
            worksheet2.Range("AJ3:AJ3").MergeCells = True
            worksheet2.Range("AJ3:AJ3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 36).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 36)
            range1.WrapText = True

            worksheet2.Cells(X, 37) = "Total"
            worksheet2.Range("AK3:AK3").MergeCells = True
            worksheet2.Range("AK3:AK3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 37).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 38) = "Balance "
            worksheet2.Range("Al3:Al3").MergeCells = True
            worksheet2.Range("Al3:Al3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 38).HorizontalAlignment = XlHAlign.xlHAlignCenter





            X = X + 1
            If cboMaterial.Text <> "" And cboQuality.Text <> "" Then

            ElseIf cboMaterial.Text <> "" Then
                vcWhere = "M28Quality='" & cboMaterial.Text & "' and M28SLocation='2040'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "RPT1"), New SqlParameter("@vcWhereClause1", vcWhere))
            ElseIf cboQuality.Text <> "" Then
                vcWhere = "M28Merchant='" & cboQuality.Text & "' and M28SLocation='2040'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "RPT1"), New SqlParameter("@vcWhereClause1", vcWhere))
            Else
                T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "RPT"))
            End If
            Z = 0
            Dim Z1 As Integer
            Dim _Shade As String
            Dim _30Class As String
            Dim _TobeDeliverd As Double
            Dim _ConFact As Double
            Dim characterToRemove As String
            Dim tmp30 As String

            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _FirstRow = X
                worksheet2.Rows(X).Font.size = 8
                worksheet2.Rows(X).Font.name = "Times New Roman"
                _Shade = ""
                _TobeDeliverd = 0
                _30Class = ""
                _ConFact = 0
                worksheet2.Cells(X, 1) = T01.Tables(0).Rows(Z)("M28Quality")
                worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If cboQuality.Text <> "" Then
                    vcWhere = "M28Quality='" & T01.Tables(0).Rows(Z)("M28Quality") & "' and M28Merchant='" & cboQuality.Text & "' "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
                Else
                    vcWhere = "M28Quality='" & T01.Tables(0).Rows(Z)("M28Quality") & "' "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
                End If
                Z1 = 0
                For Each DTRow4 As DataRow In M01.Tables(0).Rows
                    worksheet2.Cells(X, 1) = T01.Tables(0).Rows(Z)("M28Quality")
                    worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet2.Rows(X).Font.size = 8
                    worksheet2.Rows(X).Font.name = "Times New Roman"
                    vcWhere = "M22Quality='" & M01.Tables(0).Rows(Z1)("M28Quality") & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "TEC"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 2) = M02.Tables(0).Rows(0)("M22Fabric_Type")
                        worksheet2.Cells(X, 4) = M02.Tables(0).Rows(0)("M22Userble_Width")
                        range1 = worksheet2.Cells(X, 4)
                        range1.NumberFormat = "0.00"
                        worksheet2.Cells(X, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet2.Cells(X, 5) = M02.Tables(0).Rows(0)("M22Fabric_Weight")
                        range1 = worksheet2.Cells(X, 5)
                        range1.NumberFormat = "0.00"
                        worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet2.Cells(X, 6) = M02.Tables(0).Rows(0)("M22Con_Fact")
                        range1 = worksheet2.Cells(X, 6)
                        range1.NumberFormat = "0.00"
                        worksheet2.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        _ConFact = M02.Tables(0).Rows(0)("M22Con_Fact")


                        If Microsoft.VisualBasic.Left(Trim(T01.Tables(0).Rows(Z)("M28Quality")), 2) = "Y1" Then
                            worksheet2.Cells(X, 3) = "Non Preset"
                        ElseIf Microsoft.VisualBasic.Left(Trim(T01.Tables(0).Rows(Z)("M28Quality")), 2) = "Y3" Then
                            worksheet2.Cells(X, 3) = "Preset"
                        Else
                            If M02.Tables(0).Rows(0)("M22Yarn_Cons") < 100 Then
                                worksheet2.Cells(X, 3) = "Preset"
                            Else
                                worksheet2.Cells(X, 3) = "Non Preset"
                            End If
                        End If

                    End If

                  

                    worksheet2.Cells(X, 7) = M01.Tables(0).Rows(Z1)("M23Shade")
                    worksheet2.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    _Shade = M01.Tables(0).Rows(Z1)("M23Shade")

                    'COMMENT BY SURANGA ON 2016.7.16
                    'vcWhere = "M23Material='" & M01.Tables(0).Rows(Z1)("M2820Class") & "'"
                    'M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
                    'If isValidDataset(M02) Then
                    '    worksheet2.Cells(X, 7) = M02.Tables(0).Rows(0)("M23Shade")
                    '    worksheet2.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    '    _Shade = M02.Tables(0).Rows(0)("M23Shade")
                    'End If

                    worksheet2.Cells(X, 9) = M01.Tables(0).Rows(Z1)("M28Merchant")

                    vcWhere = "M29Merchant='" & Trim(M01.Tables(0).Rows(Z1)("M28Merchant")) & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "PLN"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 8) = M02.Tables(0).Rows(0)("M29Name")

                    End If

                    vcWhere = "M26Quality20='" & Trim(M01.Tables(0).Rows(Z1)("M28Quality")) & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "ALT"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        worksheet2.Cells(X, 10) = M02.Tables(0).Rows(0)("M26Quality30")

                    End If
                    Dim _AvalableGrg_Order As String

                    _AvalableGrg_Order = ""
                    vcWhere = "M28Quality='" & M01.Tables(0).Rows(Z1)("M28Quality") & "' and M28Merchant='" & Trim(M01.Tables(0).Rows(Z1)("M28Merchant")) & "'  and M28SLocation='2040' and left(M28Sales_Order,3)='200'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "LST2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    I = 0
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        If I = 0 Then
                            _AvalableGrg_Order = Trim(M02.Tables(0).Rows(I)("M28Sales_Order"))
                        Else
                            _AvalableGrg_Order = _AvalableGrg_Order & "/" & Trim(M02.Tables(0).Rows(I)("M28Sales_Order"))
                        End If
                        I = I + 1
                    Next

                    worksheet2.Cells(X, 11) = _AvalableGrg_Order

                    Dim Diff As TimeSpan
                    Dim _DateIN As Date
                    Dim _DateOUT As Date

                    vcWhere = "M28Quality='" & M01.Tables(0).Rows(Z1)("M28Quality") & "' and M28Merchant='" & Trim(M01.Tables(0).Rows(Z1)("M28Merchant")) & "' AND M23Shade='" & _Shade & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "LST3"), New SqlParameter("@vcWhereClause1", vcWhere))
                    I = 0
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        Dim _XcelCell_Value As Double

                        _XcelCell_Value = 0
                        _DateIN = M02.Tables(0).Rows(I)("M28Date")

                        _DateOUT = txtTodate.Text

                        Diff = _DateOUT.Subtract(_DateIN)
                        If Diff.Days < 30 Then
                            range1 = CType(worksheet2.Cells(X, 13), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("M28Qty")
                            worksheet2.Cells(X, 13) = _XcelCell_Value
                            range1 = worksheet2.Cells(X, 13)
                            range1.NumberFormat = "0"
                            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight

                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 13), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("Total")

                            worksheet2.Cells(X, 14) = _XcelCell_Value ' * M01.Tables(0).Rows(Z1)("M28SMC")
                            range1 = worksheet2.Cells(X, 14)
                            range1.NumberFormat = "0.00"
                            worksheet2.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignRight

                        ElseIf Diff.Days >= 30 And Diff.Days < 60 Then
                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 15), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("M28Qty")
                            worksheet2.Cells(X, 15) = _XcelCell_Value

                            range1 = worksheet2.Cells(X, 15)
                            range1.NumberFormat = "0"
                            worksheet2.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignRight

                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 16), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("Total")
                            worksheet2.Cells(X, 16) = _XcelCell_Value

                            ' worksheet2.Cells(X, 16) = M02.Tables(0).Rows(I)("Total") ' * M01.Tables(0).Rows(Z1)("M28SMC")
                            range1 = worksheet2.Cells(X, 16)
                            range1.NumberFormat = "0.00"
                            worksheet2.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignRight


                        ElseIf Diff.Days >= 60 And Diff.Days < 90 Then
                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 17), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("M28Qty")
                            worksheet2.Cells(X, 17) = _XcelCell_Value

                            ' worksheet2.Cells(X, 17) = M02.Tables(0).Rows(I)("M28Qty")
                            range1 = worksheet2.Cells(X, 17)
                            range1.NumberFormat = "0"
                            worksheet2.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight

                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 18), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("total")
                            worksheet2.Cells(X, 18) = _XcelCell_Value

                            ' worksheet2.Cells(X, 18) = M02.Tables(0).Rows(I)("Total") ' * M01.Tables(0).Rows(Z1)("M28SMC")
                            range1 = worksheet2.Cells(X, 18)
                            range1.NumberFormat = "0.00"
                            worksheet2.Cells(X, 18).HorizontalAlignment = XlHAlign.xlHAlignRight

                        ElseIf Diff.Days >= 90 And Diff.Days < 120 Then
                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 19), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("M28Qty")
                            worksheet2.Cells(X, 19) = _XcelCell_Value

                            '  worksheet2.Cells(X, 19) = M02.Tables(0).Rows(I)("M28Qty")
                            range1 = worksheet2.Cells(X, 19)
                            range1.NumberFormat = "0"
                            worksheet2.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignRight

                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 20), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("total")
                            worksheet2.Cells(X, 20) = _XcelCell_Value
                            ' worksheet2.Cells(X, 20) = M02.Tables(0).Rows(I)("Total") ' * M01.Tables(0).Rows(Z1)("M28SMC")
                            range1 = worksheet2.Cells(X, 20)
                            range1.NumberFormat = "0.00"
                            worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignRight


                        ElseIf Diff.Days >= 120 And Diff.Days < 150 Then
                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 21), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("M28Qty")
                            worksheet2.Cells(X, 21) = _XcelCell_Value

                            '  worksheet2.Cells(X, 21) = M02.Tables(0).Rows(I)("M28Qty")
                            range1 = worksheet2.Cells(X, 21)
                            range1.NumberFormat = "0"
                            worksheet2.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignRight


                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 22), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("total")
                            worksheet2.Cells(X, 22) = _XcelCell_Value
                            ' worksheet2.Cells(X, 22) = M02.Tables(0).Rows(I)("Total") ' * M01.Tables(0).Rows(Z1)("M28SMC")
                            range1 = worksheet2.Cells(X, 22)
                            range1.NumberFormat = "0.00"
                            worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignRight


                        ElseIf Diff.Days >= 150 And Diff.Days < 180 Then

                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 23), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("M28Qty")
                            worksheet2.Cells(X, 23) = _XcelCell_Value

                            '  worksheet2.Cells(X, 23) = M02.Tables(0).Rows(I)("M28Qty")
                            range1 = worksheet2.Cells(X, 23)
                            range1.NumberFormat = "0"
                            worksheet2.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignRight

                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 24), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("total")
                            worksheet2.Cells(X, 24) = _XcelCell_Value

                            '  worksheet2.Cells(X, 24) = M02.Tables(0).Rows(I)("Total") ' * M01.Tables(0).Rows(Z1)("M28SMC")
                            range1 = worksheet2.Cells(X, 24)
                            range1.NumberFormat = "0.00"
                            worksheet2.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignRight


                        ElseIf Diff.Days >= 180 Then
                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 25), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("M28Qty")
                            worksheet2.Cells(X, 25) = _XcelCell_Value

                            ' worksheet2.Cells(X, 25) = M02.Tables(0).Rows(I)("M28Qty")
                            range1 = worksheet2.Cells(X, 25)
                            range1.NumberFormat = "0"
                            worksheet2.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignRight


                            _XcelCell_Value = 0
                            range1 = CType(worksheet2.Cells(X, 26), Microsoft.Office.Interop.Excel.Range)
                            _XcelCell_Value = range1.Value()
                            _XcelCell_Value = _XcelCell_Value + M02.Tables(0).Rows(I)("total")
                            worksheet2.Cells(X, 26) = _XcelCell_Value

                            '  worksheet2.Cells(X, 26) = M02.Tables(0).Rows(I)("Total") '* M01.Tables(0).Rows(Z1)("M28SMC")
                            range1 = worksheet2.Cells(X, 26)
                            range1.NumberFormat = "0.00"
                            worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignRight
                        End If

                        I = I + 1
                    Next

                    worksheet2.Cells(X, 27) = "=M" & X & "+O" & X & "+Q" & X & "+S" & X & "+U" & X & "+W" & X & "+Y" & X
                    range1 = worksheet2.Cells(X, 27)
                    range1.NumberFormat = "0"
                    worksheet2.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignRight

                    worksheet2.Cells(X, 28) = "=N" & X & "+P" & X & "+R" & X & "+T" & X & "+V" & X & "+X" & X & "+Z" & X
                    range1 = worksheet2.Cells(X, 28)
                    range1.NumberFormat = "0.00"
                    worksheet2.Cells(X, 28).HorizontalAlignment = XlHAlign.xlHAlignRight

                    Dim _30 As String
                    Dim _Bomwst As Double
                    _30 = ""
                    _Bomwst = 0
                    'GET 30CLASS
                    If _Shade = "-" Then
                        vcWhere = "M16Quality='" & M01.Tables(0).Rows(Z1)("M28Quality") & "'"
                    ElseIf _Shade = "Light" Then
                        vcWhere = "M16Quality='" & M01.Tables(0).Rows(Z1)("M28Quality") & "' and M14Grige='L'"
                    ElseIf _Shade = "Dark" Then
                        vcWhere = "M16Quality='" & M01.Tables(0).Rows(Z1)("M28Quality") & "' and M14Grige='D'"
                    End If
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "30C"), New SqlParameter("@vcWhereClause1", vcWhere))
                    I = 0
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        If I = 0 Then
                            _30Class = Microsoft.VisualBasic.Left(Trim(M02.Tables(0).Rows(I)("M16Material")), 2) & "-" & Microsoft.VisualBasic.Right(Trim(M02.Tables(0).Rows(I)("M16Material")), 5)
                            _30 = Trim(M02.Tables(0).Rows(I)("M16Material"))
                        Else
                            _30Class = _30Class & "','" & Microsoft.VisualBasic.Left(Trim(M02.Tables(0).Rows(I)("M16Material")), 2) & "-" & Microsoft.VisualBasic.Right(Trim(M02.Tables(0).Rows(I)("M16Material")), 5)
                            _30 = _30 & "','" & Trim(M02.Tables(0).Rows(I)("M16Material"))
                        End If
                        I = I + 1
                    Next

                    Dim _Alter_Quality As String


                    _Bomwst = 0

                    _Alter_Quality = ""
                    'Tobe Plane Qty
                    _TobeDeliverd = 0
                    vcWhere = "Merchant='" & Trim(M01.Tables(0).Rows(Z1)("M28Merchant")) & "' and  Metrrial in ('" & _30Class & "') and Location='To Be planned'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        If IsDBNull(M02.Tables(0).Rows(0)("PRD_Qty")) = True Then
                        Else
                            _TobeDeliverd = M02.Tables(0).Rows(0)("PRD_Qty")
                        End If
                    End If
                    _30 = ""
                    I = 0
                    vcWhere = " Metrrial in ('" & _30Class & "') and Location='To Be planned'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        characterToRemove = "-"
                        tmp30 = M02.Tables(0).Rows(I)("Metrrial")
                        tmp30 = (Replace(tmp30, characterToRemove, ""))
                        If I = 0 Then
                            _30 = tmp30
                        Else
                            _30 = _30 & "','" & tmp30



                        End If
                        I = I + 1
                    Next

                    I = 0
                    vcWhere = "Merchant='" & Trim(M01.Tables(0).Rows(Z1)("M28Merchant")) & "' and Metrrial in ('" & _30Class & "') and Location='Dye'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        vcWhere = "Batch_No='" & M02.Tables(0).Rows(I)("Prduct_Order") & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                            Else
                                _TobeDeliverd = _TobeDeliverd + M02.Tables(0).Rows(I)("PRD_Qty")
                            End If
                        End If
                        I = I + 1
                    Next


                    I = 0
                    vcWhere = " Metrrial in ('" & _30Class & "') and Location='Dye'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD3"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        vcWhere = "Batch_No='" & M02.Tables(0).Rows(I)("Prduct_Order") & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                            Else
                                characterToRemove = "-"
                                tmp30 = M02.Tables(0).Rows(I)("Metrrial")
                                tmp30 = (Replace(tmp30, characterToRemove, ""))
                                If _30 <> "" Then
                                    _30 = _30 & "','" & tmp30
                                Else
                                    _30 = tmp30

                                End If
                            End If
                        End If
                        I = I + 1
                    Next
                    ' End If
                    'bom wast
                    '_Bomwst = 0
                    'vcWhere = " M24Material in ('" & _30 & "')"
                    'M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    'I = 0
                    'For Each DTRow5 As DataRow In M02.Tables(0).Rows
                    '    _Bomwst = _Bomwst + (1 - M02.Tables(0).Rows(I)("M24WST"))
                    '    I = I + 1
                    'Next

                    I = 0
                    vcWhere = "M26Quality20='" & Trim(M01.Tables(0).Rows(Z1)("M28Quality")) & "'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "ALT"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        If I = 0 Then
                            _Alter_Quality = M02.Tables(0).Rows(0)("M26Quality30")
                        Else
                            _Alter_Quality = _Alter_Quality & "','" & M02.Tables(0).Rows(0)("M26Quality30")
                        End If

                        I = I + 1
                    Next

                    If _Shade = "-" Then
                        vcWhere = "M16Quality in ('" & _Alter_Quality & "')"
                    ElseIf _Shade = "Light" Then
                        vcWhere = "M16Quality in ('" & _Alter_Quality & "') and M14Grige='L'"
                    ElseIf _Shade = "Dark" Then
                        vcWhere = "M16Quality in ('" & _Alter_Quality & "') and M14Grige='D'"
                    End If
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "30C"), New SqlParameter("@vcWhereClause1", vcWhere))
                    I = 0
                    _30Class = ""
                    _30 = ""
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        If I = 0 Then
                            _30Class = Microsoft.VisualBasic.Left(Trim(M02.Tables(0).Rows(I)("M16Material")), 2) & "-" & Microsoft.VisualBasic.Right(Trim(M02.Tables(0).Rows(I)("M16Material")), 5)
                            _30 = Trim(M02.Tables(0).Rows(I)("M16Material"))
                        Else
                            _30Class = _30Class & "','" & Microsoft.VisualBasic.Left(Trim(M02.Tables(0).Rows(I)("M16Material")), 2) & "-" & Microsoft.VisualBasic.Right(Trim(M02.Tables(0).Rows(I)("M16Material")), 5)
                            _30 = _30 & "','" & Trim(M02.Tables(0).Rows(I)("M16Material"))
                        End If
                        I = I + 1
                    Next

                    vcWhere = "Merchant='" & Trim(M01.Tables(0).Rows(Z1)("M28Merchant")) & "' and Metrrial in ('" & _30Class & "')  and Location='To Be planned'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M02) Then
                        If IsDBNull(M02.Tables(0).Rows(0)("PRD_Qty")) = True Then
                        Else
                            _TobeDeliverd = _TobeDeliverd + M02.Tables(0).Rows(0)("PRD_Qty")
                        End If
                    End If

                    _30 = ""
                    I = 0
                    vcWhere = " Metrrial in ('" & _30Class & "') and Location='To Be planned'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        characterToRemove = "-"
                        tmp30 = M02.Tables(0).Rows(I)("Metrrial")
                        tmp30 = (Replace(tmp30, characterToRemove, ""))
                        If I = 0 Then
                            _30 = tmp30
                        Else
                            _30 = _30 & "','" & tmp30



                        End If
                        I = I + 1
                    Next

                    I = 0
                    vcWhere = "Merchant='" & Trim(M01.Tables(0).Rows(Z1)("M28Merchant")) & "' and Metrrial in ('" & _30Class & "') and Location='Dye'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        vcWhere = "Batch_No='" & M02.Tables(0).Rows(I)("Prduct_Order") & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                _TobeDeliverd = _TobeDeliverd + M02.Tables(0).Rows(I)("PRD_Qty")
                            End If
                        End If
                        I = I + 1
                    Next

                    I = 0
                    vcWhere = " Metrrial in ('" & _30Class & "') and Location='Dye'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD3"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        vcWhere = "Batch_No='" & M02.Tables(0).Rows(I)("Prduct_Order") & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                            Else
                                characterToRemove = "-"
                                tmp30 = M02.Tables(0).Rows(I)("Metrrial")
                                tmp30 = (Replace(tmp30, characterToRemove, ""))
                                If _30 <> "" Then
                                    _30 = _30 & "','" & tmp30
                                Else
                                    _30 = tmp30

                                End If
                            End If
                        End If
                        I = I + 1
                    Next
                    '2015/9/2 request by Sameera

                    I = 0
                    vcWhere = "Merchant='" & Trim(M01.Tables(0).Rows(Z1)("M28Merchant")) & "' and Metrrial in ('" & _30Class & "') and Location='Dye'"
                    M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "OTD1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow5 As DataRow In M02.Tables(0).Rows
                        vcWhere = "Batch_No='" & M02.Tables(0).Rows(I)("Prduct_Order") & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetNo_OrderReport", New SqlParameter("@cQryType", "FRU"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            If Trim(dsUser.Tables(0).Rows(0)("Stock_Code")) <> "" Then
                                _TobeDeliverd = _TobeDeliverd + M02.Tables(0).Rows(I)("PRD_Qty")
                            End If
                        End If
                        I = I + 1
                    Next


                    If _TobeDeliverd = 0 Then
                    Else
                        _TobeDeliverd = _TobeDeliverd / _ConFact
                    End If
                    'vcWhere = " M24Material in ('" & _30 & "')"
                    'M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "BOM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    'I = 0
                    'For Each DTRow5 As DataRow In M02.Tables(0).Rows
                    '    _Bomwst = _Bomwst + (1 - M02.Tables(0).Rows(I)("M24WST"))
                    '    I = I + 1
                    'Next


                    worksheet2.Cells(X, 29) = _TobeDeliverd '/ _Bomwst
                    range1 = worksheet2.Cells(X, 29)
                    range1.NumberFormat = "0.00"
                    worksheet2.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignRight

                    worksheet2.Cells(X, 30) = "=AC" & X & "-AA" & X
                    range1 = worksheet2.Cells(X, 30)
                    range1.NumberFormat = "0.00"
                    worksheet2.Cells(X, 30).HorizontalAlignment = XlHAlign.xlHAlignRight


                    _Chr = 97
                    For I = 1 To 38
                        If I = 27 Then
                            _Chr = 97
                        End If

                        If I >= 27 Then
                            worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range("A" & Chr(_Chr) & X, "A" & Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        Else

                            worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                            worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            ' worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                        End If
                        _Chr = _Chr + 1

                    Next

                    X = X + 1
                    Z1 = Z1 + 1
                Next

                'worksheet2.Range("A" & _FirstRow & ":A" & X - 1).MergeCells = True
                ''  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                'worksheet2.Range("A" & X - 1 & ":A" & X - 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                'worksheet2.Cells(X - 1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Z = Z + 1
                Z1 = 0
            Next

            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).rowheight = 15
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True

            worksheet2.Range("M" & X & ":M" & X).MergeCells = True
            worksheet2.Range("M" & (X)).Formula = "=SUM(M3:M" & X - 1 & ")"
            worksheet2.Range("M" & X & ":M" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("M" & X & ":M" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 13)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("N" & X & ":N" & X).MergeCells = True
            worksheet2.Range("N" & (X)).Formula = "=SUM(N3:N" & X - 1 & ")"
            worksheet2.Range("N" & X & ":N" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("N" & X & ":N" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 14)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("O" & X & ":O" & X).MergeCells = True
            worksheet2.Range("O" & (X)).Formula = "=SUM(O3:O" & X - 1 & ")"
            worksheet2.Range("O" & X & ":O" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("O" & X & ":O" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 15)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 15).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("P" & X & ":P" & X).MergeCells = True
            worksheet2.Range("P" & (X)).Formula = "=SUM(P3:P" & X - 1 & ")"
            worksheet2.Range("P" & X & ":P" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("P" & X & ":P" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 16)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 16).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("Q" & X & ":Q" & X).MergeCells = True
            worksheet2.Range("Q" & (X)).Formula = "=SUM(Q3:Q" & X - 1 & ")"
            worksheet2.Range("Q" & X & ":Q" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("Q" & X & ":Q" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 17)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight


            worksheet2.Range("R" & X & ":R" & X).MergeCells = True
            worksheet2.Range("R" & (X)).Formula = "=SUM(R3:R" & X - 1 & ")"
            worksheet2.Range("R" & X & ":R" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("R" & X & ":R" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 18)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 18).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("S" & X & ":S" & X).MergeCells = True
            worksheet2.Range("S" & (X)).Formula = "=SUM(S3:S" & X - 1 & ")"
            worksheet2.Range("S" & X & ":S" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("S" & X & ":S" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 19)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 19).HorizontalAlignment = XlHAlign.xlHAlignRight


            worksheet2.Range("T" & X & ":T" & X).MergeCells = True
            worksheet2.Range("T" & (X)).Formula = "=SUM(T3:T" & X - 1 & ")"
            worksheet2.Range("T" & X & ":T" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("T" & X & ":T" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 20)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("U" & X & ":U" & X).MergeCells = True
            worksheet2.Range("U" & (X)).Formula = "=SUM(U3:U" & X - 1 & ")"
            worksheet2.Range("U" & X & ":U" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("U" & X & ":U" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 21)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("V" & X & ":V" & X).MergeCells = True
            worksheet2.Range("V" & (X)).Formula = "=SUM(V3:V" & X - 1 & ")"
            worksheet2.Range("V" & X & ":V" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("V" & X & ":V" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 22)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("W" & X & ":W" & X).MergeCells = True
            worksheet2.Range("W" & (X)).Formula = "=SUM(W3:W" & X - 1 & ")"
            worksheet2.Range("W" & X & ":W" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("W" & X & ":W" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 23)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("X" & X & ":X" & X).MergeCells = True
            worksheet2.Range("X" & (X)).Formula = "=SUM(X3:X" & X - 1 & ")"
            worksheet2.Range("X" & X & ":X" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("X" & X & ":X" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 24)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 24).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("Y" & X & ":Y" & X).MergeCells = True
            worksheet2.Range("Y" & (X)).Formula = "=SUM(Y3:Y" & X - 1 & ")"
            worksheet2.Range("Y" & X & ":Y" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("Y" & X & ":Y" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 25)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("Z" & X & ":Z" & X).MergeCells = True
            worksheet2.Range("Z" & (X)).Formula = "=SUM(Z3:Z" & X - 1 & ")"
            worksheet2.Range("Z" & X & ":Z" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("Z" & X & ":Z" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 26)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("AA" & X & ":AA" & X).MergeCells = True
            worksheet2.Range("AA" & (X)).Formula = "=SUM(AA3:AA" & X - 1 & ")"
            worksheet2.Range("AA" & X & ":AA" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AA" & X & ":AA" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 27)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 27).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("AB" & X & ":AB" & X).MergeCells = True
            worksheet2.Range("AB" & (X)).Formula = "=SUM(AB3:AB" & X - 1 & ")"
            worksheet2.Range("AB" & X & ":AB" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 28).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AB" & X & ":AB" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 28)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 28).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("AC" & X & ":AC" & X).MergeCells = True
            worksheet2.Range("AC" & (X)).Formula = "=SUM(AC3:AC" & X - 1 & ")"
            worksheet2.Range("AC" & X & ":AC" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AC" & X & ":AC" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 29)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 29).HorizontalAlignment = XlHAlign.xlHAlignRight

            worksheet2.Range("AD" & X & ":AD" & X).MergeCells = True
            worksheet2.Range("AD" & (X)).Formula = "=SUM(AD3:AD" & X - 1 & ")"
            worksheet2.Range("AD" & X & ":AD" & X).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X, 30).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("AD" & X & ":AD" & X).Interior.Color = RGB(197, 217, 241)
            range1 = worksheet2.Cells(X, 30)
            range1.NumberFormat = "0.00"
            worksheet2.Cells(X, 30).HorizontalAlignment = XlHAlign.xlHAlignRight

            MsgBox("Record successfully created", MsgBoxStyle.Information, "Information ......")

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

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

    Function upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String

        Dim _sales_Order As String
        Dim _LineItem As String
        Dim _20Class As String
        Dim _Description As String
        Dim _Department As String
        Dim _Merchant As String
        Dim _RollNo As String
        Dim _PRD_No As String
        Dim _Date As Date
        Dim _SLocation As String
        Dim _SMC As Double
        Dim _Qty As Double
        Dim _Where As String
        Dim _Quality As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim characterToRemove As String

        Dim X11 As Integer
        Try

            nvcFieldList1 = "delete from M28Stock_Grige_Price "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            strFileName = ConfigurationManager.AppSettings("FilePath") + "\Stock_Greige_Price.txt"
            pbCount.Maximum = System.IO.File.ReadAllLines(strFileName).Length
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 10 Then
                    '   MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then

                _sales_Order = (Trim(fields(0)))
                _LineItem = (Trim(fields(1)))

                _20Class = Trim(fields(2))
                _Description = Trim(fields(3))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Description = (Replace(_Description, characterToRemove, ""))



                _Department = (Trim(fields(4)))
                characterToRemove = "'"

                'MsgBox(Trim(fields(9)))
                _Department = (Replace(_Department, characterToRemove, ""))

                _Merchant = Trim(fields(5))
                _RollNo = Trim(fields(6))
                _PRD_No = Trim(fields(7))

                Dim TestString As String = _Description
                Dim TestArray() As String = Split(TestString)

                ' TestArray holds {"apple", "", "", "", "pear", "banana", "", ""} 
                Dim LastNonEmpty As Integer = -1
                For z1 As Integer = 0 To TestArray.Length - 1
                    If TestArray(z1) <> "" Then
                        LastNonEmpty += 1
                        TestArray(LastNonEmpty) = TestArray(z1)
                        ' If z1 = 2 Then
                        _Quality = TestArray(LastNonEmpty)
                        Exit For
                        'End If
                    End If
                Next

                Dim B As String
                B = Microsoft.VisualBasic.Left(Trim(fields(8)), 6)
                _Date = (Microsoft.VisualBasic.Right(B, 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(8)), 2) & "/" & Microsoft.VisualBasic.Left(B, 4))
                '_Del_Date = Trim(fields(9))
                _Qty = Trim(fields(9))
                _SLocation = Trim(fields(11))
                _SMC = Trim(fields(12))

                _Where = "M28Sales_Order='" & Trim(_sales_Order) & "' and M28LineItem='" & Trim(_LineItem) & "' and M28RollNo='" & Trim(_RollNo) & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", _Where))
                If isValidDataset(M01) Then
                    nvcFieldList1 = "update M28Stock_Grige_Price set M28Department='" & _Department & "',M28Merchant='" & _Merchant & "',M28Prod_No='" & _PRD_No & "',M28Date='" & _Date & "',M28Qty='" & _Qty & "',M28SLocation='" & _SLocation & "',M28SMC='" & _SMC & "',M28Quality='" & _Quality & "' where " & _Where
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M28Stock_Grige_Price(M28Sales_Order,M28LineItem,M2820Class,M28Description,M28Department,M28Merchant,M28RollNo,M28Prod_No,M28Date,M28Qty,M28SLocation,M28SMC,M28Quality)" & _
                                                        " values('" & Trim(_sales_Order) & "', '" & Trim(_LineItem) & "','" & Trim(_20Class) & "','" & Trim(_Description) & "','" & _Department & "','" & Trim(_Merchant) & "','" & Trim(_RollNo) & "','" & Trim(_PRD_No) & "','" & Trim(_Date) & "','" & _Qty & "','" & _SLocation & "','" & _SMC & "','" & _Quality & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                pbCount.Value = pbCount.Value + 1


                lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))

                _sales_Order = ""
                _LineItem = ""
                '_LineItem = ""
                _Department = ""
                _Description = ""
                _Qty = 0
                _20Class = ""
                _SMC = 0
                _PRD_No = ""
                _Merchant = ""

                ' End If
                X11 = X11 + 1
                'cmdEdit.Enabled = True
            Next
            'MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ....")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            connection.Close()
            '  pbCount.Value = 0
            ' lblPro.Text = "Stock_Greige_Price.txt"
            lblPro.Refresh()
            pbCount.Refresh()
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            MsgBox(X11)
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function

    Private Sub chkUpload_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUpload.CheckedChanged
        lblPro.Text = ""
        pbCount.Value = 0
        If chkUpload.Checked = True Then
            Call upload_File()
        End If
    End Sub

    Private Sub frmGrgProvision_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtTodate.Text = Today
        Call Load_merchnt()
    End Sub

    Function Load_merchnt()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try

            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetM28Stock_Grige_Price", New SqlParameter("@cQryType", "MER"))
            If isValidDataset(M01) Then
                With cboQuality
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 245
                End With
            End If


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Function
End Class