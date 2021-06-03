
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

Public Class frmBackLog
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


    Private Sub frmBackLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFrom.Text = Today
        txtPlnDate.Text = Today
        txtTo.Text = Today

    End Sub

    Function Create_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim M03 As DataSet
        Dim _p4p As Integer

        Dim _1stQualityROW As Integer
        Dim _PreShade As String
        Dim _from As Date
        Dim _to As Date
        Dim _AW1stBulkSL As Integer
        Dim _MCSL As Integer

        Dim _AW1stBulk As Integer
        Dim _AWReceipt As Integer
        Dim _AWDyeStaff As Integer
        Dim _ShortLeadTime As Integer
        Dim _28Lead As Integer
        Dim _AWGride As Integer
        Dim _AWGrige_Short As Integer
        Dim vcWhere As String
        Dim _FirstRow As Integer
        Dim _MC As Integer

        Dim exc As New Application
        Try
            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
            Dim range1 As Range
            Dim I As Integer
            Dim Z As Integer
            Dim _Material As String
            Dim characterToRemove As String

            '   Try
            exc.Visible = True

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 11
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 10
            worksheet2.Columns("B").ColumnWidth = 15
            worksheet2.Columns("C").ColumnWidth = 10
            worksheet2.Columns("D").ColumnWidth = 10
            worksheet2.Columns("E").ColumnWidth = 10
            worksheet2.Columns("F").ColumnWidth = 35
            worksheet2.Columns("G").ColumnWidth = 10
            worksheet2.Columns("H").ColumnWidth = 16
            worksheet2.Columns("I").ColumnWidth = 10
            worksheet2.Columns("J").ColumnWidth = 12
            worksheet2.Columns("K").ColumnWidth = 36
            worksheet2.Columns("L").ColumnWidth = 17

            worksheet2.Columns("M").ColumnWidth = 17
            worksheet2.Columns("N").ColumnWidth = 20
            worksheet2.Columns("O").ColumnWidth = 10
            worksheet2.Columns("P").ColumnWidth = 15
            worksheet2.Columns("Q").ColumnWidth = 15

            Dim WeekNumber As Integer = DatePart(DateInterval.WeekOfYear, Date.Today, _
    FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

            If txtPlnDate.Text = Today Then
                worksheet2.Cells(1, 1) = "Actual Back Log in week  " & WeekNumber & " Dye plan"
                worksheet2.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Range("A1:E1").Interior.Color = RGB(197, 217, 241)
            Else
                worksheet2.Cells(1, 1) = "Possible Back Log in week  " & WeekNumber & " Dye plan"
                worksheet2.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Range("A1:E1").Interior.Color = RGB(197, 217, 241)
            End If
            ' worksheet2.Range("A2:M2").Interior.Color = RGB(197, 217, 241)
            worksheet2.Rows(1).Font.size = 13
            worksheet2.Rows(1).rowheight = 35
            worksheet2.Rows(1).Font.name = "Times New Roman"
            worksheet2.Rows(1).Font.BOLD = True
            worksheet2.Range("A1:E1").MergeCells = True
            worksheet2.Range("A1:E1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(1, 8) = "Week No : "
            worksheet2.Cells(1, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("H1:H1").MergeCells = True
            worksheet2.Range("H1:H1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(1, 9) = WeekNumber
            worksheet2.Cells(1, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("I1:I1").MergeCells = True
            worksheet2.Range("I1:I1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(1, 10) = "Date"
            worksheet2.Cells(1, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("J1:J1").MergeCells = True
            worksheet2.Range("J1:J1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(1, 11) = Today
            worksheet2.Cells(1, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("K1:K1").MergeCells = True
            worksheet2.Range("K1:K1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(1, 12) = "D.Date"
            worksheet2.Cells(1, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("L1:L1").MergeCells = True
            worksheet2.Range("L1:L1").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(1, 13) = txtTo.Text
            worksheet2.Cells(1, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("M1:M1").MergeCells = True
            worksheet2.Range("M1:M1").VerticalAlignment = XlVAlign.xlVAlignCenter

            Dim x As Integer
            Dim _Chr As Integer
            x = 1
            _Chr = 97
            For I = 1 To 17



                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)

                _Chr = _Chr + 1

            Next

            x = x + 1

            worksheet2.Rows(x).Font.size = 10
            worksheet2.Rows(x).rowheight = 20
            worksheet2.Rows(x).Font.name = "Times New Roman"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 1) = "Dye Machine"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(x, 2) = "Product code"
            worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 3) = "S/Order"
            worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 4) = "L/Item"
            worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(x, 5) = "Order code"
            worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(x, 6) = "Product description"
            worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 7) = "R - Code"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 8) = "Actual St/ Code"
            worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 9) = "Batch Qty"
            worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 10) = "Delivery date"
            worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 11) = "Extended status"
            worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 12) = "Customer"
            worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 13) = "Planning Comments"
            worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 14) = "Dyeing date"
            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 15) = "N/C Comments"
            worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 16) = "Reason for Delay"
            worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(x, 17) = "LIB Dip for Delay"
            worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter

            _Chr = 97
            For I = 1 To 17



                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                _Chr = _Chr + 1

            Next
            x = x + 2
            'TO BE PLANE
            _AW1stBulk = 0
            _AWDyeStaff = 0
            _AWReceipt = 0
            _ShortLeadTime = 0
            _28Lead = 0
            _AWGride = 0
            _AWGrige_Short = 0
            _MC = 0

            worksheet2.Cells(x, 1) = "To be Planed"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            x = x + 1
            _FirstRow = x

            vcWhere = "location='To Be planned' and del_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "'  and PRD_Qty>(M01SO_Qty*Tollarance_MIN/100)"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows

                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Metrrial")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = T01.Tables(0).Rows(Z)("Prduct_Order")
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Met_Des")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = T01.Tables(0).Rows(Z)("Metrrial")
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 8) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("del_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "To be Plan"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                worksheet2.Cells(x, 12) = T01.Tables(0).Rows(Z)("Customer")
                worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) <> "" Then
                    vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                        worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                        worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                        worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If
                End If
                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "To be Planed"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next

            x = x + 1
            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next

            '  Aw Dye stuff
            Dim _DyeMC As String
            Dim _StockCode As String

            worksheet2.Cells(x, 1) = "Aw Dye Stuff"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            x = x + 1
            _FirstRow = x
            Dim _ProductNo As String
            vcWhere = "location in  ('Dye','AW Presetting','AW Preparation') and del_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "AWD"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) & "' and Recipy_Status='Aw Dye Stuffs' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")
                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "' and Recipy_Status='Aw Dye Stuffs' "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            Z = Z + 1
                            Continue For
                        End If
                    Else
                        Z = Z + 1
                        Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Metrrial")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Prduct_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                End If
                worksheet2.Cells(x, 5) = T01.Tables(0).Rows(Z)("Prduct_Order")
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Met_Des")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = T01.Tables(0).Rows(Z)("Metrrial")
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "'  and  Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "' and Delivery_Date = '" & T01.Tables(0).Rows(Z)("del_Date") & "'  and Product_Order='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "' "
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M03) Then
                    worksheet2.Cells(x, 9) = M03.Tables(0).Rows(0)("Order_Qty_Kg")
                    worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet2.Cells(x, 13) = M03.Tables(0).Rows(0)("Pln_Comment")
                    worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft

                End If

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("del_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Aw Dye Stuff"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                worksheet2.Cells(x, 12) = T01.Tables(0).Rows(Z)("Customer")
                worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'COMMENT FOR TESTING
                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "Aw Dye Stuff"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            _AWDyeStaff = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            '---------------------------------------------------------------------------
            'Aw 1st b App
            x = x + 1
            worksheet2.Cells(x, 1) = "Aw 1st b App"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            x = x + 1
            _FirstRow = x

            vcWhere = "o.location in ('Dye','AW Presetting','AW Preparation') and o.del_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and left(z.Pln_Comment,7)<>'Short L'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "AWD1"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) & "' and first_BlkApp='Yes' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")
                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "' and first_BlkApp='Yes' "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")
                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            Z = Z + 1
                            Continue For
                        End If
                    Else
                        Z = Z + 1
                        Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Metrrial")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Prduct_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                End If
                worksheet2.Cells(x, 5) = T01.Tables(0).Rows(Z)("Prduct_Order")
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Met_Des")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = T01.Tables(0).Rows(Z)("Metrrial")
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "'  and  Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "' and Delivery_Date = '" & T01.Tables(0).Rows(Z)("del_Date") & "'  and Product_Order='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "' "
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M03) Then
                    worksheet2.Cells(x, 9) = M03.Tables(0).Rows(0)("Order_Qty_Kg")
                    worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet2.Cells(x, 13) = M03.Tables(0).Rows(0)("Pln_Comment")
                    worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("del_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Aw 1st b App"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                worksheet2.Cells(x, 12) = T01.Tables(0).Rows(Z)("Customer")
                worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "Aw 1st b App"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            _AW1stBulk = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            '-----------------------------------------------------------------------------------------------------------------------------
            'short lead time
            x = x + 2
            worksheet2.Cells(x, 1) = "Aw 1st b App- Short lead time"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            _FirstRow = x

            vcWhere = "o.location in ('Dye','AW Presetting','AW Preparation') and o.del_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and left(z.Pln_Comment,7)='Short L'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "AWD1"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) & "' and first_BlkApp='Yes' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")
                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "' and first_BlkApp='Yes' "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            Z = Z + 1
                            Continue For
                        End If
                    Else
                        Z = Z + 1
                        Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Metrrial")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Prduct_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                End If
                worksheet2.Cells(x, 5) = T01.Tables(0).Rows(Z)("Prduct_Order")
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Met_Des")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = T01.Tables(0).Rows(Z)("Metrrial")
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "'  and  Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "' and Delivery_Date = '" & T01.Tables(0).Rows(Z)("del_Date") & "'  and Product_Order='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "' "
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M03) Then
                    worksheet2.Cells(x, 9) = M03.Tables(0).Rows(0)("Order_Qty_Kg")
                    worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet2.Cells(x, 13) = M03.Tables(0).Rows(0)("Pln_Comment")
                    worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("del_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Aw 1st b App- Short lead time"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                worksheet2.Cells(x, 12) = T01.Tables(0).Rows(Z)("Customer")
                worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next
            If _FirstRow = x Then
                x = x + 1
            End If
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "Aw 1st b App- Short lead time"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            _AW1stBulkSL = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            '--------------------------------------------------------------------------------------
            'Aw Recipe
            x = x + 1
            worksheet2.Cells(x, 1) = "Aw Recipe"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            x = x + 1
            _FirstRow = x

            vcWhere = "location in ('Dye','AW Preparation','AW Presetting') and del_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "AWD"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) & "' and Recipy_Status='Awaiting Recipe' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "' and Recipy_Status='Awaiting Recipe' "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            Z = Z + 1
                            Continue For
                        End If
                    Else
                        Z = Z + 1
                        Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Metrrial")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Prduct_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                End If

                worksheet2.Cells(x, 5) = T01.Tables(0).Rows(Z)("Prduct_Order")
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Met_Des")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = T01.Tables(0).Rows(Z)("Metrrial")
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "'  and  Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "' and Delivery_Date = '" & T01.Tables(0).Rows(Z)("del_Date") & "'  and Product_Order='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "' "
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M03) Then
                    worksheet2.Cells(x, 9) = M03.Tables(0).Rows(0)("Order_Qty_Kg")
                    worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet2.Cells(x, 13) = M03.Tables(0).Rows(0)("Pln_Comment")
                    worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If
                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("del_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Awaiting Recipe(Fresh)"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                worksheet2.Cells(x, 12) = T01.Tables(0).Rows(Z)("Customer")
                worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    'worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    'worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next
            '----==================
            ' x = x + 1
            vcWhere = "location in ('Dye','Finishing') and del_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and NC_Comment in ('18.OFF SHADE BULK','19.Off shade Sample','20.Off shade Yarn Dye','21.Wet form - TB reprocess','17.Held due to Trials','15.Batches Tb over Dyed','16.Stripped tb Over Dyed','28.Down  Grade') "
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "AWD"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                'vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) & "' and Recipy_Status='Awaiting Recipe' "
                'M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                'If isValidDataset(M01) Then
                '    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                '    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")
                'Else

                '    Z = Z + 1
                '    Continue For

                'End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Metrrial")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Prduct_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Prduct_Order")
                End If
                worksheet2.Cells(x, 5) = T01.Tables(0).Rows(Z)("Prduct_Order")
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Met_Des")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = T01.Tables(0).Rows(Z)("Metrrial")
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                'worksheet2.Cells(x, 8) = _StockCode
                'worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "'  and  Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "' and Delivery_Date = '" & T01.Tables(0).Rows(Z)("del_Date") & "'  and Product_Order='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "' "
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M03) Then
                    worksheet2.Cells(x, 9) = M03.Tables(0).Rows(0)("Order_Qty_Kg")
                    worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet2.Cells(x, 13) = M03.Tables(0).Rows(0)("Pln_Comment")
                    worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If
                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("del_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Awaiting Recipe for Dry and Hold Batches"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                worksheet2.Cells(x, 12) = T01.Tables(0).Rows(Z)("Customer")
                worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 15) = T01.Tables(0).Rows(Z)("NC_Comment")
                worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Prduct_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    'worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    'worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next
            Dim _Last As Integer

            _Last = x - 1
            '==================================================================
            vcWhere = "M09Oredr_Type in ('Dyeing') and M09Del_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and M09ZPL_OrderType='ZP02' "
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("M09BatchNo")) & "' and Recipy_Status='First Bulk Awaiting Recipe' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("M09BatchNo") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "' and Recipy_Status='First Bulk Awaiting Recipe' "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            Z = Z + 1
                            Continue For
                        End If
                    Else
                        Z = Z + 1
                        Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("M09BatchNo")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("M09BatchNo")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("M09BatchNo")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("M09BatchNo")
                End If

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("M09Sales_Oredr")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("M09Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = T01.Tables(0).Rows(Z)("M09Meterial")
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("M09Dis")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = T01.Tables(0).Rows(Z)("M09Meterial")
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("M09Qty_KG")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("M09Del_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Awaiting 1st Bulk Recipe"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                worksheet2.Cells(x, 12) = T01.Tables(0).Rows(Z)("M09Customer")
                worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("M09BatchNo") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(217, 217, 217)
                    'worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    'worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "Aw Recipe"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            _AWReceipt = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & _Last & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                'worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                'worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            '---------------------------------------------------------------------------
            'Aw Greige

            x = x + 1
            worksheet2.Cells(x, 1) = "Aw Greige"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            x = x + 1
            _FirstRow = x
            Dim Diff As TimeSpan
            Dim _DateIN As Date
            Dim _DateOUT As Date

            _DateOUT = Today
            vcWhere = "location in ('AW Greige') and Delivery_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and left(Pln_Comment,7)<>'Short l' " 'and left(Material_Dis,1)<>'Y'  and day(Delivery_Date-GETDATE())<7"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""
                _DateIN = T01.Tables(0).Rows(Z)("Delivery_Date")
                Diff = _DateIN.Subtract(_DateOUT)

                If Microsoft.VisualBasic.Left(T01.Tables(0).Rows(Z)("Material_Dis"), 1) = "Y" And Diff.Days > 7 Then
                    Z = Z + 1
                    Continue For
                End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Product_Order")) & "'  "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "'  "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            ' Z = Z + 1
                            ' Continue For
                        End If
                    Else
                        ' Z = Z + 1
                        ' Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Product_Order")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Product_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                End If

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = CInt(T01.Tables(0).Rows(Z)("Material"))
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Material_Dis")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = CInt(T01.Tables(0).Rows(Z)("Material"))
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("Order_Qty_Kg")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("Delivery_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Aw Greige"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "' and Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 12) = M03.Tables(0).Rows(0)("Customer")
                    worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 13) = T01.Tables(0).Rows(Z)("Pln_Comment")
                worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 15) = T01.Tables(0).Rows(Z)("NC_Comment")
                worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next


            vcWhere = "location in ('AW Greige') and Delivery_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and left(Pln_Comment,7)<>'Short l' and left(Material_Dis,1)='Y' and day(Delivery_Date-GETDATE())<7"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Product_Order")) & "'  "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "'  "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            ' Z = Z + 1
                            ' Continue For
                        End If
                    Else
                        ' Z = Z + 1
                        ' Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Product_Order")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Product_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                End If

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = CInt(T01.Tables(0).Rows(Z)("Material"))
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Material_Dis")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = CInt(T01.Tables(0).Rows(Z)("Material"))
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("Order_Qty_Kg")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("Delivery_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Aw Greige"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "' and Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 12) = M03.Tables(0).Rows(0)("Customer")
                    worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 13) = T01.Tables(0).Rows(Z)("Pln_Comment")
                worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 15) = T01.Tables(0).Rows(Z)("NC_Comment")
                worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "Aw Greige"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            _AWGride = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            '=======================================================================================================================
            x = x + 1

            'Aw Greige -Short lead time

            worksheet2.Cells(x, 1) = "Aw Greige -Short lead time"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            x = x + 1
            _FirstRow = x

            vcWhere = "location in ('AW Greige') and Delivery_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and left(Pln_Comment,7)='Short l'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Product_Order")) & "'  "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "'  "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            '  Z = Z + 1
                            '  Continue For
                        End If
                    Else
                        ' Z = Z + 1
                        ' Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Product_Order")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Product_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                End If

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = CInt(T01.Tables(0).Rows(Z)("Material"))
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Material_Dis")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = CInt(T01.Tables(0).Rows(Z)("Material"))
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("Order_Qty_Kg")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("Delivery_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Aw Greige -Short lead time"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "' and Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 12) = M03.Tables(0).Rows(0)("Customer")
                    worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 13) = T01.Tables(0).Rows(Z)("Pln_Comment")
                worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 15) = T01.Tables(0).Rows(Z)("NC_Comment")
                worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "Aw Greige -Short lead time"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            _AWGrige_Short = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            '--------------------------------------------------------------------------------------------------------
            x = x + 1

            '28 days Yarn dye orders

            worksheet2.Cells(x, 1) = "28 days Yarn dye orders"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            x = x + 1
            _FirstRow = x
            vcWhere = "location in ('AW Greige') and Delivery_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and left(Pln_Comment,7)<>'Short l' and left(Material_Dis,1)='Y' and day(Delivery_Date-GETDATE())<7"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Product_Order")) & "'  "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "'  "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            ' Z = Z + 1
                            'Continue For
                        End If
                    Else
                        '  Z = Z + 1
                        ' Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Product_Order")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Product_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                End If

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = CInt(T01.Tables(0).Rows(Z)("Material"))
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Material_Dis")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = CInt(T01.Tables(0).Rows(Z)("Material"))
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("Order_Qty_Kg")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("Delivery_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "28 days Yarn dye orders"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "' and Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 12) = M03.Tables(0).Rows(0)("Customer")
                    worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 13) = T01.Tables(0).Rows(Z)("Pln_Comment")
                worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 15) = T01.Tables(0).Rows(Z)("NC_Comment")
                worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "28 days Yarn dye orders"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            _28Lead = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next

            x = x + 1
            'M/C Backlog

            worksheet2.Cells(x, 1) = "M/C Backlog"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            x = x + 1

            _FirstRow = x
            vcWhere = "location in ('Aw Dyeing') and Delivery_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and Product_Order not in ('" & _ProductNo & "') and left(Pln_Comment,7)<>'Short L'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Product_Order")) & "' and Dye_Pln_Date>'" & txtPlnDate.Text & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    If Microsoft.VisualBasic.Left(M01.Tables(0).Rows(0)("Dye_Machine"), 4) = "FINI" Then
                        Z = Z + 1
                        Continue For
                    End If
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "' "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "' and Dye_Pln_Date> '" & txtPlnDate.Text & "' "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            If Microsoft.VisualBasic.Left(M02.Tables(0).Rows(0)("Dye_Machine"), 4) = "FINI" Then
                                Z = Z + 1
                                Continue For
                            End If

                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            Z = Z + 1
                            Continue For
                        End If
                    Else
                        Z = Z + 1
                        Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Product_Order")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Product_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                End If

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = CInt(T01.Tables(0).Rows(Z)("Material"))
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Material_Dis")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = CInt(T01.Tables(0).Rows(Z)("Material"))
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("Order_Qty_Kg")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("Delivery_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "M/C Backlog"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "' and Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 12) = M03.Tables(0).Rows(0)("Customer")
                    worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 13) = T01.Tables(0).Rows(Z)("Pln_Comment")
                worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 15) = T01.Tables(0).Rows(Z)("NC_Comment")
                worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "M/C Backlog"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            _MC = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":h" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            x = x + 1
            'M/C Backlog Shot Lead Time

            worksheet2.Cells(x, 1) = "M/C Backlog Short L"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            x = x + 1

            _FirstRow = x
            vcWhere = "location in ('Aw Dyeing') and Delivery_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and Product_Order not in ('" & _ProductNo & "') and left(Pln_Comment,7)='Short L'"
            SQL = "SELECT * FROM ZPP_DEL WHERE location in ('Aw Dyeing') and Delivery_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and Product_Order not in ('" & _ProductNo & "') and left(Pln_Comment,7)='Short L'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            'T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("Product_Order")) & "' and Dye_Pln_Date>'" & txtPlnDate.Text & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "' "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "' and Dye_Pln_Date> '" & txtPlnDate.Text & "' "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            Z = Z + 1
                            Continue For
                        End If
                    Else
                        Z = Z + 1
                        Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("Product_Order")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("Product_Order")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("Product_Order")
                End If

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("Sales_Order")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = CInt(T01.Tables(0).Rows(Z)("Material"))
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("Material_Dis")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = CInt(T01.Tables(0).Rows(Z)("Material"))
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("Order_Qty_Kg")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("Delivery_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "M/C Backlog"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("Sales_Order") & "' and Line_Item='" & T01.Tables(0).Rows(Z)("Line_Item") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 12) = M03.Tables(0).Rows(0)("Customer")
                    worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 13) = T01.Tables(0).Rows(Z)("Pln_Comment")
                worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 15) = T01.Tables(0).Rows(Z)("NC_Comment")
                worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("Product_Order") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "M/C Backlog Short L"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            _MCSL = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            'P4P
            x = x + 1
            worksheet2.Cells(x, 1) = "P4P"
            worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("A" & x & ":B" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("A" & x & ":B" & x).MergeCells = True
            worksheet2.Range("A" & x & ":B" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True
            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next
            x = x + 1

            _FirstRow = x
            vcWhere = "M09Oredr_Type in ('Dyeing') and M09Del_Date between '" & txtFrom.Text & "' and '" & txtTo.Text & "' and M09Fabric_Type='P4P'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "ZORD"), New SqlParameter("@vcWhereClause1", vcWhere))
            Z = 0
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                _StockCode = ""

                'If Trim(T01.Tables(0).Rows(Z)("Prduct_Order")) = "1601311" Then
                '    MsgBox("")
                'End If
                vcWhere = "Batch_No='" & Trim(T01.Tables(0).Rows(Z)("M09BatchNo")) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    _DyeMC = M01.Tables(0).Rows(0)("Dye_Machine")
                    _StockCode = M01.Tables(0).Rows(0)("Stock_Code")

                    worksheet2.Cells(x, 14) = M01.Tables(0).Rows(0)("Dye_Pln_Date")
                    worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    vcWhere = "m30BatchNo='" & T01.Tables(0).Rows(Z)("M09BatchNo") & "' "
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then

                        vcWhere = "Batch_No='" & M01.Tables(0).Rows(0)("m30M_No") & "'  "
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "FRS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            _DyeMC = M02.Tables(0).Rows(0)("Dye_Machine")
                            _StockCode = M02.Tables(0).Rows(0)("Stock_Code")

                            worksheet2.Cells(x, 14) = M02.Tables(0).Rows(0)("Dye_Pln_Date")
                            worksheet2.Cells(x, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            Z = Z + 1
                            Continue For
                        End If
                    Else
                        Z = Z + 1
                        Continue For
                    End If
                End If
                worksheet2.Rows(x).Font.size = 8
                worksheet2.Rows(x).Font.name = "Arial"

                worksheet2.Cells(x, 1) = _DyeMC
                worksheet2.Cells(x, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet2.Cells(x, 2) = T01.Tables(0).Rows(Z)("M09BatchNo")
                worksheet2.Cells(x, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If Z = 0 Then
                    If _ProductNo <> "" Then
                        _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("M09BatchNo")
                    Else
                        _ProductNo = T01.Tables(0).Rows(Z)("M09BatchNo")
                    End If
                Else
                    _ProductNo = _ProductNo & "','" & T01.Tables(0).Rows(Z)("M09BatchNo")
                End If

                worksheet2.Cells(x, 3) = T01.Tables(0).Rows(Z)("M09Sales_Oredr")
                worksheet2.Cells(x, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 4) = T01.Tables(0).Rows(Z)("M09Line_Item")
                worksheet2.Cells(x, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet2.Cells(x, 5) = CInt(T01.Tables(0).Rows(Z)("M09Meterial"))
                worksheet2.Cells(x, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(x, 6) = T01.Tables(0).Rows(Z)("M09Dis")
                worksheet2.Cells(x, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                _Material = CInt(T01.Tables(0).Rows(Z)("M09Meterial"))
                characterToRemove = "-"
                _Material = (Replace(_Material, characterToRemove, ""))

                vcWhere = "M16Material='" & _Material & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "LPT"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 7) = T01.Tables(0).Rows(Z)("M16R_Code")
                    worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 8) = _StockCode
                worksheet2.Cells(x, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("PRD_Qty")
                'worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 9) = T01.Tables(0).Rows(Z)("M09Qty_KG")
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                range1 = worksheet2.Cells(x, 9)
                range1.NumberFormat = "0"
                worksheet2.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight

                worksheet2.Cells(x, 10) = T01.Tables(0).Rows(Z)("M09Del_Date")
                worksheet2.Cells(x, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(x, 11) = "Tobe Dye P4P"
                worksheet2.Cells(x, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Range("K" & x, "K" & x).Font.Color = RGB(255, 0, 0)

                vcWhere = "Sales_Order='" & T01.Tables(0).Rows(Z)("M09Sales_Oredr") & "' and Line_Item='" & T01.Tables(0).Rows(Z)("M09Line_Item") & "'"
                M03 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "OTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 12) = M03.Tables(0).Rows(0)("Customer")
                    worksheet2.Cells(x, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                End If

                worksheet2.Cells(x, 13) = T01.Tables(0).Rows(Z)("M09Planning_Comm")
                worksheet2.Cells(x, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(x, 15) = T01.Tables(0).Rows(Z)("NC_Comment")
                'worksheet2.Cells(x, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter


                vcWhere = "T07BatchNo='" & T01.Tables(0).Rows(Z)("M09BatchNo") & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "COM"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    worksheet2.Cells(x, 16) = M01.Tables(0).Rows(0)("T07Reason")
                    worksheet2.Cells(x, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet2.Cells(x, 17) = M01.Tables(0).Rows(0)("T07LIB")
                    worksheet2.Cells(x, 17).HorizontalAlignment = XlHAlign.xlHAlignLeft
                End If

                _Chr = 97
                For I = 1 To 17



                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                    ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter


                    _Chr = _Chr + 1

                Next

                x = x + 1
                Z = Z + 1
            Next

            worksheet2.Rows(x).Font.size = 8
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "P4P"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            _p4p = x
            worksheet1.Range("I" & (x)).Formula = "=SUM(I" & _FirstRow & ":I" & (x - 1) & ")"
            worksheet1.Cells(x, 9).HorizontalAlignment = XlHAlign.xlHAlignRight
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            worksheet2.Range("g" & x & ":i" & x).Interior.Color = RGB(197, 217, 241)
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter

            _Chr = 97
            For I = 1 To 17

                ' worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).Interior.Color = RGB(197, 217, 241)
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).MergeCells = True
                worksheet2.Range(Chr(_Chr) & x & ":" & Chr(_Chr) & x).VerticalAlignment = XlVAlign.xlVAlignCenter

                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & x, Chr(_Chr) & x).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                _Chr = _Chr + 1

            Next

            x = x + 2
            worksheet2.Rows(x).Font.size = 12
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "Total"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            Dim _Total As Double
            Dim Val As Double
            _Total = 0
            range1 = CType(worksheet2.Cells(_28Lead, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AW1stBulk, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AWDyeStaff, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AWGride, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AWReceipt, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_MC, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_p4p, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value


            worksheet2.Cells(x, 9) = _Total '"=I" & _AW1stBulk & "+I" & _AWDyeStaff & "+I" & _28Lead & "+I" & _AWGride & "+I" & _AWReceipt & "+I" & _MC & ""
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

            x = x + 2
            worksheet2.Rows(x).Font.size = 12
            worksheet2.Rows(x).Font.name = "Arial"
            worksheet2.Rows(x).Font.BOLD = True

            worksheet2.Cells(x, 7) = "Total with Short Lead Time"
            worksheet2.Cells(x, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Range("g" & x & ":H" & x).MergeCells = True
            worksheet2.Range("g" & x & ":H" & x).VerticalAlignment = XlVAlign.xlVAlignCenter
            _Total = 0
            range1 = CType(worksheet2.Cells(_28Lead, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AW1stBulk, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AWDyeStaff, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AWGride, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AWReceipt, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_MC, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AW1stBulkSL, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_MCSL, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_AWGrige_Short, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value

            range1 = CType(worksheet2.Cells(_p4p, 9), Microsoft.Office.Interop.Excel.Range)
            _Total = _Total + range1.Value


            worksheet2.Cells(x, 9) = _Total '"=I" & _AW1stBulk & "+I" & _AWDyeStaff & "+I" & _28Lead & "+I" & _AWGride & "+I" & _AWReceipt & "+I" & _MC & ""
            range1 = worksheet1.Cells(x, 9)
            range1.NumberFormat = "0.00"

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Create_File()

    End Sub

    Function Upload_File()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String

        Dim _MNo As String
        Dim _BatchNo As String
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
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\M numbers.txt"
            pbCount.Maximum = System.IO.File.ReadAllLines(strFileName).Length
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)
                If X11 = 10 Then
                    '   MsgBox("")
                End If

                '  MsgBox(Trim(fields(0)))
                '_Location = Trim(fields(15))
                ' If _Location <> "" Then

                _MNo = (Trim(fields(0)))
                _BatchNo = (Trim(fields(1)))


                _Where = " m30BatchNo='" & Trim(_BatchNo) & "' "
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetPossibleBacklog", New SqlParameter("@cQryType", "MNO"), New SqlParameter("@vcWhereClause1", _Where))
                If isValidDataset(M01) Then
                    nvcFieldList1 = "update M30MNumber set m30M_No='" & _MNo & "' where " & _Where
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M30MNumber(m30BatchNo,m30M_No)" & _
                                                        " values('" & Trim(_BatchNo) & "', '" & Trim(_MNo) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                pbCount.Value = pbCount.Value + 1


                lblPro.Text = Trim(fields(0)) & "-" & Trim(fields(1))

                _MNo = ""
                _BatchNo = ""
                '_LineItem = ""
               
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
            Call Upload_File()
        End If
    End Sub

    Private Sub pbCount_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pbCount.ValueChanged

    End Sub
End Class