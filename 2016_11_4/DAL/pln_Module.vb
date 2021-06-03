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

Module pln_Module

    Function Create_MeetingReport(ByVal strFrom As Date, ByVal strTo As Date, ByVal strDepartment As String)
        'Project Owerner Asela - Planning
        'Project Maneger - Lalith Attapatthu

        Dim sql As String
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

        Try
            Dim exc As New Application

            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

            exc.Visible = True

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 12
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 12
            worksheet2.Columns("B").ColumnWidth = 8
            worksheet2.Columns("C").ColumnWidth = 8
            worksheet2.Columns("D").ColumnWidth = 24
            worksheet2.Columns("K").ColumnWidth = 8
            worksheet2.Columns("L").ColumnWidth = 12
            worksheet2.Columns("M").ColumnWidth = 22
            worksheet2.Columns("N").ColumnWidth = 12
            worksheet2.Columns("O").ColumnWidth = 12
            worksheet2.Columns("P").ColumnWidth = 12
            worksheet2.Columns("J").ColumnWidth = 32
            worksheet2.Columns("K").ColumnWidth = 22
            worksheet2.Columns("S").ColumnWidth = 50

            worksheet2.Rows(2).Font.size = 14
            worksheet2.Rows(2).Font.name = "Times New Roman"
            worksheet2.Rows(2).rowheight = 24.25
            worksheet2.Rows(2).Font.bold = True
            worksheet2.Cells(2, 4) = "OTD Meeting Report-" & strDepartment
            worksheet2.Cells(2, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("a2:q2").MergeCells = True
            worksheet1.Range("a2:a2").VerticalAlignment = XlVAlign.xlVAlignCenter
            Dim X As Integer
            X = 4
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).rowheight = 24.25
            worksheet2.Rows(X).Font.bold = True

            worksheet2.Cells(X, 1) = "Sales Order"
            worksheet2.Cells(X, 2) = "Line Item"
            worksheet2.Cells(X, 3) = "Material"
            worksheet2.Cells(X, 4) = "Description"
            worksheet2.Cells(X, 5) = "Delivary Date"
            worksheet2.Cells(X, 6) = "Batch No"
            worksheet2.Cells(X, 7) = "Batch Qty (Kg)"
            worksheet2.Cells(X, 8) = "Batch Qty (Mtr)"
            worksheet2.Cells(X, 9) = "No.Of.Dys In S/L"
            worksheet2.Cells(X, 10) = "Next Operation"
            worksheet2.Cells(X, 11) = "NC Comment"
            worksheet2.Cells(X, 12) = strFrom
            worksheet2.Cells(X, 13) = strTo
            worksheet2.Cells(X, 14) = "Complete Date"
            worksheet2.Cells(X, 15) = "Possible Delivary Date"
            worksheet2.Cells(X, 16) = "Reason"
            worksheet2.Cells(X, 17) = "LIB dep"
            worksheet2.Cells(X, 18) = "Week"
            worksheet2.Cells(X, 19) = "Customer"

            Dim i As Integer
            i = 4
            Dim _Char As Integer
            _Char = 97
            ' MsgBox(ChrW(_Char))
            For i = 1 To 19

                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Cells(X, i).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(X, i).WrapText = True
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Interior.Color = RGB(216, 216, 216)
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).MergeCells = True
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                _Char = _Char + 1

                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            Next
            X = X + 1
            Dim z As Integer
            i = 0
            With frmUpdate_Status
                For Each uRow As UltraGridRow In .UltraGrid4.Rows
                    worksheet2.Rows(X).Font.size = 8
                    worksheet2.Rows(X).Font.name = "Times New Roman"

                    _Char = 97
                    For z = 0 To 18
                        '  MsgBox(.UltraGrid4.Rows(i).Cells(z).Value)
                        worksheet2.Cells(X, z + 1) = .UltraGrid4.Rows(i).Cells(z).Value
                        If .UltraGrid4.Rows(i).Cells(z).Appearance.BackColor = Color.Yellow Then
                            worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Interior.Color = RGB(255, 255, 0)
                        End If
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).MergeCells = True
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        If z + 1 = 4 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 7 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ElseIf z + 1 = 8 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ElseIf z + 1 = 10 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 11 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 16 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 17 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 13 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                        End If
                        _Char = _Char + 1
                    Next

                    _Char = 97
                    For z = 1 To 19

                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet2.Cells(X, z).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        ' worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        ' worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Interior.Color = RGB(216, 216, 216)
                        'worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).MergeCells = True
                        ' worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                        _Char = _Char + 1

                        'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    Next
                    X = X + 1
                    i = i + 1
                Next

            End With

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.CLOSE()
            ' worksheet1.Cells(4, 5) = _Fail_Batch
            'worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
            ' MsgBox(_Fail_Batch)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.CLOSE()
            End If
        End Try
    End Function

    Function Create_DelivaryReport(ByVal strDate As Date)
        Dim sql As String
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

        Try
            Dim exc As New Application

            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

            exc.Visible = True

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 12
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 12
            worksheet2.Columns("B").ColumnWidth = 8
            worksheet2.Columns("C").ColumnWidth = 8
            worksheet2.Columns("D").ColumnWidth = 24
            worksheet2.Columns("K").ColumnWidth = 8
            worksheet2.Columns("L").ColumnWidth = 12
            worksheet2.Columns("M").ColumnWidth = 12
            worksheet2.Columns("N").ColumnWidth = 12
            worksheet2.Columns("O").ColumnWidth = 12
            worksheet2.Columns("P").ColumnWidth = 12
            worksheet2.Columns("J").ColumnWidth = 12
            worksheet2.Columns("K").ColumnWidth = 12
            worksheet2.Columns("V").ColumnWidth = 17

            worksheet2.Rows(2).Font.size = 14
            worksheet2.Rows(2).Font.name = "Times New Roman"
            worksheet2.Rows(2).rowheight = 24.25
            worksheet2.Rows(2).Font.bold = True
            worksheet2.Cells(2, 4) = "Delivary Report"
            worksheet2.Cells(2, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("a2:q2").MergeCells = True
            worksheet1.Range("a2:a2").VerticalAlignment = XlVAlign.xlVAlignCenter
            Dim X As Integer
            X = 4
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).rowheight = 24.25
            worksheet2.Rows(X).Font.bold = True

            worksheet2.Cells(X, 1) = "Sales Order"
            worksheet2.Cells(X, 2) = "Line Item"
            worksheet2.Cells(X, 3) = "Material"
            worksheet2.Cells(X, 4) = "Description"
            worksheet2.Cells(X, 5) = "Del.qty as today"

            Dim _Del As String

            If Month(Today) = 1 Then
                _Del = "B/T/Del December"
            Else

                _Del = "B/T/Del " & MonthName(Month(strDate) - 1)
            End If
            worksheet2.Cells(X, 6) = _Del
            worksheet2.Cells(X, 7) = "B/T/Del " & MonthName(Month(strDate))
            If Month(Today) = 12 Then
                _Del = "1st Wk of (January)"
            Else

                _Del = "1st Wk of (" & MonthName(Month(strDate) + 1) & ")"
            End If
            worksheet2.Cells(X, 8) = _Del

            If Month(Today) = 12 Then
                _Del = "2nd Wk of (January)"
            Else

                _Del = "2nd Wk of (" & MonthName(Month(strDate) + 1) & ")"
            End If

            worksheet2.Cells(X, 9) = _Del
            worksheet2.Cells(X, 10) = "Pro.to del " & MonthName(Month(strDate))
            If Month(Today) = 12 Then
                _Del = "Pro.to del January"
            Else
                _Del = "Pro.to del " & MonthName(Month(Today) + 1)
            End If

            worksheet2.Cells(X, 11) = _Del
            worksheet2.Cells(X, 12) = "Tot.Possi Qty"
            worksheet2.Cells(X, 13) = "Excess"
            If Month(Today) = 1 Then
                _Del = "Stock(December)"
            Else

                _Del = "Stock(" & MonthName(Month(Today) - 1) & ")"
            End If
            worksheet2.Cells(X, 14) = _Del

            worksheet2.Cells(X, 15) = "Stock(" & MonthName(Month(Today)) & ")"
            If Month(Today) = 12 Then
                _Del = "1st Wk (Stock)"
            Else
                _Del = "1st Wk (Stock) " '& MonthName(Month(Today) + 1)
            End If
            worksheet2.Cells(X, 16) = _Del
            If Month(Today) = 12 Then
                _Del = "2nd Wk (Stock)"
            Else
                _Del = "2nd Wk (Stock) " ' & MonthName(Month(Today) + 1)
            End If
            worksheet2.Cells(X, 17) = _Del
            worksheet2.Cells(X, 18) = "WIP(Exam)"
            worksheet2.Cells(X, 19) = "WIP(Finishing)"
            worksheet2.Cells(X, 20) = "WIP(Dyeing)"
            worksheet2.Cells(X, 21) = "Merchant"
            worksheet2.Cells(X, 22) = "Department"

            Dim i As Integer
            i = 4
            Dim _Char As Integer
            _Char = 97
            ' MsgBox(ChrW(_Char))
            For i = 1 To 22

                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Cells(X, i).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(X, i).WrapText = True
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Interior.Color = RGB(216, 216, 216)
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).MergeCells = True
                worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                _Char = _Char + 1

                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            Next
            X = X + 1

            Dim z As Integer
            i = 0
            With frmDelivary_Forcust
                For Each uRow As UltraGridRow In .UltraGrid3.Rows
                    worksheet2.Rows(X).Font.size = 8
                    worksheet2.Rows(X).Font.name = "Times New Roman"

                    _Char = 97
                    For z = 0 To 19
                        worksheet2.Cells(X, z + 1) = .UltraGrid3.Rows(i).Cells(z).Value
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).MergeCells = True
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        If z + 1 = 4 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 5 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ElseIf z + 1 = 6 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ElseIf z + 1 = 7 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ElseIf z + 1 = 8 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignRight
                        ElseIf z + 1 = 10 Then
                            worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 11 Then
                            ' worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 16 Then
                            'worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 17 Then
                            'worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        ElseIf z + 1 = 13 Then
                            ' worksheet2.Cells(X, z + 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                        End If
                    Next

                    _Char = 97
                    sql = "select * from M07TobeDelivered where M07Sales_Order='" & .UltraGrid3.Rows(i).Cells(0).Value & "' and M07Line_Item='" & .UltraGrid3.Rows(i).Cells(1).Value & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, sql)
                    If isValidDataset(dsUser) Then
                        worksheet2.Cells(X, z + 1) = dsUser.Tables(0).Rows(0)("M07Merchant")
                    Else
                        sql = "select * from M06Delivary_Qty where M06Sales_Order='" & .UltraGrid3.Rows(i).Cells(0).Value & "' and M06Line_Item='" & .UltraGrid3.Rows(i).Cells(1).Value & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, sql)
                        If isValidDataset(dsUser) Then
                            worksheet2.Cells(X, z + 1) = dsUser.Tables(0).Rows(0)("M06Merchant")
                        End If
                    End If
                    For z = 1 To 22

                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet2.Cells(X, z).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        ' worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        ' worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Interior.Color = RGB(216, 216, 216)
                        'worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).MergeCells = True
                        ' worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter

                        _Char = _Char + 1

                        'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                    Next
                    X = X + 1
                    i = i + 1
                Next
            End With

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.CLOSE()
            ' worksheet1.Cells(4, 5) = _Fail_Batch
            'worksheet1.Cells(4, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
            ' MsgBox(_Fail_Batch)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.CLOSE()
            End If
        End Try
    End Function

    Function Create_STockComm()
        Dim sql As String
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
        Dim M01 As DataSet
        Dim vcWharer As String

        Dim range1 As Range

        Try
            Dim exc As New Application

            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

            exc.Visible = True

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            '  worksheet2.Rows(2).Font.size = 12
            '  worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 10
            worksheet2.Columns("B").ColumnWidth = 8
            worksheet2.Columns("C").ColumnWidth = 30
            worksheet2.Columns("D").ColumnWidth = 10
            worksheet2.Columns("K").ColumnWidth = 8
            worksheet2.Columns("L").ColumnWidth = 16
            worksheet2.Columns("M").ColumnWidth = 12
            worksheet2.Columns("N").ColumnWidth = 12
            worksheet2.Columns("O").ColumnWidth = 12
            worksheet2.Columns("P").ColumnWidth = 12
            worksheet2.Columns("Q").ColumnWidth = 22
            worksheet2.Columns("J").ColumnWidth = 12
            worksheet2.Columns("K").ColumnWidth = 12
            worksheet2.Columns("V").ColumnWidth = 12
            worksheet2.Columns("W").ColumnWidth = 12
            worksheet2.Columns("X").ColumnWidth = 30
            worksheet2.Columns("Y").ColumnWidth = 30
            worksheet2.Columns("Z").ColumnWidth = 30
            worksheet2.Columns("T").ColumnWidth = 30
            worksheet2.Columns("U").ColumnWidth = 30
            worksheet2.Columns("V").ColumnWidth = 30
            worksheet2.Columns("W").ColumnWidth = 30



            worksheet2.Rows(2).Font.size = 14
            worksheet2.Rows(2).Font.name = "Times New Roman"
            worksheet2.Rows(2).rowheight = 24.25
            worksheet2.Rows(2).Font.bold = True
            worksheet2.Cells(2, 4) = "FG Stock Analysis"
            worksheet2.Cells(2, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("a2:q2").MergeCells = True
            worksheet1.Range("a2:a2").VerticalAlignment = XlVAlign.xlVAlignCenter
            Dim X As Integer
            X = 4
            worksheet2.Rows(X).Font.size = 8
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).rowheight = 24.25
            worksheet2.Rows(X).Font.bold = True

            worksheet2.Cells(X, 1) = "Stock Location"
            worksheet2.Cells(X, 2) = "Material"
            worksheet2.Cells(X, 3) = "Material Description"
            worksheet2.Cells(X, 4) = "Batch"
            worksheet2.Cells(X, 5) = "Sales order"
            worksheet2.Cells(X, 6) = "Line Item"
            worksheet2.Cells(X, 7) = "Qty(m)"
            worksheet2.Cells(X, 8) = "Qty(kg)"
            worksheet2.Cells(X, 9) = "Last GR"
            worksheet2.Cells(X, 10) = "Prod Order No"
            worksheet2.Cells(X, 11) = "# Days"
            worksheet2.Cells(X, 12) = "Range"
            worksheet2.Cells(X, 13) = "Retailer"
            worksheet2.Cells(X, 14) = "Merchant"
            worksheet2.Cells(X, 15) = "Std price(m)"
            worksheet2.Cells(X, 16) = "Std price(kg)"
            worksheet2.Cells(X, 17) = "Convertion Factor to Meters"
            worksheet2.Cells(X, 18) = "BU"
            worksheet2.Cells(X, 19) = "Ship to party"
            worksheet2.Cells(X, 20) = "1st Week"
            worksheet2.Cells(X, 21) = "2nd Week"
            worksheet2.Cells(X, 22) = "3rd Week"
            worksheet2.Cells(X, 23) = "4th Week"
            worksheet2.Cells(X, 24) = "5th Week"
            worksheet2.Cells(X, 25) = "latest update"
            worksheet2.Cells(X, 26) = "Dedline to clear"
            worksheet2.Cells(X, 27) = "Value m ($)"

            Dim i As Integer
            i = 4
            Dim _Char As Integer
            _Char = 97
            ' MsgBox(ChrW(_Char))
            For i = 1 To 27

                If i = 27 Then
                    _Char = 97

                    worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Cells(X, i).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, i).WrapText = True
                    worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Interior.Color = RGB(216, 216, 216)
                    worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).MergeCells = True
                    worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                Else
                    worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Cells(X, i).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(X, i).WrapText = True
                    worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Interior.Color = RGB(216, 216, 216)
                    worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).MergeCells = True
                    worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                End If
                _Char = _Char + 1

                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

            Next
            ' X = X + 1


            With frmFG_Stock
                If Trim(.txtCustomer.Text) <> "" And Trim(.txtDepartment.Text) <> "" And Trim(.txtMerchant.Text) <> "" And .txtBU.Text <> "" And Trim(.txtLocation.Text) <> "" And .cboStatus.Text <> "" Then

                ElseIf Trim(.txtCustomer.Text) <> "" And Trim(.txtDepartment.Text) <> "" And Trim(.txtMerchant.Text) <> "" And .txtBU.Text <> "" And Trim(.txtLocation.Text) <> "" Then

                ElseIf Trim(.txtCustomer.Text) <> "" And Trim(.txtDepartment.Text) <> "" And Trim(.txtMerchant.Text) <> "" And .txtBU.Text <> "" Then

                ElseIf Trim(.txtCustomer.Text) <> "" And Trim(.txtDepartment.Text) <> "" And Trim(.txtMerchant.Text) <> "" Then
                ElseIf Trim(.cboStatus.Text) <> "" And Trim(.txtLocation.Text) <> "" And Trim(.txtBU.Text) <> "" Then
                    If Trim(.cboStatus.Text) = "Below One Month" Then
                        vcWharer = "M14Name in ('" & pln_BU & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                    ElseIf Trim(.cboStatus.Text) = "Over One Months" Then
                        vcWharer = "M14Name in ('" & pln_BU & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Two Months" Then
                        vcWharer = "M14Name in ('" & pln_BU & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Three Months" Then
                        vcWharer = "M14Name in ('" & pln_BU & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Fore Months" Then
                        vcWharer = "M14Name in ('" & pln_BU & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Five Months" Then
                        vcWharer = "M14Name in ('" & pln_BU & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                    ElseIf Trim(.cboStatus.Text) = "Over Six Months" Then
                        vcWharer = "M14Name in ('" & pln_BU & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over One Year" Then
                        vcWharer = "M14Name in ('" & pln_BU & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    End If


                ElseIf Trim(.cboStatus.Text) <> "" And Trim(.txtLocation.Text) <> "" And Trim(.txtMerchant.Text) <> "" Then
                    If Trim(.cboStatus.Text) = "Below One Month" Then
                        vcWharer = "M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                    ElseIf Trim(.cboStatus.Text) = "Over One Months" Then
                        vcWharer = "M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Two Months" Then
                        vcWharer = "M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Three Months" Then
                        vcWharer = "M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Fore Months" Then
                        vcWharer = "M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Five Months" Then
                        vcWharer = "M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                    ElseIf Trim(.cboStatus.Text) = "Over Six Months" Then
                        vcWharer = "M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over One Year" Then
                        vcWharer = "M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    End If
                ElseIf Trim(.txtCustomer.Text) <> "" And Trim(.txtDepartment.Text) <> "" Then
                ElseIf Trim(.txtCustomer.Text) <> "" And Trim(.cboStatus.Text) <> "" Then
                    If Trim(.cboStatus.Text) = "Below One Month" Then
                        vcWharer = "M01Cuatomer_Name in ('" & pln_Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                    ElseIf Trim(.cboStatus.Text) = "Over One Months" Then
                        vcWharer = "M01Cuatomer_Name in ('" & pln_Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Two Months" Then
                        vcWharer = "M01Cuatomer_Name in ('" & pln_Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Three Months" Then
                        vcWharer = "M01Cuatomer_Name in ('" & pln_Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Fore Months" Then
                        vcWharer = "M01Cuatomer_Name in ('" & pln_Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Five Months" Then
                        vcWharer = "M01Cuatomer_Name in ('" & pln_Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                    ElseIf Trim(.cboStatus.Text) = "Over Six Months" Then
                        vcWharer = "M01Cuatomer_Name in ('" & pln_Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over One Year" Then
                        vcWharer = "M01Cuatomer_Name in ('" & pln_Customer & "') and DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    End If
                ElseIf Trim(.txtCustomer.Text) <> "" And Trim(.txtLocation.Text) <> "" Then
                    vcWharer = " M01Cuatomer_Name in ('" & pln_Customer & "') and M08Location in ('" & pln_Location & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(.txtCustomer.Text) <> "" And Trim(.txtLocation.Text) <> "" Then
                    vcWharer = " M01Cuatomer_Name in ('" & pln_Customer & "') and M08Location in ('" & pln_Location & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(.txtMerchant.Text) <> "" And Trim(.txtLocation.Text) <> "" Then
                    vcWharer = " M08Merchant in ('" & pln_Merchnt & "') and M08Location in ('" & pln_Location & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(.txtCustomer.Text) <> "" Then
                    vcWharer = " M01Cuatomer_Name in ('" & pln_Customer & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(.txtDepartment.Text) <> "" And Trim(.txtLocation.Text) <> "" Then
                    vcWharer = " M08Retailer in ('" & pln_Retailer & "') AND  M08Location in ('" & pln_Location & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(.txtDepartment.Text) <> "" Then
                    vcWharer = " M08Retailer in ('" & pln_Retailer & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(.txtMerchant.Text) <> "" Then
                    vcWharer = " M08Merchant in ('" & pln_Merchnt & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(.txtBU.Text) <> "" Then
                ElseIf Trim(.txtLocation.Text) <> "" Then
                    vcWharer = " M08Location in ('" & pln_Location & "')"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                ElseIf Trim(.cboStatus.Text) <> "" Then
                    If Trim(.cboStatus.Text) = "Below One Month" Then
                        vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) <=30 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                    ElseIf Trim(.cboStatus.Text) = "Over One Months" Then
                        vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >30 and DATEDIFF(day,  m08TR_Date,GETDATE()) <60 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Two Months" Then
                        vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=60 and DATEDIFF(day,  m08TR_Date,GETDATE()) <90 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Three Months" Then
                        vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=90 and DATEDIFF(day,  m08TR_Date,GETDATE()) <120 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Fore Months" Then
                        vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=120 and DATEDIFF(day,  m08TR_Date,GETDATE()) <150 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over Five Months" Then
                        vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=150 and DATEDIFF(day,  m08TR_Date,GETDATE()) <180 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))

                    ElseIf Trim(.cboStatus.Text) = "Over Six Months" Then
                        vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=180 and DATEDIFF(day,  m08TR_Date,GETDATE()) <360 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    ElseIf Trim(.cboStatus.Text) = "Over One Year" Then
                        vcWharer = "DATEDIFF(day,  m08TR_Date,GETDATE()) >=365 "
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetStockCom", New SqlParameter("@cQryType", "LST"), New SqlParameter("@vcWhereClause1", vcWharer))
                    End If
                Else
                    sql = "select M14Name,M08Batch_No,M08TR_Date,M08Location,M08Meterial,M08Dis,M01Cuatomer_Name,M08Retailer,M08Merchant,M08Sales_Order,M08Line_Item,M08RollNo,M08Qty_KG,M08Qty_Mtr from M08Stock inner join M01Sales_Order_SAP on M01Sales_Order=M08Sales_Order and M08Line_Item=M01Line_Item inner join M13Biz_Unit on M13Merchant=M08Merchant inner join M14Retailer on M13Department=M14Code order by M08Meterial"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, sql)
                End If

            End With

            Dim diff1 As System.TimeSpan
            Dim date2 As System.DateTime
            Dim date1 As System.DateTime
            Dim _RowFirst As Integer

            i = 0
            X = X + 1
            _RowFirst = X
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                worksheet2.Rows(X).Font.size = 8
                worksheet2.Rows(X).Font.name = "Times New Roman"

                worksheet2.Cells(X, 1) = M01.Tables(0).Rows(i)("M08Location")
                worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(X, 2) = M01.Tables(0).Rows(i)("M08Meterial")
                worksheet2.Cells(X, 3) = M01.Tables(0).Rows(i)("M08Dis")
                worksheet2.Cells(X, 4) = M01.Tables(0).Rows(i)("M08RollNo")
                worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(X, 5) = M01.Tables(0).Rows(i)("M08Sales_Order")
                worksheet2.Cells(X, 6) = M01.Tables(0).Rows(i)("M08Line_Item")
                worksheet2.Cells(X, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(X, 7) = M01.Tables(0).Rows(i)("M08Line_Item")
                'worksheet2.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                'worksheet2.Cells(X, 8) = M01.Tables(0).Rows(i)("M08RollNo")
                'worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(X, 7) = M01.Tables(0).Rows(i)("M08Qty_Mtr")
                worksheet2.Cells(X, 7).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet2.Cells(X, 7)
                range1.NumberFormat = "0"

                worksheet2.Cells(X, 8) = M01.Tables(0).Rows(i)("M08Qty_KG")
                worksheet2.Cells(X, 8).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet2.Cells(X, 8)
                range1.NumberFormat = "0"

                If IsDBNull(M01.Tables(0).Rows(i)("M08TR_Date")) = True Then
                Else
                    worksheet2.Cells(X, 9) = Month(M01.Tables(0).Rows(i)("M08TR_Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M08TR_Date")) & "/" & Year(M01.Tables(0).Rows(i)("M08TR_Date"))
                    worksheet2.Cells(X, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet2.Cells(X, 10) = M01.Tables(0).Rows(i)("M08Batch_No")
                    worksheet2.Cells(X, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    date2 = M01.Tables(0).Rows(i)("M08TR_Date")
                    date1 = Today
                    diff1 = date1.Subtract(date2)
                    worksheet2.Cells(X, 11) = diff1.Days
                    worksheet2.Cells(X, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If diff1.Days < 30 Then
                        worksheet2.Cells(X, 12) = "Below One Month"
                    ElseIf diff1.Days >= 30 And diff1.Days < 60 Then
                        worksheet2.Cells(X, 12) = "Over One Months"
                    ElseIf diff1.Days >= 60 And diff1.Days < 90 Then
                        worksheet2.Cells(X, 12) = "Over Two Months"
                    ElseIf diff1.Days >= 90 And diff1.Days < 120 Then
                        worksheet2.Cells(X, 12) = "Over Three Months"
                    ElseIf diff1.Days >= 120 And diff1.Days < 150 Then
                        worksheet2.Cells(X, 12) = "Over Four Months"
                    ElseIf diff1.Days >= 150 And diff1.Days < 180 Then
                        worksheet2.Cells(X, 12) = "Over Five Months"
                    ElseIf diff1.Days >= 180 And diff1.Days < 365 Then
                        worksheet2.Cells(X, 12) = "Over Six Month"

                    ElseIf diff1.Days >= 365 Then
                        worksheet2.Cells(X, 12) = "Over One Year"

                    End If
                End If
                worksheet2.Cells(X, 13) = M01.Tables(0).Rows(i)("M08Retailer")
                worksheet2.Cells(X, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(X, 14) = M01.Tables(0).Rows(i)("M08Merchant")
                worksheet2.Cells(X, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet2.Cells(X, 17) = M01.Tables(0).Rows(i)("M08Qty_Mtr") / M01.Tables(0).Rows(i)("M08Qty_KG")
                worksheet2.Cells(X, 17).HorizontalAlignment = XlHAlign.xlHAlignRight
                range1 = worksheet2.Cells(X, 17)
                range1.NumberFormat = "0.0"

                worksheet2.Cells(X, 18) = M01.Tables(0).Rows(i)("M14Name") ' / M01.Tables(0).Rows(i)("M08Qty_KG")
                worksheet2.Cells(X, 18).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'newRow("Qty(Kg)") = M01.Tables(0).Rows(i)("M08Qty_KG")
                'newRow("GRN Date") = Month(M01.Tables(0).Rows(i)("M08TR_Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("M08TR_Date")) & "/" & Year(M01.Tables(0).Rows(i)("M08TR_Date"))
                'date2 = M01.Tables(0).Rows(i)("M08TR_Date")
                'date1 = Today
                'diff1 = date1.Subtract(date2)
                '_Date = diff1.Days
                'If _Date < 30 Then
                '    newRow("Ageing") = "Over One Month"
                'ElseIf _Date > 30 And _Date < 60 Then
                '    newRow("Ageing") = "Over Two Months"
                'ElseIf _Date > 60 And _Date < 90 Then
                '    newRow("Ageing") = "Over Three Months"
                'ElseIf _Date > 90 And _Date < 120 Then
                '    newRow("Ageing") = "Over Four Months"
                'ElseIf _Date > 120 And _Date < 150 Then
                '    newRow("Ageing") = "Over Five Months"
                'ElseIf _Date > 150 And _Date < 180 Then
                '    newRow("Ageing") = "Over Six Months"
                'ElseIf _Date > 180 Then
                '    newRow("Ageing") = "Over One Year"


                'End If

                sql = "select * from T09Stock_Comments where  T09Year=" & Year(Today) & " and T09Location='" & M01.Tables(0).Rows(i)("M08Location") & "' and T09Roll_No='" & M01.Tables(0).Rows(i)("M08RollNo") & "' order by T09Date DEsc"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, sql)
                If isValidDataset(dsUser) Then
                    If IsDBNull(dsUser.Tables(0).Rows(0)("T09Ded_Date")) = True Then
                    Else
                        If Year(dsUser.Tables(0).Rows(0)("T09Ded_Date")) = "1900" Then

                        Else
                            worksheet2.Cells(X, 26) = Month(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Microsoft.VisualBasic.Day(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Year(dsUser.Tables(0).Rows(0)("T09Ded_Date"))
                            worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter

                            worksheet2.Cells(X, 25) = dsUser.Tables(0).Rows(0)("T09Comment")
                            worksheet2.Cells(X, 25).HorizontalAlignment = XlHAlign.xlHAlignLeft
                        End If
                    End If
                End If

                Dim _Frm As Date
                Dim _To As Date

                _Frm = Month(Today) & "/1/" & Year(Today)
                _To = _Frm.AddDays(+7)

                '1ST WEEK
                sql = "select * from T09Stock_Comments where  T09DATE BETWEEN '" & _Frm & "' AND '" & _To & "' and T09Location='" & M01.Tables(0).Rows(i)("M08Location") & "' and T09Roll_No='" & M01.Tables(0).Rows(i)("M08RollNo") & "' order by T09Date DEsc"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, sql)
                If isValidDataset(dsUser) Then
                    If Year(dsUser.Tables(0).Rows(0)("T09Ded_Date")) = "1900" Then

                    Else
                        'worksheet2.Cells(X, 26) = Month(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Microsoft.VisualBasic.Day(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Year(dsUser.Tables(0).Rows(0)("T09Ded_Date"))
                        'worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet2.Cells(X, 20) = dsUser.Tables(0).Rows(0)("T09Comment")
                        worksheet2.Cells(X, 20).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If
                End If
                '2ND WEEK
                _Frm = _To.AddDays(+1)
                _To = _Frm.AddDays(+7)
                sql = "select * from T09Stock_Comments where  T09DATE BETWEEN '" & _Frm & "' AND '" & _To & "' and T09Location='" & M01.Tables(0).Rows(i)("M08Location") & "' and T09Roll_No='" & M01.Tables(0).Rows(i)("M08RollNo") & "' order by T09Date DEsc"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, sql)
                If isValidDataset(dsUser) Then
                    If Year(dsUser.Tables(0).Rows(0)("T09Ded_Date")) = "1900" Then

                    Else
                        'worksheet2.Cells(X, 26) = Month(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Microsoft.VisualBasic.Day(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Year(dsUser.Tables(0).Rows(0)("T09Ded_Date"))
                        'worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet2.Cells(X, 21) = dsUser.Tables(0).Rows(0)("T09Comment")
                        worksheet2.Cells(X, 21).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If
                End If

                '3RD WEEK
                _Frm = _To.AddDays(+1)
                _To = _Frm.AddDays(+7)
                sql = "select * from T09Stock_Comments where  T09DATE BETWEEN '" & _Frm & "' AND '" & _To & "' and T09Location='" & M01.Tables(0).Rows(i)("M08Location") & "' and T09Roll_No='" & M01.Tables(0).Rows(i)("M08RollNo") & "' order by T09Date DEsc"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, sql)
                If isValidDataset(dsUser) Then
                    If Year(dsUser.Tables(0).Rows(0)("T09Ded_Date")) = "1900" Then

                    Else
                        'worksheet2.Cells(X, 26) = Month(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Microsoft.VisualBasic.Day(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Year(dsUser.Tables(0).Rows(0)("T09Ded_Date"))
                        'worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet2.Cells(X, 22) = dsUser.Tables(0).Rows(0)("T09Comment")
                        worksheet2.Cells(X, 22).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If
                End If

                '4TH WEEK
                _Frm = _To.AddDays(+1)
                _To = _Frm.AddDays(+7)
                sql = "select * from T09Stock_Comments where  T09DATE BETWEEN '" & _Frm & "' AND '" & _To & "' and T09Location='" & M01.Tables(0).Rows(i)("M08Location") & "' and T09Roll_No='" & M01.Tables(0).Rows(i)("M08RollNo") & "' order by T09Date DEsc"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, sql)
                If isValidDataset(dsUser) Then
                    If Year(dsUser.Tables(0).Rows(0)("T09Ded_Date")) = "1900" Then

                    Else
                        'worksheet2.Cells(X, 26) = Month(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Microsoft.VisualBasic.Day(dsUser.Tables(0).Rows(0)("T09Ded_Date")) & "/" & Year(dsUser.Tables(0).Rows(0)("T09Ded_Date"))
                        'worksheet2.Cells(X, 26).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet2.Cells(X, 23) = dsUser.Tables(0).Rows(0)("T09Comment")
                        worksheet2.Cells(X, 23).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    End If
                End If
                _Char = 97
                Dim Y As Integer

                ' MsgBox(ChrW(_Char))
                For Y = 1 To 27

                    If Y = 27 Then
                        _Char = 97

                        worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        'worksheet2.Cells(X, Y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        ' worksheet1.Cells(X, Y).WrapText = True
                        worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        ' worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).Interior.Color = RGB(216, 216, 216)
                        worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).MergeCells = True
                        worksheet2.Range("A" & ChrW(_Char) & X, "A" & ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    Else
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        ' worksheet2.Cells(X, Y).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        ' worksheet1.Cells(X, Y).WrapText = True
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                        ' worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).Interior.Color = RGB(216, 216, 216)
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).MergeCells = True
                        worksheet2.Range(ChrW(_Char) & X, ChrW(_Char) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                    End If
                    _Char = _Char + 1

                    'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                    'worksheet1.Range(ChrW(_Char) & i, ChrW(_Char) & i).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                Next

                X = X + 1
                i = i + 1

            Next

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.CLOSE()
            MsgBox("Report Genarated successfully", MsgBoxStyle.Information, "Technova ....")
            ' MsgBox(_Fail_Batch)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.CLOSE()
            End If
        End Try
    End Function
End Module
