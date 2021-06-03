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
Public Class frmExcess_Analysis
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Private Sub frmExcess_Analysis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        cboMaterial.Text = ""
        cboQuality.Text = ""
    End Sub

    Private Sub chkReq_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReq.CheckedChanged
        If chkReq.Checked = True Then
            chkStock.Checked = False
            chkWIP.Checked = False
            Call Load_Quality()
            Call Load_Material()

        End If
    End Sub

    Function Load_Quality()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try
            If chkReq.Checked = True Then
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "CQTB"))
                If isValidDataset(M01) Then
                    With cboQuality
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 245
                    End With
                End If
            ElseIf chkStock.Checked = True Then
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "SMM"))
                If isValidDataset(M01) Then
                    With cboQuality
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 245
                    End With
                End If
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

    Function Load_Material()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try
            If chkReq.Checked = True Then
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "CMTB"))
                If isValidDataset(M01) Then
                    With cboMaterial
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 245
                    End With
                End If
            ElseIf chkStock.Checked = True Then
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "SMS"))
                If isValidDataset(M01) Then
                    With cboMaterial
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 245
                    End With
                End If

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

    Function Create_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim M01 As DataSet

        Dim vcWhere As String
        Try
            Dim exc As New Application

            Dim workbooks As Workbooks = exc.Workbooks
            Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            Dim sheets As Sheets = workbook.Worksheets
            Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
            Dim range1 As Range

            exc.Visible = True

            Dim sheets1 As Sheets = workbook.Worksheets
            Dim worksheet2 As _Worksheet = CType(sheets1.Item(1), _Worksheet)
            worksheet2.Rows(2).Font.size = 11
            worksheet2.Rows(2).Font.Bold = True
            worksheet2.Columns("A").ColumnWidth = 10
            worksheet2.Columns("B").ColumnWidth = 30
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

            If chkReq.Checked = True Then
                worksheet2.Cells(1, 1) = "Production Excess Analysis -  Requrment basis "
                worksheet2.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ElseIf chkStock.Checked = True Then
                worksheet2.Cells(1, 1) = "Production Excess Analysis -  Stock basis "
                worksheet2.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

            ElseIf chkWIP.Checked = True Then
                worksheet2.Cells(1, 1) = "Production Excess Analysis -  WIP basis "
                worksheet2.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            End If
            worksheet2.Range("A1:Q1").Interior.Color = RGB(197, 217, 241)
            ' worksheet2.Range("A2:M2").Interior.Color = RGB(197, 217, 241)
            worksheet2.Rows(1).Font.size = 13
            worksheet2.Rows(1).rowheight = 55
            worksheet2.Rows(1).Font.name = "Times New Roman"
            worksheet2.Rows(1).Font.BOLD = True
            worksheet2.Range("A1:Q1").MergeCells = True
            worksheet2.Range("A1:Q1").VerticalAlignment = XlVAlign.xlVAlignCenter

            Dim X As Integer
            X = 3

            worksheet2.Rows(X).rowheight = 22
            worksheet2.Cells(X, 1) = "Material"
            worksheet2.Rows(X).Font.size = 9
            worksheet2.Rows(X).Font.name = "Times New Roman"
            worksheet2.Rows(X).Font.BOLD = True
            worksheet2.Range("A3:A3").MergeCells = True
            worksheet2.Range("A3:A3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 2) = "Description"
            worksheet2.Range("B3:B3").MergeCells = True
            worksheet2.Range("B3:B3").VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet2.Cells(X, 3) = "S/order"
            worksheet2.Range("C3:C3").MergeCells = True
            worksheet2.Range("C3:C3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 4) = "Line Item"
            worksheet2.Range("D3:D3").MergeCells = True
            worksheet2.Range("D3:D3").VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet2.Cells(X, 5) = " Balace to be  Deliver qty"
            worksheet1.Cells(X, 5).WrapText = True
            worksheet2.Range("E3:E3").MergeCells = True
            worksheet2.Range("E3:E3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(X, 6) = "Total To be deliverd Qty"
            worksheet1.Cells(X, 6).WrapText = True
            worksheet2.Range("F3:F3").MergeCells = True
            worksheet2.Range("F3:F3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 7) = "Qty in 2060 / 2059"
            worksheet1.Cells(X, 7).WrapText = True
            worksheet2.Range("G3:G3").MergeCells = True
            worksheet2.Range("G3:G3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 8) = "Qty in 2065"
            worksheet1.Cells(X, 8).WrapText = True
            worksheet2.Range("H3:H3").MergeCells = True
            worksheet2.Range("H3:H3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 9) = "Qty in 2070"
            worksheet1.Cells(X, 9).WrapText = True
            worksheet2.Range("I3:I3").MergeCells = True
            worksheet2.Range("I3:I3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(X, 10) = "Qty in 2062"
            worksheet1.Cells(X, 10).WrapText = True
            worksheet2.Range("J3:J3").MergeCells = True
            worksheet2.Range("J3:J3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Cells(X, 11) = "Qty in 2055"
            worksheet1.Cells(X, 11).WrapText = True
            worksheet2.Range("K3:K3").MergeCells = True
            worksheet2.Range("K3:K3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 12) = "Qty in  Exam"
            worksheet1.Cells(X, 12).WrapText = True
            worksheet2.Range("L3:L3").MergeCells = True
            worksheet2.Range("L3:L3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 13) = "Qty in  Finishing"
            worksheet1.Cells(X, 13).WrapText = True
            worksheet2.Range("M3:M3").MergeCells = True
            worksheet2.Range("M3:M3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 14) = "Qty in Dyeing"
            worksheet1.Cells(X, 14).WrapText = True
            worksheet2.Range("N3:N3").MergeCells = True
            worksheet2.Range("N3:N3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 15) = " Total  produce Qty"
            worksheet1.Cells(X, 15).WrapText = True
            worksheet2.Range("O3:O3").MergeCells = True
            worksheet2.Range("O3:O3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 16) = "Deference"
            worksheet1.Cells(X, 16).WrapText = True
            worksheet2.Range("P3:P3").MergeCells = True
            worksheet2.Range("P3:P3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X, 17) = "Excess / Shortage"
            worksheet1.Cells(X, 17).WrapText = True
            worksheet2.Range("Q3:Q3").MergeCells = True
            worksheet2.Range("Q3:Q3").VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(3, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            Dim _Chr As Integer
            Dim I As Integer

            _Chr = 97
            For I = 1 To 17

                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                _Chr = _Chr + 1

            Next

            If chkReq.Checked = True Then
                If Trim(cboMaterial.Text) <> "" Then
                    vcWhere = "M07Material='" & Trim(cboMaterial.Text) & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "TB2"), New SqlParameter("@vcWhereClause1", vcWhere))
                ElseIf Trim(cboQuality.Text) <> "" Then
                    vcWhere = "M07Quality='" & Trim(cboQuality.Text) & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "TB2"), New SqlParameter("@vcWhereClause1", vcWhere))
                Else
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "TB1"))
                End If
            ElseIf chkStock.Checked = True Then
                If Trim(cboMaterial.Text) <> "" Then
                    vcWhere = "M08Meterial='" & Trim(cboMaterial.Text) & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "ST1"), New SqlParameter("@vcWhereClause1", vcWhere))
                ElseIf Trim(cboQuality.Text) <> "" Then
                    'vcWhere = "M07Quality='" & Trim(cboQuality.Text) & "'"
                    'T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "TB2"), New SqlParameter("@vcWhereClause1", vcWhere))
                Else
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "STK"))
                End If

            ElseIf chkWIP.Checked = True Then
                If Trim(cboMaterial.Text) <> "" Then
                    vcWhere = "M08Meterial='" & Trim(cboMaterial.Text) & "'"
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "WIP"), New SqlParameter("@vcWhereClause1", vcWhere))
                ElseIf Trim(cboQuality.Text) <> "" Then
                    'vcWhere = "M07Quality='" & Trim(cboQuality.Text) & "'"
                    'T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "TB2"), New SqlParameter("@vcWhereClause1", vcWhere))
                Else
                    T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "WIP1"))
                End If
            End If

            Dim _Fist As Integer
            Dim _Last As Integer
            Dim _Material As String
            Dim Y As Integer

            I = 0
            X = X + 1
            _Fist = X
            _Last = X
            _Material = ""
            For Each DTRow3 As DataRow In T01.Tables(0).Rows
                worksheet2.Rows(X).Font.size = 10
                worksheet2.Rows(X).Font.name = "Times New Roman"
                If _Material = T01.Tables(0).Rows(I)("M07Material") Then

                Else
                    If _Material <> "" Then
                        worksheet2.Range("F" & (X - 1)).Formula = "=SUM(e" & _Fist & ":e" & _Last & ")"
                        worksheet2.Range("F" & _Fist & ":F" & _Last).MergeCells = True
                        '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
                        worksheet2.Range("f" & (X - 1) & ":F" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(X - 1, 6)
                        range1.NumberFormat = "0.00"

                        '2060/2059

                        vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2060','2059')"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            worksheet2.Cells(X - 1, 7) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                            worksheet2.Range("G" & (X - 1) & ":G" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(X - 1, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(X - 1, 7)
                            range1.NumberFormat = "0.00"

                        End If
                        worksheet2.Range("G" & _Fist & ":G" & _Last).MergeCells = True
                        worksheet2.Range("G" & (X - 1) & ":G" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        '2065
                        worksheet2.Range("H" & _Fist & ":H" & _Last).MergeCells = True
                        vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2065')"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            worksheet2.Cells(X - 1, 8) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                            worksheet2.Cells(X - 1, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(X - 1, 8)
                            range1.NumberFormat = "0.00"

                        End If

                        worksheet2.Range("H" & (X - 1) & ":H" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        '2070
                        worksheet2.Range("I" & _Fist & ":I" & _Last).MergeCells = True
                        vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2070')"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            worksheet2.Cells(X - 1, 9) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                            worksheet2.Cells(X - 1, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(X - 1, 9)
                            range1.NumberFormat = "0.00"

                        End If

                        worksheet2.Range("I" & (X - 1) & ":I" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        '2062
                        worksheet2.Range("J" & _Fist & ":J" & _Last).MergeCells = True
                        vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2062')"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            worksheet2.Cells(X - 1, 10) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                            worksheet2.Cells(X - 1, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(X - 1, 10)
                            range1.NumberFormat = "0.00"

                        End If

                        worksheet2.Range("J" & (X - 1) & ":J" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        '2055
                        ' worksheet2.Range("K" & _Fist & ":K" & _Last).MergeCells = True
                        vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2055')"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            worksheet2.Cells(X - 1, 11) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                            worksheet2.Cells(X - 1, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(X - 1, 11)
                            range1.NumberFormat = "0.00"

                        End If
                        worksheet2.Range("K" & _Fist & ":K" & _Last).MergeCells = True
                        worksheet2.Range("K" & (X - 1) & ":K" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        'EXAM
                        worksheet2.Range("L" & _Fist & ":L" & _Last).MergeCells = True
                        vcWhere = " M09Meterial='" & Trim(_Material) & "' AND M09Oredr_Type='Exam'"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            worksheet2.Cells(X - 1, 12) = M01.Tables(0).Rows(0)("M09Qty_Mtr")
                            worksheet2.Cells(X - 1, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(X - 1, 12)
                            range1.NumberFormat = "0.00"

                        End If

                        worksheet2.Range("L" & _Fist & ":L" & _Last).MergeCells = True
                        worksheet2.Range("L" & (X - 1) & ":L" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        'FINISHING
                        ' worksheet2.Range("M" & _Fist & ":M" & _Last).MergeCells = True
                        vcWhere = " M09Meterial='" & Trim(_Material) & "' AND M09Oredr_Type='Finishing'"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            worksheet2.Cells(X - 1, 13) = M01.Tables(0).Rows(0)("M09Qty_Mtr")
                            worksheet2.Cells(X - 1, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(X - 1, 13)
                            range1.NumberFormat = "0.00"

                        End If
                        worksheet2.Range("M" & _Fist & ":M" & _Last).MergeCells = True
                        worksheet2.Range("M" & (X - 1) & ":M" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter


                        'DYEING
                        ' worksheet2.Range("N" & _Fist & ":N" & _Last).MergeCells = True
                        vcWhere = " M09Meterial='" & Trim(_Material) & "' AND M09Oredr_Type='Dyeing'"
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            worksheet2.Cells(X - 1, 14) = M01.Tables(0).Rows(0)("M09Qty_Mtr")
                            worksheet2.Cells(X - 1, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            range1 = worksheet2.Cells(X - 1, 14)
                            range1.NumberFormat = "0.00"

                        End If
                        worksheet2.Range("N" & _Fist & ":N" & _Last).MergeCells = True
                        worksheet2.Range("N" & (X - 1) & ":N" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter


                        worksheet2.Range("O" & (X - 1)).Formula = "=SUM(G" & _Fist & ":N" & X - 1 & ")"
                        worksheet2.Range("O" & _Fist & ":O" & _Last).MergeCells = True
                        worksheet2.Range("O" & (X - 1) & ":O" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(X - 1, 15)
                        range1.NumberFormat = "0.00"


                        worksheet2.Cells(X - 1, 16) = "=F" & _Fist & "-O" & _Fist
                        worksheet2.Range("P" & _Fist & ":P" & _Last).MergeCells = True
                        worksheet2.Range("P" & (X - 1) & ":P" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                        worksheet2.Cells(X - 1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

                        worksheet2.Cells(X - 1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        range1 = worksheet2.Cells(X - 1, 16)
                        range1.NumberFormat = "0.00"

                        range1 = worksheet2.Cells(_Fist, 16)

                        If range1.Value > 0 Then
                            worksheet2.Cells(X - 1, 17) = "Shortage"
                            worksheet2.Range("Q" & _Fist & ":Q" & _Last).MergeCells = True
                            worksheet2.Range("Q" & (X - 1) & ":Q" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        Else
                            worksheet2.Cells(X - 1, 17) = "Excess"
                            worksheet2.Range("Q" & _Fist & ":Q" & _Last).MergeCells = True
                            worksheet2.Range("Q" & (X - 1) & ":Q" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                            worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        End If
                        _Fist = X
                    End If
                    _Material = T01.Tables(0).Rows(I)("M07Material")

                End If
                Dim _Qty As Double
                _Qty = 0

                worksheet2.Cells(X, 1) = T01.Tables(0).Rows(I)("M07Material")
                worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(X, 2) = T01.Tables(0).Rows(I)("M07Met_Dis")
                worksheet2.Cells(X, 3) = T01.Tables(0).Rows(I)("M07Sales_Order")
                worksheet2.Cells(X, 4) = T01.Tables(0).Rows(I)("M07Line_Item")
                worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                If Not IsDBNull(T01.Tables(0).Rows(I)("M07Qty_Mtr")) = True Then


                    _Qty = T01.Tables(0).Rows(I)("M07Qty_Mtr")
                    vcWhere = "M01Sales_Order='" & T01.Tables(0).Rows(I)("M07Sales_Order") & "' and M01Line_Item='" & T01.Tables(0).Rows(I)("M07Line_Item") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "DEL"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        '   _Qty = _Qty + M01.Tables(0).Rows(0)("M01SO_Qty")
                    End If
                    ' Dim Qty1 As Double
                    '  Qty1 = 0
                    vcWhere = "Sales_Order='" & T01.Tables(0).Rows(I)("M07Sales_Order") & "' and Line_Item='" & T01.Tables(0).Rows(I)("M07Line_Item") & "'"
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "TOL"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        Dim _Qty1 As Double

                        _Qty1 = _Qty * M01.Tables(0).Rows(0)("Tollarance_PLS")
                        _Qty1 = _Qty / 100
                        _Qty = _Qty + _Qty1
                    End If
                    worksheet2.Cells(X, 5) = _Qty
                    worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet2.Cells(X, 5)
                    range1.NumberFormat = "0.00"
                End If
                _Chr = 97
                For Y = 1 To 17

                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                    worksheet2.Range(Chr(_Chr) & X, Chr(_Chr) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                    ' worksheet2.Range(Chr(_Chr) & X & ":" & Chr(_Chr) & X).Interior.Color = RGB(197, 217, 241)
                    _Chr = _Chr + 1

                Next

                _Last = X
                X = X + 1
                I = I + 1
            Next

            worksheet2.Range("F" & (X - 1)).Formula = "=SUM(e" & _Fist & ":e" & _Last & ")"
            worksheet2.Range("F" & _Fist & ":F" & _Last).MergeCells = True
            '  worksheet2.Range("F" & (X)).Formula = "=SUM(F" & _Fist & ":F" & _Last & ")"
            worksheet2.Range("f" & (X - 1) & ":F" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X - 1, 6)
            range1.NumberFormat = "0.00"

            '2060/2059

            vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2060','2059')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                worksheet2.Cells(X - 1, 7) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                worksheet2.Range("G" & (X - 1) & ":G" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet2.Cells(X - 1, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X - 1, 7)
                range1.NumberFormat = "0.00"

            End If
            worksheet2.Range("G" & _Fist & ":G" & _Last).MergeCells = True
            worksheet2.Range("G" & (X - 1) & ":G" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '2065
            worksheet2.Range("H" & _Fist & ":H" & _Last).MergeCells = True
            vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2065')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                worksheet2.Cells(X - 1, 8) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                worksheet2.Cells(X - 1, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X - 1, 8)
                range1.NumberFormat = "0.00"

            End If

            worksheet2.Range("H" & (X - 1) & ":H" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '2070
            worksheet2.Range("I" & _Fist & ":I" & _Last).MergeCells = True
            vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2070')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                worksheet2.Cells(X - 1, 9) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                worksheet2.Cells(X - 1, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X - 1, 9)
                range1.NumberFormat = "0.00"

            End If

            worksheet2.Range("I" & (X - 1) & ":I" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '2062
            worksheet2.Range("J" & _Fist & ":J" & _Last).MergeCells = True
            vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2062')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                worksheet2.Cells(X - 1, 10) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                worksheet2.Cells(X - 1, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X - 1, 10)
                range1.NumberFormat = "0.00"

            End If

            worksheet2.Range("J" & (X - 1) & ":J" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

            '2055
            ' worksheet2.Range("K" & _Fist & ":K" & _Last).MergeCells = True
            vcWhere = "M08Meterial='" & Trim(_Material) & "' AND M08Location IN('2055')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "FGS"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                worksheet2.Cells(X - 1, 11) = M01.Tables(0).Rows(0)("M08Qty_Mtr")
                worksheet2.Cells(X - 1, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X - 1, 11)
                range1.NumberFormat = "0.00"

            End If
            worksheet2.Range("K" & _Fist & ":K" & _Last).MergeCells = True
            worksheet2.Range("K" & (X - 1) & ":K" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

            'EXAM
            worksheet2.Range("L" & _Fist & ":L" & _Last).MergeCells = True
            vcWhere = " M09Meterial='" & Trim(_Material) & "' AND M09Oredr_Type='Exam'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                worksheet2.Cells(X - 1, 12) = M01.Tables(0).Rows(0)("M09Qty_Mtr")
                worksheet2.Cells(X - 1, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X - 1, 12)
                range1.NumberFormat = "0.00"

            End If

            worksheet2.Range("L" & _Fist & ":L" & _Last).MergeCells = True
            worksheet2.Range("L" & (X - 1) & ":L" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

            'FINISHING
            ' worksheet2.Range("M" & _Fist & ":M" & _Last).MergeCells = True
            vcWhere = " M09Meterial='" & Trim(_Material) & "' AND M09Oredr_Type='Finishing'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                worksheet2.Cells(X - 1, 13) = M01.Tables(0).Rows(0)("M09Qty_Mtr")
                worksheet2.Cells(X - 1, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X - 1, 13)
                range1.NumberFormat = "0.00"

            End If
            worksheet2.Range("M" & _Fist & ":M" & _Last).MergeCells = True
            worksheet2.Range("M" & (X - 1) & ":M" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter


            'DYEING
            ' worksheet2.Range("N" & _Fist & ":N" & _Last).MergeCells = True
            vcWhere = " M09Meterial='" & Trim(_Material) & "' AND M09Oredr_Type='Dyeing'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetExcessAnalysis", New SqlParameter("@cQryType", "ZPL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                worksheet2.Cells(X - 1, 14) = M01.Tables(0).Rows(0)("M09Qty_Mtr")
                worksheet2.Cells(X - 1, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet2.Cells(X - 1, 14)
                range1.NumberFormat = "0.00"

            End If
            worksheet2.Range("N" & _Fist & ":N" & _Last).MergeCells = True
            worksheet2.Range("N" & (X - 1) & ":N" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter


            worksheet2.Range("O" & (X - 1)).Formula = "=SUM(G" & _Fist & ":N" & X - 1 & ")"
            worksheet2.Range("O" & _Fist & ":O" & _Last).MergeCells = True
            worksheet2.Range("O" & (X - 1) & ":O" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X - 1, 15)
            range1.NumberFormat = "0.00"


            worksheet2.Cells(X - 1, 16) = "=F" & _Fist & "-O" & _Fist
            worksheet2.Range("P" & _Fist & ":P" & _Last).MergeCells = True
            worksheet2.Range("P" & (X - 1) & ":P" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
            worksheet2.Cells(X - 1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet2.Cells(X - 1, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X - 1, 16)
            range1.NumberFormat = "0.00"

            range1 = worksheet2.Cells(_Fist, 16)

            If range1.Value > 0 Then
                worksheet2.Cells(X - 1, 17) = "Shortage"
                worksheet2.Range("Q" & _Fist & ":Q" & _Last).MergeCells = True
                worksheet2.Range("Q" & (X - 1) & ":Q" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            ElseIf range1.Value = 0 Then
                worksheet2.Cells(X - 1, 17) = "OK"
                worksheet2.Range("Q" & _Fist & ":Q" & _Last).MergeCells = True
                worksheet2.Range("Q" & (X - 1) & ":Q" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            Else
                worksheet2.Cells(X - 1, 17) = "Excess"
                worksheet2.Range("Q" & _Fist & ":Q" & _Last).MergeCells = True
                worksheet2.Range("Q" & (X - 1) & ":Q" & (X - 1)).VerticalAlignment = XlVAlign.xlVAlignCenter
                worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet2.Cells(X - 1, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter
            End If


            worksheet2.Cells(X, 1) = T01.Tables(0).Rows(I)("M07Material")
            worksheet2.Cells(X, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Cells(X, 2) = T01.Tables(0).Rows(I)("M07Met_Dis")
            worksheet2.Cells(X, 3) = T01.Tables(0).Rows(I)("M07Sales_Order")
            worksheet2.Cells(X, 4) = T01.Tables(0).Rows(I)("M07Line_Item")
            worksheet2.Cells(X, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet2.Cells(X, 5) = T01.Tables(0).Rows(I)("M07Qty_Mtr")
            worksheet2.Cells(X, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet2.Cells(X, 5)
            range1.NumberFormat = "0.00"
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

    Private Sub chkStock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkStock.CheckedChanged
        If chkStock.Checked = True Then
            chkReq.Checked = False
            chkWIP.Checked = False
            Call Load_Quality()
            Call Load_Material()

        End If
    End Sub

    Private Sub chkWIP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkWIP.CheckedChanged
        If chkWIP.Checked = True Then
            chkStock.Checked = False
            chkReq.Checked = False
            Call Load_Quality()
            Call Load_Material()

        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call Create_File()
    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click

    End Sub
End Class