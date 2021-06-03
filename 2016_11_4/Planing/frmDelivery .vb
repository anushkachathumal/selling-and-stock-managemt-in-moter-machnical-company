Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO.StreamReader
Imports Microsoft.VisualBasic.FileIO
Imports System.IO

Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports DBLotVbnet.Quarrys
Imports System.IO.File
'Imports System.IO.StreamWriter
Imports System.Net.Mail
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Spire.XlS


Public Class frmDelivery
    Inherits System.Windows.Forms.Form
    Dim Clicked As String

    Dim strLine As String
    Dim strLineflu As String
    Dim strDash As String
    Dim StrDisCode As String
    Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter


    Private Sub frmDelivery_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFromDate.Text = Today
        txtTodate.Text = Today
        lblA.Visible = False
        txtMC.Visible = False
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
        'Call Daily_Boliout()
    End Sub

    Function Upload_ZPL()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim strFileName1 As String

        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim p_Date As Date
        Dim _Convenshion As Double
        Dim T01 As DataSet
        Dim T02 As DataSet
        Dim T03 As DataSet

        Dim _BatchNo As String
        Dim _Customer As String
        Dim _Material As String
        Dim _Dis As String
        Dim _DDate As Date
        Dim _LCDate As Date
        Dim _QtyKG As Double
        Dim _NextOP As String
        Dim _PLCom As String
        Dim _OrderType As String
        Dim _SalesOrder As String
        Dim _LineItem As String
        Dim _QtyMtr As Double
        Dim _Merchant As String


        Dim t_Date As Date
        Dim _WeekNo As Integer
        Dim X11 As Integer
        Dim Y As Integer
        Dim _Status As Boolean


        Try
            'Me.Cursor = Cursors.WaitCursor
            strFileName = ConfigurationManager.AppSettings("FilePath") + "\ZPL_ORDER.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)



                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                If Microsoft.VisualBasic.Left((Trim(fields(0))), 5) = "00000" Then
                    _BatchNo = CInt(Trim(fields(0)))   '0
                Else
                    _BatchNo = (Trim(fields(0)))
                End If
                _Customer = (Trim(fields(1))) '1
                _Material = Trim(fields(2)) '3
                _Dis = Trim(fields(3)) '4
                _DDate = Microsoft.VisualBasic.Left(Trim(fields(4)), 4) & "/" & Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(4)), 6), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(4)), 2)
                If Microsoft.VisualBasic.Left(Trim(fields(5)), 2) = "00" Then
                    _LCDate = "1900/1/1"
                Else
                    _LCDate = Microsoft.VisualBasic.Left(Trim(fields(5)), 4) & "/" & Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(5)), 6), 2) & "/" & Microsoft.VisualBasic.Right(Trim(fields(5)), 2)
                End If
                _QtyKG = Trim(fields(6)) '7
                _NextOP = Trim(fields(7)) '8
                'If Microsoft.VisualBasic.Left(Trim(fields(8)), 6) = "?Don't" Then
                '    _PLCom = "?Dont " & Microsoft.VisualBasic.Right(Trim(fields(8)), Microsoft.VisualBasic.Len(Trim(fields(8))) - 6)
                'Else
                'For Y = 1 To Microsoft.VisualBasic.Len(Trim(fields(8)))
                '    'MsgBox(Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(8)), Y), 1))
                '    If Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Trim(fields(8)), Y), 1) = "'" Then
                '        Exit For
                '    Else
                '        _PLCom = Microsoft.VisualBasic.Left(Trim(fields(8)), Y)
                '    End If
                'Next
                Dim stringToCleanUp As String
                Dim characterToRemove As String

                stringToCleanUp = Trim(fields(8))
                characterToRemove = "'"
                '  _PLCom = Replace(stringToCleanUp, characterToRemove, "")

                ' _PLCom = Microsoft.VisualBasic.Left(Trim(fields(8)), Y - 1) & Microsoft.VisualBasic.Right(Trim(fields(8)), Microsoft.VisualBasic.Len(Trim(fields(8))) - (Y - 1)) '9
                ' End If
                _OrderType = Trim(fields(9)) '10
                _SalesOrder = Trim(fields(10)) '11
                _LineItem = Trim(fields(11)) '11
                _QtyMtr = Trim(fields(12)) '12
                _Merchant = Trim(fields(13)) '13


                nvcFieldList1 = "select * from M16Meterial where Dis='" & CInt(Trim(_Material)) & "'"
                T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(T01) Then

                Else
                    Y = 0
                    _Status = False
                    nvcFieldList1 = "select * from M17Planing_Comment "
                    T02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    For Each DTRow4 As DataRow In T02.Tables(0).Rows
                        ' MsgBox(UCase(Microsoft.VisualBasic.Left(_PLCom, T02.Tables(0).Rows(Y)("M17Lenth"))))
                        If UCase(Microsoft.VisualBasic.Left(_PLCom, T02.Tables(0).Rows(Y)("M17Lenth"))) = Trim(T02.Tables(0).Rows(Y)("M17Dis")) Then
                            _Status = True
                            Exit For
                        Else

                        End If
                        Y = Y + 1
                    Next

                    If _Status = False Then

                        nvcFieldList1 = "select * from M18WIP where M18Batch='" & _BatchNo & "'"
                        T03 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(T03) Then

                        Else
                            _WeekNo = DatePart(DateInterval.WeekOfYear, _DDate)
                            nvcFieldList1 = "Insert Into M18WIP(M18Batch,M18Customer,M18Material,M18Dis,M18DDate,M18LCDate,M18QtyKG,M18NextOparation,M18PComment,M18OrderType,M18SalesOrder,M18LineItem,M18Qty,M18Merchant,M18Week,M18Year)" & _
                                                        " values('" & _BatchNo & "', '" & _Customer & "'," & _Material & ",'" & _Dis & "','" & _DDate & "','" & _LCDate & "','" & _QtyKG & "','" & _NextOP & "','" & _PLCom & "','" & _OrderType & "','" & _SalesOrder & "','" & _LineItem & "','" & _QtyMtr & "','" & _Merchant & "'," & _WeekNo & "," & Year(_DDate) & " )"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        End If

                        nvcFieldList1 = "select * from M19Segrigrade where M19Dis='" & _NextOP & "'"
                        T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        If isValidDataset(T01) Then
                        Else
                            strFileName1 = ConfigurationManager.AppSettings("UploadPath") + "\segrigrade" & Year(Today) & Month(Today) & Microsoft.VisualBasic.Day(Today) & ".txt"
                            FileOpen(1, strFileName1, OpenMode.Append)

                            FileClose(1)

                            Dim TargetFile As StreamWriter
                            TargetFile = New StreamWriter(strFileName, True)
                            TargetFile.Write(_NextOP & vbTab)
                            TargetFile.WriteLine()
                            TargetFile.Close()
                        End If
                    End If
                End If



                _BatchNo = ""
                _Customer = ""
                _Material = ""
                _Dis = ""

                _QtyKG = 0
                _NextOP = ""
                _PLCom = ""
                _OrderType = ""
                _SalesOrder = ""
                _LineItem = ""
                _QtyMtr = 0
                _Merchant = ""

                X11 = X11 + 1
                ' pbCount.Value = pbCount.Value + 1

            Next
            '  MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            Kill(strFileName)
            Call Upload_NC()

        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            ' MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")

        Catch ex As Exception
            ' MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            '  MsgBox("Error Record in txt File-Line -" & X11, MsgBoxStyle.Critical, "Error Upload File/" & strFileName)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function
    Function Upload_NC()
        Dim sr As System.IO.StreamReader
        Dim strFileName As String
        Dim MTNo As String
        Dim _Plant As String
        Dim _StockLoc As String
        Dim PONo As String
        Dim _Sold As String
        Dim _Sales_Doc As String
        Dim _Dis As String
        Dim _Item As String
        Dim _Meterial As String
        Dim _Create As Date
        Dim _Dilivary As String
        Dim _Qty As Double
        Dim _CreateBy As String
        Dim _Su As String
        Dim _Reject As String
        Dim _Db As String
        Dim _DS As String

        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim p_Date As Date
        Dim _Convenshion As Double
        Dim T01 As DataSet

        Dim t_Date As Date
        Dim _WeekNo As Integer
        Dim X11 As Integer
        Dim dsUser As DataSet

        Try
            Me.Cursor = Cursors.WaitCursor


            strFileName = ConfigurationManager.AppSettings("FilePath") + "\NC.txt"
            For Each line As String In System.IO.File.ReadAllLines(strFileName)
                Dim fields() As String = line.Split(vbTab)

                ' Declare a variable named currentField of type String.
                Dim currentField As String
                Dim i As Integer
                ' Use the currentField variable to loop
                ' through fields in the currentRow.



                nvcFieldList1 = "select * from M18WIP where M18Batch='" & Trim(fields(0)) & "' "
                dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(dsUser) Then
                    nvcFieldList1 = "update M18WIP set M18NonCon='" & Trim(fields(1)) & "',M18NonConDis='" & Trim(fields(2)) & "' where M18Batch='" & Trim(fields(0)) & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else

                End If



                X11 = X11 + 1
            Next

            ' MsgBox("File updated successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            Kill(strFileName)
            Me.Cursor = Cursors.Arrow
        Catch malFormLineEx As Microsoft.VisualBasic.FileIO.MalformedLineException
            ' MessageBox.Show("Line " & malFormLineEx.Message & "is not valid and will be skipped.", "Malformed Line Exception")
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
        Catch ex As Exception
            ' MessageBox.Show(ex.Message & " exception has occurred.", "Exception")
            ' MsgBox("Error Record in txt File-Line -" & X11, MsgBoxStyle.Critical, "Error Upload File/" & strFileName)
        Finally
            ' If successful or if an exception is thrown,
            ' close the TextFieldParser.
            ' theTextFieldParser.Close()
        End Try
    End Function
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        OPR0.Enabled = True
        OPR1.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False

        txtFromDate.Text = Today
        txtTodate.Text = Today

        cmdSave.Enabled = True
        '  Call Upload_ZPL()
    End Sub

    Function Create_Report()
        Dim exc As New Application

        Dim workbooks As Workbooks = exc.Workbooks
        Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
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
        Dim _weekNo As Integer
        Dim _DayOfweek As Integer
        Dim K As Integer
        Dim L As Integer
        Dim K1 As Integer
        Dim L1 As Integer
        Dim K2 As Integer
        Dim L2 As Integer
        Dim K3 As Integer
        Dim L3 As Integer
        Dim K4 As Integer
        Dim L4 As Integer
        Dim KK As Integer
        Dim K5 As Integer
        Dim L5 As Integer
        Dim K6 As Integer
        Dim L6 As Integer
        Dim K7 As Integer
        Dim L7 As Integer
        Dim k8 As Integer
        Dim L8 As Integer
        Dim K9 As Integer
        Dim L9 As Integer
        Dim L10 As Integer
        Dim K10 As Integer
        Dim KK1 As Integer
        Dim L11 As Integer

        Dim K11 As Integer
        Dim K12 As Integer
        Dim L12 As Integer
        Dim L13 As Integer
        Dim K13 As Integer
        Dim L14 As Integer
        Dim K14 As Integer
        Dim L15 As Integer
        Dim K15 As Integer

        Try
            '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
            '  Workbooks.Application.Sheets.Add()
            Dim sheets1 As Sheets = Workbook.Worksheets
            '   Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)

            n_Date = txtFromDate.Text
            _DayOfweek = (n_Date.DayOfWeek)
            If _DayOfweek = 1 Then
            Else
                txtFromDate.Text = CDate(n_Date).AddDays(-_DayOfweek)
            End If

            N_Date1 = CDate(txtFromDate.Text).AddDays(+8)
            _weekNo = DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text))
            worksheet1.Name = "Week" & _weekNo & "Delivery Batches"


            worksheet1.Rows(5).Font.size = 18
            worksheet1.Rows(5).Font.bold = True
            worksheet1.Cells(5, 1) = " Week " & _weekNo & " Delivery batches in urgent list as at " & txtFromDate.Text
            worksheet1.Columns("A").ColumnWidth = 15

            worksheet1.Rows(7).Font.size = 10
            worksheet1.Rows(7).Font.bold = True
            worksheet1.Columns("A").ColumnWidth = 15
            worksheet1.Columns("B").ColumnWidth = 15
            worksheet1.Columns("C").ColumnWidth = 10
            worksheet1.Columns("D").ColumnWidth = 15
            worksheet1.Columns("E").ColumnWidth = 10
            worksheet1.Columns("F").ColumnWidth = 20
            worksheet1.Columns("G").ColumnWidth = 10
            worksheet1.Columns("H").ColumnWidth = 10
            worksheet1.Columns("I").ColumnWidth = 10
            worksheet1.Columns("J").ColumnWidth = 10
            worksheet1.Columns("K").ColumnWidth = 10
            worksheet1.Columns("L").ColumnWidth = 10
            worksheet1.Columns("M").ColumnWidth = 25
            worksheet1.Columns("N").ColumnWidth = 45
            worksheet1.Columns("O").ColumnWidth = 20
            worksheet1.Columns("P").ColumnWidth = 10


            worksheet1.Cells(7, 1) = " Order No "
            worksheet1.Cells(7, 1).WrapText = True
            worksheet1.Cells(7, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 2) = "  Sales Order No"

            worksheet1.Cells(7, 2).WrapText = True
            worksheet1.Cells(7, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 2).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 3) = " Line Item No"

            worksheet1.Cells(7, 3).WrapText = True
            worksheet1.Cells(7, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 3).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 4) = " Customer"

            worksheet1.Cells(7, 4).WrapText = True
            worksheet1.Cells(7, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 4).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 5) = " Meterial"

            worksheet1.Cells(7, 5).WrapText = True
            worksheet1.Cells(7, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 5).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 6) = " Quality"

            worksheet1.Cells(7, 6).WrapText = True
            worksheet1.Cells(7, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 6).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 7) = "No:of days  delay for Delivery"

            worksheet1.Cells(7, 7).WrapText = True
            worksheet1.Cells(7, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 7).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 8) = "D date"

            worksheet1.Cells(7, 8).WrapText = True
            worksheet1.Cells(7, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 8).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 9) = "Last cnf date"

            worksheet1.Cells(7, 9).WrapText = True
            worksheet1.Cells(7, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 9).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 10) = "No:of days in same opparation"

            worksheet1.Cells(7, 10).WrapText = True
            worksheet1.Cells(7, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 10).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 11) = "Qty(kg)"

            worksheet1.Cells(7, 11).WrapText = True
            worksheet1.Cells(7, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 11).VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet1.Cells(7, 12) = "Qty (M)"

            worksheet1.Cells(7, 12).WrapText = True
            worksheet1.Cells(7, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 12).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 13) = "Status"

            worksheet1.Cells(7, 13).WrapText = True
            worksheet1.Cells(7, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 13).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 14) = "Planning Comments"

            worksheet1.Cells(7, 14).WrapText = True
            worksheet1.Cells(7, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 14).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 15) = "N/C Comments"
            worksheet1.Columns("O").ColumnWidth = 20
            worksheet1.Cells(7, 15).WrapText = True
            worksheet1.Cells(7, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 15).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 16) = "Order Type"
            worksheet1.Columns("P").ColumnWidth = 10
            worksheet1.Cells(7, 16).WrapText = True
            worksheet1.Cells(7, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 16).VerticalAlignment = XlVAlign.xlVAlignCenter
            '------------------------------------------------------------------------------------------------------
            worksheet1.Range("A7", "a7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b7", "b7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c7", "c7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d7", "d7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e7", "e7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f7", "f7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g7", "g7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h7", "h7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i7", "i7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j7", "j7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k7", "k7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l7", "l7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m7", "m7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n7", "n7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o7", "o7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p7", "p7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("p7", "p7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p7", "p7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o7", "o7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o7", "o7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n7", "n7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n7", "n7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m7", "m7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m7", "m7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l7", "l7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l7", "l7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k7", "k7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k7", "k7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j7", "j7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j7", "j7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i7", "i7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i7", "i7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h7", "h7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h7", "h7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g7", "g7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g7", "g7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f7", "f7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f7", "f7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e7", "e7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e7", "e7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d7", "d7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d7", "d7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c7", "c7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c7", "c7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b7", "b7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b7", "b7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("a7", "a7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("a7", "a7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            '--------------------------------------------------------------------------------------------------------------------
            worksheet1.Range("A7:a7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("b7:b7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("c7:c7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("d7:d7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("e7:e7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("f7:f7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("g7:g7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("h7:h7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("i7:i7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("j7:j7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("k7:k7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("l7:l7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("m7:m7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("n7:n7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("o7:o7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("p7:p7").Interior.Color = RGB(141, 180, 227)
            '----------------------------------------
            worksheet1.Range("p8", "p8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p8", "p8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o8", "o8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o8", "o8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n8", "n8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n8", "n8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m8", "m8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m8", "m8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l8", "l8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l8", "l8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k8", "k8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k8", "k8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j8", "j8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j8", "j8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i8", "i8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i8", "i8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h8", "h8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h8", "h8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g8", "g8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g8", "g8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f8", "f8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f8", "f8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e8", "e8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e8", "e8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d8", "d8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d8", "d8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c8", "c8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c8", "c8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b8", "b8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b8", "b8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a8", "a8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a8", "a8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Y = 9

            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1) = "Reprocess - Dyeing"
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)


            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            '------------------------------------------------------------------------------------------------------
            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation  where M18NonCon in ('Bulk Batches to be Reshedule','Aw OD recipe','Sample Batches to be Reshedule','Down Grade','Need to strip','Striped') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('dyeing','Finishng') and M18Status='Y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next
            '-----------------------------------------------------------------------------------------------------------
            K = Y
            L = Y
            worksheet1.Cells(Y, 8) = "Reprocess - Dyeing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k9:k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l9:l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '---------------------------------------------------------------------------------------------------------
            Y = Y + 1
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1) = "Shortage"
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-----------------------------------------------------------------------------------------------------------
            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where left(M18PComment,4)='?Sho' and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='dyeing' and M18Status='y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Shortage"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L1 = Y
            K1 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            '------------------------------------------------------------------------------------------------------
            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '------------------------------------------------------------------------------------------------------------
            'REPLACEMENT

            Y = Y + 1
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1) = "Replacement"
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-----------------------------------------------------------------------------------------------------------
            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where left(M18PComment,4)='?REP' and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='dyeing' and M18Status='Y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Replacement"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L2 = Y
            K2 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K1 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L1 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            '--------------------------------------------------------------------------------------------------------------
            'DYEING
            Y = Y + 1
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1) = "Dyeing"
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where left(M18PComment,4) not in ('?Sho','?Ref') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='dyeing' and M18NonCon IS NULL  and M18Status='Y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Dyeing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L3 = Y
            K3 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K2 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L2 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            '------------------------------------------------------------------------------------------------------
            Y = Y + 1
            KK = Y
            worksheet1.Cells(Y, 8) = "Total Dyeing "
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K & "+k" & K1 & "+k" & K2 & "+k" & K3 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(148, 139, 84)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L & "+l" & L1 & "+l" & L2 & "+l" & L3 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(148, 139, 84)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '----------------------------------------------------------------------------------------------------------

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '-------------------------------------------------------------------------------------------------------------
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "Awaiting Greige"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '----------------------------------------------------------------------------------------------------------
            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('Aw Greige') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='dyeing' and M18Status='Y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Awaiting Greige"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L4 = Y
            K4 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & KK + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & KK + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '------------------------------------------------------------------------------------------------------

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            '--------------------------------------------------------------------------------------------------------------
            'N/C Held & waiting for App
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "N/C Held & waiting for App"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('Aw Pilot','Aw Shade comments','Need to  finish  pending App','Finished pending App','Aw cus App','Aw N/C App','Sub as ongoing','Aw Cus care comments') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='Finishng' and M18Status='Y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next
            ' Y = Y + 1
            worksheet1.Cells(Y, 8) = "N/C Held & waiting for App"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L5 = Y
            K5 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & KK + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & KK + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '--------------------------------------------------------------------------------------------------------------
            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            'N/C Held Due to Dyeing Issue
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "N/C Held Due to Dyeing Issue"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-------------------------------------------------------------------------------------------------------

            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('Off shade','Aw Pigment','held due to dyeing issues','Aw PAD UV','OFF shade SA') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng','dyeing') and M18Status='Y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next
            worksheet1.Cells(Y, 8) = "N/C Held Due to Dyeing Issue"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L6 = Y
            K6 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K5 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L5 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '--------------------------------------------------------------------------------------------------------
            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            'N/C held due to  Finishing Issues
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "N/C held due to  Finishing Issues"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('held due to  finishing issue') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M18Status='Y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next
            '-----------------------------------------------------------------------------------------------------------
            worksheet1.Cells(Y, 8) = "N/C held due to  Finishing Issues"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L7 = Y
            K7 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K6 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L6 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '--------------------------------------------------------------------------------------------------------------------
            'N/C held due to  Knitting Issues
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "N/C held due to  Knitting Issues"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('held due to  Knitting issue') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M18Status='y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next
            '--------------------------------------------------------------------------------------------------------------
            worksheet1.Cells(Y, 8) = "N/C held due to  Knitting Issues"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L8 = Y
            k8 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K7 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L7 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '------------------------------------------------------------------------------------------------------------
            'Awaiting Printing
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "Awaiting Printing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            'worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType='ZP22' and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng') and M18Status='Y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Awaiting Printing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L9 = Y
            K9 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & k8 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L8 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            '------------------------------------------------------------------------------------------------------------
            'Awaiting Finishing
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "Awaiting Finishing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            'worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            ' SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType not in ('ZP22') and M18NonCon not in ('held due to  finishing issue','held due to  Knitting issue','Off shade','Aw Pigment','held due to  dyeing issues','Aw PAD UV','OFF shade SA','Aw Pilot','Aw Shade comments','Need to  finish  pending App','Finished pending App','Aw cus App','Aw N/C App','Sub as ongoing','Aw Cus care comments','Bulk Batches to be Reshedule','Aw OD recipe','Sample Batches to be Reshedule') and left(M18PComment,4) not in ('?Sho','?Rep') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng')"
            'SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType not in ('ZP22') and  left(M18PComment,4)  in ('?Sho','?Rep') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng') and M18NonCon IS NULL"
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType not in ('ZP22') and   M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng') and M18NonCon IS NULL and M18Status='y' and left(M18Material,2) in ('30','26')"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Awaiting Finishing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L10 = Y
            K10 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K9 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L9 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            KK1 = Y

            worksheet1.Cells(Y, 8) = "Total Finishing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K5 & "+k" & K6 & "+k" & K7 & "+k" & k8 & "+k" & K9 & "+k" & K10 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(148, 139, 84)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L & "+l" & L1 & "+l" & L2 & "+l" & L3 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(148, 139, 84)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-----------------------------------------------------------------------------------------------------------

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            '--------------------------------------------------------------------------------------------------------
            'EXAM

            Y = Y + 1
            worksheet1.Cells(Y, 1) = "Exam"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            'worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-------------------------------------------------------------------------------------------------------
            Y = Y + 1
            ' SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType not in ('ZP22') and M18NonCon not in ('held due to  finishing issue','held due to  Knitting issue','Off shade','Aw Pigment','held due to  dyeing issues','Aw PAD UV','OFF shade SA','Aw Pilot','Aw Shade comments','Need to  finish  pending App','Finished pending App','Aw cus App','Aw N/C App','Sub as ongoing','Aw Cus care comments','Bulk Batches to be Reshedule','Aw OD recipe','Sample Batches to be Reshedule') and left(M18PComment,4) not in ('?Sho','?Rep') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng')"
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where  M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Exam & Lab') and M18Status='Y' and left(M18Material,2) in ('30','26') "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows

                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("M18Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            '  Dim L11 As Integer
            ' Dim K11 As Integer

            worksheet1.Cells(Y, 8) = "Exam"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L11 = Y
            K11 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & KK1 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & KK1 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            ' KK1 = Y
            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            Y = Y + 1
            'BLOCK STOCK PRINT - 2055 Developed by suranga 11/12/2013

            worksheet1.Cells(Y, 1) = "2055"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-------------------------------------------------------------------------------------------------------
            Y = Y + 1

            SQL = "select * from  BLOCK_STOCK  where  Dilivary_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Status='Y' and Stock_Loc='2055' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows

                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("Sales_order")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("Mat_Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("Dilivary_Date")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("Dilivary_Date")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("GRN_date") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("GRN_date")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("GRN_date") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("GRN_date")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("Qty_Kg")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("Qty_Mtr")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    'worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    'worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    'worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    'worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    'worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            ' Y = Y + 1


            worksheet1.Cells(Y, 8) = "2055"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L12 = Y
            K12 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K11 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L11 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            Y = Y + 1

            'BLOCK STOCK PRINT - 2065 Developed by suranga 11/12/2013

            worksheet1.Cells(Y, 1) = "2062"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-------------------------------------------------------------------------------------------------------
            Y = Y + 1

            SQL = "select * from  BLOCK_STOCK  where  Dilivary_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Status='Y' and Stock_Loc='2062' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows

                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("Sales_order")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("Mat_Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("Dilivary_Date")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("Dilivary_Date")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("GRN_date") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("GRN_date")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("GRN_date") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("GRN_date")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("Qty_Kg")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("Qty_Mtr")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    'worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    'worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    'worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    'worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    'worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            Y = Y + 1


            worksheet1.Cells(Y, 8) = "2062"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L13 = Y
            K13 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K12 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L12 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            Y = Y + 1
            'BLOCK STOCK PRINT - 2065 Developed by suranga 11/12/2013

            worksheet1.Cells(Y, 1) = "2065"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-------------------------------------------------------------------------------------------------------
            Y = Y + 1

            SQL = "select * from  BLOCK_STOCK  where  Dilivary_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Status='Y' and Stock_Loc='2065' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows

                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("Sales_order")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("Mat_Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("Dilivary_Date")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("Dilivary_Date")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("GRN_date") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("GRN_date")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("GRN_date") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("GRN_date")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("Qty_Kg")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("Qty_Mtr")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    'worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    'worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    'worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    'worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    'worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            '   Y = Y + 1


            worksheet1.Cells(Y, 8) = "2065"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L14 = Y
            K14 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K13 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L13 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(Y, 11)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            Y = Y + 1

            'BLOCK STOCK PRINT - 2070 Developed by suranga 11/12/2013

            worksheet1.Cells(Y, 1) = "2070"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-------------------------------------------------------------------------------------------------------
            Y = Y + 1

            SQL = "select * from  BLOCK_STOCK  where  Dilivary_Date between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "'  and Status='Y' and Stock_Loc='2070' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows

                SQL = "select * from M16Meterial where Dis='" & Trim(dsUser.Tables(0).Rows(i)("Material")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    worksheet1.Rows(Y).Font.size = 10

                    worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("Batch")
                    worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("Sales_order")
                    worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("LineItem")
                    worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                    worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("Customer")
                    worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("Material")
                    worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("Mat_Dis")
                    worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                    worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("Dilivary_Date")), Today)
                    worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("Dilivary_Date")
                    worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    If dsUser.Tables(0).Rows(i)("GRN_date") = "1900/1/1" Then
                        worksheet1.Cells(Y, 9) = "00/00/0000"
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Else
                        worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("GRN_date")
                        worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    If dsUser.Tables(0).Rows(i)("GRN_date") = "1900/1/1" Then

                    Else
                        worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("GRN_date")), Today)
                        worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                    End If

                    worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("Qty_Kg")
                    worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 11)
                    range1.NumberFormat = "0"

                    worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("Qty_Mtr")
                    worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(Y, 12)
                    range1.NumberFormat = "0"

                    'worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                    'worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                    'worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonConDis")
                    'worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                    'worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                    'worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                    worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                    Y = Y + 1
                End If
                i = i + 1
            Next

            '   Y = Y + 1


            worksheet1.Cells(Y, 8) = "2070"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L15 = Y
            K15 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K14 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L14 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(Y, 11)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            Y = Y + 1

            worksheet1.Cells(Y, 8) = "Total Qty"
            worksheet1.Rows(Y).Font.size = 12
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & KK & "+k" & KK1 & "+k" & K11 & "+k" & K12 & "+k" & K13 & "+k" & K14 & "+k" & K15 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(217, 151, 149)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & KK & "+l" & KK1 & "+l" & K11 & "+l" & K12 & "+l" & K13 & "+l" & K14 & "+l" & K15 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(217, 151, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If chk1.Checked = False And chk2.Checked = False And chk3.Checked = False Then
            Call Create_Report()
        ElseIf chk1.Checked = True Then
            ' Call Create_Report_Customer(txtMC.Text)

        End If
    End Sub

    Private Sub chk1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk1.CheckedChanged
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try




            If chk1.Checked = True Then
                chk2.Checked = False
                chk3.Checked = False

                lblA.Visible = True
                txtMC.Visible = True
                lblA.Text = "Customer Name"

                txtMC.Text = ""
                Sql = "select M18customer as [Customer Name] from M18WIP where M18customer <>'' group by M18customer  "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    With txtMC
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 475
                    End With
                End If
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub chk2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk2.CheckedChanged
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try




            If chk2.Checked = True Then
                chk1.Checked = False
                chk3.Checked = False

                lblA.Visible = True
                txtMC.Visible = True
                lblA.Text = "Merchant Name"
                txtMC.Text = ""
                Sql = "select M18Merchant from M18WIP where M18Merchant <>'' group by M18Merchant "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    With txtMC
                        .DataSource = M01
                        .Rows.Band.Columns(0).Width = 475
                    End With
                End If
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub chk3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk3.CheckedChanged
        If chk3.Checked = True Then
            chk2.Checked = False
            chk1.Checked = False
        End If
    End Sub

    Function Create_Report_Customer(ByVal strCusname As String)
        Dim exc As New Application

        Dim workbooks As Workbooks = exc.Workbooks
        Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)

        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
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
        Dim _weekNo As Integer
        Dim _DayOfweek As Integer
        Dim K As Integer
        Dim L As Integer
        Dim K1 As Integer
        Dim L1 As Integer
        Dim K2 As Integer
        Dim L2 As Integer
        Dim K3 As Integer
        Dim L3 As Integer
        Dim K4 As Integer
        Dim L4 As Integer
        Dim KK As Integer
        Dim K5 As Integer
        Dim L5 As Integer
        Dim K6 As Integer
        Dim L6 As Integer
        Dim K7 As Integer
        Dim L7 As Integer
        Dim k8 As Integer
        Dim L8 As Integer
        Dim K9 As Integer
        Dim L9 As Integer
        Dim L10 As Integer
        Dim K10 As Integer
        Dim KK1 As Integer

        Try
            '  Dim worksheet11 As _worksheet1 = CType(sheets.Item(2), _worksheet1)
            Workbooks.Application.Sheets.Add()
            Dim sheets1 As Sheets = Workbook.Worksheets
            '  Dim worksheet1 As _Worksheet = CType(sheets1.Item(1), _Worksheet)

            n_Date = txtFromDate.Text
            _DayOfweek = (n_Date.DayOfWeek)
            If _DayOfweek = 1 Then
            Else
                txtFromDate.Text = CDate(n_Date).AddDays(-_DayOfweek)
            End If

            N_Date1 = CDate(txtFromDate.Text).AddDays(+8)
            _weekNo = DatePart(DateInterval.WeekOfYear, CDate(txtFromDate.Text))
            worksheet1.Name = "Week" & _weekNo & "Delivery Batches"


            worksheet1.Rows(5).Font.size = 18
            worksheet1.Rows(5).Font.bold = True
            worksheet1.Cells(5, 1) = " Week " & _weekNo & " Delivery batches in urgent list as at " & txtFromDate.Text
            worksheet1.Columns("A").ColumnWidth = 15

            worksheet1.Rows(7).Font.size = 10
            worksheet1.Rows(7).Font.bold = True
            worksheet1.Columns("A").ColumnWidth = 15
            worksheet1.Columns("B").ColumnWidth = 15
            worksheet1.Columns("C").ColumnWidth = 10
            worksheet1.Columns("D").ColumnWidth = 15
            worksheet1.Columns("E").ColumnWidth = 10
            worksheet1.Columns("F").ColumnWidth = 20
            worksheet1.Columns("G").ColumnWidth = 10
            worksheet1.Columns("H").ColumnWidth = 10
            worksheet1.Columns("I").ColumnWidth = 10
            worksheet1.Columns("J").ColumnWidth = 10
            worksheet1.Columns("K").ColumnWidth = 10
            worksheet1.Columns("L").ColumnWidth = 10
            worksheet1.Columns("M").ColumnWidth = 25
            worksheet1.Columns("N").ColumnWidth = 45
            worksheet1.Columns("O").ColumnWidth = 20
            worksheet1.Columns("P").ColumnWidth = 10


            worksheet1.Cells(7, 1) = " Order No "
            worksheet1.Cells(7, 1).WrapText = True
            worksheet1.Cells(7, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 2) = "  Sales Order No"

            worksheet1.Cells(7, 2).WrapText = True
            worksheet1.Cells(7, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 2).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 3) = " Line Item No"

            worksheet1.Cells(7, 3).WrapText = True
            worksheet1.Cells(7, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 3).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 4) = " Customer"

            worksheet1.Cells(7, 4).WrapText = True
            worksheet1.Cells(7, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 4).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 5) = " Meterial"

            worksheet1.Cells(7, 5).WrapText = True
            worksheet1.Cells(7, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 5).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 6) = " Quality"

            worksheet1.Cells(7, 6).WrapText = True
            worksheet1.Cells(7, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 6).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 7) = "No:of days  delay for Delivery"

            worksheet1.Cells(7, 7).WrapText = True
            worksheet1.Cells(7, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 7).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 8) = "D date"

            worksheet1.Cells(7, 8).WrapText = True
            worksheet1.Cells(7, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 8).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 9) = "Last cnf date"

            worksheet1.Cells(7, 9).WrapText = True
            worksheet1.Cells(7, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 9).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 10) = "No:of days in same opparation"

            worksheet1.Cells(7, 10).WrapText = True
            worksheet1.Cells(7, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 10).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 11) = "Qty(kg)"

            worksheet1.Cells(7, 11).WrapText = True
            worksheet1.Cells(7, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 11).VerticalAlignment = XlVAlign.xlVAlignCenter


            worksheet1.Cells(7, 12) = "Qty (M)"

            worksheet1.Cells(7, 12).WrapText = True
            worksheet1.Cells(7, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 12).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 13) = "Status"

            worksheet1.Cells(7, 13).WrapText = True
            worksheet1.Cells(7, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 13).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 14) = "Planning Comments"

            worksheet1.Cells(7, 14).WrapText = True
            worksheet1.Cells(7, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 14).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 15) = "N/C Comments"
            worksheet1.Columns("O").ColumnWidth = 20
            worksheet1.Cells(7, 15).WrapText = True
            worksheet1.Cells(7, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 15).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Cells(7, 16) = "Order Type"
            worksheet1.Columns("P").ColumnWidth = 10
            worksheet1.Cells(7, 16).WrapText = True
            worksheet1.Cells(7, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(7, 16).VerticalAlignment = XlVAlign.xlVAlignCenter
            '------------------------------------------------------------------------------------------------------
            worksheet1.Range("A7", "a7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b7", "b7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c7", "c7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d7", "d7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e7", "e7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f7", "f7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g7", "g7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h7", "h7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i7", "i7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j7", "j7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k7", "k7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l7", "l7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m7", "m7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n7", "n7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o7", "o7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p7", "p7").Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous

            worksheet1.Range("p7", "p7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("p7", "p7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o7", "o7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("o7", "o7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n7", "n7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("n7", "n7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m7", "m7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("m7", "m7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l7", "l7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("l7", "l7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k7", "k7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("k7", "k7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j7", "j7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("j7", "j7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i7", "i7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("i7", "i7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h7", "h7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("h7", "h7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g7", "g7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("g7", "g7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f7", "f7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("f7", "f7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e7", "e7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("e7", "e7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d7", "d7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("d7", "d7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c7", "c7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("c7", "c7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b7", "b7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("b7", "b7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("a7", "a7").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            worksheet1.Range("a7", "a7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            '--------------------------------------------------------------------------------------------------------------------
            worksheet1.Range("A7:a7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("b7:b7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("c7:c7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("d7:d7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("e7:e7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("f7:f7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("g7:g7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("h7:h7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("i7:i7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("j7:j7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("k7:k7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("l7:l7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("m7:m7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("n7:n7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("o7:o7").Interior.Color = RGB(141, 180, 227)
            worksheet1.Range("p7:p7").Interior.Color = RGB(141, 180, 227)
            '----------------------------------------
            worksheet1.Range("p8", "p8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p8", "p8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o8", "o8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o8", "o8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n8", "n8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n8", "n8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m8", "m8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m8", "m8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l8", "l8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l8", "l8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k8", "k8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k8", "k8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j8", "j8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j8", "j8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i8", "i8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i8", "i8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h8", "h8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h8", "h8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g8", "g8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g8", "g8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f8", "f8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f8", "f8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e8", "e8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e8", "e8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d8", "d8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d8", "d8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c8", "c8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c8", "c8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b8", "b8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b8", "b8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a8", "a8").Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a8", "a8").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Y = 9

            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1) = "Reprocess - Dyeing"
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)


            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            '------------------------------------------------------------------------------------------------------
            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation  where M18NonCon in ('Bulk Batches to be Reshedule','Aw OD recipe','Sample Batches to be Reshedule','Down Grade') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group ='dyeing' and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next
            '-----------------------------------------------------------------------------------------------------------
            K = Y
            L = Y
            worksheet1.Cells(Y, 8) = "Reprocess - Dyeing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k9:k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l9:l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '---------------------------------------------------------------------------------------------------------
            Y = Y + 1
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1) = "Shortage"
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-----------------------------------------------------------------------------------------------------------
            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where left(M18PComment,4)='?Sho' and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='dyeing' and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Shortage"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L1 = Y
            K1 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            '------------------------------------------------------------------------------------------------------
            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '------------------------------------------------------------------------------------------------------------
            'REPLACEMENT

            Y = Y + 1
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1) = "Replacement"
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-----------------------------------------------------------------------------------------------------------
            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where left(M18PComment,4)='?REP' and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='dyeing' and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Replacement"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L2 = Y
            K2 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K1 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L1 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            '--------------------------------------------------------------------------------------------------------------
            'DYEING
            Y = Y + 1
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1) = "Dyeing"
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where  M18NonCon not in ('Bulk Batches to be Reshedule','Aw OD recipe','Sample Batches to be Reshedule','Down Grade','Aw Pilot','Aw Shade comments','Need to  finish  pending App','Finished pending App','Aw cus App','Aw N/C App','Sub as ongoing','Aw Cus care comments','Off shade','Aw Pigment','held due to  dyeing issues','Aw PAD UV','OFF shade SA','held due to  finishing issue','held due to  Knitting issue') and left(M18PComment,4) not in ('?Sho','?Ref') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='dyeing' and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Dyeing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L3 = Y
            K3 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K2 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L2 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            '------------------------------------------------------------------------------------------------------
            Y = Y + 1
            KK = Y
            worksheet1.Cells(Y, 8) = "Total Dyeing "
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K & "+k" & K1 & "+k" & K2 & "+k" & K3 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(148, 139, 84)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L & "+l" & L1 & "+l" & L2 & "+l" & L3 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(148, 139, 84)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '----------------------------------------------------------------------------------------------------------

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '-------------------------------------------------------------------------------------------------------------
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "Awaiting Greige"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '----------------------------------------------------------------------------------------------------------
            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('Aw Greige') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='dyeing' and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Awaiting Greige"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L4 = Y
            K4 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & KK + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & KK + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '------------------------------------------------------------------------------------------------------

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            '--------------------------------------------------------------------------------------------------------------
            'N/C Held & waiting for App
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "N/C Held & waiting for App"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('Aw Pilot','Aw Shade comments','Need to  finish  pending App','Finished pending App','Aw cus App','Aw N/C App','Sub as ongoing','Aw Cus care comments') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group='Finishng' and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next
            ' Y = Y + 1
            worksheet1.Cells(Y, 8) = "N/C Held & waiting for App"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L5 = Y
            K5 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & KK + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & KK + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '--------------------------------------------------------------------------------------------------------------
            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            'N/C Held Due to Dyeing Issue
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "N/C Held Due to Dyeing Issue"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-------------------------------------------------------------------------------------------------------

            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('Off shade','Aw Pigment','held due to  dyeing issues','Aw PAD UV','OFF shade SA') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng','dyeing') and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next
            worksheet1.Cells(Y, 8) = "N/C Held Due to Dyeing Issue"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L6 = Y
            K6 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K5 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L5 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '--------------------------------------------------------------------------------------------------------
            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            'N/C held due to  Finishing Issues
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "N/C held due to  Finishing Issues"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('held due to  finishing issue') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M18Customer='" & strCusname & "' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next
            '-----------------------------------------------------------------------------------------------------------
            worksheet1.Cells(Y, 8) = "N/C held due to  Finishing Issues"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L7 = Y
            K7 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K6 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L6 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '--------------------------------------------------------------------------------------------------------------------
            'N/C held due to  Knitting Issues
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "N/C held due to  Knitting Issues"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Cells(Y, 1).WrapText = True
            worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18NonCon in ('held due to  Knitting issue') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next
            '--------------------------------------------------------------------------------------------------------------
            worksheet1.Cells(Y, 8) = "N/C held due to  Knitting Issues"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L8 = Y
            k8 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K7 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L7 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next
            '------------------------------------------------------------------------------------------------------------
            'Awaiting Printing
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "Awaiting Printing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            'worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType='ZP22' and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng') and M18Customer='" & strCusname & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Awaiting Printing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L9 = Y
            K9 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & k8 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L8 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            '------------------------------------------------------------------------------------------------------------
            'Awaiting Finishing
            Y = Y + 1
            worksheet1.Cells(Y, 1) = "Awaiting Finishing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            'worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            ' SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType not in ('ZP22') and M18NonCon not in ('held due to  finishing issue','held due to  Knitting issue','Off shade','Aw Pigment','held due to  dyeing issues','Aw PAD UV','OFF shade SA','Aw Pilot','Aw Shade comments','Need to  finish  pending App','Finished pending App','Aw cus App','Aw N/C App','Sub as ongoing','Aw Cus care comments','Bulk Batches to be Reshedule','Aw OD recipe','Sample Batches to be Reshedule') and left(M18PComment,4) not in ('?Sho','?Rep') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng')"
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType not in ('ZP22') and  left(M18PComment,4) not in ('?Sho','?Rep') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng') and M18Customer='" & strCusname & "' and M18NonCon IS NULL"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next

            worksheet1.Cells(Y, 8) = "Awaiting Finishing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L10 = Y
            K10 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K9 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L9 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            KK1 = Y

            worksheet1.Cells(Y, 8) = "Total Finishing"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & K5 & "+k" & K6 & "+k" & K7 & "+k" & k8 & "+k" & K9 & "+k" & K10 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(148, 139, 84)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & L & "+l" & L1 & "+l" & L2 & "+l" & L3 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(148, 139, 84)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-----------------------------------------------------------------------------------------------------------

            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            '--------------------------------------------------------------------------------------------------------
            'EXAM

            Y = Y + 1
            worksheet1.Cells(Y, 1) = "Exam"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            ' worksheet1.Cells(Y, 1).WrapText = True
            'worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
            'worksheet1.Cells(Y, 1).VerticalAlignment = XlVAlign.xlVAlignCenter

            worksheet1.Range("A" & Y & ":a" & Y).Interior.Color = RGB(14, 216, 149)
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            '-------------------------------------------------------------------------------------------------------
            Y = Y + 1
            ' SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where M18OrderType not in ('ZP22') and M18NonCon not in ('held due to  finishing issue','held due to  Knitting issue','Off shade','Aw Pigment','held due to  dyeing issues','Aw PAD UV','OFF shade SA','Aw Pilot','Aw Shade comments','Need to  finish  pending App','Finished pending App','Aw cus App','Aw N/C App','Sub as ongoing','Aw Cus care comments','Bulk Batches to be Reshedule','Aw OD recipe','Sample Batches to be Reshedule') and left(M18PComment,4) not in ('?Sho','?Rep') and M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Finishng')"
            SQL = "select * from  M18WIP inner join M19Segrigrade on M19Dis=M18NextOparation where  M18DDate between '" & txtFromDate.Text & "' and '" & txtTodate.Text & "' and M19Group in ('Exam & Lab') and M18Customer='" & strCusname & "' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow4 As DataRow In dsUser.Tables(0).Rows
                worksheet1.Rows(Y).Font.size = 10

                worksheet1.Cells(Y, 1) = dsUser.Tables(0).Rows(i)("M18Batch")
                worksheet1.Cells(Y, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Cells(Y, 2) = dsUser.Tables(0).Rows(i)("M18SalesOrder")
                worksheet1.Cells(Y, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 3) = dsUser.Tables(0).Rows(i)("M18LineItem")
                worksheet1.Cells(Y, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter


                worksheet1.Cells(Y, 4) = dsUser.Tables(0).Rows(i)("M18Customer")
                worksheet1.Cells(Y, 4).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 5) = dsUser.Tables(0).Rows(i)("M18Material")
                worksheet1.Cells(Y, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                worksheet1.Cells(Y, 6) = dsUser.Tables(0).Rows(i)("M18Dis")
                worksheet1.Cells(Y, 6).HorizontalAlignment = XlHAlign.xlHAlignLeft

                'MsgBox(dsUser.Tables(0).Rows(i)("M18DDate"))
                worksheet1.Cells(Y, 7) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18DDate")), Today)
                worksheet1.Cells(Y, 7).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(Y, 8) = dsUser.Tables(0).Rows(i)("M18DDate")
                worksheet1.Cells(Y, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then
                    worksheet1.Cells(Y, 9) = "00/00/0000"
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter
                Else
                    worksheet1.Cells(Y, 9) = dsUser.Tables(0).Rows(i)("M18LCDate")
                    worksheet1.Cells(Y, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                If dsUser.Tables(0).Rows(i)("M18LCDate") = "1900/1/1" Then

                Else
                    worksheet1.Cells(Y, 10) = DateDiff("d", CDate(dsUser.Tables(0).Rows(i)("M18LCDate")), Today)
                    worksheet1.Cells(Y, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

                End If

                worksheet1.Cells(Y, 11) = dsUser.Tables(0).Rows(i)("M18QtyKG")
                worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 11)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 12) = dsUser.Tables(0).Rows(i)("M18Qty")
                worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(Y, 12)
                range1.NumberFormat = "0"

                worksheet1.Cells(Y, 13) = dsUser.Tables(0).Rows(i)("M18NextOparation")
                worksheet1.Cells(Y, 13).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 14) = dsUser.Tables(0).Rows(i)("M18PComment")
                worksheet1.Cells(Y, 14).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 15) = dsUser.Tables(0).Rows(i)("M18NonCon")
                worksheet1.Cells(Y, 15).HorizontalAlignment = XlHAlign.xlHAlignLeft
                worksheet1.Cells(Y, 16) = dsUser.Tables(0).Rows(i)("M18OrderType")
                worksheet1.Cells(Y, 16).HorizontalAlignment = XlHAlign.xlHAlignLeft

                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                Y = Y + 1
                i = i + 1
            Next

            Dim L12 As Integer
            Dim K12 As Integer

            worksheet1.Cells(Y, 8) = "Exam"
            worksheet1.Rows(Y).Font.size = 10
            worksheet1.Rows(Y).Font.bold = True
            L12 = Y
            K12 = Y
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & KK1 + 2 & ":k" & Y - 1 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(255, 192, 0)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & KK1 + 2 & ":l" & Y - 1 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(255, 192, 0)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash


            Y = Y + 1
            For i = Y To Y + 2
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
                worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            Next

            Y = Y + 1
            worksheet1.Cells(Y, 8) = "Total Qty"
            worksheet1.Rows(Y).Font.size = 12
            worksheet1.Rows(Y).Font.bold = True
            worksheet1.Range("K" & (Y)).Formula = "=SUM(k" & KK & "+k" & KK1 & "+k" & K12 & ")"
            worksheet1.Cells(Y, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("K" & Y & ":K" & Y).Interior.Color = RGB(217, 151, 149)
            worksheet1.Range("l" & (Y)).Formula = "=SUM(l" & KK & "+l" & KK1 & "+l" & K12 & ")"
            range1 = worksheet1.Cells(Y, 12)
            range1.NumberFormat = "0"
            worksheet1.Cells(Y, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter
            worksheet1.Range("L" & Y & ":L" & Y).Interior.Color = RGB(217, 151, 149)

            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("p" & Y, "p" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("o" & Y, "o" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("n" & Y, "n" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("m" & Y, "m" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("l" & Y, "l" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("k" & Y, "k" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("j" & Y, "j" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("i" & Y, "i" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("h" & Y, "h" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("g" & Y, "g" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("f" & Y, "f" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("e" & Y, "e" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("d" & Y, "d" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("c" & Y, "c" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("b" & Y, "b" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlDash
            worksheet1.Range("a" & Y, "a" & Y).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDash

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Function

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click

    End Sub
End Class