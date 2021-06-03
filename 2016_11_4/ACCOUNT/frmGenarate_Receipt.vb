Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmGenarate_Receipt
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Comcode As String



    Private Sub frmGenarate_Receipt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDate3.Text = Today
        txtDate4.Text = Today
        Call Load_Gride_Recipt()
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
    End Sub

    Function Load_Gride_Recipt()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Receipt
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 120
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False


            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(0).Columns(1).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            ' .DisplayLayout.Bands(0).Columns(1).
            ' .DisplayLayout.Bands(0).Header.Height = 60

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Gride_Recipt()
        Call Load_Data()

    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Gride_Recipt()
    End Sub

    Function Load_Data()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Amount As Double

        Try
            Sql = "select *  from View_Receipt where t01Date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and T01FromLoc_Code='" & _Comcode & "' order by t01ref_no"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            _Amount = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_no")
                Value = M01.Tables(0).Rows(i)("T03cash")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Amount = _Amount + CDbl(M01.Tables(0).Rows(i)("T03cash"))
                newRow("Amount") = _St
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            _St = (_Amount.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Amount))
            '  _Amount = Value + CDbl(M01.Tables(0).Rows(i)("T01cash"))
            newRow1("Amount") = _St
            c_dataCustomer1.Rows.Add(newRow1)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(3).Appearance.BackColor = Color.Gold
            ' UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(3).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
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
        Dim i As Integer
        Dim result1 As String
        Dim M01 As DataSet
        Dim Value As Double
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument
        Dim _ReceiptNo As String
        Dim _Ref As Integer
        Dim _sT As String
        Value = 0

        Try
            If UltraGrid1.Rows.Count > 1 Then
                nvcFieldList1 = "select * from P01Parameter where P01Code='RCP'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    _Ref = M01.Tables(0).Rows(0)("P01LastNo")
                End If


                nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='IN' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "select * from P01Parameter where P01Code='RCP'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                        _ReceiptNo = "REC00" & M01.Tables(0).Rows(0)("P01LastNo")
                    ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                        _ReceiptNo = "REC0" & M01.Tables(0).Rows(0)("P01LastNo")
                    Else
                        _ReceiptNo = "REC" & M01.Tables(0).Rows(0)("P01LastNo")
                    End If
                End If
            End If

            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='RCP' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Value = 0
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If Trim(UltraGrid1.Rows(i).Cells(0).Text) <> "" Then
                    nvcFieldList1 = "UPDATE T01Transaction_Header SET T01Receipt='YES' WHERE T01Trans_Type='DR' AND T01Ref_No='" & UltraGrid1.Rows(i).Cells(0).Text & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into T16Genarate_Receipt(T16Receipt,T16Date,T16Ref,T16Invoice,T16Amount,T16Com_Code,T16User,T16Gen_Date)" & _
                                                                      " values('" & _ReceiptNo & "','" & (UltraGrid1.Rows(i).Cells(1).Value) & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & _Comcode & "','" & strDisname & "','" & Today & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    Value = Value + CDbl(UltraGrid1.Rows(i).Cells(3).Text)
                End If
                i = i + 1
            Next
            _sT = "Cash Sales Receipt - " & _ReceiptNo
            nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Invo,T05Com_Code,T05User,T05Status)" & _
                                                              " values('" & _Ref & "','RECEIPT', '" & Today & "','CB001','" & _sT & "','" & Value & "','0','" & _ReceiptNo & "','" & _Comcode & "','" & strDisname & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Receipt genareted successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()
            A = MsgBox("Are you sure you want to print Cash Sale Receipt", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Sales Receipt .....")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\Cash_Receipt.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T16Genarate_Receipt.T16Receipt}='" & _ReceiptNo & "' and {T16Genarate_Receipt.T16Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            connection.Close()
            Call Load_Gride_Recipt()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try

    End Sub
End Class