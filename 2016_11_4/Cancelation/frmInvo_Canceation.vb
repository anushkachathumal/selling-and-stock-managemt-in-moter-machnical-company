Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmInvo_Canceation
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _MainStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Supplier As String
    Dim _Category As String
    Dim _Comcode As String
    Dim _EDITSTATUS As Boolean
    Dim _Root As String
    Dim _Ref_1 As Integer
    Dim _Loc_Code As String
    Dim _Ref As String

    Private Sub frmInvo_Canceation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Loc_Code = ConfigurationManager.AppSettings("LOCCODE")
        txtGross.ReadOnly = True
        txtGross.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
       

        ' Call Load_Loading()
        txtDis1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtTotal.ReadOnly = True
        'txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.ReadOnly = True
        'Call Load_Item()
        'Call Load_SALES_REF()
        txtI_Date.Text = Today
        txtI_Add1.ReadOnly = True
        ' txtLocation.ReadOnly = True
        txtRoot.ReadOnly = True
        txtSales.ReadOnly = True
        txtCustomer.ReadOnly = True


        txtDis1.ReadOnly = True
        txtDiscount.ReadOnly = True
        ' Call Load_Root()
        ' Call LOAD_GRIDE()
        Call Load_Invoice()
        Call Load_Gride1()

    End Sub

    Function Load_Invoice()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select T01Invoice_No AS [##] from T01Transaction_Header where T01Status='A' and T01Com_Code='" & _Loc_Code & "' and T01Trans_Type='DR'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboInvoice
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 310
                '  .Rows.Band.Columns(1).Width = 260


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub


    Function Load_Gride1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_SalesTR
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 180
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 110
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False

            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 90
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        End With
    End Function

    Function SEARCH_RECORDS()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer
        Dim Value As Double

        Dim _St As String

        Try
            Sql = "select * from View_T01Sales where  T01Com_Code='" & _Loc_Code & "' AND T01Invoice_No='" & cboInvoice.Text & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Ref = Trim(M01.Tables(0).Rows(0)("T01Ref_No"))
                txtCustomer.Text = Trim(M01.Tables(0).Rows(0)("M17Name"))
                txtI_Add1.Text = Trim(M01.Tables(0).Rows(0)("Expr1"))
                txtI_Add2.Text = Trim(M01.Tables(0).Rows(0)("Expr2"))
                txtI_Tp.Text = Trim(M01.Tables(0).Rows(0)("TP"))
                txtRoot.Text = Trim(M01.Tables(0).Rows(0)("M02Name"))
                Value = Trim(M01.Tables(0).Rows(0)("net"))
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtSales.Text = Trim(M01.Tables(0).Rows(0)("T01User"))

                Value = Trim(M01.Tables(0).Rows(0)("disc"))
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = CDbl(txtNett.Text) - CDbl(txtDiscount.Text)
                txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                cmdDelete.Enabled = True
            End If
            i = 0
            Call Load_Gride1()
            Sql = "select * from View_T02Transaction where  T01Com_Code='" & _Loc_Code & "' AND T02ref_no='" & _Ref & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow2 As DataRow In M01.Tables(0).Rows

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = Trim(M01.Tables(0).Rows(i)("M03Item_Code"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("M03Item_Name"))
                Value = Trim(M01.Tables(0).Rows(i)("T02Cost"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St
                Value = Trim(M01.Tables(0).Rows(i)("T02Retail_Price"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St
                newRow("Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                Value = Trim(M01.Tables(0).Rows(i)("T02MTR"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Mtr") = _St
                ' newRow("Rec.Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))

                Value = Trim(M01.Tables(0).Rows(i)("T02Total"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                'Value = CDbl(txtNett.Text) + Value
                'txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cboInvoice_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboInvoice.AfterCloseUp
        Call SEARCH_RECORDS()
    End Sub

    Private Sub cboInvoice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboInvoice.KeyUp
        If e.KeyCode = 13 Then
            Call SEARCH_RECORDS()
            cmdDelete.Focus()
        End If
    End Sub

    Function Clear_Text()
        Me.txtNett.Text = ""
        Me.txtGross.Text = ""
        Me.txtDiscount.Text = ""
        Me.txtI_Add1.Text = ""
        Me.txtI_Add2.Text = ""
        Me.txtRemark.Text = ""
        Me.txtRoot.Text = ""
        Me.txtSales.Text = ""
        Me.cboInvoice.Text = ""
        Call Load_Gride1()

    End Function
    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
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
        Dim A As String

        Try
            A = MsgBox("Are you sure you want to delete this Invoice", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information ..........")
            If A = vbYes Then
                nvcFieldList1 = "UPDATE T01Transaction_Header SET T01Status='CANCEL' WHERE T01Invoice_No='" & cboInvoice.Text & "' AND T01Com_Code='" & _Loc_Code & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T02Transaction_flutter SET T02Status='CANCEL' WHERE T02Ref_No='" & _Ref & "' AND T02Com_Code='" & _Loc_Code & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T02Transaction_flutter SET T02Status='CANCEL' WHERE T02Ref_No='" & _Ref & "' AND T02Com_Code='" & _Loc_Code & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S01Stock_Balance SET S01Status='CANCEL' WHERE S01Ref_No='" & _Ref & "' AND S01Com_Code='" & _Loc_Code & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "UPDATE S04Stock_Price SET S04Status='CANCEL' WHERE S04Ref_no='" & _Ref & "' AND S04Com_Code='" & _Loc_Code & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update t06OutStanding_Balance set T06Status='I'  where T06Invoice_No='" & cboInvoice.Text & "' and T06Com_Code ='" & _Loc_Code & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update T03Pay_Main set T03Status='I'  where T03Ref_No='" & _Ref & "' and T03Com_Code ='" & _Loc_Code & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            MsgBox("Records deleted successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            Call Clear_Text()
            connection.Close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim B As New ReportDocument
        Dim A1 As String
        Try
            A1 = ConfigurationManager.AppSettings("ReportPath") + "\InvoTm.rpt"
            B.Load(A1.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            ' B.SetParameterValue("Customer", cboCustomer.Text)
            'B.SetParameterValue("From", _From)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{View_T01Sales.T01Ref_No}=" & _Ref & ""
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'connection.Close()
            End If
        End Try
    End Sub
End Class