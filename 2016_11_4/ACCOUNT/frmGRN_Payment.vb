Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Configuration
Imports DBLotVbnet.modlVar1
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmGRN_Payment
    Dim _Acc_Type As String
    Dim _Comcode As String
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable
    Dim c_dataCustomer4 As DataTable
    Dim _From As Date
    Dim _To As Date

    Dim _Suppcode As String
    Dim _Search_Status As String


    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Gride_Data_Supplier()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            Call Load_Gride2()
            Sql = "select *  from View_Non_Paid_Invoice where T01com_code='" & _Comcode & "' and M09Name='" & Trim(cboSupplier.Text) & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("##") = False
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Com.Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Paid = Value
                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St
                _Paid = _Paid - Value
                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St
                _Paid = _Paid + Value
                If IsNumeric(M01.Tables(0).Rows(i)("T01NBT")) Then
                    Value = M01.Tables(0).Rows(i)("T01NBT")
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("NBT") = _St
                    _Paid = _Paid + Value
                Else
                    newRow("NBT") = "00.00"
                End If

                Value = M01.Tables(0).Rows(i)("T01Market_Return")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("MK Return") = _St
                _Paid = _Paid - Value
                Value = _Paid
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Paid Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Paid Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Supplier4()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Piad_Invoice_Summery where T14Com_Code='" & _Comcode & "' and M09Name='" & Trim(cboSupplier.Text) & "' order by T14Ref desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                ' newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Supplier_Print_Chq()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Piad_Invoice_Summery where T14Com_Code='" & _Comcode & "' and T14Chq_Print='NO' order by T14Ref desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                ' newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Supplier_Print_Sum()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Piad_Invoice_Summery where T14Com_Code='" & _Comcode & "'  order by T14Ref desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                ' newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Print_All()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Paid_Invoice_Detailes where T14Com_Code='" & _Comcode & "' order by T14Ref desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Supplier1()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Paid_Invoice_Detailes where T14Com_Code='" & _Comcode & "' and M09Name='" & Trim(cboSupplier.Text) & "' order by T14Ref desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Supplier_All()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Paid_Invoice_Detailes where T14Com_Code='" & _Comcode & "'  order by T14Ref "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Supplier_Paid()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Piad_Invoice_Summery where T14Com_Code='" & _Comcode & "' and M09Name='" & Trim(cboSupp1.Text) & "' and T14Date between '" & txtC1.Text & "' and '" & txtC2.Text & "' order by T14Ref "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                ' newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function
    Function Load_Gride_Data_Supplier_Paid_Date1()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Piad_Invoice_Summery where T14Com_Code='" & _Comcode & "' and T14Date between '" & txtE1.Text & "' and '" & txtE2.Text & "' and T15Bank_code='" & cboAccount.Text & "' order by T14Ref "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                ' newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Supplier_Paid_Date()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Piad_Invoice_Summery where T14Com_Code='" & _Comcode & "' and T14Date between '" & txtD1.Text & "' and '" & txtD2.Text & "' order by T14Ref "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                ' newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function
    Function Load_Gride_Data_Supplier2()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_CANCEL_GRN_VOUCHER where T14Com_Code='" & _Comcode & "' and M09Name='" & Trim(cboSupplier.Text) & "' order by T14Ref desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")
                newRow("Canceld by") = M01.Tables(0).Rows(i)("tmpuser")
                newRow("Cancel Time") = M01.Tables(0).Rows(i)("tmptime")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data_Supplier3()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Paid_Invoice_Detailes where T14Com_Code='" & _Comcode & "' and M09Name='" & Trim(cboSupp1.Text) & "' and T14date between '" & txtC1.Text & "' and '" & txtC2.Text & "' order by T14Ref desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Load_Gride_Data()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            Call Load_Gride2()
            Sql = "select *  from View_Non_Paid_Invoice where T01To_Loc_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("##") = False
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Com.Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Paid = Value
                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St
                _Paid = _Paid - Value
                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St
                _Paid = _Paid + Value
                If IsNumeric(M01.Tables(0).Rows(i)("T01NBT")) Then
                    Value = M01.Tables(0).Rows(i)("T01NBT")
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("NBT") = _St
                    _Paid = _Paid + Value
                Else
                    newRow("NBT") = "00.00"
                End If

                Value = M01.Tables(0).Rows(i)("T01Market_Return")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("MK Return") = _St
                _Paid = _Paid - Value
                Value = _Paid
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Paid Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Paid Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    

    Function Load_Gride_Data_date1()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            '  Call Load_Gride2()
            Sql = "select *  from View_Paid_Invoice_Detailes where T14Com_Code='" & _Comcode & "' and T14date between '" & txtD1.Text & "' and '" & txtD2.Text & "' order by T14Ref "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref_No") = M01.Tables(0).Rows(i)("T14Ref")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T14Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T14Date")) & "/" & Year(M01.Tables(0).Rows(i)("T14Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Voucher No") = M01.Tables(0).Rows(i)("T14voucher")
                newRow("GRN No") = M01.Tables(0).Rows(i)("T14GRN")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09name")

                Value = M01.Tables(0).Rows(i)("T14amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Total = Value + _Total
                newRow("Net Amount") = _St
                newRow("Chq No") = M01.Tables(0).Rows(i)("T04Chq_No")
                newRow("Bank Name") = M01.Tables(0).Rows(i)("M01Acc_Name")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try

    End Function

    Function Clear_Panal2()
        Me.txtVoucher_1.Text = ""
        Me.txtDate_1.Text = ""
        Me.txtDue_1.Text = ""
        Me.txtB_Name1.Text = ""
        Me.txtPay_1.Text = ""
        Me.txtB_Code1.Text = ""
        Me.txtRemark_1.Text = ""
        Me.txtnet_1.Text = ""
        Call Load_Gride4()

    End Function
    Private Sub frmGRN_Payment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride2()
        Call Load_Gride_Data()
        Call Load_Supplier()
        txtRef.ReadOnly = True
        txtRef.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtName.ReadOnly = True
        txtTotal.ReadOnly = True
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        Call Load_Voucher()
        Call Load_Gride3()
        '  txtPay_C.ReadOnly = True
        txtPay_Amount.ReadOnly = True
        txtPay_Dis.ReadOnly = True
        Call Load_Gride4()

        Call Load_Bank()
        Call Load_Payee()
        txtDate.Text = Today

        txtVoucher_1.ReadOnly = True
        txtVoucher_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate_1.ReadOnly = True
        txtDate_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDue_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDue_1.ReadOnly = True

        txtB_Code1.ReadOnly = True
        txtB_Name1.ReadOnly = True
        txtChq_1.ReadOnly = True
        txtPay_1.ReadOnly = True
        txtnet_1.ReadOnly = True
        txtnet_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Right

    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_GRN_PAY
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 60
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 180
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 110
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 90
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Width = 90
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Width = 90
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(10).Width = 90
            .DisplayLayout.Bands(0).Columns(10).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(8).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(9).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(10).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            .DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride_2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_GRN_CANCEL_INVOICE
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 60
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 80
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 180
            '.DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(8).Width = 90
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 90
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(10).Width = 90
            '.DisplayLayout.Bands(0).Columns(10).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(0).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(8).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(9).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(10).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride_3()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_GRN_PAYID_INVOICE_SUMMERY
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 60
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 80
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 180
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(7).Width = 180
            '.DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(8).Width = 90
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 90
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(10).Width = 90
            '.DisplayLayout.Bands(0).Columns(10).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(0).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(8).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(9).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(10).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function


    Function Load_Gride_1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_GRN_PAYID_INVOICE
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 60
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 80
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 180
            '.DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(8).Width = 90
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 90
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(10).Width = 90
            '.DisplayLayout.Bands(0).Columns(10).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(0).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(8).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(9).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(10).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Search_Supplier() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Status='A' and M01Acc_Type='SP' and M01Acc_Name='" & Trim(txtP_Name.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Supplier = True
                _Suppcode = Trim(M01.Tables(0).Rows(0)("M01Acc_Code"))
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try


    End Function

    Function Search_Bank() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Status='A' and M01Acc_Type='3' and M01Acc_Code='" & Trim(txtBank.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Bank = True
                txtName.Text = Trim(M01.Tables(0).Rows(0)("M01Acc_Name"))
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try


    End Function

    Function Load_Gride3()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_GRN_Payment1
        UltraGrid2.DataSource = c_dataCustomer2
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 180
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
           

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(5).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(6).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(8).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(9).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(10).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride4()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer4 = CustomerDataClass.MakeDataTable_GRN_Payment1
        UltraGrid3.DataSource = c_dataCustomer4
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 180
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90


            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(5).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(6).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(8).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(9).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(10).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function


    Private Sub ByDateToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem1.Click
        Panel4.Visible = True
        _Search_Status = "A"
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Loc_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupplier
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                ' .Rows.Band.Columns(1).Width = 180


            End With

            With cboSupp1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Function Load_Bank()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Code as [##] from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Status='A' and M01Acc_Type='3'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtBank
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 130
                ' .Rows.Band.Columns(1).Width = 180
            End With

            With cboAccount
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 130
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function


    Function Load_Payee()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Name as [##] from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Status='A' and M01Acc_Type='SP'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtP_Name
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 290
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        If _Search_Status = "A" Then
            Call Load_Gride2()
            Call Load_Gride_Data_Supplier()
            Panel4.Visible = False
        ElseIf _Search_Status = "B" Then
            Call Load_Gride_1()
            Call Load_Gride_Data_Supplier1()
            Panel4.Visible = False
        ElseIf _Search_Status = "C" Then
            Call Load_Gride_3()
            Call Load_Gride_Data_Supplier3()
            Panel4.Visible = False

        ElseIf _Search_Status = "D" Then
            Call Load_Gride_2()
            Call Load_Gride_Data_Supplier2()
            Panel4.Visible = False
        ElseIf _Search_Status = "P2" Then
            Call Load_Gride2()
            Call Load_Gride_Data_Supplier()
            Panel4.Visible = False
            _Suppcode = Trim(cboSupplier.Text)
        ElseIf _Search_Status = "P3" Then
            Call Load_Gride_1()
            Call Load_Gride_Data_Supplier1()
            Panel4.Visible = False
            _Suppcode = Trim(cboSupplier.Text)
        ElseIf _Search_Status = "P4" Then
            Call Load_Gride_1()
            Call Load_Gride_Data_Supplier1()
            Panel4.Visible = False
            _Suppcode = Trim(cboSupplier.Text)
        ElseIf _Search_Status = "P8" Then
            Call Load_Gride_3()
            Call Load_Gride_Data_Supplier4()
            _Suppcode = Trim(cboSupplier.Text)
            Panel4.Visible = False
        End If
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        OPR3.Visible = False
        Call Clear_Text()
    End Sub

    Function Clear_Text()
        Me.txtTotal.Text = ""
        Me.txtName.Text = ""
        Me.txtRemark.Text = ""
        Me.cboSupplier.Text = ""
        Me.txtChq.Text = ""
        Me.txtBank.Text = ""
        Me.txtP_Name.Text = ""
    End Function

    Function Load_Voucher()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where P01Code='VOU' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    txtRef.Text = _Comcode & "/VU00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    txtRef.Text = _Comcode & "/VU0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtRef.Text = _Comcode & "/VU" & M01.Tables(0).Rows(0)("P01LastNo")

                End If
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        Dim _RowIndex As Integer
        Dim _Status As Integer
        Dim I As Integer
        Dim _Net As Double
        Dim _Last As Integer
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _Ref As Integer

        Try

            If _Search_Status = "A" Then
                _Net = 0
                I = 0
                _Status = 0
                _Last = UltraGrid1.Rows.Count

                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    If I < _Last - 1 Then
                        If UltraGrid1.Rows(I).Cells(0).Value = True Then
                            _Status = _Status + 1
                        End If
                    End If
                    I = I + 1
                Next

                If _Status >= 1 Then
                    OPR3.Visible = True
                    Call Load_Voucher()
                    Call Load_Gride3()
                    Call Clear_Text()
                    I = 0
                    For Each uRow As UltraGridRow In UltraGrid1.Rows
                        If I < _Last - 1 Then
                            If UltraGrid1.Rows(I).Cells(0).Value = True Then
                                Dim newRow As DataRow = c_dataCustomer2.NewRow

                                newRow("Supplier Name") = UltraGrid1.Rows(I).Cells(1).Value

                                ' newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                                newRow("GRN No") = UltraGrid1.Rows(I).Cells(3).Value
                                newRow("Date") = (UltraGrid1.Rows(I).Cells(2).Value)
                                newRow("Com.Invoice") = (UltraGrid1.Rows(I).Cells(4).Value)
                                newRow("Paid Amount") = (UltraGrid1.Rows(I).Cells(10).Value)
                                _Net = _Net + CDbl(UltraGrid1.Rows(I).Cells(10).Value)

                                c_dataCustomer2.Rows.Add(newRow)
                            End If
                        End If
                        I = I + 1
                    Next


                    txtTotal.Text = (_Net.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Net))

                    txtBank.ToggleDropdown()
                End If
            ElseIf _Search_Status = "B" Then
                Call Load_Gride4()
                OPR5.Visible = True
                txtVoucher_1.Text = (UltraGrid1.Rows(I).Cells(3).Value)
                txtDate_1.Text = (UltraGrid1.Rows(I).Cells(2).Value)
                txtB_Name1.Text = (UltraGrid1.Rows(I).Cells(7).Value)
                txtChq_1.Text = (UltraGrid1.Rows(I).Cells(6).Value)
                txtPay_1.Text = (UltraGrid1.Rows(I).Cells(1).Value)

                Sql = "select * from View_Paid_Invoice_Detailes where T14Voucher='" & Trim(txtVoucher_1.Text) & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                I = 0
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer4.NewRow

                    newRow("Supplier Name") = M01.Tables(0).Rows(I)("M09name")
                    ' newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                    newRow("GRN No") = M01.Tables(0).Rows(I)("T14Grn")
                    Sql = " select *  from T01Transaction_Header where t01grn_no='" & M01.Tables(0).Rows(I)("T14Grn") & "' "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        newRow("Date") = Month(M02.Tables(0).Rows(0)("T01date")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(0)("T01date")) & "/" & Year(M02.Tables(0).Rows(0)("T01date"))
                        newRow("Com.Invoice") = M02.Tables(0).Rows(0)("T01invoice_no")
                        newRow("Paid Amount") = M02.Tables(0).Rows(0)("T01net_amount")

                        _Net = _Net + CDbl(M02.Tables(0).Rows(0)("T01net_amount"))
                    End If
                    c_dataCustomer4.Rows.Add(newRow)
                    I = I + 1
                Next

                txtnet_1.Text = (_Net.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtnet_1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Net))

                Sql = "select * from T15Bank_Transaction where T15Pay_No='" & Trim(txtVoucher_1.Text) & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    txtRemark_1.Text = M02.Tables(0).Rows(0)("T15Remark")
                    txtB_Code1.Text = M02.Tables(0).Rows(0)("T15bank_code")
                    _Ref = M02.Tables(0).Rows(0)("T15Ref")
                End If

                Sql = "select * from T04Chq_Trans where T04Ref_No='" & _Ref & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    txtDue_1.Text = M02.Tables(0).Rows(0)("T04DOR")
                    ' txtB_Code1.Text = M02.Tables(0).Rows(0)("T15bank_code")
                End If


            ElseIf _Search_Status = "C" Then
                Call Load_Gride4()
                OPR5.Visible = True
                txtVoucher_1.Text = (UltraGrid1.Rows(I).Cells(3).Value)
                txtDate_1.Text = (UltraGrid1.Rows(I).Cells(2).Value)
                txtB_Name1.Text = (UltraGrid1.Rows(I).Cells(6).Value)
                txtChq_1.Text = (UltraGrid1.Rows(I).Cells(5).Value)
                txtPay_1.Text = (UltraGrid1.Rows(I).Cells(1).Value)

                Sql = "select * from View_Paid_Invoice_Detailes where T14Voucher='" & Trim(txtVoucher_1.Text) & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                I = 0
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer4.NewRow

                    newRow("Supplier Name") = M01.Tables(0).Rows(I)("M09name")
                    ' newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                    newRow("GRN No") = M01.Tables(0).Rows(I)("T14Grn")
                    Sql = " select *  from T01Transaction_Header where t01grn_no='" & M01.Tables(0).Rows(I)("T14Grn") & "' "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        newRow("Date") = Month(M02.Tables(0).Rows(0)("T01date")) & "/" & Microsoft.VisualBasic.Day(M02.Tables(0).Rows(0)("T01date")) & "/" & Year(M02.Tables(0).Rows(0)("T01date"))
                        newRow("Com.Invoice") = M02.Tables(0).Rows(0)("T01invoice_no")
                        newRow("Paid Amount") = M02.Tables(0).Rows(0)("T01net_amount")

                        _Net = _Net + CDbl(M02.Tables(0).Rows(0)("T01net_amount"))
                    End If
                    c_dataCustomer4.Rows.Add(newRow)
                    I = I + 1
                Next

                txtnet_1.Text = (_Net.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtnet_1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Net))

                Sql = "select * from T15Bank_Transaction where T15Pay_No='" & Trim(txtVoucher_1.Text) & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    txtRemark_1.Text = M02.Tables(0).Rows(0)("T15Remark")
                    txtB_Code1.Text = M02.Tables(0).Rows(0)("T15bank_code")
                    _Ref = M02.Tables(0).Rows(0)("T15Ref")
                End If

                Sql = "select * from T04Chq_Trans where T04Ref_No='" & _Ref & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    txtDue_1.Text = M02.Tables(0).Rows(0)("T04DOR")
                    ' txtB_Code1.Text = M02.Tables(0).Rows(0)("T15bank_code")
                End If


            ElseIf _Search_Status = "C1" Then
                Panel5.Visible = True
                I = UltraGrid1.ActiveRow.Index
                txtPay_C.Text = UltraGrid1.Rows(I).Cells(1).Text
                _Net = UltraGrid1.Rows(I).Cells(4).Text
                txtPay_Amount.Text = (_Net.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtPay_Amount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Net))
                txtPay_Dis.Text = Num2String(CDbl(txtPay_Amount.Text))
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub

    Private Sub txtBank_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBank.AfterCloseUp
        Call Search_Bank()
    End Sub

    
   

    Private Sub txtBank_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBank.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Bank()
            txtChq.Focus()
        End If
    End Sub

    Private Sub txtChq_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChq.KeyUp
        If e.KeyCode = 13 Then
            txtDue.Focus()
        End If
    End Sub

    Private Sub txtDue_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDue.KeyUp
        If e.KeyCode = 13 Then
            txtP_Name.ToggleDropdown()
        End If
    End Sub

    Private Sub txtP_Name_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtP_Name.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click

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
        Dim _Status As Boolean
        Dim _RefNo As Integer
        Dim A1 As String
        Dim B As New ReportDocument

        Try
            If Search_Bank() = True Then
            Else
                MsgBox("Please select the correct Bank Code", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                txtBank.ToggleDropdown()
                Exit Sub
            End If

            If Search_Supplier() = True Then
            Else
                MsgBox("Please select the correct supplier", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                txtP_Name.ToggleDropdown()
                Exit Sub
            End If

            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = "-"
            End If
            If UltraGrid2.Rows.Count >= 1 Then

            Else
                MsgBox("You can't create this voucher ", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                Exit Sub
            End If

            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                If Trim(UltraGrid2.Rows(i).Cells(0).Text) = Trim(txtP_Name.Text) Then

                Else
                    _Status = True
                End If
                i = i + 1
            Next

            If _Status = True Then
                MsgBox("Please check the Supplier", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub
            End If

            If IsDate(txtDue.Text) Then
            Else
                MsgBox("Please enter the correct due date", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub

            End If

            If txtChq.Text <> "" Then
            Else
                MsgBox("Please enter the chq No", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                Exit Sub
            End If

            nvcFieldList1 = "select * from P01Parameter where P01Code='IN'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                _RefNo = MB51.Tables(0).Rows(0)("P01LastNo")
            End If

            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo+" & 1 & " WHERE P01Code='VOU' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo+" & 1 & " WHERE P01Code='IN' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            i = 0
            'TRANSACTION HEADER
            nvcFieldList1 = "Insert Into T15Bank_Transaction(T15Ref,T15Date,T15TR_Type,T15Bank_Code,T15Remark,T15Tr_Status,T15Cr,T15Dr,T15Status,T15Com_Code,T15User,T15Pay_No)" & _
                                                                      " values(" & _RefNo & ", '" & txtDate.Text & "','GRN_PAY','" & Trim(txtBank.Text) & "','" & Trim(txtRemark.Text) & "','NO','0','" & CDbl(txtTotal.Text) & "','A','" & _Comcode & "','" & strDisname & "','" & txtRef.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            'ACCOUNT HEADER
            nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Date,T05Acc_Type,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Status,T05Com_Code,T05User,T05Invo)" & _
                                                                      " values(" & _RefNo & ", '" & txtDate.Text & "','GRN_PAY','" & Trim(txtBank.Text) & "','" & Trim(txtRemark.Text) & "','0','" & CDbl(txtTotal.Text) & "','A','" & _Comcode & "','" & strDisname & "','" & txtRef.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Date,T05Acc_Type,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Status,T05Com_Code,T05User,T05Invo)" & _
                                                                      " values(" & _RefNo & ", '" & txtDate.Text & "','GRN_PAY','" & _Suppcode & "','" & Trim(txtRemark.Text) & "','" & CDbl(txtTotal.Text) & "','0','A','" & _Comcode & "','" & strDisname & "','" & txtRef.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'CHQ HEADER
            nvcFieldList1 = "Insert Into T04Chq_Trans(T04Ref_No,T04Acc_Type,T04Chq_no,T04ACC_No,T04Amount,T04DOR,T04Status,T04Com_Code)" & _
                                                                     " values(" & _RefNo & ",'GRN_PAY','" & Trim(txtChq.Text) & "','" & Trim(txtBank.Text) & "','" & CDbl(txtTotal.Text) & "','" & txtDue.Text & "','A','" & _Comcode & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'TRANSACTION LOG
            nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpUser,tmpLog)" & _
                                                                     " values('GRN_PAY','SAVE','" & Trim(txtRef.Text) & "','" & Now & "','" & strDisname & "','" & _Comcode & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                nvcFieldList1 = "Insert Into T14GRN_Pay(T14Ref,T14Tr_Type,T14Date,T14Voucher,T14Supplier,T14GRN,T14Amount,T14Status,T14Com_Code,T14Chq_Print,T14Count)" & _
                                                                    " values(" & _RefNo & ",'GRN_PAY','" & txtDate.Text & "','" & Trim(txtRef.Text) & "','" & _Suppcode & "','" & UltraGrid2.Rows(i).Cells(1).Value & "','" & UltraGrid2.Rows(i).Cells(4).Value & "','A','" & _Comcode & "','NO','" & i + 1 & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T01Transaction_Header SET T01Paid='1' WHERE T01Grn_No='" & UltraGrid2.Rows(i).Cells(1).Value & "' AND T01To_Loc_Code='" & _Comcode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = i + 1
            Next

            transaction.Commit()
            connection.Close()
            result1 = MsgBox("Are you sure you want to print this Voucher", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Voucher .....")
            If result1 = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\pay_Voucher.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("Amount", txtTotal.Text)
                'B.SetParameterValue("Dis", txtName.Text)
                'B.SetParameterValue("Voucher", _Voucher)
                'frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T15Bank_Transaction.T15Ref}=" & _RefNo & " and {T15Bank_Transaction.T15Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If
            Call Clear_Text()
            Call Load_Voucher()
            OPR3.Visible = False
            Call Load_Gride2()
            Call Load_Gride3()
            Call Load_Gride_Data()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Panel4.Visible = False
        OPR3.Visible = False
        cmdAdd.Enabled = True
        Panel5.Visible = False
        Call Clear_Text()
        Call Load_Voucher()
        Call Load_Gride2()
        Call Load_Gride3()
        Call Load_Gride_Data()
        Call Clear_Panal2()
        Call Load_Gride4()
        _Search_Status = ""
        Panel2.Visible = False
        AllTobePrintChequeToolStripMenuItem.Checked = False
        Panel1.Visible = False
        OPR5.Visible = False
    End Sub

 


    Private Sub DetailsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetailsToolStripMenuItem1.Click
        Panel4.Visible = True
        _Search_Status = "B"
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub


    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Clear_Panal2()
        OPR5.Visible = False
    End Sub

    Private Sub UltraButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click
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
        Dim _Status As Boolean
        Dim _RefNo As Integer
        Dim A1 As String
        Dim B As New ReportDocument
        Try
            A1 = MsgBox("Are you sure you want to cancel this voucher", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Cancel Voucher ........")
            If A1 = vbYes Then
                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpAthz,tmpUser,tmpLog)" & _
                                                                    " values('GRN_PAY', 'DELETE','" & Trim(txtVoucher_1.Text) & "','" & Now & "','" & strDisname & "','" & strDisname & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T14GRN_Pay SET T14Status='I' WHERE T14Voucher='" & txtVoucher_1.Text & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "SELECT * FROM T14GRN_Pay WHERE  T14Voucher='" & txtVoucher_1.Text & "'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                    nvcFieldList1 = "UPDATE T15Bank_Transaction SET T15Status='I' WHERE T15Ref='" & Trim(MB51.Tables(0).Rows(0)("T14Ref")) & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "UPDATE T04Chq_Trans SET T04Status='I' WHERE T04Ref_No='" & Trim(MB51.Tables(0).Rows(0)("T14Ref")) & "' AND T04Acc_Type='GRN_PAY' AND T04Status='A'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "UPDATE T05Acc_Trans SET T05Status='I' WHERE T05Ref_No='" & Trim(MB51.Tables(0).Rows(0)("T14Ref")) & "' AND T05Acc_Type='GRN_PAY' AND T05Status='A' AND T05Com_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    'nvcFieldList1 = "UPDATE T01Transaction_Header SET T01Paid='0' WHERE T01Grn_No='" & Trim(MB51.Tables(0).Rows(0)("T14GRN")) & "' AND T01To_Loc_Code='" & _Comcode & "'"
                    'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                End If

                nvcFieldList1 = "SELECT * FROM T14GRN_Pay WHERE  T14Voucher='" & txtVoucher_1.Text & "'"
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                i = 0
                For Each DTRow1 As DataRow In MB51.Tables(0).Rows
                    nvcFieldList1 = "UPDATE T01Transaction_Header SET T01Paid='0' WHERE T01Grn_No='" & Trim(MB51.Tables(0).Rows(i)("T14GRN")) & "' AND T01To_Loc_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    i = i + 1
                Next

                MsgBox("GRN Voucher cancel successfully", MsgBoxStyle.Information, "Information .........")
                transaction.Commit()
                connection.Close()
                Call Clear_Panal2()
                Call Load_Gride4()
                OPR5.Visible = False
                Call Load_Gride2()
                Call Load_Gride_Data()
            End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Dim A1 As String
        Dim B As New ReportDocument
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select * from T14GRN_Pay where T14Voucher='" & Trim(txtVoucher_1.Text) & "' and T14Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\pay_Voucher.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("Amount", txtTotal.Text)
                'B.SetParameterValue("Dis", txtName.Text)
                'B.SetParameterValue("Voucher", _Voucher)
                'frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T15Bank_Transaction.T15Ref}=" & M01.Tables(0).Rows(0)("T14Ref") & " and {T15Bank_Transaction.T15Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Sub

    Private Sub CancelVoucherToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelVoucherToolStripMenuItem.Click
        Panel4.Visible = True
        _Search_Status = "D"
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub SummeryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SummeryToolStripMenuItem.Click
        Panel4.Visible = True
        _Search_Status = "C"
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub AllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllToolStripMenuItem.Click
        _Search_Status = "P1"
        Call Load_Gride2()
        Call Load_Gride_Data()
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub PrintReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintReportToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String

        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"

            If _Search_Status = "P1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Non_Paid_GRN.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                ' frmReport.CrystalReportViewer1.SelectionFormula = "{View_CashierSummery.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_CashierSummery.t01User}='" & cboCashier.Text & "' and {View_CashierSummery.Location} ='" & _Comcode & "'"
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Non_Paid_Invoice.T01To_Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Search_Status = "P2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Non_Paid_GRN.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Non_Paid_Invoice.M09Name}='" & _Suppcode & "' and {View_Non_Paid_Invoice.T01To_Loc_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Search_Status = "P3" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Paid_Invoice_Detailes.T14Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()

            ElseIf _Search_Status = "P4" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Paid_Invoice_Detailes.T14Com_Code}='" & _Comcode & "' and {View_Paid_Invoice_Detailes.M09Name}='" & _Suppcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Search_Status = "P5" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Paid_Invoice_Detailes.T14Com_Code}='" & _Comcode & "' and {View_Paid_Invoice_Detailes.M09Name}='" & _Suppcode & "' and {View_Paid_Invoice_Detailes.T14Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Search_Status = "P6" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Paid_Invoice_Detailes.T14Com_Code}='" & _Comcode & "'  and {View_Paid_Invoice_Detailes.T14Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()

            ElseIf _Search_Status = "P7" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher_Summery.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Piad_Invoice_Summery.T14Com_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Search_Status = "P8" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher_Summery.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Piad_Invoice_Summery.T14Com_Code}='" & _Comcode & "' and {View_Piad_Invoice_Summery.M09Name}='" & _Suppcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Search_Status = "P9" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher_Summery.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Piad_Invoice_Summery.T14Com_Code}='" & _Comcode & "' and {View_Piad_Invoice_Summery.M09Name}='" & _Suppcode & "' and {View_Piad_Invoice_Summery.T14Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Search_Status = "P11" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher_Summery.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Piad_Invoice_Summery.T14Com_Code}='" & _Comcode & "' and  {View_Piad_Invoice_Summery.T14Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & ""
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _Search_Status = "P10" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Paid_GRN_Voucher_Summery.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = " {View_Piad_Invoice_Summery.T14Com_Code}='" & _Comcode & "' and  {View_Piad_Invoice_Summery.T14Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_Piad_Invoice_Summery.T15Bank_Code}='" & _Suppcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub

    Private Sub BySupplierToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem2.Click
        Panel4.Visible = True
        _Search_Status = "P2"
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

   
    Private Sub AllTransactionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllTransactionToolStripMenuItem.Click
        ' Panel4.Visible = True
        Call Load_Gride_1()
        Call Load_Gride_Data_Print_All()
        ' Panel4.Visible = False
        _Search_Status = "P3"
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub AllTransactionToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllTransactionToolStripMenuItem2.Click
        Panel4.Visible = True
        _Search_Status = "P4"
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub ByDateToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem4.Click
        _Search_Status = "P5"
        txtC1.Text = Today
        txtC2.Text = Today
        ' _Suppcode = cboSupp1.Text
        Panel1.Visible = True
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        If _Search_Status = "P5" Then
            _Suppcode = cboSupp1.Text
            _From = txtC1.Text
            _To = txtC2.Text
            Call Load_Gride_1()
            Call Load_Gride_Data_Supplier3()
            Panel1.Visible = False
        ElseIf _Search_Status = "P9" Then
            _Suppcode = cboSupp1.Text
            _From = txtC1.Text
            _To = txtC2.Text
            Call Load_Gride_3()
            Call Load_Gride_Data_Supplier_Paid()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub ByDateToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem2.Click
        _Search_Status = "P6"
        txtD1.Text = Today
        txtD2.Text = Today
        Panel2.Visible = True
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub UltraButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton7.Click
        If _Search_Status = "P6" Then
            _From = txtD1.Text
            _To = txtD2.Text
            Call Load_Gride_1()
            Call Load_Gride_Data_date1()
            Panel2.Visible = False
        ElseIf _Search_Status = "P11" Then
            _From = txtD1.Text
            _To = txtD2.Text
            Call Load_Gride_3()
            Call Load_Gride_Data_Supplier_Paid_Date()
            Panel2.Visible = False
        End If
    End Sub

    Private Sub AllTransactionToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllTransactionToolStripMenuItem1.Click
        Call Load_Gride_3()
        Call Load_Gride_Data_Supplier_Print_Sum()
        _Search_Status = "P7"
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub AllTransactionToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllTransactionToolStripMenuItem3.Click
        _Search_Status = "P8"
        Panel4.Visible = True
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub ByDateToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem5.Click
        _Search_Status = "P9"
        txtC1.Text = Today
        txtC2.Text = Today
        ' _Suppcode = cboSupp1.Text
        Panel1.Visible = True
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub ByDateToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem3.Click
        _Search_Status = "P11"
        txtD1.Text = Today
        txtD2.Text = Today
        Panel2.Visible = True
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub UltraButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton8.Click
        _From = txtE1.Text
        _To = txtE2.Text
        _Suppcode = Trim(cboAccount.Text)
        Call Load_Gride_3()
        Call Load_Gride_Data_Supplier_Paid_Date1()
        Panel3.Visible = False
    End Sub

    Private Sub ByBankToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByBankToolStripMenuItem1.Click
        _Search_Status = "P10"
        Panel3.Visible = True
        txtE1.Text = Today
        txtE2.Text = Today
        AllTobePrintChequeToolStripMenuItem.Checked = False
    End Sub

    Private Sub AllTobePrintChequeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllTobePrintChequeToolStripMenuItem.Click
        Call Load_Gride_3()
        Call Load_Gride_Data_Supplier_Print_Chq()
        _Search_Status = "C1"
        AllTobePrintChequeToolStripMenuItem.Checked = True
    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub Panel5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel5.Paint

    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class