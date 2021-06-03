Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmCr_Received
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable
    Dim c_dataCustomer4 As DataTable

    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Comcode As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub frmCr_Received_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride1()
        Call Load_Gride1_Data()
        Call Load_Customer()
        txtDate.Text = Today
        Call Load_Gride_Pay_Invo()
        '  txtDue.Text = Today
        Call Load_Gride_Invo()
        txtCash.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPayment.ReadOnly = True
        txtPayment.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Bank()
        txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtPayment1.ReadOnly = True
        txtPayment1.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtRef1.ReadOnly = True
        txtRef1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPay.ReadOnly = True
        txtPay.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCus_Code1.ReadOnly = True
        txtName1.ReadOnly = True
        cboStatus1.ReadOnly = True
        txtCash1.ReadOnly = True
        txtCash1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtAmount1.ReadOnly = True
        txtAmount1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtChq1.ReadOnly = True
        txtDue1.ReadOnly = True

    End Sub
    Function Load_Bank()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select M33Name as [##] from M33Bank_Name"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboBank
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 190
            End With

            With cboBank1
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 190
            End With
            'SQL = "select M01Acc_Code as [Account No] from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Acc_Type='BN'"
            'T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            'With cboBank
            '    .DataSource = T01
            '    .Rows.Band.Columns(0).Width = 130
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride_Invo()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_Inv
        UltraGrid2.DataSource = c_dataCustomer2
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Function Search_Records()
        Dim I As Integer
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim _St As String
        Dim _Total As Double
        Dim Value As Double

        Try
            Sql = "select * from T04Chq_Trans where T04Ref_No='" & txtRef1.Text & "' and T04Acc_Type='PR' and T04Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtDue1.Text = M01.Tables(0).Rows(0)("T04DOR")
                cboBank1.Text = M01.Tables(0).Rows(0)("T04Bank_Name")
                txtChq1.Text = M01.Tables(0).Rows(0)("T04Chq_no")

            End If
            con.close()
            Call Load_Gride_Pay_Invo()
            I = 0
            _Total = 0
            Sql = " select * from t06OutStanding_Balance where T06Tr_Type='PR' and T06RefNo='" & txtRef1.Text & "' and T06Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer4.NewRow

                newRow("Invo No") = M01.Tables(0).Rows(I)("T06Invoice_No")
                Value = M01.Tables(0).Rows(I)("T06Pay_amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Pay Amount") = _St
                _Total = _Total + Value
                
                c_dataCustomer4.Rows.Add(newRow)

                I = I + 1
            Next
            Value = _Total
            txtTotal1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride_Pay_Invo()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer4 = CustomerDataClass.MakeDataTable_CR_PAY_Invo
        UltraGrid3.DataSource = c_dataCustomer4
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function


    Function Load_Gride1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Outstanding
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Function Load_Gride1_Data()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Double
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Try
            Sql = "select *  from View_Outstanding where   T06Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17name")
                Value = M01.Tables(0).Rows(i)("balance")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Credit Amount") = _St

                _Qty = _Qty + Value

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Qty
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Credit Amount") = _St
            c_dataCustomer1.Rows.Add(newRow1)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(1).Appearance.BackColor = Color.Gold
            ' UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            ' UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer3 = CustomerDataClass.MakeDataTable_CR_RECEIVED
        UltraGrid1.DataSource = c_dataCustomer3
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 190
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 90
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(8).Width = 70
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 90
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right




        End With
    End Function

    Function Load_Gride_PAY()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer3 = CustomerDataClass.MakeDataTable_CR_PAY
        UltraGrid1.DataSource = c_dataCustomer3
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 50
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 170
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90
            '.DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(8).Width = 70
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 90
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right



        End With
    End Function


    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Call Load_Gride1()
        Call Load_Gride1_Data()

    End Sub

    Private Sub UsingSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingSupplierToolStripMenuItem.Click

        Call Load_Gride()
        Call Load_Customer()
        Panel4.Visible = True
        cboItem.ToggleDropdown()
    End Sub

    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M17Name as [##] from M17Customer where M17Active='A' and M17Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 220
                ' .Rows.Band.Columns(1).Width = 180


            End With

            With cboCus
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 220
                ' .Rows.Band.Columns(1).Width = 180


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Call Load_Gride()
        Call Load_Gride_Data_1(Trim(cboItem.Text))
        _PrintStatus = "A"
        Panel4.Visible = False
    End Sub

    Function Load_Gride_Data_1(ByVal strCode As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Double
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim TM As TimeSpan
        Try
            Sql = "select *  from View_Outstanding_1 where BALANCE<>'0' AND M17name='" & strCode & "' and T06Com_Code='" & _Comcode & "' order by T06Cus_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer3.NewRow
                newRow("##") = False
                newRow("Customer Code") = M01.Tables(0).Rows(i)("T06Cus_Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17name")
                newRow("Invo Date") = Month(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Year(M01.Tables(0).Rows(i)("Invo_date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T06Invoice_no")

                Value = M01.Tables(0).Rows(i)("Total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Invo Amount") = _St

                Value = M01.Tables(0).Rows(i)("Paid")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Paid Amount") = _St

                Value = M01.Tables(0).Rows(i)("Balance")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Balance Amount") = _St

                _Qty = _Qty + Value

                c_dataCustomer3.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer3.NewRow
            Value = _Qty
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Balance Amount") = _St
            c_dataCustomer3.Rows.Add(newRow1)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(7).Appearance.BackColor = Color.Gold
            ' UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(7).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            ' UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Gride_Data_Pay2()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Total As Double
        Dim _Cash As Double
        Dim _Chq As Double

        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim TM As TimeSpan
        Try
            Sql = "select *  from View_CR_Received where BALANCE<>'0' AND T22Date between '" & txtH1.Text & "' and '" & txtH2.Text & "' and T22Status='A' and M17name='" & Trim(cboCus.Text) & "' order by T22Ref_no"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Cash = 0
            _Chq = 0

            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer3.NewRow
                newRow("Ref No") = M01.Tables(0).Rows(i)("T22Ref_no")
                newRow("Pay.Doc") = M01.Tables(0).Rows(i)("T22Pay_no")
                newRow("Pay Date") = Month(M01.Tables(0).Rows(i)("T22date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T22Date")) & "/" & Year(M01.Tables(0).Rows(i)("T22Date"))
                ' newRow("Customer Code") = M01.Tables(0).Rows(i)("T22Cus_Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17name")
                ' newRow("Invoice No") = M01.Tables(0).Rows(i)("T06Invoice_no")

                Value = M01.Tables(0).Rows(i)("T22Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                _Total = _Total + Value
                newRow("Pay Amount") = _St

                Value = M01.Tables(0).Rows(i)("T03Cash")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                _Cash = _Cash + Value

                newRow("Cash") = _St

                Value = M01.Tables(0).Rows(i)("T03Chq")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                _Chq = _Chq + Value
                newRow("Cheque") = _St

                ' _Qty = _Qty + Value

                c_dataCustomer3.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer3.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Pay Amount") = _St

            Value = _Cash
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Cash") = _St

            Value = _Chq
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Cheque") = _St
            c_dataCustomer3.Rows.Add(newRow1)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(6).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(6).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            UltraGrid1.Rows(_Rowcount).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            UltraGrid1.Rows(_Rowcount).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            ' UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Gride_Data_Pay1()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Total As Double
        Dim _Cash As Double
        Dim _Chq As Double

        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim TM As TimeSpan
        Try
            Sql = "select *  from View_CR_Received where BALANCE<>'0' AND T22Date between '" & txtCh1.Text & "' and '" & txtCh2.Text & "' and T22Status='A' and T22Com_Code='" & _Comcode & "' order by T22Ref_no"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Cash = 0
            _Chq = 0

            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer3.NewRow
                newRow("Ref No") = M01.Tables(0).Rows(i)("T22Ref_no")
                newRow("Pay.Doc") = M01.Tables(0).Rows(i)("T22Pay_no")
                newRow("Pay Date") = Month(M01.Tables(0).Rows(i)("T22date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T22Date")) & "/" & Year(M01.Tables(0).Rows(i)("T22Date"))
                ' newRow("Customer Code") = M01.Tables(0).Rows(i)("T22Cus_Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17name")
                ' newRow("Invoice No") = M01.Tables(0).Rows(i)("T06Invoice_no")

                Value = M01.Tables(0).Rows(i)("T22Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                _Total = _Total + Value
                newRow("Pay Amount") = _St

                Value = M01.Tables(0).Rows(i)("T03Cash")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                _Cash = _Cash + Value

                newRow("Cash") = _St

                Value = M01.Tables(0).Rows(i)("T03Chq")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                _Chq = _Chq + Value
                newRow("Cheque") = _St

                ' _Qty = _Qty + Value

                c_dataCustomer3.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer3.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Pay Amount") = _St

            Value = _Cash
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Cash") = _St

            Value = _Chq
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Cheque") = _St
            c_dataCustomer3.Rows.Add(newRow1)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(6).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(6).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            UltraGrid1.Rows(_Rowcount).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            UltraGrid1.Rows(_Rowcount).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            ' UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call Load_Gride1()
        Call Load_Gride1_Data()
        Panel4.Visible = False
        Panel5.Visible = False
        OPR2.Visible = False
        OPR10.Visible = False
        Call Clear1()
    End Sub

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        Dim i As Integer
        Dim Value As Double
        Dim _Status As Boolean
        Dim _RowIndex As Integer

        i = 0
        Try
            If _PrintStatus = "C" Or _PrintStatus = "D" Then
                OPR10.Visible = True
                _RowIndex = UltraGrid1.ActiveRow.Index
                txtRef1.Text = UltraGrid1.Rows(_RowIndex).Cells(0).Text
                txtPay.Text = UltraGrid1.Rows(_RowIndex).Cells(1).Text
                txtName1.Text = UltraGrid1.Rows(_RowIndex).Cells(3).Text
                txtPayment1.Text = UltraGrid1.Rows(_RowIndex).Cells(4).Text
                txtCash1.Text = UltraGrid1.Rows(_RowIndex).Cells(5).Text
                txtAmount1.Text = UltraGrid1.Rows(_RowIndex).Cells(6).Text

                cboStatus1.Text = "M/S"
                Call Load_Gride_Pay_Invo()
                Call Search_Records()
            Else
                _Status = False
                Value = 0
                Call Load_Gride_Invo()
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    If i = UltraGrid1.Rows.Count - 1 Then
                        Exit For
                    End If
                    If UltraGrid1.Rows(i).Cells(0).Value = True Then
                        Dim newRow As DataRow = c_dataCustomer2.NewRow
                        newRow("Invo Date") = UltraGrid1.Rows(i).Cells(3).Text
                        newRow("Invoice No") = UltraGrid1.Rows(i).Cells(4).Text
                        newRow("Tobe Paid") = UltraGrid1.Rows(i).Cells(7).Text
                        Value = Value + CDbl(UltraGrid1.Rows(i).Cells(7).Text)
                        c_dataCustomer2.Rows.Add(newRow)

                        txtCus_Code.Text = UltraGrid1.Rows(i).Cells(1).Text
                        txtCus_Name.Text = UltraGrid1.Rows(i).Cells(2).Text
                        cboStatus.Text = "M/S"
                        _Status = True
                    End If
                    i = i + 1
                Next

                txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                If _Status = True Then
                    OPR2.Visible = True
                End If
            End If
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub
    Function Clear1()
        Me.txtTotal.Text = "00.00"
        Me.txtPayment.Text = ""
        Me.txtCash.Text = ""
        Me.txtChq.Text = ""
        Me.cboBank.Text = ""
        Me.txtRef.Text = ""
        Me.txtName.Text = ""
        Me.txtCus_Name.Text = ""
        Me.txtCus_Code.Text = ""
        OPR2.Visible = False
    End Function
   
    
    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Me.txtTotal.Text = "00.00"
        Me.txtPayment.Text = ""
        Me.txtCash.Text = ""
        Me.txtChq.Text = ""
        Me.cboBank.Text = ""
        Me.txtRef.Text = ""
        Me.txtName.Text = ""
        Me.txtCus_Name.Text = ""
        Me.txtCus_Code.Text = ""
        OPR2.Visible = False
    End Sub

    Private Sub txtCash_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCash.KeyUp
        If e.KeyCode = 13 Then
            Call Calculation()
            cboBank.ToggleDropdown()
        End If
    End Sub

    Function Calculation()
        On Error Resume Next
        Dim value As Double
        value = 0
        If IsNumeric(txtCash.Text) Then
            value = CDbl(txtCash.Text)
        End If

        If IsNumeric(txtAmount.Text) Then
            value = value + CDbl(txtAmount.Text)
        End If

        txtPayment.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtPayment.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

    End Function
    Private Sub txtCash_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCash.ValueChanged
        Call Calculation()
    End Sub

   
    Private Sub cboBank_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBank.KeyUp
        If e.KeyCode = 13 Then
            txtChq.Focus()
        End If
    End Sub

    Private Sub txtChq_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChq.KeyUp
        If e.KeyCode = 13 Then
            txtDue.Focus()
        End If
    End Sub

    Private Sub txtChq_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChq.ValueChanged

    End Sub

    Private Sub txtDue_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDue.KeyUp
        If e.KeyCode = 13 Then
            txtAmount.Focus()
        End If
    End Sub

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        End If
    End Sub

    Private Sub txtAmount_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.ValueChanged
        Call Calculation()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If txtCash.Text <> "" Then
            If IsNumeric(txtCash.Text) Then
            Else
                MsgBox("Please enter the correct Cash Amount", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If
        Else
            txtCash.Text = "0"
        End If

      

        If txtName.Text <> "" Then
        Else
            txtName.Text = "-"
        End If

        If txtAmount.Text <> "" Then
            If cboBank.Text <> "" Then
            Else
                MsgBox("Please select the Bank Name", MsgBoxStyle.Information, "Informaton .......")
                Exit Sub
            End If

            If txtDue.Text <> "" Then
            Else
                MsgBox("Please select the Due Date", MsgBoxStyle.Information, "Informaton .......")
                Exit Sub
            End If

            If IsDate(txtDue.Text) Then
            Else
                MsgBox("Please select the correct due date", MsgBoxStyle.Information, "Informaton .......")
                Exit Sub
            End If

            If txtChq.Text <> "" Then
            Else
                MsgBox("Please select the chq No", MsgBoxStyle.Information, "Informaton .......")
                Exit Sub
            End If
        End If

        If cboBank.Text <> "" Then
            If txtAmount.Text <> "" Then
            Else
                MsgBox("Please select the chq amount", MsgBoxStyle.Information, "Informaton .......")
                Exit Sub
            End If

            If txtDue.Text <> "" Then
            Else
                MsgBox("Please select the Due Date", MsgBoxStyle.Information, "Informaton .......")
                Exit Sub
            End If

            If IsDate(txtDue.Text) Then
            Else
                MsgBox("Please select the correct due date", MsgBoxStyle.Information, "Informaton .......")
                Exit Sub
            End If

            If txtChq.Text <> "" Then
            Else
                MsgBox("Please select the chq No", MsgBoxStyle.Information, "Informaton .......")
                Exit Sub
            End If
        End If

        If UltraGrid2.Rows.Count > 0 Then
        Else
            MsgBox("Please add to the invoice No", MsgBoxStyle.Information, "Information .........")
            Exit Sub
        End If

        If txtAmount.Text <> "" Then
            If IsNumeric(txtAmount.Text) Then
            Else
                MsgBox("Please enter the correct Chq Amount", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If
        Else
            txtAmount.Text = "0"
        End If

        Call Save_Data()
    End Sub

    Function Save_Data()
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
        Dim _Balance As Double
        Dim _RefNo As Integer
        Dim _Pay As String
        Dim M01 As DataSet
        Try
            nvcFieldList1 = "SELECT * FROM P01Parameter WHERE P01Code='IN'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = M01.Tables(0).Rows(0)("P01LastNo")
            End If

            nvcFieldList1 = "SELECT * FROM P01Parameter WHERE P01Code='PR'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    _Pay = "PAY-SD/00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    _Pay = "PAY-SD/0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    _Pay = "PAY-SD/" & M01.Tables(0).Rows(0)("P01LastNo")
                End If
            End If

            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo +" & 1 & " WHERE P01Code='IN'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo +" & 1 & " WHERE P01Code='PR'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert T03Pay_Main(T03Ref_No,T03Trans_Type,T03Net_Amt,T03Credit,T03Cash,T03Chq,T03Status,T03Com_Code)" & _
                                                                            " values( '" & _RefNo & "','PR','" & txtPayment.Text & "','0','" & txtCash.Text & "','" & txtAmount.Text & "','A','" & _Comcode & "' )"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            If cboBank.Text <> "" And txtChq.Text <> "" And txtDue.Text <> "" Then
                nvcFieldList1 = "Insert T04Chq_Trans(T04Ref_No,T04Acc_Type,T04Chq_no,T04Bank_Name,T04Amount,T04DOR,T04Status,T04Com_Code)" & _
                                                                             " values('" & _RefNo & "', 'PR','" & Trim(txtChq.Text) & "','" & cboBank.Text & "','" & txtAmount.Text & "','" & txtDue.Text & "','A','" & _Comcode & "' )"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            nvcFieldList1 = "Insert T22CR_Received(T22Ref_No,T22Pay_No,T22Cus_Code,T22Date,T22Amount,T22Status,T22User,T22Remark,T22Com_Code)" & _
                                                                            " values( '" & _RefNo & "','" & _Pay & "','" & txtCus_Code.Text & "','" & txtDate.Text & "','" & txtPayment.Text & "','A','" & strDisname & "','" & txtName.Text & "','" & _Comcode & "' )"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            Dim _Remark As String
            _Balance = txtPayment.Text
            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                'If _Balance < 0 Then
                '    Exit For
                'End If
                _Remark = "Credit Invoice-" & Trim(UltraGrid2.Rows(i).Cells(1).Text)
                If _Balance >= CDbl(UltraGrid2.Rows(i).Cells(2).Text) Then
                    nvcFieldList1 = "Insert T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Invo,T05Com_Code,T05User,T05Status)" & _
                                                                      " values( '" & _RefNo & "','PR','" & txtDate.Text & "','" & txtCus_Code.Text & "','" & _Remark & "','0','" & Trim(UltraGrid2.Rows(i).Cells(2).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Text) & "','" & _Comcode & "','" & strDisname & "','A' )"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'OUT STANDING -----------------------------------------------------------------
                    nvcFieldList1 = "Insert t06OutStanding_Balance(T06RefNo,T06Tr_Type,T06Date,T06Cus_Code,T06Invoice_no,T06Bill_Amount,T06Pay_Amount,T06Pay_RefNo,T06Com_Code,T06Remark,T06Status)" & _
                                                                      " values( '" & _RefNo & "','PR','" & txtDate.Text & "','" & txtCus_Code.Text & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Text) & "','0','" & Trim(UltraGrid2.Rows(i).Cells(2).Text) & "','" & _Pay & "','" & _Comcode & "','" & _Remark & "','A' )"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    _Balance = _Balance - CDbl(UltraGrid2.Rows(i).Cells(2).Text)
                Else
                    nvcFieldList1 = "Insert T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Invo,T05Com_Code,T05User,T05Status)" & _
                                                                     " values( '" & _RefNo & "','PR','" & txtDate.Text & "','" & txtCus_Code.Text & "','" & _Remark & "','0','" & _Balance & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Text) & "','" & _Comcode & "','" & strDisname & "','A' )"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    'OUT STANDING -----------------------------------------------------------------
                    nvcFieldList1 = "Insert t06OutStanding_Balance(T06RefNo,T06Tr_Type,T06Date,T06Cus_Code,T06Invoice_no,T06Bill_Amount,T06Pay_Amount,T06Pay_RefNo,T06Com_Code,T06Remark,T06Status)" & _
                                                                      " values( '" & _RefNo & "','PR','" & txtDate.Text & "','" & txtCus_Code.Text & "','" & Trim(UltraGrid2.Rows(i).Cells(1).Text) & "','0','" & _Balance & "','" & _Pay & "','" & _Comcode & "','" & _Remark & "','A' )"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    _Balance = _Balance - CDbl(UltraGrid2.Rows(i).Cells(2).Text)
                End If
                i = i + 1
            Next

            MsgBox("Records upddate successfully", MsgBoxStyle.Information, "Information .........")
            transaction.Commit()
            connection.Close()
            Me.txtTotal.Text = "00.00"
            Me.txtPayment.Text = ""
            Me.txtCash.Text = ""
            Me.txtChq.Text = ""
            Me.cboBank.Text = ""
            Me.txtRef.Text = ""
            Me.txtName.Text = ""
            Me.txtCus_Name.Text = ""
            Me.txtCus_Code.Text = ""
            OPR2.Visible = False
            Call Load_Gride1()
            Call Load_Gride1_Data()
            Panel4.Visible = False

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        _PrintStatus = "C"
        txtCh1.Text = Today
        txtCh2.Text = Today
        Panel5.Visible = True
        Panel1.Visible = False
        Panel4.Visible = False
        Call Load_Gride_PAY()
    End Sub

    Private Sub UltraButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click
        Call Load_Gride_PAY()
        OPR2.Visible = False
        Call Clear1()
        Call Load_Gride_Data_Pay1()
        Panel5.Visible = False
    End Sub

    Private Sub ByCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByCustomerToolStripMenuItem.Click
        _PrintStatus = "D"
        txtH1.Text = Today
        txtH2.Text = Today
        Panel5.Visible = False
        Panel1.Visible = True
        Panel4.Visible = False
        Call Load_Gride_PAY()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Load_Gride_PAY()
        OPR2.Visible = False
        Call Clear1()
        Call Load_Gride_Data_Pay2()
        Panel1.Visible = False
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        OPR10.Visible = False
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim A As String
        Try
            A = MsgBox("Are you sure you want to cancel this Transaction", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Cancel Transaction ......")
            If A = vbYes Then
                nvcFieldList1 = "UPDATE T03Pay_Main SET t03STATUS='I' WHERE T03Ref_No='" & txtRef1.Text & "' AND T03Trans_Type='PR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T22CR_Received SET T22Status='I' WHERE T22Ref_No='" & txtRef1.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE t06OutStanding_Balance SET T06Status='I' WHERE T06RefNo='" & txtRef1.Text & "' AND T06Tr_Type='PR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T05Acc_Trans SET T05Status='I' WHERE T05Ref_No='" & txtRef1.Text & "' AND T05Acc_Type='PR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T04Chq_Trans SET T04Status='I' WHERE T04Ref_No='" & txtRef1.Text & "' AND T04Acc_Type='PR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records Canceled successfully", MsgBoxStyle.Information, "Information .......")
            End If
            transaction.Commit()
            connection.Close()
            OPR10.Visible = False
            Call Load_Gride1()
            Call Load_Gride1_Data()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub
End Class