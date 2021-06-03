Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptSales
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Cashier As String
    Dim _Comcode As String
    Dim _Catcode As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Call Load_Gride()
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        Panel6.Visible = False
    End Sub
    Function Clear_Text()
        Call Load_Gride()
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Function

    Function Load_Gride_Cashier_Summery()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer3 = CustomerDataClass.MakeDataTable_Cashier_Summery
        UltraGrid1.DataSource = c_dataCustomer3
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(7).Width = 90
            '.DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(8).Width = 90
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

        End With
    End Function


    Function Load_Gride_Customer_Summery()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer3 = CustomerDataClass.MakeDataTable_Customer_Summery
        UltraGrid1.DataSource = c_dataCustomer3
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 220
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(5).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(6).Width = 90
            '.DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            ''.DisplayLayout.Bands(0).Columns(7).Width = 90
            ''.DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            ''.DisplayLayout.Bands(0).Columns(8).Width = 90
            ''.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            ''.DisplayLayout.Bands(0).Columns(9).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

        End With
    End Function

    Function Load_Gride_Daily_Summery()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer3 = CustomerDataClass.MakeDataTable_Daily_Summery
        UltraGrid1.DataSource = c_dataCustomer3
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(7).Width = 90
            '.DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(8).Width = 90
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(9).Width = 70
            '.DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

        End With
    End Function

    Function Load_Gride_Cashier()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Cashier
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
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
            .DisplayLayout.Bands(0).Columns(8).Width = 90
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Width = 90
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(10).Width = 70
            .DisplayLayout.Bands(0).Columns(10).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Sales
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 130
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 70
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Width = 70
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Width = 90
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        End With
    End Function

    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M17Name as [##] from M17Customer where M17Active='A' and M17Com_Code='" & _Comcode & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
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

    Function Load_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Code as [##] from M03Item_Master where m03Status='A' and M03Com_Code='" & _Comcode & "' order by M03Item_Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
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

    Function Load_Cashier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Employee_Name as [##] from M01Employee_Master where M01Status='A'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCashier
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

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer3 = CustomerDataClass.MakeDataTable_SalesTR
        UltraGrid2.DataSource = c_dataCustomer3
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            '.DisplayLayout.Bands(0).Columns(5).Width = 70
            '.DisplayLayout.Bands(0).Columns(6).Width = 80
            '.DisplayLayout.Bands(0).Columns(7).Width = 90
            ' .DisplayLayout.Bands(0).Columns(8).Width = 90

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit


            '.DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function


    Private Sub frmrptSales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
        txtDate3.Text = Today
        txtDate4.Text = Today
        txtA1.Text = Today
        txtA2.Text = Today
        Call Load_Supplier()
        txtB1.Text = Today
        txtB2.Text = Today
        Call Load_Category()
        txtC1.Text = Today
        txtC2.Text = Today
        Call Load_Item()

        txtEntry.ReadOnly = True
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate.ReadOnly = True
        txtDate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCount.ReadOnly = True
        txtCount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtNett.ReadOnly = True
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtCom_Invoice.ReadOnly = True
    End Sub

    Function Load_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M02Cat_Name as [##] from M02Category where M02Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
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

    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Loc_Code='" & _Comcode & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupplier
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

  

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _PrintStatus = "A1" Then
            _From = txtDate3.Text
            _To = txtDate4.Text
            Call Load_Gride()
            ' _PrintStatus = "A1"
            Call Load_Data_A1()
            Panel3.Visible = False
        ElseIf _PrintStatus = "A2" Then
            _From = txtDate3.Text
            _To = txtDate4.Text
            Call Load_Gride_Cashier()
            Call Load_Data_TotalSales()
            Panel3.Visible = False
        ElseIf _PrintStatus = "A3" Then
            _From = txtDate3.Text
            _To = txtDate4.Text
            ' Call Load_Gride_Daily_Summery()
            ' Call Load_Gride_Daily_Summery()
            Call Load_Data_Daily_Summery()
            Panel3.Visible = False
        End If
    End Sub

    Function Load_Data_A1()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _proft As Double
        Dim _Total As Double
        Try

            _proft = 0
            _Total = 0

            Sql = "select *  from View_Sales where t01date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and T01Com_Code='" & _Comcode & "' and M03Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Customer") = M01.Tables(0).Rows(i)("M17Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
                _proft = _proft + M01.Tables(0).Rows(i)("Profit")
                Value = M01.Tables(0).Rows(i)("Profit")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Profit") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                _Total = _Total + M01.Tables(0).Rows(i)("total")
                Value = M01.Tables(0).Rows(i)("total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                newRow("Terminal") = M01.Tables(0).Rows(i)("T01Terminal")
                newRow("User") = M01.Tables(0).Rows(i)("T01User")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _proft
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Profit") = _St
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_A2()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _profit As Double
        Dim _total As Double
        Try
            Sql = "select *  from View_Sales where t01date between '" & txtA1.Text & "' and '" & txtA2.Text & "' and M09Name='" & Trim(cboSupplier.Text) & "' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            _profit = 0
            _total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Customer") = M01.Tables(0).Rows(i)("M17Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
                _profit = _profit + M01.Tables(0).Rows(i)("Profit")
                Value = M01.Tables(0).Rows(i)("Profit")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Profit") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                _total = _total + M01.Tables(0).Rows(i)("total")
                Value = M01.Tables(0).Rows(i)("total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                newRow("Terminal") = M01.Tables(0).Rows(i)("T01Terminal")
                newRow("User") = M01.Tables(0).Rows(i)("T01User")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _profit
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Profit") = _St
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_A3()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _profit As Double
        Dim _total As Double

        _profit = 0
        _total = 0
        Try
            Sql = "select *  from View_Sales where t01date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and M02Cat_Name='" & Trim(cboCategory.Text) & "' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
         
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Customer") = M01.Tables(0).Rows(i)("M17Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
                _profit = _profit + M01.Tables(0).Rows(i)("Profit")
                Value = M01.Tables(0).Rows(i)("Profit")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Profit") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                _total = _total + M01.Tables(0).Rows(i)("total")
                Value = M01.Tables(0).Rows(i)("total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                newRow("Terminal") = M01.Tables(0).Rows(i)("T01Terminal")
                newRow("User") = M01.Tables(0).Rows(i)("T01User")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _profit
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Profit") = _St
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_A5(ByVal strInv0 As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String


        Try

            Sql = "select *  from View_Purchasing where  T01Ref_No='" & strInv0 & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Customer") = M01.Tables(0).Rows(i)("M17Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Free Issue") = _St
                Value = M01.Tables(0).Rows(i)("T02Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St

                Value = M01.Tables(0).Rows(i)("total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_A4()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _profit As Double
        Dim _total As Double
        Try
            Sql = "select *  from View_Sales where t01date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and M03Item_Code='" & Trim(cboItem.Text) & "' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            _profit = 0
            _total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Customer") = M01.Tables(0).Rows(i)("M17Name")
                newRow("Item Code") = M01.Tables(0).Rows(i)("m03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
                _profit = _profit + M01.Tables(0).Rows(i)("Profit")
                Value = M01.Tables(0).Rows(i)("Profit")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Profit") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                _Total = _Total + M01.Tables(0).Rows(i)("total")
                Value = M01.Tables(0).Rows(i)("total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                newRow("Terminal") = M01.Tables(0).Rows(i)("T01Terminal")
                newRow("User") = M01.Tables(0).Rows(i)("T01User")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _profit
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Profit") = _St
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Blue
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String

        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"

            If _PrintStatus = "CH" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Sales_Cashier.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_CashierSummery.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_CashierSummery.t01User}='" & cboCashier.Text & "' and {View_CashierSummery.Location} ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "CH1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Sales_Cashier_Sum.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Cashier_Sales_Summery.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_Cashier_Sales_Summery.t01User}='" & cboCashier.Text & "' and {View_Cashier_Sales_Summery.Location} ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A3" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Sales_Daily_Sum.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Daily_Sales_Summery.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  and {View_Daily_Sales_Summery.Location}  ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()

            ElseIf _PrintStatus = "A2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Sales_Cashier.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_CashierSummery.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_CashierSummery.Location} ='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Sales.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Sales.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  and {View_Sales.T01Com_Code} ='" & _Comcode & "' and {View_Sales.M03Com_Code} ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A3" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Sales.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Sales.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  and {View_Sales.T01Com_Code} ='" & _Comcode & "' and {View_Sales.M02Cat_Name}='" & _Catcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A4" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Sales.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Sales.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & "  and {View_Sales.T01Com_Code} ='" & _Comcode & "' and {View_Sales.M03Item_Code}='" & _Itemcode & "'"
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

    Private Sub UsingSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingSupplierToolStripMenuItem.Click
        'Panel3.Visible = False
        'Panel1.Visible = False
        'Panel2.Visible = False
        'Panel4.Visible = False
        'Panel5.Visible = True
        'Call Load_Gride_Cashier()
        'Call Load_Cashier()
        'txtCh1.Text = Today
        'txtCh2.Text = Today
        'cboCashier.Text = ""
    End Sub

    Private Sub UsingCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingCategoryToolStripMenuItem.Click
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = True
        Panel4.Visible = False
        Panel5.Visible = False
        _PrintStatus = "A3"
        txtB1.Text = Today
        txtB2.Text = Today
    End Sub

    Private Sub UsingItemNameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingItemNameToolStripMenuItem.Click
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = True
        Panel5.Visible = False
        ' _PrintStatus = "B2"
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Call Load_Gride()
        _PrintStatus = "A4"
        Call Load_Data_A4()
        _Itemcode = Trim(cboItem.Text)
        Panel4.Visible = False
    End Sub

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        'On Error Resume Next
        'Dim _RowIndex As Integer
        'Dim strInv0 As String
        '_RowIndex = UltraGrid1.ActiveRow.Index
        'strInv0 = UltraGrid1.Rows(_RowIndex).Cells(0).Text
        'Call Load_Gride()
        'Call Load_Data_A5(strInv0)
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub UltraButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click
        If _PrintStatus = "CH" Then
            Call Load_Gride_Cashier()
            Call Load_Data_CashierPC()
            _From = txtCh1.Text
            _To = txtCh2.Text
            _Cashier = cboCashier.Text
            Panel5.Visible = False
        ElseIf _PrintStatus = "CH1" Then
            Call Load_Gride_Cashier_Summery()
            Call Load_Data_Cashier_Summery()
            _From = txtCh1.Text
            _To = txtCh2.Text
            _Cashier = cboCashier.Text
            Panel5.Visible = False
        End If
    End Sub

    Function Load_Data_Cashier_Summery()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _St As String
        Dim _Net As Double
        Dim _Cash As Double
        Dim _VISA As Double
        Dim _Master As Double
        Dim _Amex As Double
        Dim Value As Double
        Dim M02 As DataSet
        Dim _LastRow As Integer

        Try
            i = 0
            _Cash = 0
            _VISA = 0
            _Master = 0
            _Amex = 0
            _Net = 0
            Sql = "select * from View_Cashier_Sales_Summery where T01Date between '" & txtCh1.Text & "' and '" & txtCh2.Text & "' and T01User='" & Trim(cboCashier.Text) & "' and Location='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows

                Dim _Unet As Double
                Dim _UCard As Double

                _UCard = 0
                _Unet = 0
                Dim newRow As DataRow = c_dataCustomer3.NewRow
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                'newRow("MC No") = M01.Tables(0).Rows(i)("T01Terminal")
                Value = M01.Tables(0).Rows(i)("Net")
                _Unet = M01.Tables(0).Rows(i)("Net")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Net = _Net + Value

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("Cash")
                ' _Cash = M01.Tables(0).Rows(i)("Cash")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Cash = _Cash + Value

                newRow("Cash Amount") = _St

                Value = M01.Tables(0).Rows(i)("visa")
                _UCard = _UCard + M01.Tables(0).Rows(i)("visa")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _VISA = _VISA + Value
                newRow("VISA") = _St

                Value = M01.Tables(0).Rows(i)("Masterc")
                _UCard = _UCard + M01.Tables(0).Rows(i)("Masterc")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Master = _Master + Value
                newRow("Master") = _St

                Value = M01.Tables(0).Rows(i)("Amex")
                _UCard = _UCard + M01.Tables(0).Rows(i)("Amex")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Amex = _Amex + Value
                newRow("AMEX") = _St

                _Unet = _Unet - _UCard
                '_Cash = _Cash + _Unet
                'Value = _Unet
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                '' _Amex = _Amex + Value
                'newRow("Cash") = _St
                newRow("User") = M01.Tables(0).Rows(i)("T01User")

                c_dataCustomer3.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer3.NewRow
            Value = _Net
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            Value = _VISA
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("VISA") = _St

            Value = _Master
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Master") = _St

            Value = _Amex
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("AMEX") = _St

            Value = _Cash
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Cash Amount") = _St

            c_dataCustomer3.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_LastRow).Cells(1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(2).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(3).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            UltraGrid1.Rows(_LastRow).Cells(1).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(2).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(3).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_Customer_Summery()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _St As String
        Dim _Net As Double
        Dim _Cash As Double
        Dim _VISA As Double
        Dim _Master As Double
        Dim _Amex As Double
        Dim Value As Double
        Dim M02 As DataSet
        Dim _LastRow As Integer

        Try
            i = 0
            _Cash = 0
            _VISA = 0
            _Master = 0
            _Amex = 0
            _Net = 0
            Sql = "select * from View_T01Sales where T01Date between '" & txtCu1.Text & "' and '" & txtCu2.Text & "' and M17Name='" & Trim(cboCustomer.Text) & "' and T01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows

                Dim _Unet As Double
                Dim _UCard As Double

                _UCard = 0
                _Unet = 0
                Dim newRow As DataRow = c_dataCustomer3.NewRow

                newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                'newRow("MC No") = M01.Tables(0).Rows(i)("T01Terminal")
                Value = M01.Tables(0).Rows(i)("Net")
                _Unet = M01.Tables(0).Rows(i)("Net")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Net = _Net + Value

                newRow("Invo.Value") = _St

                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17Name")

                c_dataCustomer3.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer3.NewRow
            Value = _Net
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Invo.Value") = _St

            'Value = _VISA
            '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            'newRow1("VISA") = _St

            'Value = _Master
            '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            'newRow1("Master") = _St

            'Value = _Amex
            '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            'newRow1("AMEX") = _St

            'Value = _Cash
            '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            'newRow1("Cash Amount") = _St

            c_dataCustomer3.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count - 1
            'UltraGrid1.Rows(_LastRow).Cells(1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid1.Rows(_LastRow).Cells(2).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(3).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            'UltraGrid1.Rows(_LastRow).Cells(1).Appearance.BackColor = Color.Gold
            'UltraGrid1.Rows(_LastRow).Cells(2).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(3).Appearance.BackColor = Color.Gold
            'UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            'UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Data_Daily_Summery()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _St As String
        Dim _Net As Double
        Dim _Cash As Double
        Dim _VISA As Double
        Dim _Master As Double
        Dim _Amex As Double
        Dim Value As Double
        Dim M02 As DataSet
        Dim _LastRow As Integer
        Dim _Credit As Double

        Try
            i = 0
            _Cash = 0
            _VISA = 0
            _Master = 0
            _Amex = 0
            _Net = 0
            _Credit = 0

            Sql = "select * from View_Daily_Sales_Summery where T01Date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "'  and Location='" & _Comcode & "' order by T01Date"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows

                Dim _Unet As Double
                Dim _UCard As Double

                _UCard = 0
                _Unet = 0
                Dim newRow As DataRow = c_dataCustomer3.NewRow
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                'newRow("MC No") = M01.Tables(0).Rows(i)("T01Terminal")
                Value = M01.Tables(0).Rows(i)("Net")
                _Unet = M01.Tables(0).Rows(i)("Net")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Net = _Net + Value

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("Cash")
                ' _Cash = M01.Tables(0).Rows(i)("Cash")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Cash = _Cash + Value

                newRow("Cash Amount") = _St

                Value = M01.Tables(0).Rows(i)("credit")
                _Credit = _Credit + Value
                ' _Cash = M01.Tables(0).Rows(i)("Cash")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                '_Credit = _Credit + Value

                newRow("Credit Amount") = _St

                Value = M01.Tables(0).Rows(i)("visa")
                _UCard = _UCard + M01.Tables(0).Rows(i)("visa")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _VISA = _VISA + Value
                newRow("VISA") = _St

                Value = M01.Tables(0).Rows(i)("Masterc")
                _UCard = _UCard + M01.Tables(0).Rows(i)("Masterc")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Master = _Master + Value
                newRow("Master") = _St

                Value = M01.Tables(0).Rows(i)("Amex")
                _UCard = _UCard + M01.Tables(0).Rows(i)("Amex")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Amex = _Amex + Value
                newRow("AMEX") = _St

                _Unet = _Unet - _UCard
                '_Cash = _Cash + _Unet
                'Value = _Unet
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                '' _Amex = _Amex + Value
                'newRow("Cash") = _St
                'newRow("User") = M01.Tables(0).Rows(i)("T01User")

                c_dataCustomer3.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer3.NewRow
            Value = _Net
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            Value = _VISA
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("VISA") = _St

            Value = _Master
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Master") = _St

            Value = _Amex
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("AMEX") = _St

            Value = _Cash
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Cash Amount") = _St

            Value = _Credit
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Credit Amount") = _St

            c_dataCustomer3.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_LastRow).Cells(1).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(2).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(3).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(6).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            UltraGrid1.Rows(_LastRow).Cells(1).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(2).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(3).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(6).Appearance.BackColor = Color.Gold

            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Data_CashierPC()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _St As String
        Dim _Net As Double
        Dim _Cash As Double
        Dim _VISA As Double
        Dim _Master As Double
        Dim _Amex As Double
        Dim Value As Double
        Dim M02 As DataSet
        Dim _LastRow As Integer

        Try
            i = 0
            _Cash = 0
            _VISA = 0
            _Master = 0
            _Amex = 0
            _Net = 0
            Sql = "select * from View_CashierSummery where T01Date between '" & txtCh1.Text & "' and '" & txtCh2.Text & "' and T01User='" & Trim(cboCashier.Text) & "' and Location='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows

                Dim _Unet As Double
                Dim _UCard As Double

                _UCard = 0
                _Unet = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("MC No") = M01.Tables(0).Rows(i)("T01Terminal")
                Value = M01.Tables(0).Rows(i)("Net")
                _Unet = M01.Tables(0).Rows(i)("Net")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Net = _Net + Value

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("visa")
                _UCard = _UCard + M01.Tables(0).Rows(i)("visa")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _VISA = _VISA + Value
                newRow("VISA") = _St

                Value = M01.Tables(0).Rows(i)("Masterc")
                _UCard = _UCard + M01.Tables(0).Rows(i)("Masterc")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Master = _Master + Value
                newRow("Master") = _St

                Value = M01.Tables(0).Rows(i)("Amex")
                _UCard = _UCard + M01.Tables(0).Rows(i)("Amex")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Amex = _Amex + Value
                newRow("AMEX") = _St

                _Unet = _Unet - _UCard
                _Cash = _Cash + _Unet
                Value = _Unet
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                ' _Amex = _Amex + Value
                newRow("Cash") = _St
                newRow("User") = M01.Tables(0).Rows(i)("T01User")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Net
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            Value = _VISA
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("VISA") = _St

            Value = _Master
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Master") = _St

            Value = _Amex
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("AMEX") = _St

            Value = _Cash
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Cash") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(6).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(7).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            UltraGrid1.Rows(_LastRow).Cells(7).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(6).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Data_TotalSales()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _St As String
        Dim _Net As Double
        Dim _Cash As Double
        Dim _VISA As Double
        Dim _Master As Double
        Dim _Amex As Double
        Dim Value As Double
        Dim M02 As DataSet
        Dim _LastRow As Integer
        Dim _Credit As Double

        Try
            i = 0
            _Cash = 0
            _VISA = 0
            _Master = 0
            _Amex = 0
            _Net = 0
            _Credit = 0

            Sql = "select * from View_CashierSummery where T01Date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "'  and Location='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows

                Dim _Unet As Double
                Dim _UCard As Double

                _UCard = 0
                _Unet = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("MC No") = M01.Tables(0).Rows(i)("T01Terminal")
                Value = M01.Tables(0).Rows(i)("Net")
                _Unet = M01.Tables(0).Rows(i)("Net")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Net = _Net + Value

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("Credit")
                ' _UCard = _UCard + M01.Tables(0).Rows(i)("credit")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Credit = _Credit + Value
                newRow("Credit") = _St


                Value = M01.Tables(0).Rows(i)("visa")
                _UCard = _UCard + M01.Tables(0).Rows(i)("visa")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _VISA = _VISA + Value
                newRow("VISA") = _St

                Value = M01.Tables(0).Rows(i)("Masterc")
                _UCard = _UCard + M01.Tables(0).Rows(i)("Masterc")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Master = _Master + Value
                newRow("Master") = _St

                Value = M01.Tables(0).Rows(i)("Amex")
                _UCard = _UCard + M01.Tables(0).Rows(i)("Amex")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Amex = _Amex + Value
                newRow("AMEX") = _St

           
                _Unet = _Unet - _UCard
                _Unet = _Unet - _Credit
                _Cash = _Cash + _Unet
                Value = M01.Tables(0).Rows(i)("cash")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                ' _Amex = _Amex + Value
                newRow("Cash") = _St
                newRow("User") = M01.Tables(0).Rows(i)("T01User")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Net
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Net Amount") = _St

            Value = _VISA
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("VISA") = _St

            Value = _Master
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Master") = _St

            Value = _Amex
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("AMEX") = _St

            Value = _Net - (_Credit + _VISA + _Amex + _Master)
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Cash") = _St

            Value = _Credit
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Credit") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(6).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(7).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            UltraGrid1.Rows(_LastRow).Cells(7).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(6).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(4).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Private Sub DetailleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetailleToolStripMenuItem.Click
        Panel3.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        _PrintStatus = "A1"
    End Sub

    Private Sub SummeryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SummeryToolStripMenuItem.Click
        Panel3.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        _PrintStatus = "A2"
        Call Load_Gride_Cashier()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        _From = txtB1.Text
        _To = txtB2.Text
        _Catcode = Trim(cboCategory.Text)
        Call Load_Gride()
        ' _PrintStatus = "A1"
        Call Load_Data_A3()
        Panel2.Visible = False
    End Sub

    Private Sub DetailesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetailesToolStripMenuItem.Click
        Call Clear_Text()
        _PrintStatus = "CH"
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = True
        UltraGrid1.Refresh()
        Call Load_Gride_Cashier()
        Call Load_Cashier()
        txtCh1.Text = Today
        txtCh2.Text = Today
        cboCashier.Text = ""
    End Sub

    Private Sub SummeryToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SummeryToolStripMenuItem1.Click
        Call Clear_Text()
        _PrintStatus = "CH1"
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = True
        Call Load_Gride_Cashier_Summery()
        Call Load_Cashier()
        txtCh1.Text = Today
        txtCh2.Text = Today
        cboCashier.Text = ""
    End Sub

    Function Search_RecordsUsing_Entry()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet
        Dim i As Integer
        Dim Value As Double
        Dim _RefNo As Integer

        Try
            Sql = "select * from T01Transaction_Header inner join T02Transaction_Flutter  on T01Ref_No=T02Ref_No  inner join M03Item_Master on T02Item_Code=M03Item_Code  inner join M04Location on M04Loc_Code=T01Fromloc_code where T01Invoice_no='" & Trim(txtEntry.Text) & "' and T01Trans_Type='DR' and T01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboLocation.Text = Trim(M01.Tables(0).Rows(0)("M04Loc_name"))
                '  cboTo.Text = Trim(M01.Tables(0).Rows(0)("M04Loc_Name"))
                txtCom_Invoice.Text = Trim(M01.Tables(0).Rows(0)("T01User"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("T01Date"))
                '  txtEntry.Text = Trim(M01.Tables(0).Rows(0)("T01Grn_No"))
                ' txtRemark.Text = Trim(M01.Tables(0).Rows(0)("T01Remark"))
                _RefNo = Trim(M01.Tables(0).Rows(0)("T01Ref_No"))

                Value = Trim(M01.Tables(0).Rows(0)("T01Net_Amount"))
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))



                txtCount.Text = M01.Tables(0).Rows.Count

                Dim _St As String
                Call Load_Gride2()

                i = 0
                For Each DTRow2 As DataRow In M01.Tables(0).Rows

                    Dim newRow As DataRow = c_dataCustomer3.NewRow
                    newRow("Item Code") = Trim(M01.Tables(0).Rows(i)("M03Item_Code"))
                    newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("M03Item_Name"))
                    ' newRow("Cost Price") = _St
                    newRow("Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Retail_Price"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Retail Price") = _St
                    newRow("Total") = _St

                    c_dataCustomer3.Rows.Add(newRow)


                    i = i + 1
                Next
              

            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub DateWiseSummeryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateWiseSummeryToolStripMenuItem.Click
        Panel3.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        _PrintStatus = "A3"
        Call Load_Gride_Daily_Summery()
    End Sub

    Private Sub UltraGrid1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid1.MouseDoubleClick
        On Error Resume Next
        Dim _RowIndex As Integer
        Panel6.Visible = True
        If _PrintStatus = "A3" Then
        Else
            _RowIndex = UltraGrid1.ActiveRow.Index
            txtEntry.Text = UltraGrid1.Rows(_RowIndex).Cells(2).Text
            Call Load_Gride2()
            Call Search_RecordsUsing_Entry()
        End If


       

    End Sub

    Private Sub AllRootsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllRootsToolStripMenuItem.Click
        Panel3.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        _PrintStatus = "O1"
    End Sub

    Private Sub CustomerWiseReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerWiseReportToolStripMenuItem.Click
        _PrintStatus = "CU1"
        Panel7.Visible = True
        Panel3.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        txtCu1.Text = Today
        txtCu2.Text = Today
        Call Load_Customer()
    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        If _PrintStatus = "CU1" Then
            Call Load_Gride_Customer_Summery()
            Call Load_Data_Customer_Summery()
            _From = txtCu1.Text
            _To = txtCu2.Text
            _Cashier = cboCustomer.Text
            Panel5.Visible = False
        End If
    End Sub
End Class