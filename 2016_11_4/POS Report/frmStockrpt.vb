Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmStockrpt
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Comcode As String
    Dim _Catcode As String
    Dim _Suplier As String

    Function Load_Gride_Movement()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_StockMovement
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 190
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
            '.DisplayLayout.Bands(0).Columns(8).Width = 90
            '.DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ' .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_StockItemsrpt
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 140
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 100
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        End With
    End Function

    Function Load_Gride1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_StockItemsrpt2
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 140
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 100
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 190
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).Width = 90
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Width = 90
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Private Sub frmStockrpt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
        Call Load_Category()
        Call Load_Supplier()
        txtDate3.Text = Today
        txtDate4.Text = Today
        Call Load_Item()
    End Sub

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

    Function Load_Grid_Category3(ByVal strCat As String)
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
            Sql = "select *  from View_StockBalance where Category='" & strCat & "' and Reorder>qty and S01Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S01Item_code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")
                Value = M01.Tables(0).Rows(i)("Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
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

    Function Load_Grid_Category1(ByVal strCat As String)
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
            Sql = "select *  from View_StockBalance1 where Category='" & strCat & "' and S04Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S04Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")
                Value = M01.Tables(0).Rows(i)("Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
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

    Function Load_Grid_Supp1(ByVal strCat As String)
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
            Sql = "select *  from View_StockBalance where Supplier='" & strCat & "'  order by Category "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S01Item_code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")
                Value = M01.Tables(0).Rows(i)("Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
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

    Function Load_Grid_All1()
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
            Sql = "select *  from View_StockBalance where S04Com_Code='" & _Comcode & "' order by Category "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S01Item_code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")
                Value = M01.Tables(0).Rows(i)("Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
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

    Function Load_Grid_All2()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Cost As Double
        Dim _Retail As Double

        Try
            Sql = "select *  from View_StockBalance1 where S04Com_Code='" & _Comcode & "' order by Category,S04Item_Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            _Cost = 0
            _Retail = 0

            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S04Item_code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")

                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                Value = M01.Tables(0).Rows(i)("S04cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost") = _St
                ' c_dataCustomer1.Rows.Add(newRow)
                Value = M01.Tables(0).Rows(i)("Rate")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail") = _St
                Value = M01.Tables(0).Rows(i)("Qty")
                newRow("Qty") = CInt(M01.Tables(0).Rows(i)("Qty"))
                Value = M01.Tables(0).Rows(i)("Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Cost = _Cost + Value


                newRow("Total Cost") = _St
                ' c_dataCustomer1.Rows.Add(newRow)
                Value = M01.Tables(0).Rows(i)("Retail")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Retail = _Retail + Value
                newRow("Total Retail") = _St
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Category") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            _St = (_Cost.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Cost))
            newRow2("Total Cost") = _St

            _St = (_Retail.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Retail))
            newRow2("Total Retail") = _St

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(7).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount).Cells(7).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Category2(ByVal strCat As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Cost As Double
        Dim _Retail As Double

        Try
            Sql = "select *  from View_StockBalance1 where Category='" & strCat & "' and S04Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            _Cost = 0
            _Retail = 0

            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S04Item_code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")

                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                Value = M01.Tables(0).Rows(i)("S04cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost") = _St
                ' c_dataCustomer1.Rows.Add(newRow)
                Value = M01.Tables(0).Rows(i)("Rate")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail") = _St
                Value = M01.Tables(0).Rows(i)("Qty")
                newRow("Qty") = CInt(M01.Tables(0).Rows(i)("Qty"))
                Value = M01.Tables(0).Rows(i)("Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Cost = _Cost + Value


                newRow("Total Cost") = _St
                ' c_dataCustomer1.Rows.Add(newRow)
                Value = M01.Tables(0).Rows(i)("Retail")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Retail = _Retail + Value
                newRow("Total Retail") = _St
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Category") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            _St = (_Cost.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Cost))
            newRow2("Total Cost") = _St

            _St = (_Retail.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Retail))
            newRow2("Total Retail") = _St

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(7).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount).Cells(7).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            ' con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Supp2(ByVal strCat As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Cost As Double
        Dim _Retail As Double

        Try
            Sql = "select *  from View_StockBalance where Supplier='" & strCat & "' and S01Com_Code='" & _Comcode & "' order by Category "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            _Cost = 0
            _Retail = 0

            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S01Item_code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")
                Value = M01.Tables(0).Rows(i)("Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
                Value = M01.Tables(0).Rows(i)("Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Cost") = _St
                ' c_dataCustomer1.Rows.Add(newRow)
                Value = M01.Tables(0).Rows(i)("Retail")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Retail") = _St
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            newRow1("Category") = ""
            c_dataCustomer1.Rows.Add(newRow1)

            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
            _St = (_Cost.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Cost))
            newRow2("Cost") = _St

            _St = (_Retail.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Retail))
            newRow2("Retail") = _St

            c_dataCustomer1.Rows.Add(newRow2)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(5).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(6).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid1.Rows(_Rowcount).Cells(6).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Ng()
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
            Sql = "select *  from View_StockBalance where Qty<0 and S01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S01Item_code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")
                Value = M01.Tables(0).Rows(i)("Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
                Value = M01.Tables(0).Rows(i)("Cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
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

    Function Load_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M02Cat_Name as [##] from M02Category where M02Com_Code='" & _Comcode & "'  "
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
            Sql = "select M09Name as [##] from M09Supplier where M09Loc_Code='" & _Comcode & "' "
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

    Function Load_Grid_Supp3(ByVal strCat As String)
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
            Sql = "select *  from View_StockBalance where Supplier='" & strCat & "' and Reorder>Qty and S01Com_Code='" & _Comcode & "' order by Category "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("Category") = M01.Tables(0).Rows(i)("Category")
                newRow("Supplier") = M01.Tables(0).Rows(i)("Supplier")
                newRow("Item Code") = M01.Tables(0).Rows(i)("S01Item_code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("Item")
                Value = M01.Tables(0).Rows(i)("Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Qty") = _St
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

    Private Sub WithoutValuationToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithoutValuationToolStripMenuItem2.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Call Load_Gride()
        Call Load_Grid_All1()
        _PrintStatus = "A2"
    End Sub

    Private Sub WithValuationToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithValuationToolStripMenuItem1.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Call Load_Gride1()
        Call Load_Grid_All2()
        _PrintStatus = "A1"
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Call Load_Gride()
        Call Load_Grid_Ng()
        _PrintStatus = "F1"
    End Sub

    Private Sub WithValuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithValuToolStripMenuItem.Click
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        _PrintStatus = "B2"
        Call Load_Gride1()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        If _PrintStatus = "B1" Then
            If cboCategory.Text <> "" Then
                Call Load_Grid_Category1(Trim(cboCategory.Text))
                _Catcode = Trim(cboCategory.Text)
                Panel1.Visible = False
                cboCategory.Text = ""
            Else
                MsgBox("Please select the Category", MsgBoxStyle.Information, "Information ......")
                Exit Sub
            End If
        ElseIf _PrintStatus = "B2" Then
            If cboCategory.Text <> "" Then
                Call Load_Grid_Category2(Trim(cboCategory.Text))
                _Catcode = Trim(cboCategory.Text)
                Panel1.Visible = False
                cboCategory.Text = ""
            Else
                MsgBox("Please select the Category", MsgBoxStyle.Information, "Information ......")
                Exit Sub
            End If
        ElseIf _PrintStatus = "D1" Then
            If cboCategory.Text <> "" Then
                Call Load_Grid_Category3(Trim(cboCategory.Text))
                _Catcode = Trim(cboCategory.Text)
                Panel1.Visible = False
                cboCategory.Text = ""
            Else
                MsgBox("Please select the Category", MsgBoxStyle.Information, "Information ......")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub ExitToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem2.Click
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        cboCategory.Text = ""
        cboSupplier.Text = ""
        cboItem.Text = ""
        Call Load_Gride()

    End Sub

    Private Sub WithoutValuationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithoutValuationToolStripMenuItem.Click
        _PrintStatus = "B1"
        Call Load_Gride()
        Panel1.Visible = True
        Panel3.Visible = False
        Panel2.Visible = False
    End Sub

    Private Sub WithValuationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithValuationToolStripMenuItem.Click
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        _PrintStatus = "C1"
        Call Load_Gride1()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _PrintStatus = "C2" Then
            If cboSupplier.Text <> "" Then
                Call Load_Grid_Supp1(Trim(cboSupplier.Text))
                _Suplier = Trim(cboSupplier.Text)
                Panel2.Visible = False
                cboSupplier.Text = ""
            Else
                MsgBox("Please select the Supplier Name", MsgBoxStyle.Information, "Information .....")
            End If
        ElseIf _PrintStatus = "C1" Then
            If cboSupplier.Text <> "" Then
                Call Load_Grid_Supp2(Trim(cboSupplier.Text))
                _Suplier = Trim(cboSupplier.Text)
                Panel2.Visible = False
                cboSupplier.Text = ""
            Else
                MsgBox("Please select the Supplier Name", MsgBoxStyle.Information, "Information .....")
            End If
        ElseIf _PrintStatus = "D2" Then
            If cboSupplier.Text <> "" Then
                Call Load_Grid_Supp3(Trim(cboSupplier.Text))
                _Suplier = Trim(cboSupplier.Text)
                Panel2.Visible = False
                cboSupplier.Text = ""
            Else
                MsgBox("Please select the Supplier Name", MsgBoxStyle.Information, "Information .....")
            End If
        End If
    End Sub

    Private Sub WithoutValuationToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithoutValuationToolStripMenuItem1.Click
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        _PrintStatus = "C2"
        Call Load_Gride()
    End Sub

    Private Sub CategoryWiseToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CategoryWiseToolStripMenuItem1.Click
        Panel1.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        _PrintStatus = "D1"
        Call Load_Gride()
    End Sub

    Private Sub SupplierWiseToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierWiseToolStripMenuItem1.Click
        Panel1.Visible = False
        Panel2.Visible = True
        Panel3.Visible = False
        _PrintStatus = "D2"
        Call Load_Gride()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Panel3.Visible = True
        Panel2.Visible = False
        Panel1.Visible = False
        _PrintStatus = "E"
        Call Load_Gride_Movement()
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If Trim(cboItem.Text) <> "" Then
            If Search_Item() = True Then
                Call Load_Grid_StockMoveData(txtDate3.Text, txtDate4.Text)
                _From = txtDate3.Text
                _To = txtDate4.Text
                '_Itemcode=Trim(cboItem.Text )
                cboItem.Text = ""
                Panel3.Visible = False
            Else
                MsgBox("Please enter the correct Item Name", MsgBoxStyle.Information, "Information .....")
            End If
        Else
            MsgBox("Please enter the Item Name", MsgBoxStyle.Information, "Information .......")
        End If
    End Sub

    Function Search_Item() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M03Item_Master where m03Status='A' and  M03Item_Code='" & Trim(cboItem.Text) & "' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Item = True
                _Itemcode = M01.Tables(0).Rows(0)("M03Item_Code")
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
            End If
        End Try
    End Function

    Function Save_Report_StockMovement()
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

        Try
            Cursor = Cursors.WaitCursor
            nvcFieldList1 = "DELETE FROM R05Report"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                nvcFieldList1 = "Insert Into R05Report(R05Date,R05Item,R05Remark,R05Sales,R05GRN,R05MRT,R05Trnsfer,R05Wst,R05Balance,R05Location)" & _
                                                                " values('" & UltraGrid1.Rows(i).Cells(0).Value & "', '" & _Itemcode & "','" & (UltraGrid1.Rows(i).Cells(1).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & (UltraGrid1.Rows(i).Cells(4).Value) & "','" & (UltraGrid1.Rows(i).Cells(5).Value) & "','" & (UltraGrid1.Rows(i).Cells(6).Value) & "','" & (UltraGrid1.Rows(i).Cells(7).Value) & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                i = i + 1
            Next

            transaction.Commit()
            MsgBox("Report genarated successfully", MsgBoxStyle.Information, "Information .......")
            Cursor = Cursors.Arrow
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Function Load_Grid_StockMoveData(ByVal strFrom As Date, ByVal strTo As Date)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _Qty As Double
        Dim Value As Double
        Dim _Rowcount As Double
        Dim _St As String
        Dim _Date As Date
        Dim X As Integer
        Dim _Invoice As Integer
        Dim _LastOB_Date As Date

        Try
            _Date = strFrom.AddDays(-1)
            _LastOB_Date = _Date
            _Qty = 0
            Sql = "select s01Qty as Qty,S01Date  from S01Stock_Balance  where S01Date<'" & strFrom & "' and S01Item_Code='" & _Itemcode & "' and S01Trans_Type='OB' and S01Com_Code='" & _Comcode & "' order by s01id desc "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Qty = M01.Tables(0).Rows(0)("Qty")
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Date") = Month(_Date) & "/" & Microsoft.VisualBasic.Day(_Date) & "/" & Year(_Date)
                newRow("Remark") = "Operning Balance-" & cboItem.Text
                _LastOB_Date = M01.Tables(0).Rows(0)("S01Date")

                Sql = "select sum(S01Qty) as Qty  from S01Stock_Balance where S01Date between '" & _LastOB_Date & "' and '" & strFrom & "' and S01Trans_Type<>'OB' and S01Item_Code='" & _Itemcode & "'  and S01Loc_Code='" & _Comcode & "' group by S01Item_Code "
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    _Qty = _Qty + CDbl(M02.Tables(0).Rows(0)("Qty"))
                End If
                Value = _Qty
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Balance") = _St
                c_dataCustomer1.Rows.Add(newRow)

            Else
                Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Date<'" & strFrom & "' and S01Item_Code='" & _Itemcode & "' and  S01Com_Code='" & _Comcode & "' group by S01Item_Code "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    _Qty = M01.Tables(0).Rows(0)("Qty")
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Date") = Month(_Date) & "/" & Microsoft.VisualBasic.Day(_Date) & "/" & Year(_Date)
                    newRow("Remark") = "Operning Balance-" & cboItem.Text
                    ' _LastOB_Date = M01.Tables(0).Rows(0)("S01Date")

                  
                    Value = _Qty
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                    newRow("Balance") = _St
                    c_dataCustomer1.Rows.Add(newRow)
                End If
                End If


            Sql = "select S01Date  from S01Stock_Balance where S01Date between '" & txtDate3.Text & "' and '" & txtDate4.Text & "' and S01Item_Code='" & _Itemcode & "' and S01Status not in ('CLOSE','CANCEL') and S01Com_Code='" & _Comcode & "' group by S01Date "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Sql = "select sum(S01Qty) as Qty ,S01STATUS  from S01Stock_Balance where S01Date='" & M01.Tables(0).Rows(i)("S01Date") & "' and S01Item_Code='" & _Itemcode & "'  and S01Com_Code='" & _Comcode & "' and S01Trans_Type='OB' group by S01Item_Code,S01STATUS "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        If Trim(M02.Tables(0).Rows(0)("S01STATUS")) = "I" Then
                            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
                            newRow2("Date") = Month(M01.Tables(0).Rows(i)("S01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("S01Date")) & "/" & Year(M01.Tables(0).Rows(i)("S01Date"))
                            newRow2("Remark") = "Previous O/b-" & cboItem.Text

                            Value = M02.Tables(0).Rows(0)("Qty")
                            _Qty = Value
                            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            newRow2("Balance") = _St
                            c_dataCustomer1.Rows.Add(newRow2)
                        Else
                            Dim newRow2 As DataRow = c_dataCustomer1.NewRow
                            newRow2("Date") = Month(M01.Tables(0).Rows(i)("S01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("S01Date")) & "/" & Year(M01.Tables(0).Rows(i)("S01Date"))
                            newRow2("Remark") = "Change O/b-" & cboItem.Text
                            _Qty = 0
                            Value = M02.Tables(0).Rows(0)("Qty")
                            _Qty = Value
                            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                            newRow2("Balance") = _St
                            c_dataCustomer1.Rows.Add(newRow2)
                        End If
                    End If

                Sql = "select count(S01Ref_No) as Qty  from S01Stock_Balance where S01Date='" & M01.Tables(0).Rows(i)("S01Date") & "' and S01Item_Code='" & _Itemcode & "' and S01Status NOT IN ('CLOSE','CANCEL') and S01Trans_Type='DR' and S01Com_Code='" & _Comcode & "' group by S01Item_Code,S01Ref_No "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    X = 0
                    _Invoice = 0
                    For Each DTRow2 As DataRow In M02.Tables(0).Rows
                        _Invoice = _Invoice + M02.Tables(0).Rows(X)("Qty")
                        X = X + 1
                    Next

                    _Rowcount = 0
                    Dim newRow1 As DataRow = c_dataCustomer1.NewRow
                    newRow1("Date") = Month(M01.Tables(0).Rows(i)("S01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("S01Date")) & "/" & Year(M01.Tables(0).Rows(i)("S01Date"))
                    newRow1("Remark") = _Invoice & " Sales Invoice"
                    'Sales
                Sql = "select sum(S01Qty) as Qty  from S01Stock_Balance where S01Date='" & M01.Tables(0).Rows(i)("S01Date") & "' and S01Item_Code='" & _Itemcode & "' and S01Status  NOT IN ('CLOSE','CANCEL') and S01Trans_Type='DR' and S01Com_Code='" & _Comcode & "' group by S01Item_Code "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _Qty = _Qty + M02.Tables(0).Rows(0)("Qty")
                        _Rowcount = M02.Tables(0).Rows(0)("Qty")
                        Value = -(_Rowcount)
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        newRow1("Sales Qty") = _St
                    End If




                    'GRN
                Sql = "select sum(S01Qty) as Qty  from S01Stock_Balance where S01Date='" & M01.Tables(0).Rows(i)("S01Date") & "' and S01Item_Code='" & _Itemcode & "' and S01Status NOT IN ('CLOSE','CANCEL') and S01Trans_Type='GRN' and S01Com_Code='" & _Comcode & "' group by S01Item_Code "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _Qty = _Qty + M02.Tables(0).Rows(0)("Qty")
                        Value = M02.Tables(0).Rows(0)("Qty")
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        _Rowcount = M02.Tables(0).Rows(0)("Qty") + _Rowcount
                        ' M02.Tables(0).Rows(0)("GRN Qty") = _St
                        newRow1("GRN Qty") = _St
                    End If
                    'MK Return
                Sql = "select sum(S01Qty) as Qty  from S01Stock_Balance where S01Date='" & M01.Tables(0).Rows(i)("S01Date") & "' and S01Item_Code='" & _Itemcode & "' and S01Status NOT IN ('CLOSE','CANCEL') and S01Trans_Type='MR' and S01Com_Code='" & _Comcode & "' group by S01Item_Code "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _Qty = _Qty + M02.Tables(0).Rows(0)("Qty")
                        _Rowcount = M02.Tables(0).Rows(0)("Qty") + _Rowcount
                        Value = -(M02.Tables(0).Rows(0)("Qty"))
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        ' M02.Tables(0).Rows(0)("Mkt Return Qty") = _St
                        newRow1("Mkt Return Qty") = _St
                    End If

                    'Transfer
                Sql = "select sum(S01Qty) as Qty  from S01Stock_Balance where S01Date='" & M01.Tables(0).Rows(i)("S01Date") & "' and S01Item_Code='" & _Itemcode & "' and S01Status NOT IN ('CLOSE','CANCEL') and S01Trans_Type='TR' and S01Com_Code='" & _Comcode & "' group by S01Item_Code "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _Qty = _Qty + M02.Tables(0).Rows(0)("Qty")
                        _Rowcount = M02.Tables(0).Rows(0)("Qty") + _Rowcount
                        Value = -(M02.Tables(0).Rows(0)("Qty"))
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        ' M02.Tables(0).Rows(0)("Mkt Return Qty") = _St
                        newRow1("Transfer") = _St
                    End If

                    'Wastage
                Sql = "select sum(S01Qty) as Qty  from S01Stock_Balance where S01Date='" & M01.Tables(0).Rows(i)("S01Date") & "' and S01Item_Code='" & _Itemcode & "' and S01Status NOT IN ('CLOSE','CANCEL') and S01Trans_Type='WT' and S01Com_Code='" & _Comcode & "' group by S01Item_Code "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        _Qty = _Qty + M02.Tables(0).Rows(0)("Qty")
                        _Rowcount = M02.Tables(0).Rows(0)("Qty") + _Rowcount
                        Value = -(M02.Tables(0).Rows(0)("Qty"))
                        _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        ' M02.Tables(0).Rows(0)("Mkt Return Qty") = _St
                        newRow1("Wastage") = _St
                    End If
                    Value = _Qty
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    ' M02.Tables(0).Rows(0)("Balance") = _St
                    newRow1("Balance") = _St
                    c_dataCustomer1.Rows.Add(newRow1)
                    i = i + 1
                Next

                'i = 0
                '_Qty = 0
                'For Each DTRow1 As DataRow In M01.Tables(0).Rows
                'Dim newRow3 As DataRow = c_dataCustomer1.NewRow

                'Value = _Qty
                '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'newRow3("Balance") = _St
                'c_dataCustomer1.Rows.Add(newRow3)

                '    i = i + 1
                'Next
                con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String

        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"

            If _PrintStatus = "F1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance.Qty} < 0 and {View_StockBalance.S01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock3.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance1.S04Com_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance.S01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "B1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance.S04Com_Code}='" & _Comcode & "' and {View_StockBalance.Category} ='" & _Catcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "B2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock3.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance1.S04Com_Code}='" & _Comcode & "' and {View_StockBalance1.Category} ='" & _Catcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "C1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance.S01Com_Code}='" & _Comcode & "' and {View_StockBalance.Supplier} ='" & _Suplier & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "C2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance.Supplier} ='" & _Suplier & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "E" Then
                Call Save_Report_StockMovement()
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{R05Report.R05Location}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "D1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance.Qty} < 0  and {View_StockBalance.S01Com_Code}='" & _Comcode & "' and {View_StockBalance.Category} ='" & _Catcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "D2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Stock.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_StockBalance.Qty} < {View_StockBalance.reorder} and {View_StockBalance.S01Com_Code}='" & _Comcode & "' and {View_StockBalance.Supplier} ='" & _Suplier & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'con.close()
            End If
        End Try
    End Sub
End Class