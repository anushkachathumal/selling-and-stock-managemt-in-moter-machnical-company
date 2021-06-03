Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptProduct_Items
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _Category As String
    Dim _Comcode As String
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
    Private Sub frmrptProduct_Items_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        If strUGroup = "ADMIN" Or strUGroup = "MD" Or strUGroup = "DIRECTOR" Or strUGroup = "ACCOUNTANT" Or strUGroup = "ADD_ACCO" Then
            WithPriceToolStripMenuItem.Enabled = True
            WithCostPriceToolStripMenuItem.Enabled = True
        Else
            WithPriceToolStripMenuItem.Visible = False
            WithCostPriceToolStripMenuItem.Visible = False
        End If

        Call Load_Gride_Production()
        Call Load_Category()
        Call Load_Supplier()
    End Sub

   

    

    Private Sub WithPriceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithPriceToolStripMenuItem.Click
        Call Load_Gride_ProductionSet()
        _PrintStatus = "B"
        Call Load_Item_WithCost()
        Panel3.Visible = False
        Panel1.Visible = False
    End Sub


    Function Load_Gride_Production()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_ProductionItem
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 320
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False

            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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

    Function Load_Gride_ProductionSet()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_ProductionSet
        UltraGrid1.DataSource = c_dataCustomer2
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 320
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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

    Private Sub WithOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithOutToolStripMenuItem.Click
        Call Load_Gride_Production()
        _PrintStatus = "A"
        Call Load_Item_WithoutCost()
        Panel3.Visible = False
        Panel1.Visible = False
    End Sub

    Function Load_Item_WithoutCost_Supplier(ByVal strCat As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As String
        Dim Value As Double
        Try
            Sql = "select * from View_Items where  M09name='" & strCat & "' and M03Status='A' and M03Com_Code='" & _Comcode & "' order by M02Cat_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                _Qty = ""
                newRow("Item Code") = M01.Tables(0).Rows(i)("M03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                newRow("Product Category") = M01.Tables(0).Rows(i)("M02Cat_Name")
                Value = M01.Tables(0).Rows(i)("M03Retail_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _Qty
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

    Function Load_Item_WithoutCost_Category(ByVal strCat As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As String
        Dim Value As Double
        Try
            Sql = "select * from View_Items where  M02Cat_Name='" & strCat & "' and M03Status='A' and M03Com_Code='" & _Comcode & "' order by M02Cat_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                _Qty = ""
                newRow("Item Code") = M01.Tables(0).Rows(i)("M03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                newRow("Product Category") = M01.Tables(0).Rows(i)("M02Cat_Name")
                Value = M01.Tables(0).Rows(i)("M03Retail_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _Qty
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

    Function Load_Item_WithoutCost()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As String
        Dim Value As Double
        Try
            Sql = "select * from View_Items where M03Status='A' and M03Com_Code='" & _Comcode & "' order by M02Cat_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                _Qty = ""
                newRow("Item Code") = M01.Tables(0).Rows(i)("M03Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                newRow("Product Category") = M01.Tables(0).Rows(i)("M02Cat_Name")
                Value = M01.Tables(0).Rows(i)("M03Retail_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _Qty
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

    Function Load_Item_WithCost()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As String
        Dim Value As Double
        Try
            Sql = "select * from View_Items where M03Status='A' and M03Com_Code='" & _Comcode & "' order by M02Cat_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                _Qty = ""
                newRow("Item Code") = M01.Tables(0).Rows(i)("M03Item_Code")
                newRow("Product Category") = M01.Tables(0).Rows(i)("M02Cat_Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("M03Retail_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _Qty
                Value = M01.Tables(0).Rows(i)("M03Cost_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _Qty
                c_dataCustomer2.Rows.Add(newRow)

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

    Function Load_Item_WithCost_Supplier(ByVal strCat As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As String
        Dim Value As Double
        Try
            Sql = "select * from View_Items where M09name='" & strCat & "' and M03Status='A' and M03Com_Code='" & _Comcode & "' order by M02Cat_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                _Qty = ""
                newRow("Item Code") = M01.Tables(0).Rows(i)("M03Item_Code")
                newRow("Product Category") = M01.Tables(0).Rows(i)("M02Cat_Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("M03Retail_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _Qty
                Value = M01.Tables(0).Rows(i)("M03Cost_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _Qty
                c_dataCustomer2.Rows.Add(newRow)

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

    Function Load_Item_WithCost_Category(ByVal strCat As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As String
        Dim Value As Double
        Try
            Sql = "select * from View_Items where M02Cat_Name='" & strCat & "' and M03Status='A' and M03Com_Code='" & _Comcode & "' order by M02Cat_Name"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                _Qty = ""
                newRow("Item Code") = M01.Tables(0).Rows(i)("M03Item_Code")
                newRow("Product Category") = M01.Tables(0).Rows(i)("M02Cat_Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03Item_Name")
                Value = M01.Tables(0).Rows(i)("M03Retail_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _Qty
                Value = M01.Tables(0).Rows(i)("M03Cost_Price")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _Qty
                c_dataCustomer2.Rows.Add(newRow)

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

    Function Load_Production_WithoutCost()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As String
        Dim Value As Double
        Try
            Sql = "select * from View_Production_Items where M14Status='A' and category='PS'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                _Qty = ""
                newRow("Item Code") = M01.Tables(0).Rows(i)("M14Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                Value = M01.Tables(0).Rows(i)("M14Retail")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _Qty
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


    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Try
            If _PrintStatus = "A" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Item2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Items.m03Status} = 'A' and {View_Items.M03Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()

                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "B" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Item1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Items.m03Status} = 'A' and {View_Items.M03Location}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "C1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Item2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Items.m03Status} = 'A' and {View_Items.M03Com_Code}='" & _Comcode & "' and {View_Items.M02Cat_Name} ='" & _Category & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "C2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Item1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Items.m03Status} = 'A' and {View_Items.M03Com_Code}='" & _Comcode & "' and {View_Items.M02Cat_Name} ='" & _Category & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "D1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Item2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Items.m03Status} = 'A' and {View_Items.M03Com_Code}='" & _Comcode & "' and {View_Items.M09Name}  ='" & _Category & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "D2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Item1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Items.m03Status} = 'A' and {View_Items.M03Com_Code}='" & _Comcode & "' and {View_Items.M09Name}  ='" & _Category & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                '   MsgBox(i)
            End If
        End Try

    End Sub

    Function Load_Production_WithCost()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As String
        Dim Value As Double
        Try
            Sql = "select * from View_Production_Items where M14Status='A' and category='PS'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer2.NewRow
                _Qty = ""
                newRow("Item Code") = M01.Tables(0).Rows(i)("M14Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                Value = M01.Tables(0).Rows(i)("M14Retail")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _Qty
                Value = M01.Tables(0).Rows(i)("M14Cost")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _Qty
                c_dataCustomer2.Rows.Add(newRow)

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

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
       
    End Sub

  
    Private Sub WithoutCostPriceToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_Gride_Production()
        _PrintStatus = "C"
        Call Load_Item_WithoutCost_Category("BOYES")
        _Category = "BOYES"

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
    Private Sub WithCostPriceToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_Gride_ProductionSet()
        _PrintStatus = "D"
        Call Load_Item_WithCost_Category("BOYES")
        _Category = "BOYES"
    End Sub

    Private Sub WithoutCostPriceToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_Gride_Production()
        _PrintStatus = "E"
        Call Load_Item_WithoutCost_Category("GENTS")
        _Category = "GENTS"

    End Sub

    Private Sub WithCostPriceToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Load_Gride_ProductionSet()
        _PrintStatus = "F"
        Call Load_Item_WithCost_Category("GENTS")
        _Category = "GENTS"
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub WithoutCostPriceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithoutCostPriceToolStripMenuItem.Click
        _PrintStatus = "C1"
        Panel3.Visible = True
        Panel1.Visible = False
    End Sub

    Private Sub WithCostPriceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithCostPriceToolStripMenuItem.Click
        _PrintStatus = "C2"
        Panel3.Visible = True
        Panel1.Visible = False
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _PrintStatus = "C1" Then
            Call Load_Gride_Production()
            Call Load_Item_WithoutCost_Category(cboCategory.Text)
            _Category = Trim(cboCategory.Text)
            cboCategory.Text = ""
            Panel3.Visible = False

        ElseIf _PrintStatus = "C2" Then
            Call Load_Gride_ProductionSet()
            Call Load_Item_WithCost_Category(cboCategory.Text)
            _Category = Trim(cboCategory.Text)
            cboCategory.Text = ""
            Panel3.Visible = False
        End If
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Call Load_Gride_Production()
        _PrintStatus = ""
        Panel3.Visible = False
        cboCategory.Text = ""
    End Sub

    Private Sub WithoutCostPriceToolStripMenuItem1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithoutCostPriceToolStripMenuItem1.Click
        _PrintStatus = "D1"
        Panel3.Visible = False
        Panel1.Visible = True
    End Sub

    Private Sub WithCostPriceToolStripMenuItem1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithCostPriceToolStripMenuItem1.Click
        _PrintStatus = "D2"
        Panel3.Visible = False
        Panel1.Visible = True
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _PrintStatus = "D1" Then
            Call Load_Gride_Production()
            Call Load_Item_WithoutCost_Supplier(cboSupplier.Text)
            _Category = cboSupplier.Text
            cboSupplier.Text = ""
            Panel1.Visible = False
        ElseIf _PrintStatus = "D2" Then
            Call Load_Gride_ProductionSet()
            Call Load_Item_WithCost_Supplier(cboSupplier.Text)
            _Category = cboSupplier.Text
            cboSupplier.Text = ""
            Panel1.Visible = False
        End If
    End Sub
End Class