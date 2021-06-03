Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class rptOutstanding
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Comcode As String
    Dim _Supplier As String
    Dim _Category As String
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub rptOutstanding_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
        Call Load_Customer()
        Call Load_Sales_Ref()
    End Sub


    Function Load_Gride_Data_1(ByVal strCode As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim TM As TimeSpan
        Try
            Sql = "select *  from View_Outstanding_1 where balance>0 and M17name='" & strCode & "' order by T06Cus_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Customer Code") = M01.Tables(0).Rows(i)("T06Cus_Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17name")
                TM = Today.Subtract(CDate(M01.Tables(0).Rows(i)("Invo_date")))
                newRow("##") = TM.Days
                newRow("Invoice Date") = Month(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Year(M01.Tables(0).Rows(i)("Invo_date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T06Invoice_no")

                Value = M01.Tables(0).Rows(i)("Total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Invoice Value") = _St

                Value = M01.Tables(0).Rows(i)("Paid")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Paid Amount") = _St

                Value = M01.Tables(0).Rows(i)("Balance")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Balance Amount") = _St

                _Qty = _Qty + Value

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Qty
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Balance Amount") = _St
            c_dataCustomer1.Rows.Add(newRow1)

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

    Function Load_Gride_Data_2(ByVal strCode As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim TM As TimeSpan
        Try
            Sql = "select *  from View_Outstanding_1 where balance>0 and T01user='" & strCode & "' order by T06Cus_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Customer Code") = M01.Tables(0).Rows(i)("T06Cus_Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17name")
                TM = Today.Subtract(CDate(M01.Tables(0).Rows(i)("Invo_date")))
                newRow("##") = TM.Days
                newRow("Invoice Date") = Month(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Year(M01.Tables(0).Rows(i)("Invo_date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T06Invoice_no")

                Value = M01.Tables(0).Rows(i)("Total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Invoice Value") = _St

                Value = M01.Tables(0).Rows(i)("Paid")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Paid Amount") = _St

                Value = M01.Tables(0).Rows(i)("Balance")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Balance Amount") = _St

                _Qty = _Qty + Value

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Qty
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Balance Amount") = _St
            c_dataCustomer1.Rows.Add(newRow1)

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

    Function Load_Gride_Data_Age(ByVal strCode As String)
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim TM As TimeSpan
        Try
            Sql = "select *  from View_AG_ALL where balance>0 and T01user='" & strCode & "' order by  p_days desc,st_30 desc,st_40 desc,st_60 desc"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Customer Code") = M01.Tables(0).Rows(i)("T06Cus_Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17name")
                TM = Today.Subtract(CDate(M01.Tables(0).Rows(i)("Invo_date")))
                newRow("1-30") = M01.Tables(0).Rows(i)("P_Days")
                newRow("30-45") = M01.Tables(0).Rows(i)("St_30")
                newRow("45-60") = M01.Tables(0).Rows(i)("ST_40")
                newRow("Grater than 60") = M01.Tables(0).Rows(i)("st_60")
                newRow("Inv.Date") = Month(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Year(M01.Tables(0).Rows(i)("Invo_date"))
                newRow("Inv.No") = M01.Tables(0).Rows(i)("T06Invoice_no")

                Value = M01.Tables(0).Rows(i)("Total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Inv.Value") = _St

                Value = M01.Tables(0).Rows(i)("Paid")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Paid Amount") = _St

                Value = M01.Tables(0).Rows(i)("Balance")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Balance") = _St

                _Qty = _Qty + Value

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Qty
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Balance") = _St
            c_dataCustomer1.Rows.Add(newRow1)

            _Rowcount = UltraGrid1.Rows.Count - 1
            UltraGrid1.Rows(_Rowcount).Cells(6).Appearance.BackColor = Color.Gold
            ' UltraGrid1.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_Rowcount).Cells(6).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            ' UltraGrid1.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
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
        Dim TM As TimeSpan
        Try
            Sql = "select *  from View_Outstanding_1 where balance>0 order by T06Cus_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Customer Code") = M01.Tables(0).Rows(i)("T06Cus_Code")
                newRow("Customer Name") = M01.Tables(0).Rows(i)("M17name")
                TM = Today.Subtract(CDate(M01.Tables(0).Rows(i)("Invo_date")))
                newRow("##") = TM.Days
                newRow("Invoice Date") = Month(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("Invo_date")) & "/" & Year(M01.Tables(0).Rows(i)("Invo_date"))
                newRow("Invoice No") = M01.Tables(0).Rows(i)("T06Invoice_no")

                Value = M01.Tables(0).Rows(i)("Total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Invoice Value") = _St

                Value = M01.Tables(0).Rows(i)("Paid")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Paid Amount") = _St

                Value = M01.Tables(0).Rows(i)("Balance")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Balance Amount") = _St

                _Qty = _Qty + Value

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Qty
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Balance Amount") = _St
            c_dataCustomer1.Rows.Add(newRow1)

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

    Function Load_Gride1()
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
            Sql = "select *  from View_Outstanding where balance>0"
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


    Function Load_Gride_age()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Outstanding_Age
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 80
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 80
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 60
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Width = 60
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Width = 60
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(10).Width = 60
            .DisplayLayout.Bands(0).Columns(10).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function
    Function Load_Gride_Deatiels()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Outstanding_detailes
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 220
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 60
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 80
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 80
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 90
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function
    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M17Name as [##] from M17Customer where M17Active='A' "
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

    Function Load_Sales_Ref()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Employee_Name as [##] from M01Employee_Master "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRef
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


    Private Sub SummeryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SummeryToolStripMenuItem.Click
        _PrintStatus = "A"
        Call Load_Gride()
        Call Load_Gride1()

    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String
        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"

            If _PrintStatus = "A" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\outstanding.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                ' B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Outstanding.Balance} > 0"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()


            ElseIf _PrintStatus = "B" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\outstanding1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                ' B.SetParameterValue("To", _To)
               ' B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Outstanding_1.Balance} > 0"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "C" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\outstanding1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                ' B.SetParameterValue("To", _To)
                ' B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Outstanding_1.Balance} > 0 and {View_Outstanding_1.M17name}='" & _Supplier & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "D" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\outstanding1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                ' B.SetParameterValue("To", _To)
                ' B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Outstanding_1.Balance} > 0 and {View_Outstanding_1.T01User}='" & _Supplier & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "AX" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\Ag_Report.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                ' B.SetParameterValue("To", _To)
                ' B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_AG_ALL.balance}  > 0 and {View_AG_ALL.T06Com_Code} = 'DS' and {View_AG_ALL.T01User}='" & _Category & "'"
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

    Private Sub DeatilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeatilesToolStripMenuItem.Click
        _PrintStatus = "B"
        Call Load_Gride_Deatiels()
        Call Load_Gride_Data()
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        cboItem.Text = ""
        Panel4.Visible = False
        Call Load_Gride()
    End Sub

    Private Sub TobeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TobeToolStripMenuItem.Click
        _PrintStatus = "C"
        Panel4.Visible = True
        Call Load_Gride_Deatiels()

    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Call Load_Gride_Deatiels()
        Call Load_Gride_Data_1(Trim(cboItem.Text))
        _Supplier = Trim(cboItem.Text)
        Panel4.Visible = False
    End Sub

    Private Sub BySalesRefToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySalesRefToolStripMenuItem.Click
        _PrintStatus = "D"
        Panel1.Visible = True
        Panel4.Visible = False
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _PrintStatus = "D" Then
            Call Load_Gride_Deatiels()
            Call Load_Gride_Data_2(Trim(cboRef.Text))
            _Supplier = Trim(cboRef.Text)
            Panel1.Visible = False
        ElseIf _PrintStatus = "AX" Then
            Call Load_Gride_age()
            Call Load_Gride_Data_Age(Trim(cboRef.Text))
            _Supplier = Trim(cboRef.Text)
            Panel1.Visible = False
        End If
    End Sub

    Private Sub AgeAnalysisToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AgeAnalysisToolStripMenuItem.Click
        _PrintStatus = "AX"
        Panel1.Visible = True
        Panel4.Visible = False
        Call Load_Gride_age()
        _Category = Trim(cboRef.Text)

    End Sub
End Class