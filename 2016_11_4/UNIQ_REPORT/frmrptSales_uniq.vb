Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptSales_uniq
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable
    Dim _PrintStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Cashier As String
    Dim _Comcode As String
    Dim _dIS As String

    Dim _Catcode As String
    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub Panel5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel5.Paint

    End Sub

    Private Sub AllInvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllInvoiceToolStripMenuItem.Click
        Call Load_Gride_Detailes()
        Panel5.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        _PrintStatus = "A1"
        txtCh1.Text = Today
        txtCh2.Text = Today

    End Sub

    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Sales
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 230
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 110
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        End With
    End Function


    Function Load_Gride_Detailes()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Sales_Detailes
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 100
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 180
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 70
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Width = 90
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Width = 110
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(10).Width = 110
            '.DisplayLayout.Bands(0).Columns(10).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        End With
    End Function

    Function Load_Gride_Detailes_1()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Sales_Detailes_1
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 280
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(6).Width = 70
            '.DisplayLayout.Bands(0).Columns(6).AutoEdit = False
          

            '.DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ''.DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        End With
    End Function

    Function Load_Gride_Detailes_Summery()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Sales_Summery
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 110
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 110
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 110
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 110
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            '.DisplayLayout.Bands(0).Columns(6).Width = 70
            '.DisplayLayout.Bands(0).Columns(6).AutoEdit = False


            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ''.DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        End With
    End Function

    
    Private Sub frmrptSales_uniq_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        Call Load_Type()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Panel5.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        _PrintStatus = ""
        Call Load_Gride()
    End Sub

    'Private Sub UltraButton5(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click, UltraButton1.Click
    '    If _PrintStatus = "A1" Then
    '        Call Load_Gride_Detailes()
    '        Call Load_Data_A1()
    '        _From = txtCh1.Text
    '        _To = txtCh2.Text
    '        Panel5.Visible = False
    '    ElseIf _PrintStatus = "A3" Then
    '        Call Load_Gride_Detailes_1()
    '        Call Load_Data_A3()
    '        _From = txtCh1.Text
    '        _To = txtCh2.Text
    '        Panel5.Visible = False
    '    ElseIf _PrintStatus = "C1" Then
    '        Call Load_Gride_Detailes_Summery()
    '        _From = txtCh1.Text
    '        _To = txtCh2.Text
    '        Panel5.Visible = False

    '    End If
    'End Sub
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

            Sql = "select T08Date,T08Tr_Type,T08Job_No,T08Invo_No,T09Item_Code,T09Item_Name,T09Qty,T09Retail,T09Discount,(T09Qty*T09Retail)-(T09Qty*T09Retail)*T09Discount/100 as Total from T08Sales_Header inner join T09Sales_Flutter on T08Invo_No=T09Inv_No  where t08date between '" & txtCh1.Text & "' and '" & txtCh2.Text & "' and T08Status='A' and T09Department='-' order by T08ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                ' newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T08Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T08Date")) & "/" & Year(M01.Tables(0).Rows(i)("T08Date"))
                newRow("#Invoice Type") = Trim(M01.Tables(0).Rows(i)("T08Tr_Type"))
                newRow("Job No") = Trim(M01.Tables(0).Rows(i)("T08Job_No"))
                newRow("Invoice No") = Trim(M01.Tables(0).Rows(i)("T08Invo_No"))
                newRow("Part No") = Trim(M01.Tables(0).Rows(i)("T09Item_Code"))
                newRow("Part Name") = M01.Tables(0).Rows(i)("T09Item_Name")
                Value = M01.Tables(0).Rows(i)("T09Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Qty") = _St

                Value = M01.Tables(0).Rows(i)("T09Retail")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Rate") = _St
                Value = M01.Tables(0).Rows(i)("T09Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Discount%") = _St

                _Total = _Total + CDbl(M01.Tables(0).Rows(i)("total"))
                Value = M01.Tables(0).Rows(i)("total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Amount") = _St
                'newRow("Terminal") = M01.Tables(0).Rows(i)("T01Terminal")
                'newRow("User") = M01.Tables(0).Rows(i)("T01User")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
           
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("#Total Amount") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid2.Rows.Count - 1
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.ClearAllPools()
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
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
        Dim _proft As Double
        Dim _Total As Double
        Try

            _proft = 0
            _Total = 0

            Sql = "select T08Date,T08Tr_Type,T08Job_No,T08Invo_No,T09Department,T09Item_Name,T09Qty,T09Retail from T08Sales_Header inner join T09Sales_Flutter on T08Invo_No=T09Inv_No  where t08date between '" & txtCh1.Text & "' and '" & txtCh2.Text & "' and T08Status='A' and T09Department<>'-' order by T08ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                ' newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T08Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T08Date")) & "/" & Year(M01.Tables(0).Rows(i)("T08Date"))
                '  newRow("#Invoice Type") = Trim(M01.Tables(0).Rows(i)("T08Tr_Type"))
                newRow("Job No") = Trim(M01.Tables(0).Rows(i)("T08Job_No"))
                newRow("Invoice No") = Trim(M01.Tables(0).Rows(i)("T08Invo_No"))
                newRow("Department") = Trim(M01.Tables(0).Rows(i)("T09Department"))
                newRow("Description") = Trim(M01.Tables(0).Rows(i)("T09Item_Name"))
                Value = M01.Tables(0).Rows(i)("T09Retail")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Rate") = _St

               
                _Total = _Total + CDbl(M01.Tables(0).Rows(i)("T09Retail"))
                
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow

            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("#Rate") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid2.Rows.Count - 1
            UltraGrid2.Rows(_Rowcount).Cells(5).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid2.Rows(_Rowcount).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.ClearAllPools()
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
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
        Dim _proft As Double
        Dim _Total As Double
        Try

            _proft = 0
            _Total = 0

            Sql = "select T08Date,T08Tr_Type,T08Job_No,T08Invo_No,T09Department,T09Item_Name,T09Qty,T09Retail from T08Sales_Header inner join T09Sales_Flutter on T08Invo_No=T09Inv_No  where t08date between '" & txtD1.Text & "' and '" & txtD2.Text & "' and T08Status='A' and T09Department='" & Trim(cboDepartment.Text) & "' order by T08ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                ' newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T08Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T08Date")) & "/" & Year(M01.Tables(0).Rows(i)("T08Date"))
                '  newRow("#Invoice Type") = Trim(M01.Tables(0).Rows(i)("T08Tr_Type"))
                newRow("Job No") = Trim(M01.Tables(0).Rows(i)("T08Job_No"))
                newRow("Invoice No") = Trim(M01.Tables(0).Rows(i)("T08Invo_No"))
                newRow("Department") = Trim(M01.Tables(0).Rows(i)("T09Department"))
                newRow("Description") = Trim(M01.Tables(0).Rows(i)("T09Item_Name"))
                Value = M01.Tables(0).Rows(i)("T09Retail")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Rate") = _St


                _Total = _Total + CDbl(M01.Tables(0).Rows(i)("T09Retail"))

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow

            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("#Rate") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid2.Rows.Count - 1
            UltraGrid2.Rows(_Rowcount).Cells(5).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid2.Rows(_Rowcount).Cells(5).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.ClearAllPools()
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
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
        Dim _proft As Double
        Dim _Total As Double
        Try

            _proft = 0
            _Total = 0

            If Trim(cboType.Text) = "DIRECT SALES" Then
                _dIS = "DIRECT_SALE"
            ElseIf Trim(cboType.Text) = "JOB INVOICE" Then
                _dIS = "JOB_INVOICE"
            End If
            Sql = "select T08Date,T08Tr_Type,T08Job_No,T08Invo_No,T09Item_Code,T09Item_Name,T09Qty,T09Retail,T09Discount,(T09Qty*T09Retail)-(T09Qty*T09Retail)*T09Discount/100 as Total from T08Sales_Header inner join T09Sales_Flutter on T08Invo_No=T09Inv_No  where t08date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and T08Status='A' AND T08Tr_Type='" & _dIS & "' and T09Department='-' order by T08ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                ' newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T08Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T08Date")) & "/" & Year(M01.Tables(0).Rows(i)("T08Date"))
                newRow("#Invoice Type") = Trim(M01.Tables(0).Rows(i)("T08Tr_Type"))
                newRow("Job No") = Trim(M01.Tables(0).Rows(i)("T08Job_No"))
                newRow("Invoice No") = Trim(M01.Tables(0).Rows(i)("T08Invo_No"))
                newRow("Part No") = Trim(M01.Tables(0).Rows(i)("T09Item_Code"))
                newRow("Part Name") = M01.Tables(0).Rows(i)("T09Item_Name")
                Value = M01.Tables(0).Rows(i)("T09Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Qty") = _St

                Value = M01.Tables(0).Rows(i)("T09Retail")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Rate") = _St
                Value = M01.Tables(0).Rows(i)("T09Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Discount%") = _St

                _Total = _Total + CDbl(M01.Tables(0).Rows(i)("total"))
                Value = M01.Tables(0).Rows(i)("total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Amount") = _St
                'newRow("Terminal") = M01.Tables(0).Rows(i)("T01Terminal")
                'newRow("User") = M01.Tables(0).Rows(i)("T01User")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow

            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("#Total Amount") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid2.Rows.Count - 1
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.ClearAllPools()
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()
            End If
        End Try
    End Function

  

    Function Load_Type()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M11Name as [##] from M11Common WHERE M11Status='SD' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboType
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275

            End With

           

            con.ClearAllPools()
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function


    Function Load_Department()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M08Description as [##] from M08Department WHERE M08Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDepartment
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 275

            End With



            con.ClearAllPools()
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
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

            If _PrintStatus = "A1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\SalesD1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T08Sales_Header.T08Date}   in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T08Sales_Header.T08Status}='A' and {T09Sales_Flutter.T09Department} = '-'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\SalesD1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T08Sales_Header.T08Date}   in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T08Sales_Header.T08Status}='A' and {T09Sales_Flutter.T09Department} = '-' AND {T08Sales_Header.T08Tr_Type}='" & _dIS & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A3" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\SalesD2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T08Sales_Header.T08Date}   in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T08Sales_Header.T08Status}='A' and {T09Sales_Flutter.T09Department} <> '-' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "A4" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\SalesD2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "sainfinity")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T08Sales_Header.T08Date}   in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {T08Sales_Header.T08Status}='A'  and {T09Sales_Flutter.T09Department}='" & _dIS & "'"
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

    Private Sub ByDepartmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDepartmentToolStripMenuItem.Click
        Call Load_Gride_Detailes()
        Panel5.Visible = False
        Panel1.Visible = True
        Panel2.Visible = False
        _PrintStatus = "A2"
        txtC1.Text = Today
        txtC2.Text = Today
    End Sub

    Private Sub UltraButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        If _PrintStatus = "A2" Then
            _From = txtC1.Text
            _To = txtC2.Text
            Call Load_Gride_Detailes()
            Call Load_Data_A2()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub AllDepartmentToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllDepartmentToolStripMenuItem1.Click
        Call Load_Gride_Detailes_1()
        Panel5.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        _PrintStatus = "A3"
        txtCh1.Text = Today
        txtCh2.Text = Today

    End Sub

    Private Sub ByDepartmentToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDepartmentToolStripMenuItem1.Click
        Call Load_Gride_Detailes_1()
        Panel5.Visible = False
        Panel1.Visible = False
        Panel2.Visible = True
        Call Load_Department()
        cboDepartment.ToggleDropdown()
        _PrintStatus = "A4"
        txtD1.Text = Today
        txtD2.Text = Today

    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        If _PrintStatus = "A4" Then
            Call Load_Gride_Detailes_1()
            Call Load_Data_A4()
            _From = txtD1.Text
            _To = txtD2.Text
            _dIS = Trim(cboDepartment.Text)
            Panel2.Visible = False
        End If
    End Sub

    Private Sub LaborChargersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaborChargersToolStripMenuItem.Click

    End Sub

    Private Sub SummeryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SummeryToolStripMenuItem.Click
        Call Load_Gride_Detailes()
        Panel5.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        _PrintStatus = "c1"
        txtCh1.Text = Today
        txtCh2.Text = Today
    End Sub

    Private Sub DateWiseSummeryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateWiseSummeryToolStripMenuItem.Click

        Call Load_Gride_Detailes_Summery1()
        Panel5.Visible = True
        Panel1.Visible = False
        Panel2.Visible = False
        _PrintStatus = "C"
        txtD1.Text = Today
        txtD2.Text = Today
    End Sub
    
    Function Load_Gride_Detailes_Summery1()
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

            Sql = "select T08Date,T08Tr_Type,T08Job_No,T08Invo_No,T09Item_Code,T09Item_Name,T09Qty,T09Retail,T09Discount,(T09Qty*T09Retail)-(T09Qty*T09Retail)*T09Discount/100 as Total from T08Sales_Header inner join T09Sales_Flutter on T08Invo_No=T09Inv_No  where t08date between '" & txtCh1.Text & "' and '" & txtCh2.Text & "' and T08Status='A' and T09Department='-' order by T08ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                ' newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T08Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T08Date")) & "/" & Year(M01.Tables(0).Rows(i)("T08Date"))
                newRow("#Invoice Type") = Trim(M01.Tables(0).Rows(i)("T08Tr_Type"))
                newRow("Job No") = Trim(M01.Tables(0).Rows(i)("T08Job_No"))
                newRow("Invoice No") = Trim(M01.Tables(0).Rows(i)("T08Invo_No"))
                newRow("Part No") = Trim(M01.Tables(0).Rows(i)("T09Item_Code"))
                newRow("Part Name") = M01.Tables(0).Rows(i)("T09Item_Name")
                Value = M01.Tables(0).Rows(i)("T09Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Qty") = _St

                Value = M01.Tables(0).Rows(i)("T09Retail")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Rate") = _St
                Value = M01.Tables(0).Rows(i)("T09Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Discount%") = _St

                _Total = _Total + CDbl(M01.Tables(0).Rows(i)("total"))
                Value = M01.Tables(0).Rows(i)("total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Total Amount") = _St
                'newRow("Terminal") = M01.Tables(0).Rows(i)("T01Terminal")
                'newRow("User") = M01.Tables(0).Rows(i)("T01User")
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow

            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            'newRow1("#Total Amount") = _St
            c_dataCustomer1.Rows.Add(newRow1)
            con.close()

            _Rowcount = UltraGrid2.Rows.Count - 1
            UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            'UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.ClearAllPools()
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()
            End If
        End Try
    End Function


    Function Load_Gride_Date_Summery1()
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

            Sql = "select t11date,sum(t11net_Amount) as Net_Amount,sum(T11Cash) as T11Cash,sum(T11Credit ) as T11Credit,sum(T11Chq) as T11Chq,sum(T11Card) as T11Card from T11Income_Summery where T11Status='A' and T11Date between '" & txtCh1.Text & "' and '" & txtCh2.Text & "' group by T11Date "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Qty = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                '  newRow("Ref No") = M01.Tables(0).Rows(i)("T01Ref_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T11Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T11Date")) & "/" & Year(M01.Tables(0).Rows(i)("T11Date"))
                Value = Trim(M01.Tables(0).Rows(i)("Net_Amount"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("#Net Amount") = _St
                Value = Trim(M01.Tables(0).Rows(i)("T11Cash"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Cash") = _St
                Value = Trim(M01.Tables(0).Rows(i)("T11Credit"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Credit") = _St
                Value = Trim(M01.Tables(0).Rows(i)("T11Chq"))
                Value = Trim(M01.Tables(0).Rows(i)("T11Credit"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Chq") = _St
                Value = Trim(M01.Tables(0).Rows(i)("T11Card"))
                Value = Trim(M01.Tables(0).Rows(i)("T11Credit"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("#Chq") = _St


             
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            'Dim newRow1 As DataRow = c_dataCustomer1.NewRow

            'Value = _Total
            '_St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '_St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            'newRow1("#Total Amount") = _St
            'c_dataCustomer1.Rows.Add(newRow1)
            'con.close()

            '_Rowcount = UltraGrid2.Rows.Count - 1
            'UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.BackColor = Color.Gold
            ''UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.BackColor = Color.Gold
            ''UltraGrid2.Rows(_Rowcount).Cells(8).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            'UltraGrid2.Rows(_Rowcount).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True

            con.ClearAllPools()
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()
            End If
        End Try
    End Function


    Private Sub UltraButton5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click
        If _PrintStatus = "C" Then
            Call Load_Gride_Detailes_Summery()
            Call Load_Gride_Date_Summery1()
        End If
    End Sub
End Class