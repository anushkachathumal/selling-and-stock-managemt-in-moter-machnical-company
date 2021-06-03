Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmrptGRN_POS
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable
    Dim _PrintStatus As String
    Dim _MainStatus As String
    Dim _From As Date
    Dim _To As Date
    Dim _Itemcode As String
    Dim _Supplier As String
    Dim _Category As String
    Dim _Comcode As String

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Gride_Det()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_GRN_rptD
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 150
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 170
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 70
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Width = 60
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Width = 60
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(10).Width = 110
            .DisplayLayout.Bands(0).Columns(10).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(10).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function


    Function Load_Gride()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_GRN_rpt
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
            .DisplayLayout.Bands(0).Columns(4).Width = 220
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(5).Width = 90
            .DisplayLayout.Bands(0).Columns(5).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            .DisplayLayout.Bands(0).Columns(6).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(7).Width = 90
            .DisplayLayout.Bands(0).Columns(7).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(8).Width = 90
            .DisplayLayout.Bands(0).Columns(8).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(9).Width = 110
            .DisplayLayout.Bands(0).Columns(9).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(9).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        End With
    End Function

    Private Sub frmrptGRN_POS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
        Call Load_Supplier()
        Call Load_Gride_Item()
        Call Load_Item()
        Call Load_Category()
        txtEntry.ReadOnly = True
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate.ReadOnly = True
        txtCom_Invoice.ReadOnly = True
        cboLocation.ReadOnly = True
        cboTo.ReadOnly = True
        txtRemark.ReadOnly = True
        txtNett.ReadOnly = True
        txtMarket.ReadOnly = True
        txtCount.ReadOnly = True
        txtVAT.ReadOnly = True
        txtNbt.ReadOnly = True
        txtDis_Rate.ReadOnly = True
        txtDiscount.ReadOnly = True
        txtGross.ReadOnly = True
        txtGross.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtMarket.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDis_Rate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtVAT.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNbt.Appearance.TextHAlign = Infragistics.Win.HAlign.Right

        txtCount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
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
            Sql = "select M03Item_Code as [##] from M03Item_Master where m03Status='A' and M03Com_Code='" & _Comcode & "' "
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

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Call Load_Gride()
        _MainStatus = ""
        _PrintStatus = ""
        Panel2.Visible = False
        Panel1.Visible = False
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        txtDate.Text = ""
        txtDis_Rate.Text = ""
        txtEntry.Text = ""
        cboLocation.Text = ""
        cboTo.Text = ""
        txtNett.Text = ""
        txtCount.Text = ""
        txtCom_Invoice.Text = ""
        txtDiscount.Text = ""
        txtGross.Text = ""
        txtVAT.Text = ""
        txtCount.Text = ""
        ' Panel5.Visible = True
    End Sub

    Private Sub ByDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem.Click
        _MainStatus = "SUMMERY"
        _PrintStatus = "S1"
        Panel2.Visible = True
        Panel1.Visible = False
        Panel3.Visible = False
        OPR5.Visible = False
        txtB1.Text = Today
        txtB2.Text = Today
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Function Load_Gride_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name] from M03Item_Master where M03Status='A' and M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 130
            UltraGrid3.Rows.Band.Columns(1).Width = 370
            ' UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        _From = txtB1.Text
        _To = txtB2.Text

        If _PrintStatus = "S1" Then
            Call Load_Grid_Data()
            Panel2.Visible = False
        ElseIf _PrintStatus = "S3" Then
            Call Load_Grid_Data_S3()
            Panel2.Visible = False
        ElseIf _PrintStatus = "S5" Then
            Call Load_Grid_Data_S5()
            Panel2.Visible = False
        ElseIf _PrintStatus = "D1" Then
            Call Load_Grid_Data_D1()
            Panel2.Visible = False
        ElseIf _PrintStatus = "D5" Then
            Call Load_Grid_Data_D4()
            Panel2.Visible = False
        End If
    End Sub

    Function Load_Grid_Data_S8()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select *  from View_GRN_Header where t01date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M09Name='" & Trim(cboSupplier.Text) & "' and T01PO_NO='' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St

                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01NBT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("NBT") = _St

                Value = M01.Tables(0).Rows(i)("Gross")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Gross Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Gross Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function
    Function Load_Grid_Data_S7()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select *  from View_GRN_Header where t01date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M09Name='" & Trim(cboSupplier.Text) & "' and T01PO_NO<>'' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St

                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01NBT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("NBT") = _St

                Value = M01.Tables(0).Rows(i)("Gross")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Gross Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Gross Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data_S6()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select *  from View_GRN_Header where t01date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M09Name='" & Trim(cboSupplier.Text) & "' and T01Com_Discount>0 and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St

                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01NBT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("NBT") = _St


                Value = M01.Tables(0).Rows(i)("Gross")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Gross Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Gross Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data_S4()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select *  from View_GRN_Header where t01date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M09Name='" & Trim(cboSupplier.Text) & "' and T01Vat>0 and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St

                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01NBT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("NBT") = _St


                Value = M01.Tables(0).Rows(i)("Gross")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Gross Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Gross Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data_S2()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select *  from View_GRN_Header where t01date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M09Name='" & Trim(cboSupplier.Text) & "' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St

                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01NBT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("NBT") = _St


                Value = M01.Tables(0).Rows(i)("Gross")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Gross Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Gross Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select *  from View_GRN_Header where t01date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St

                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01NBT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("NBT") = _St


                Value = M01.Tables(0).Rows(i)("Gross")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Gross Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Gross Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function


    Function Load_Grid_Data_S3()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select *  from View_GRN_Header where t01date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and T01Vat>0 and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St

                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01NBT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("NBT") = _St


                Value = M01.Tables(0).Rows(i)("Gross")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Gross Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Gross Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data_D4()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride_Det()
            Sql = "select *  from View_GRN_Header inner join View_T02Transaction on T02Ref_No=T01Ref_no where t01date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and T02Free_Issue>0 and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03item_Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T02cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Cost Price") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Free Issue") = _St


                Value = M01.Tables(0).Rows(i)("T02Total")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data_D1()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride_Det()
            Sql = "select *  from View_GRN_Header inner join View_T02Transaction on T02Ref_No=T01Ref_no where t01date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03item_Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T02cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Cost Price") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Free Issue") = _St


                Value = M01.Tables(0).Rows(i)("T02Total")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data_D2()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride_Det()
            Sql = "select *  from View_GRN_Header inner join View_T02Transaction on T02Ref_No=T01Ref_no where t01date between '" & txtDate1.Text & "' and '" & txtDate2.Text & "' and M09Name='" & _Supplier & "' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03item_Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T02cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Cost Price") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Free Issue") = _St


                Value = M01.Tables(0).Rows(i)("T02Total")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data_D5()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride_Det()
            Sql = "select *  from View_GRN_Header inner join View_T02Transaction on T02Ref_No=T01Ref_no where t01date between '" & txtD1.Text & "' and '" & txtD2.Text & "' and M03Cat_Code='" & _Category & "' and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03item_Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T02cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Cost Price") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Free Issue") = _St


                Value = M01.Tables(0).Rows(i)("T02Total")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data_S5()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride()
            Sql = "select *  from View_GRN_Header where t01date between '" & txtB1.Text & "' and '" & txtB2.Text & "' and T01Com_Discount>0 and T01Com_Code='" & _Comcode & "' order by T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T01Net_Amount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Net Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01Com_Discount")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount") = _St

                Value = M01.Tables(0).Rows(i)("T01Vat")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("VAT Amount") = _St

                Value = M01.Tables(0).Rows(i)("T01NBT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("NBT") = _St


                Value = M01.Tables(0).Rows(i)("Gross")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Gross Amount") = _St

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Gross Amount") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(9).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function


    Function Load_Grid_Data_D3()
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

        Try
            Me.Cursor = Cursors.WaitCursor
            Call Load_Gride_Det()
            Sql = "select *  from View_GRN_Header T inner join View_T02Transaction T1 on t1.T02Ref_No=t.T01Ref_no where t.t01date between '" & txtC1.Text & "' and '" & txtC2.Text & "' and t1.T02Item_Code='" & _Itemcode & "' and t.T01Com_Code='" & _Comcode & "' order by t.T01Ref_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow

                newRow("GRN No") = M01.Tables(0).Rows(i)("T01Grn_No")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T01Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01Date")) & "/" & Year(M01.Tables(0).Rows(i)("T01Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("Supplier Name") = M01.Tables(0).Rows(i)("M09Name")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M03item_Name")
                newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")
                newRow("PO#") = M01.Tables(0).Rows(i)("T01PO_NO")
                Value = M01.Tables(0).Rows(i)("T02cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("Cost Price") = _St
                Value = M01.Tables(0).Rows(i)("T02Retail_Price")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                Value = M01.Tables(0).Rows(i)("T02Qty")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Qty") = _St

                Value = M01.Tables(0).Rows(i)("T02Free_Issue")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Free Issue") = _St


                Value = M01.Tables(0).Rows(i)("T02Total")
                _Total = _Total + Value
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

            Dim newRow1 As DataRow = c_dataCustomer1.NewRow
            Value = _Total
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow1("Total") = _St

            c_dataCustomer1.Rows.Add(newRow1)

            _LastRow = UltraGrid1.Rows.Count
            _LastRow = _LastRow - 1
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.BackColor = Color.Gold
            UltraGrid1.Rows(_LastRow).Cells(10).Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            con.close()
            Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim B As New ReportDocument
        Dim A As String
        Dim StrFromDate As String
        Dim StrToDate As String

        Try
            StrFromDate = "(" & Year(_From) & ", " & VB6.Format(Month(_From), "0#") & ", " & VB6.Format(CDate(_From).Day, "0#") & ", 00, 00, 00)"
            StrToDate = "(" & Year(_To) & ", " & VB6.Format(Month(_To), "0#") & ", " & VB6.Format(CDate(_To).Day, "0#") & ", 00, 00, 00)"

            If _PrintStatus = "S1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.T01Com_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "S2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.M09Name}='" & _Supplier & "' and {View_GRN_Header.T01To_Loc_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "S3" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.T01Vat}>0 and {View_GRN_Header.T01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "S4" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.M09Name}='" & _Supplier & "' and {View_GRN_Header.T01Vat}>0 and {View_GRN_Header.T01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "S5" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.T01Com_Discount}>0 and {View_GRN_Header.T01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "S6" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.M09Name}='" & _Supplier & "' and {View_GRN_Header.T01Com_Discount}>0 and {View_GRN_Header.T01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "S7" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.M09Name}='" & _Supplier & "' and {View_GRN_Header.T01PO_NO}<>'' and {View_GRN_Header.T01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "S8" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN1.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.M09Name}='" & _Supplier & "' and {View_GRN_Header.T01PO_NO}='' and {View_GRN_Header.TT01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()

            ElseIf _PrintStatus = "D1" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.T01Com_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "D2" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_GRN_Header.M09Name}='" & _Supplier & "' and {View_GRN_Header.T01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "D3" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_T02Transaction.T02Item_Code}='" & _Itemcode & "' and {View_GRN_Header.T01Com_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            ElseIf _PrintStatus = "D4" Then
                A = ConfigurationManager.AppSettings("ReportPath") + "\GRN2.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("To", _To)
                B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Date}  in DateTime " & StrFromDate & " to DateTime " & StrToDate & " and {View_T02Transaction.M03Cat_Code}='" & _Category & "' and {View_GRN_Header.T01Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub

    Private Sub BySupplerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplerToolStripMenuItem.Click
        _MainStatus = "SUMMERY"
        _PrintStatus = "S2"
        Panel2.Visible = False
        Panel1.Visible = True
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Active='A' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupplier
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

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        _From = txtDate1.Text
        _To = txtDate2.Text

        If _PrintStatus = "S2" Then
            If Search_Supplier() = True Then
                Call Load_Grid_Data_S2()
                _Supplier = Trim(cboSupplier.Text)
            Else
                MsgBox("Please select the correct supplier", MsgBoxStyle.Information, "Information ......")
                cboSupplier.ToggleDropdown()
                Exit Sub
            End If
        ElseIf _PrintStatus = "S4" Then
            If Search_Supplier() = True Then
                Call Load_Grid_Data_S4()
                _Supplier = Trim(cboSupplier.Text)
            Else
                MsgBox("Please select the correct supplier", MsgBoxStyle.Information, "Information ......")
                cboSupplier.ToggleDropdown()
                Exit Sub
            End If
        ElseIf _PrintStatus = "S6" Then
            If Search_Supplier() = True Then
                Call Load_Grid_Data_S6()
                _Supplier = Trim(cboSupplier.Text)
            Else
                MsgBox("Please select the correct supplier", MsgBoxStyle.Information, "Information ......")
                cboSupplier.ToggleDropdown()
                Exit Sub
            End If
        ElseIf _PrintStatus = "S7" Then
            If Search_Supplier() = True Then
                Call Load_Grid_Data_S7()
                _Supplier = Trim(cboSupplier.Text)
            Else
                MsgBox("Please select the correct supplier", MsgBoxStyle.Information, "Information ......")
                cboSupplier.ToggleDropdown()
                Exit Sub
            End If
        ElseIf _PrintStatus = "S8" Then
            If Search_Supplier() = True Then
                Call Load_Grid_Data_S8()
                _Supplier = Trim(cboSupplier.Text)
            Else
                MsgBox("Please select the correct supplier", MsgBoxStyle.Information, "Information ......")
                cboSupplier.ToggleDropdown()
                Exit Sub
            End If
        ElseIf _PrintStatus = "D2" Then
            If Search_Supplier() = True Then
                Call Load_Grid_Data_D2()
                _Supplier = Trim(cboSupplier.Text)
            Else
                MsgBox("Please select the correct supplier", MsgBoxStyle.Information, "Information ......")
                cboSupplier.ToggleDropdown()
                Exit Sub
            End If
        End If

        cboSupplier.Text = ""
        Panel1.Visible = False
    End Sub

    Function Search_Category() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M02Category where  M02Cat_Name='" & Trim(cboCategory.Text) & "' and M02Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Category = True
                _Category = Trim(M01.Tables(0).Rows(0)("M02Cat_Code"))
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

    Function Search_Supplier() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M09Supplier where M09Active='A' and m09Name='" & Trim(cboSupplier.Text) & "' and M09Loc_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Supplier = True
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

    Function Search_Item() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M03Item_Master where m03Status='A' and M03Item_Code='" & Trim(cboItem.Text) & "' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Item = True
                _Itemcode = Trim(M01.Tables(0).Rows(0)("M03Item_Code"))
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

    Private Sub AllVATToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllVATToolStripMenuItem.Click
        _MainStatus = "SUMMERY"
        _PrintStatus = "S3"
        Panel2.Visible = True
        Panel1.Visible = False
        txtB1.Text = Today
        txtB2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Private Sub UsingSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingSupplierToolStripMenuItem.Click
        _MainStatus = "SUMMERY"
        _PrintStatus = "S4"
        Panel2.Visible = False
        Panel1.Visible = True
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Private Sub AllDiscountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllDiscountToolStripMenuItem.Click
        _MainStatus = "SUMMERY"
        _PrintStatus = "S5"
        Panel2.Visible = True
        Panel1.Visible = False
        txtB1.Text = Today
        txtB2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub UsingSupplierToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingSupplierToolStripMenuItem1.Click
        _MainStatus = "SUMMERY"
        _PrintStatus = "S6"
        Panel2.Visible = False
        Panel1.Visible = True
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
    End Sub

    Private Sub POBaseInvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBaseInvoiceToolStripMenuItem.Click
        _MainStatus = "SUMMERY"
        _PrintStatus = "S7"
        Panel2.Visible = False
        Panel1.Visible = True
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Private Sub WithoutPOInvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithoutPOInvoiceToolStripMenuItem.Click
        _MainStatus = "SUMMERY"
        _PrintStatus = "S8"
        Panel2.Visible = False
        Panel1.Visible = True
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Private Sub ByDateToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem1.Click
        _MainStatus = "DETAILES"
        _PrintStatus = "D1"
        Panel2.Visible = True
        Panel1.Visible = False
        txtB1.Text = Today
        txtB2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Private Sub BySupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BySupplierToolStripMenuItem.Click
        _MainStatus = "DETAILES"
        _PrintStatus = "D2"
        Panel2.Visible = False
        Panel1.Visible = True
        txtDate1.Text = Today
        txtDate2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Private Sub ByItemToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByItemToolStripMenuItem.Click
        _MainStatus = "DETAILES"
        _PrintStatus = "D3"
        Panel2.Visible = False
        Panel1.Visible = False
        txtC1.Text = Today
        txtC2.Text = Today
        Panel3.Visible = True
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

  

    Private Sub cboItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItem.KeyUp
        If e.KeyCode = Keys.F1 Then
            OPR5.Visible = True
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            cboItem.Focus()
        End If
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        _From = txtC1.Text
        _To = txtC2.Text

        If _PrintStatus = "D3" Then
            If Search_Item() = True Then
                Call Load_Grid_Data_D3()
                ' _Itemcode = Trim(cboItem.Text)
            Else
                MsgBox("Please select the correct Item", MsgBoxStyle.Information, "Information ......")
                cboSupplier.ToggleDropdown()
                Exit Sub
            End If
        End If

        Panel3.Visible = False
    End Sub

    Private Sub UltraGrid3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid3.KeyUp
        On Error Resume Next
        If e.KeyCode = 13 Then

            Dim _RowIndex As Integer
            _RowIndex = UltraGrid3.ActiveRow.Index
            cboItem.Text = Trim(UltraGrid3.Rows(_RowIndex).Cells(1).Text)
            OPR5.Visible = False
        End If
    End Sub

  

    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _RowIndex As Integer
        _RowIndex = UltraGrid3.ActiveRow.Index
        cboItem.Text = Trim(UltraGrid3.Rows(_RowIndex).Cells(1).Text)
        OPR5.Visible = False
    End Sub

    Private Sub txtFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyUp
        If e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            cboItem.Focus()
        End If
    End Sub

    Private Sub txtFind_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFind.TextChanged
        Call Load_Gride_Item3()
    End Sub

    Function Load_Gride_Item3()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Item_Name  like '%" & txtFind.Text & "%' and M03Status='A' and M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 130
            UltraGrid3.Rows.Band.Columns(1).Width = 370
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub ByCategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByCategoryToolStripMenuItem.Click
        _MainStatus = "DETAILES"
        _PrintStatus = "D4"
        Panel2.Visible = False
        Panel1.Visible = False
        txtD1.Text = Today
        txtD2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = True
        Panel5.Visible = False
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        If Search_Category() = True Then
            _Category = Trim(cboCategory.Text)
            Call Load_Grid_Data_D5()
        Else
            MsgBox("Please select the correct Category", MsgBoxStyle.Information, "Information ......")
            cboCategory.ToggleDropdown()
            Exit Sub
        End If

        Panel4.Visible = False
    End Sub

    Private Sub txtD1_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtD1.BeforeDropDown

    End Sub

    Private Sub ByDateToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateToolStripMenuItem2.Click
        _MainStatus = "DETAILES"
        _PrintStatus = "D5"
        Panel2.Visible = True
        Panel1.Visible = False
        txtB1.Text = Today
        txtB2.Text = Today
        Panel3.Visible = False
        OPR5.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub UltraGrid1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid1.MouseDoubleClick
        Dim _Rowindex As Integer
        _Rowindex = UltraGrid1.ActiveRow.Index
        If _PrintStatus = "S1" Then
            If Panel5.Visible = True Then
                Call Clear_Gride3()
                Panel5.Visible = False

            Else
                Call Clear_Gride3()
                Panel5.Visible = True
                Call Search_RecordsUsing_Entry(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
            End If
        ElseIf _PrintStatus = "S2" Then

            If Panel5.Visible = True Then
                Call Clear_Gride3()
                Panel5.Visible = False

            Else
                Call Clear_Gride3()
                Panel5.Visible = True
                Call Search_RecordsUsing_Entry(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
            End If
        ElseIf _PrintStatus = "S3" Then

            If Panel5.Visible = True Then
                Call Clear_Gride3()
                Panel5.Visible = False

            Else
                Call Clear_Gride3()
                Panel5.Visible = True
                Call Search_RecordsUsing_Entry(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
            End If
        ElseIf _PrintStatus = "S4" Then

            If Panel5.Visible = True Then
                Call Clear_Gride3()
                Panel5.Visible = False

            Else
                Call Clear_Gride3()
                Panel5.Visible = True
                Call Search_RecordsUsing_Entry(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
            End If
        ElseIf _PrintStatus = "S5" Then

            If Panel5.Visible = True Then
                Call Clear_Gride3()
                Panel5.Visible = False

            Else
                Call Clear_Gride3()
                Panel5.Visible = True
                Call Search_RecordsUsing_Entry(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
            End If
        ElseIf _PrintStatus = "S6" Then

            If Panel5.Visible = True Then
                Call Clear_Gride3()
                Panel5.Visible = False

            Else
                Call Clear_Gride3()
                Panel5.Visible = True
                Call Search_RecordsUsing_Entry(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
            End If
        ElseIf _PrintStatus = "S7" Then

            If Panel5.Visible = True Then
                Call Clear_Gride3()
                Panel5.Visible = False

            Else
                Call Clear_Gride3()
                Panel5.Visible = True
                Call Search_RecordsUsing_Entry(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
            End If
        ElseIf _PrintStatus = "S8" Then

            If Panel5.Visible = True Then
                Call Clear_Gride3()
                Panel5.Visible = False

            Else
                Call Clear_Gride3()
                Panel5.Visible = True
                Call Search_RecordsUsing_Entry(UltraGrid1.Rows(_Rowindex).Cells(1).Text)
            End If
        End If
    End Sub

    Function Clear_Gride3()
        txtDate.Text = ""
        txtDis_Rate.Text = ""
        txtEntry.Text = ""
        cboLocation.Text = ""
        cboTo.Text = ""
        txtNett.Text = ""
        txtCount.Text = ""
        txtCom_Invoice.Text = ""
        txtDiscount.Text = ""
        txtGross.Text = ""
        txtVAT.Text = ""
        txtCount.Text = ""
    End Function

    Function Search_RecordsUsing_Entry(ByVal strCode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet
        Dim i As Integer
        Dim Value As Double


        Try
            Sql = "select * from T01Transaction_Header inner join T02Transaction_Flutter  on T01Ref_No=T02Ref_No  inner join M03Item_Master on T02Item_Code=M03Item_Code  inner join M04Location on M04Loc_Code=T01To_Loc_Code  inner join M09Supplier on M09Code=T01FromLoc_Code where T01Grn_No='" & strCode & "' and T01Trans_Type='GRN' and T01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboLocation.Text = Trim(M01.Tables(0).Rows(0)("M09Name"))
                cboTo.Text = Trim(M01.Tables(0).Rows(0)("M04Loc_Name"))
                txtCom_Invoice.Text = Trim(M01.Tables(0).Rows(0)("T01Invoice_No"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("T01Date"))
                txtEntry.Text = Trim(M01.Tables(0).Rows(0)("T01Grn_No"))
                txtRemark.Text = Trim(M01.Tables(0).Rows(0)("T01Remark"))
                ' _RefNo = Trim(M01.Tables(0).Rows(0)("T01Ref_No"))

                Value = Trim(M01.Tables(0).Rows(0)("T01Net_Amount"))
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Vat"))
                txtVAT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtVAT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01NBT"))
                txtNbt.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNbt.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Market_Return"))
                txtMarket.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMarket.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Com_Discount"))
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtDis_Rate.Text = Trim(M01.Tables(0).Rows(0)("T01DisRate"))

                'Value = Trim(M01.Tables(0).Rows(0)("T01FreeIssue"))
                'txtFree_Amount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtFree_Amount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = (CDbl(txtNett.Text) + CDbl(txtVAT.Text)) - (CDbl(txtDiscount.Text) + Val(txtMarket.Text))
                txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtCount.Text = M01.Tables(0).Rows.Count

                Dim _St As String
                Call Load_Gride2()

                i = 0
                For Each DTRow2 As DataRow In M01.Tables(0).Rows

                    Dim newRow As DataRow = c_dataCustomer3.NewRow
                    newRow("Item Code") = Trim(M01.Tables(0).Rows(i)("M03Item_Code"))
                    newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("M03Item_Name"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Cost Price") = _St
                    newRow("Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    If IsDate(Trim(M01.Tables(0).Rows(i)("T02Ex_Date"))) Then
                        newRow("Ex Date") = Trim(M01.Tables(0).Rows(i)("T02Ex_Date"))
                    End If
                    Value = Trim(M01.Tables(0).Rows(i)("T02Retail_Price"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Retail Price") = _St
                    ' newRow("Rec.Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    newRow("Free Issue") = Trim(M01.Tables(0).Rows(i)("T02Free_Issue"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Qty")) * Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Total") = _St
                    newRow("##") = False
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

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer3 = CustomerDataClass.MakeDataTableGRN
        UltraGrid2.DataSource = c_dataCustomer3
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(6).Width = 80
            .DisplayLayout.Bands(0).Columns(7).Width = 90
            ' .DisplayLayout.Bands(0).Columns(8).Width = 90

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

   
End Class