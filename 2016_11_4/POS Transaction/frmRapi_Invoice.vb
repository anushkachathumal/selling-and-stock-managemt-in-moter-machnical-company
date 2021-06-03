
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmRapi_Invoice
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

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Exit Sub
    End Sub

    Function LOAD_GRIDE()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select T01REF_NO as [##],T01Invoice_No as [Invoice No],t01date as [Date],M17name as [Customer],CAST(T01net_amount AS DECIMAL(16,2)) as [Net Amount],CAST(T01com_discount AS DECIMAL(16,2)) as [Discount],CAST(T03Cash AS DECIMAL(16,2)) as [Cash],CAST(T03credit AS DECIMAL(16,2)) as [Credit],CAST(T03chq AS DECIMAL(16,2)) as [Cheque] from View_SALES_SUMMERY where t01date='" & Today & "' and  T01com_code='" & _Comcode & "' ORDER BY T01DATE DESC "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 70
            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 90
            UltraGrid1.Rows.Band.Columns(3).Width = 170
            UltraGrid1.Rows.Band.Columns(4).Width = 90
            UltraGrid1.Rows.Band.Columns(5).Width = 90
            UltraGrid1.Rows.Band.Columns(6).Width = 90
            UltraGrid1.Rows.Band.Columns(7).Width = 90
            UltraGrid1.Rows.Band.Columns(8).Width = 90
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Function

    Private Sub frmRapi_Invoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        _Comcode = ConfigurationManager.AppSettings("LOCCODE")
        txtGross.ReadOnly = True
        txtGross.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMtr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        ' Call Load_Loading()
        txtDis1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.ReadOnly = True
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.ReadOnly = True
        Call Load_Item()
        Call Load_SALES_REF()
        txtI_Date.Text = Today
        txtI_Add1.ReadOnly = True
        Call Load_Customer()
        'Call Load_Issue()
        Call Load_Gride1()
        txtEntry.ReadOnly = True
        Call Load_Entry1()

        Call Load_Location()

        Call Load_Combo_Name()
        Call Load_Combo_Status()
        'Call Load_Combo_Type()
        Call Load_Supp_Code()
        '  Call Load_Grid()
        txtCode.ReadOnly = True

        ' txtDis1.ReadOnly = True
        txtDiscount.ReadOnly = True
        Call Load_Root()
        Call LOAD_GRIDE()
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
    Function Load_SALES_REF()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Employee_Name AS [##] from M01Employee_Master"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRef
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


    Function Load_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Name AS [Item Name] from M03Item_Master where m03Status='A' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
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

    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M04Location where M04Loc_Name='" & cboLocation.Text & "' and M04Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Location = True
                _Loc_Code = Trim(M01.Tables(0).Rows(0)("M04Loc_Code"))
            End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Location()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [##] from M04Location where M04Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
                ' .Rows.Band.Columns(1).Width = 160


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Customer_Root()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M17Name as [##] from View_Customer_Root where M17Com_Code='" & _Comcode & "' and M02Name='" & cboRoot.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 410
                ' .Rows.Band.Columns(1).Width = 160


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M17Name as [##] from M17Customer where M17Active='A' and M17Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 410
                ' .Rows.Band.Columns(1).Width = 160


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Function Load_Entry1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select  * from P01Parameter where P01Code='SL'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    txtEntry.Text = "INV/SD/-00" & Trim(M01.Tables(0).Rows(0)("P01LastNo"))
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") > 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    txtEntry.Text = "INV/SD-0" & Trim(M01.Tables(0).Rows(0)("P01LastNo"))
                Else
                    txtEntry.Text = "INV/SD-" & Trim(M01.Tables(0).Rows(0)("P01LastNo"))
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

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        OPR1.Visible = True
        cboLocation.ToggleDropdown()
    End Sub

    Function Calculation_Final()
        On Error Resume Next
        Dim _St As String
        Dim Value As Double
        If txtNett.Text <> "" Then
        Else
            txtNett.Text = "0"
        End If

        If txtDiscount.Text <> "" Then
        Else
            txtDiscount.Text = "0"
        End If
        Value = CDbl(txtNett.Text) - CDbl(txtDiscount.Text)

        txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
    End Function

    Function CLEAR_TEXT()
        Me.txtI_Date.Text = Today

        Me.txtI_Add2.Text = ""
        Me.txtI_Add1.Text = ""
        Me.txtGross.Text = ""
        Me.txtDis1.Text = ""
        Me.txtRemark.Text = ""
        'Me.txtI_Test.Text = ""
        Me.txtI_Tp.Text = ""
        ' Me.txtI_Vehicle.Text = ""
        ' Me.txtLaber.Text = ""
        ' Me.txtTransport.Text = ""
        ' Me.cboI_Job.Text = ""
        Me.txtNett.Text = ""
        Me.txtGross.Text = ""
        Me.txtDiscount.Text = ""
        Me.txtGross.Text = ""
        Me.cboName.Text = ""
        'Me.cboI_Job.ToggleDropdown()
        Call Load_Entry1()
        Call Load_Gride()
    End Function

    Private Sub txtTransport_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Calculation_Final()
    End Sub

    Private Sub txtLaber_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Calculation_Final()
    End Sub

    Private Sub cboCustomer_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustomer.AfterCloseUp
        Call Search_Customer()
    End Sub

    Private Sub cboCustomer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCustomer.KeyUp
        If e.KeyCode = 13 Then
            cboItem.ToggleDropdown()
        End If
    End Sub

    Function Search_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Dim _St As String

        Try
            Sql = "select * from M17Customer where M17Name='" & cboCustomer.Text & "' AND M17Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtI_Add1.Text = Trim(M01.Tables(0).Rows(0)("M17Address"))
                txtI_Add2.Text = Trim(M01.Tables(0).Rows(0)("M17Address1"))
                txtI_Tp.Text = Trim(M01.Tables(0).Rows(0)("M17TP"))
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Private Sub txtDiscount_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDiscount.ValueChanged
        On Error Resume Next
        Dim Value As Double
        If IsNumeric(txtDiscount.Text) And IsNumeric(txtNett.Text) Then
            Value = CDbl(txtNett.Text) - CDbl(txtDiscount.Text)
            txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Sub

    Private Sub txtTotal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotal.KeyUp
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Dim _St As String
        Dim _Cost As Double
        Dim _Qty As Double

        Try
            _Qty = 0
            If e.KeyCode = 13 Then
                Sql = "select * from M03Item_Master where M03Item_Name='" & Trim(cboItem.Text) & "' and M03Com_Code='" & _Comcode & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                Else
                    MsgBox("Please enter the Correct Item code", MsgBoxStyle.Information, "Information .........")
                    con.close()
                    Exit Sub
                End If

                Sql = "select * from View_StockBalance1 where S04Item_Code='" & _Itemcode & "' and Rate='" & txtRate.Text & "' and S04Com_Code='" & _Comcode & "' and Qty>0 "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    _Cost = M01.Tables(0).Rows(0)("S04Cost")
                    _Qty = M01.Tables(0).Rows(0)("Qty")
                End If

               
                If txtNett.Text <> "" Then

                Else
                    txtNett.Text = "0"
                End If

                If txtQty.Text <> "" Then
                Else
                    MsgBox("Please enter the Qty", MsgBoxStyle.Information, "Information .......")
                    con.close()
                    Exit Sub
                End If

                If txtRate.Text <> "" Then
                Else
                    MsgBox("Please enter the Rate", MsgBoxStyle.Information, "Information .......")
                    con.close()
                    Exit Sub
                End If


                If IsNumeric(txtQty.Text) Then
                Else
                    MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information .......")
                    con.close()
                    Exit Sub
                End If

                If IsNumeric(txtRate.Text) Then
                Else
                    MsgBox("Please enter the correct Rate", MsgBoxStyle.Information, "Information .......")
                    con.close()
                    Exit Sub
                End If

                If _Qty < CDbl(txtQty.Text) Then
                    MsgBox("Stock Qty less than ented qty", MsgBoxStyle.Information, "Information ........")
                    con.close()
                    Exit Sub
                End If
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = _Itemcode
                newRow("Item Name") = Trim(cboItem.Text)
                Value = _Cost
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St
                Value = txtRate.Text
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St
                newRow("Qty") = txtQty.Text
                newRow("Mtr") = txtMtr.Text
                ' newRow("Rec.Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))

                Value = CDbl(txtRate.Text) * CDbl(txtQty.Text)
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                Value = CDbl(txtNett.Text) + Value
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                c_dataCustomer1.Rows.Add(newRow)
                Call Calculation_Final()
                'txtCode.Text = ""
                txtRate.Text = ""
                txtQty.Text = ""
                txtTotal.Text = ""
                cboItem.Text = ""
                lblDis.Text = ""
                txtMtr.Text = ""
                cboItem.ToggleDropdown()
            End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Sub

    Private Sub txtDis1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDis1.ValueChanged
        On Error Resume Next
        Dim Value As Double
        If IsNumeric(txtDis1.Text) Then
            If IsNumeric(txtNett.Text) Then
                Value = CDbl(txtNett.Text)
                Value = Value * CDbl(txtDis1.Text)
                Value = Value / 100
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
        End If
    End Sub

    Private Sub txtQty_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQty.ValueChanged
        On Error Resume Next
        Dim Value As Double
        If IsNumeric(txtRate.Text) And IsNumeric(txtQty.Text) Then
            Value = CDbl(txtRate.Text) * CDbl(txtQty.Text)
            txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Sub


    Private Sub txtRate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next
        Dim Value As Double
        If IsNumeric(txtRate.Text) And IsNumeric(txtQty.Text) Then
            Value = CDbl(txtRate.Text) * CDbl(txtQty.Text)
            txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        If e.KeyCode = 13 Then
            txtTotal.Focus()
        End If
    End Sub
    Function Load_Gride_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name] from M03Item_Master where M03Com_Code='" & _Comcode & "' order by M03Item_Code"
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

    Function Load_Gride_Item_Find()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name] from M03Item_Master where M03Com_Code='" & _Comcode & "' and M03Item_Name like '%" & txtFind.Text & "%' order by M03Item_Code"
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

    Private Sub cboItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItem.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemName()
            Call Load_Retail_Price(_Itemcode)
            txtRate.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F1 Then
            frmView_Sales.Show()
        ElseIf e.KeyCode = Keys.F2 Then
            OPR5.Visible = True
            Call Load_Gride_Item()
            txtFind.Focus()
        End If
    End Sub

    Private Sub txtRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            txtQty.Focus()
        End If
    End Sub

    Private Sub cboItem_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItem.AfterCloseUp
        Call Search_ItemName()
        Call Load_Retail_Price(_Itemcode)
    End Sub

    Function Search_ItemName()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Try
            Sql = "select *from M03Item_Master where m03Status='A' and M03Item_Name='" & Trim(cboItem.Text) & "' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            If isValidDataset(M01) Then
                ' lblDis.Text = Trim(M01.Tables(0).Rows(0)("M03Item_Name"))
                Value = M01.Tables(0).Rows(0)("M03Retail_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Itemcode = M01.Tables(0).Rows(0)("M03Item_Code")

                Value = M01.Tables(0).Rows(0)("M03MRP")
                txtMtr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMtr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Dim i As Integer
        Dim _St As String

        Try
            Sql = "select *from View_T01Sales where T01Invoice_No='" & Trim(txtEntry.Text) & "' and T01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)

            If isValidDataset(M01) Then
                ' lblDis.Text = Trim(M01.Tables(0).Rows(0)("M03Item_Name"))
                Value = M01.Tables(0).Rows(0)("net")
                txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtNett.Text = txtGross.Text
                _Ref_1 = M01.Tables(0).Rows(0)("T01Ref_No")

                cboCustomer.Text = Trim(M01.Tables(0).Rows(0)("M17Name"))
                Call Search_Customer()
                txtI_Date.Text = Trim(M01.Tables(0).Rows(0)("T01Date"))
                cboRef.Text = Trim(M01.Tables(0).Rows(0)("T01User"))
                cboLocation.Text = Trim(M01.Tables(0).Rows(0)("T01Com_Code"))
            End If
            Call LOAD_GRIDE()
            Sql = "select * from View_T02Transaction where T02Ref_No='" & _Ref_1 & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
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


    Private Sub UltraGrid2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid2.AfterRowsDeleted
        On Error Resume Next
        Dim I As Integer
        Dim Value As Double

        I = 0

        Value = 0

        For Each uRow As UltraGridRow In UltraGrid2.Rows
            Value = Value + CDbl(UltraGrid2.Rows(I).Cells(6).Value)
            I = I + 1
        Next

        txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        Call Calculation_Final()


    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            If Search_Location() = True Then
            Else
                MsgBox("Please select the location", MsgBoxStyle.Information, "Information .......")
                con.close()
                Exit Sub
            End If
            Sql = "select * from M17Customer where M17Name='" & cboCustomer.Text & "' and M17Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
            Else
                MsgBox("Please select the Customer", MsgBoxStyle.Information, "Information ........")
                con.close()
                Exit Sub
            End If
            If UltraGrid2.Rows.Count > 0 Then
            Else
                MsgBox("Please enter the Transaction detailes", MsgBoxStyle.Information, "Information ......")
                con.close()
                Exit Sub
            End If

            If txtDiscount.Text <> "" Then
            Else
                txtDiscount.Text = "0"
            End If
            Call Load_Entry1()
            Call Calculation_Final()

            Call Save_DAte()
            'frmPay_Main.Show()
            'With frmPay_Main
            '    .lblBill.Text = txtGross.Text
            '    .lblBalance.Text = "000.00"
            '    .lblOutstanding.Text = "000.00"
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Sub

    Function Save_DAte()
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
        Dim M01 As DataSet
        Dim _EMP As String
        Dim _ST As String
        Dim A As String
        Dim Value As Double
        '  Dim i As Integer
        Dim _cost As Double
        Dim _Customer As String
        Dim _Cash As Double
        Dim _Remark As String
        Dim B As New ReportDocument
        Dim A1 As String
        Dim _LOC As String
        Dim _RefNo As Integer

        Try

            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = "-"
            End If
           
            _Cash = 0
            '  _Cash = CDbl(lblBill.Text) - CDbl(txtTotal_C.Text)

            nvcFieldList1 = "SELECT * FROM P01Parameter WHERE P01Code='IN'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = M01.Tables(0).Rows(0)("P01LastNo")
            End If

            'nvcFieldList1 = "SELECT * FROM M04Location WHERE M04Loc_Name='" & _VEHICLEnO & "'"
            'M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            'If isValidDataset(M01) Then
            '    _LOC = Trim(M01.Tables(0).Rows(0)("M04Loc_Code"))
            'End If

            nvcFieldList1 = "SELECT * FROM M17Customer WHERE M17Name='" & cboCustomer.Text & "' and M17Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _Customer = Trim(M01.Tables(0).Rows(0)("M17Code"))
            End If
            If txtDiscount.Text <> "" Then
            Else
                txtDiscount.Text = "0"
            End If
            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo +" & 1 & " WHERE P01Code='IN'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo +" & 1 & " WHERE P01Code='SL'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            i = 0



            For Each uRow As UltraGridRow In UltraGrid2.Rows

                'nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & UltraGrid2.Rows(i).Cells(0).Text & "' AND M03Com_Code='" & _Comcode & "'"
                'M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(M01) Then
                '    _cost = M01.Tables(0).Rows(0)("M03Cost_Price")
                'End If

                nvcFieldList1 = "Insert T02Transaction_flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Qty,T02Free_Issue,T02Total,T02Status,T02Count,T02Com_Code,T02MTR)" & _
                                                                      " values('" & _RefNo & "', '" & UltraGrid2.Rows(i).Cells(0).Text & "','" & UltraGrid2.Rows(i).Cells(2).Text & "','" & UltraGrid2.Rows(i).Cells(3).Text & "','" & UltraGrid2.Rows(i).Cells(5).Text & "','0','" & UltraGrid2.Rows(i).Cells(6).Text & "','A','" & i + 1 & "','" & _Comcode & "','" & UltraGrid2.Rows(i).Cells(4).Text & "' )"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Ref_No,S01Com_Code,S01STATUS)" & _
                                                     " values('" & _Loc_Code & "', '" & (UltraGrid2.Rows(i).Cells(0).Value) & "','" & txtI_Date.Text & "','DR','" & -(UltraGrid2.Rows(i).Cells(5).Value) & "','" & _RefNo & "','" & _Comcode & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "Insert Into S04Stock_Price(S04Location,S04Tr_Type,S04Date,S04Item_Code,S04Qty,S04Ref_No,S04Status,S04Rate,S04Cost,S04Com_Code)" & _
                '                                               " values('" & _Loc_Code & "','DR','" & txtI_Date.Text & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & -(UltraGrid1.Rows(i).Cells(5).Text) & "','" & _RefNo & "','A','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & _Comcode & "')"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S04Stock_Price(S04Location,S04Tr_Type,S04Date,S04Item_Code,S04Qty,S04Ref_No,S04Status,S04Rate,S04Cost,S04Com_Code)" & _
                                                            " values('" & _Loc_Code & "','DR','" & txtI_Date.Text & "', '" & (UltraGrid2.Rows(i).Cells(0).Value) & "','" & -(UltraGrid2.Rows(i).Cells(5).Text) & "','" & _RefNo & "','A','" & (UltraGrid2.Rows(i).Cells(3).Value) & "','" & (UltraGrid2.Rows(i).Cells(2).Value) & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                i = i + 1
            Next

            'Transaction Header
            nvcFieldList1 = "Insert T01Transaction_Header(T01Trans_Type,T01Ref_No,T01Date,T01Invoice_No,T01PO_NO,T01Grn_No,T01FromLoc_Code,T01Customer,T01Net_Amount,T01Com_Code,T01User,T01Transport,T01Status,T01Paid,T01Time,T01Br_Charge,T01Remark,T01Km,T01Com_Discount,T01Terminal)" & _
                                                                     " values('DR', '" & _RefNo & "','" & txtI_Date.Text & "','" & txtEntry.Text & "','-','-','" & _Loc_Code & "','" & _Customer & "','" & txtGross.Text & "','" & _Comcode & "','" & cboRef.Text & "','0','A','" & txtGross.Text & "','" & Now & "','0','" & txtRemark.Text & "','0','" & txtDiscount.Text & "','-' )"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert T03Pay_Main(T03Ref_No,T03Trans_Type,T03Net_Amt,T03Credit,T03Cash,T03Chq,T03Status,T03Com_Code)" & _
                                                                    " values( '" & _RefNo & "','DR','" & txtGross.Text & "','" & txtGross.Text & "','0','0','A','" & _Comcode & "' )"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            '  If Value > 0 Then
            _Remark = "Credit Invoice-" & txtEntry.Text
            nvcFieldList1 = "Insert T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Invo,T05Com_Code,T05User,T05Status)" & _
                                                              " values( '" & _RefNo & "','DR','" & txtI_Date.Text & "','" & _Customer & "','" & _Remark & "','" & txtGross.Text & "','0','" & txtEntry.Text & "','" & _Comcode & "','" & strDisname & "','A' )"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert t06OutStanding_Balance(T06RefNo,T06Cus_Code,T06Date,T06Bill_Amount,T06Invoice_No,T06Pay_Amount,T06Pay_RefNo,T06Remark,T06Com_Code,T06Status,T06Tr_Type)" & _
                                                          " values( '" & _RefNo & "','" & _Customer & "','" & txtI_Date.Text & "','" & txtGross.Text & "','" & txtEntry.Text & "','0','0','" & _Remark & "','" & _Comcode & "','A','DR' )"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            '  End If



            transaction.Commit()
            connection.Close()
            A = MsgBox("Are you sure you want to print Invoice", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Invoice ........")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\InvoTm.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                ' B.SetParameterValue("Customer", cboCustomer.Text)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_T01Sales.T01Ref_No}=" & _RefNo & ""
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            ' Call Load_Entry1()
            Call Load_Entry1()
            Call Load_Gride()
            ' Me.txtDoc.Text = ""
            Me.cboCustomer.Text = ""
            Me.txtI_Add1.Text = ""
            Me.txtI_Add2.Text = ""
            Me.txtI_Tp.Text = ""
            Me.txtDis1.Text = ""
            Me.txtGross.Text = ""
            Me.txtNett.Text = ""
            Me.txtDiscount.Text = ""

            Call Load_Gride1()
            cboCustomer.ToggleDropdown()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

   

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call CLEAR_TEXT()
        strSales_Status = False
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        OPR1.Visible = False
        Call CLEAR_TEXT()
        Call Load_Entry1()
        Call Load_Gride1()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        OPR1.Visible = False
        Call CLEAR_TEXT()
        Call Load_Entry1()
        Call Load_Gride1()
        ' CancelInvoiceToolStripMenuItem.Checked = False
        Call Claer_Text()
        _EDITSTATUS = False
    End Sub


    Function Load_Supp_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='CUM'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") <= 10 Then
                    txtCode.Text = "CU/SD/00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") > 10 And M01.Tables(0).Rows(0)("P01LastNo") <= 100 Then
                    txtCode.Text = "CU/SD/0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtCode.Text = "CU/SD/" & M01.Tables(0).Rows(0)("P01LastNo")
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


    Private Sub cboStatus_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboStatus.KeyUp
        If e.KeyCode = 13 Then
            If cboStatus.Text <> "" Then
                cboName.ToggleDropdown()
            End If
        ElseIf e.KeyCode = Keys.F1 Then
            'OPR4.Visible = True
            'strWindowName = "frmNewCustomer"
            'txtFind.Focus()
        End If
    End Sub

    Private Sub cboName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboName.KeyUp
        If e.KeyCode = 13 Then
            txtAddress.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            'OPR4.Visible = True
            'strWindowName = "frmNewCustomer"
            'txtFind.Focus()
        End If
    End Sub

    Private Sub txtAddress_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyUp
        If e.KeyCode = 13 Then
            txtAdd1.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            'OPR4.Visible = True
            'strWindowName = "frmNewCustomer"
            'txtFind.Focus()
        End If
    End Sub

    Private Sub txtAdd1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAdd1.KeyUp
        If e.KeyCode = 13 Then
            txtTp.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            'OPR4.Visible = True
            'strWindowName = "frmNewCustomer"
            'txtFind.Focus()
        End If
    End Sub

    Private Sub txtTp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTp.KeyUp
        If e.KeyCode = 13 Then
            txtFax.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            'OPR4.Visible = True
            'strWindowName = "frmNewCustomer"
            'txtFind.Focus()
        End If
    End Sub

    Private Sub txtFax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFax.KeyUp
        If e.KeyCode = 13 Then
            txtContact.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            'OPR4.Visible = True
            'strWindowName = "frmNewCustomer"
            'txtFind.Focus()
        End If
    End Sub

    Private Sub txtContact_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtContact.KeyUp
        If e.KeyCode = 13 Then
            txtVAT.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            'OPR4.Visible = True
            'strWindowName = "frmNewCustomer"
            'txtFind.Focus()
        End If
    End Sub

    Private Sub txtVAT_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVAT.KeyUp
        If e.KeyCode = 13 Then
            cmdSave.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            'OPR4.Visible = True
            'strWindowName = "frmNewCustomer"
            'txtFind.Focus()
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim t01 As DataSet

        Try
            If cboName.Text <> "" Then
            Else
                MsgBox("Please enter the Customer Name", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                cboName.ToggleDropdown()
                Exit Sub
            End If

            If cboStatus.Text <> "" Then
            Else
                MsgBox("Please enter the Status", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                cboStatus.ToggleDropdown()
                Exit Sub
            End If

            'If cboType.Text <> "" Then
            'Else
            '    MsgBox("Please enter the Type", MsgBoxStyle.Information, "Information .......")
            '    connection.Close()
            '    cboType.ToggleDropdown()
            '    Exit Sub
            'End If


            If txtAdd1.Text <> "" Then
            Else
                txtAdd1.Text = " "
            End If


            If txtAddress.Text <> "" Then
            Else
                txtAddress.Text = " "
            End If

            If txtVAT.Text <> "" Then
            Else
                txtVAT.Text = " "
            End If

            If txtContact.Text <> "" Then
            Else
                txtContact.Text = " "
            End If

            If txtTp.Text <> "" Then
            Else
                txtTp.Text = " "
            End If

            If txtFax.Text <> "" Then
            Else
                txtFax.Text = " "
            End If

            nvcFieldList1 = "SELECT * FROM M17Customer where M17Code='" & txtCode.Text & "' AND M17Com_Code='" & _Comcode & "'"
            t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(t01) Then

                nvcFieldList1 = "UPDATE M17Customer SET M17Status='" & cboStatus.Text & "',M17Name='" & cboName.Text & "',M17Address='" & txtAddress.Text & "',M17Address1='" & txtAdd1.Text & "',M17TP='" & txtTp.Text & "',M17VAT='" & txtVAT.Text & "',M17Fax='" & txtFax.Text & "',M17Contact_On='" & txtContact.Text & "',M17Active='A' WHERE M17Code='" & txtCode.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE M01Account_Master SET M01Acc_Name='" & cboName.Text & "',M01Address='" & txtAddress.Text & "',M01Address2='" & txtAdd1.Text & "',M01TP='" & txtTp.Text & "',M01Status='A' WHERE M01Acc_Code='" & txtCode.Text & "' AND M01Acc_Type='CU'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            Else
                Call Load_Supp_Code()

                nvcFieldList1 = "UPDATE P01PARAMETER SET P01LastNo=P01LastNo +" & 1 & " WHERE P01CODE='CUM'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M17Customer(M17Code,M17Status,M17Name,M17Address,M17Address1,M17TP,M17VAT,M17Fax,M17Contact_On,M17Time,M17User,M17Active,M17Com_Code)" & _
                                                                  " values('" & Trim(txtCode.Text) & "', '" & Trim(cboStatus.Text) & "','" & Trim(cboName.Text) & "','" & txtAddress.Text & "','" & txtAdd1.Text & "','" & txtTp.Text & "','" & txtVAT.Text & "','" & txtFax.Text & "','" & txtContact.Text & "','" & Now & "','" & strDisname & "','A','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M01Account_Master(M01Acc_Type,M01Acc_Code,M01Acc_Name,M01Address,M01Address2,M01TP,M01Acc_Limit,M01DOC,M01User,M01Status,M01year,M01Comm,M01Com_Code,M01ACC_OF,M01OB_Chq)" & _
                                                                  " values('CU', '" & Trim(txtCode.Text) & "','" & Trim(cboName.Text) & "','" & txtAddress.Text & "','" & txtAdd1.Text & "','" & txtTp.Text & "','0','" & Today & "','" & strDisname & "','A','" & Year(Today) & "','0','" & _Comcode & "','" & _Comcode & "','0')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()

            Call Claer_Text()
            Call Load_Supp_Code()
            Call Load_Combo_Name()
            cboStatus.ToggleDropdown()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Function Load_Root()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M02Name as [##] from M02Root where M02Status='A' AND m02Come_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRoot
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 370
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


    Function Load_Combo_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M17Name as [##] from M17Customer where M17Active='A' AND M17Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 370
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
    Function Load_Combo_Status()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M10Dis as [##] from M10Status  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboStatus
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 90
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
    Function Claer_Text()
        Me.txtCode.Text = ""
        Me.txtContact.Text = ""
        Me.txtAdd1.Text = ""
        Me.txtAddress.Text = ""
        Me.txtVAT.Text = ""
        Me.txtFax.Text = ""
        Me.txtTp.Text = ""
        Me.cboName.Text = ""
        Me.cboStatus.Text = ""
        Me.txtMtr.Text = ""
        'Me.lblDis.Text = ""
        ' Me.cboType.Text = ""
        'Call Load_Grid()
    End Function

    Private Sub CreateNewCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        OPR0.Visible = True
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Claer_Text()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        OPR0.Visible = False
    End Sub

    Private Sub CancelInvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' CancelInvoiceToolStripMenuItem.Checked = True
        _EDITSTATUS = True

    End Sub

   

    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            cboRef.ToggleDropdown()
        End If
    End Sub

    Private Sub cboRef_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboRef.KeyUp
        If e.KeyCode = 13 Then
            cboRoot.ToggleDropdown()
        End If
    End Sub

    Private Sub cboRoot_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoot.AfterCloseUp
        Call Load_Customer_Root()
    End Sub

   
    Private Sub cboRef_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboRef.InitializeLayout

    End Sub

    Private Sub cboRoot_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboRoot.KeyUp
        If e.KeyCode = 13 Then
            cboCustomer.ToggleDropdown()
        End If
    End Sub

    Function Load_Retail_Price(ByVal strCode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Call Search_Location()
            Sql = "select CAST(Rate AS DECIMAL(16,2)) as [##Retail],max(CAST(S04Cost AS DECIMAL(16,2))) as [##Cost],SUM(Qty) as Qty from View_StockBalance1 where S04Item_Code='" & strCode & "' and S04Com_Code='" & _Comcode & "' and qty>0 and S04Location='" & _Loc_Code & "' group by rate"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtRate
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
                .Rows.Band.Columns(1).Width = 110
                .Rows.Band.Columns(2).Width = 90

            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function


    Private Sub txtRate_KeyUp1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyUp
        If e.KeyCode = 13 Then
            txtMtr.Focus()
        End If
    End Sub

    Private Sub txtMtr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMtr.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtMtr.Text) Then
                Value = txtMtr.Text
                txtMtr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMtr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtQty.Focus()
        End If
    End Sub

  

    Private Sub txtFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyUp
        If e.KeyCode = 13 Then
            UltraGrid3.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
        End If
    End Sub

    Private Sub txtFind_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFind.ValueChanged
        Call Load_Gride_Item_Find()
    End Sub

    Private Sub UltraGrid3_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid3.DoubleClickRow
        On Error Resume Next
        Dim _Rowindex As Integer
        _Rowindex = UltraGrid1.ActiveRow.Index


        cboItem.Text = Trim(UltraGrid3.Rows(_Rowindex).Cells(1).Text)
        Search_ItemName()
        OPR5.Visible = False
    End Sub


    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer
        _Row = UltraGrid1.ActiveRow.Index
        txtEntry.Text = UltraGrid1.Rows(_Row).Cells(1).Text
        OPR1.Visible = True
        Call Search_Records()
    End Sub

   
    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument
        Try
            A = MsgBox("Are you sure you want to print Invoice", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Invoice ........")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\InvoTm.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                ' B.SetParameterValue("Customer", cboCustomer.Text)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_T01Sales.T01Ref_No}=" & _Ref_1 & ""
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

            End If
        End Try
    End Sub

    Private Sub cboItem_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboItem.InitializeLayout

    End Sub

    Private Sub txtTotal_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotal.ValueChanged

    End Sub

    Private Sub UltraGrid2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid2.InitializeLayout

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

    End Sub
End Class