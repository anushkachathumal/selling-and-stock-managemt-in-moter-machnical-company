Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmGRN_uniq
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Supplier As String
    Dim _Location As Integer
    Dim _LogStaus As Boolean
    Dim _UserLevel As String
    Dim _Itemcode As String

    Private Sub frmGRN_uniq_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmSupplier_Cnt.Close()
        frmItem_cnt.Close()
        frmView_GRN_cnt.Close()
        frmView_Items_cnt.Close()
    End Sub

    Function Load_Grid_ROW()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select m05ref_no as  ##, max(m05item_code) as [Part No],max(M05Brand_Name) as [Brand Name],MAX(tmpDescription) as [Description],max(CAST(Retail AS DECIMAL(16,2))) as [Retail Price],sum(qty) as [Current Stock],max(rack) as [Rack No],max(cell) as [Cell No] from View_Product_Stock  group by m05ref_no having  max(m05item_code) like '" & Trim(cboCode.Text) & "%'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            UltraGrid2.Rows.Band.Columns(0).Width = 40

            UltraGrid2.Rows.Band.Columns(1).Width = 90
            UltraGrid2.Rows.Band.Columns(2).Width = 110
            UltraGrid2.Rows.Band.Columns(3).Width = 210
            UltraGrid2.Rows.Band.Columns(4).Width = 80
            UltraGrid2.Rows.Band.Columns(5).Width = 80
            UltraGrid2.Rows.Band.Columns(6).Width = 80
            UltraGrid2.Rows.Band.Columns(7).Width = 80
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid2.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid2.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
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


    Private Sub frmGRN_uniq_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCell.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRack.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.ReadOnly = True
        txtSales.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        ' Call Load_Item()
        txtName.ReadOnly = True
        Call Load_Gride2()
        Call Load_Supplier()
        Call Load_Location()
        Call Load_EntryNo()
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtGross.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtDis_Rate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtVAT.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtNBT.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDiscount.ReadOnly = True
        txtNett.ReadOnly = True
        txtGross.ReadOnly = True
        txtDate.Text = Today
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim i As Integer
        Dim Value As Double
        Dim _St As String
        Dim M02 As DataSet

        Try
            Sql = "select *  from View_GRN_Header where T01Ref_No='" & txtEntry.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtDate.Text = M01.Tables(0).Rows(0)("T01Date")
                txtCom_Invoice.Text = M01.Tables(0).Rows(0)("T01Com_Invoice")
                cboLocation.Text = M01.Tables(0).Rows(0)("M04Name")
                cboTo.Text = M01.Tables(0).Rows(0)("M11Name")
                Value = M01.Tables(0).Rows(0)("T01Net_Amount")
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M01.Tables(0).Rows(0)("T01Dis_Rate")
                txtDis_Rate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDis_Rate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M01.Tables(0).Rows(0)("Discount")
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))


                Value = M01.Tables(0).Rows(0)("Gross")
                txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                cmdDelete.Enabled = True
                cmdEdit.Enabled = True
            End If

            Sql = "select * from View_GRN_Flutter where T01Ref_No='" & Trim(txtEntry.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("#Ref.No") = Trim(M01.Tables(0).Rows(i)("T02Item_Ref"))
                newRow("#Part No") = Trim(M01.Tables(0).Rows(i)("M05Item_Code"))
                newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("tmpDescription"))
                Value = Trim(M01.Tables(0).Rows(i)("T02Cost"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Cost Price") = _St
                newRow("Qty") = CInt(M01.Tables(0).Rows(i)("T02Qty"))

                Value = Trim(M01.Tables(0).Rows(i)("T02Cost"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                Value = Trim(M01.Tables(0).Rows(i)("T02Retail"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St

                'newRow("Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))

                Sql = "select * from M12Store_Location where M12Item_Code='" & Trim(M01.Tables(0).Rows(i)("T02Item_Ref")) & "'"
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    newRow("#Rack No") = Trim(M02.Tables(0).Rows(0)("M12Rack"))
                    newRow("Cell No") = Trim(M02.Tables(0).Rows(0)("M12Cell"))
                End If
                Value = Trim(M01.Tables(0).Rows(i)("Total"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                newRow("##") = False
                c_dataCustomer1.Rows.Add(newRow)


                i = i + 1
            Next

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
    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGRN
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 210
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(7).Width = 80
            .DisplayLayout.Bands(0).Columns(8).Width = 90
            .DisplayLayout.Bands(0).Columns(9).Width = 40

            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(8).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(8).CellActivation = Activation.NoEdit

            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M05Item_Code as [Part No],M05Description as [Item Name] from M05Item_Master where M05Status='A' order by M05ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCode
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 212
                .Rows.Band.Columns(1).Width = 360

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

    Function Serch_Item() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select *  from View_Product_Item where M05Status='A' and M05Ref_No='" & _Itemcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Serch_Item = True
                cboCode.Text = Trim(M01.Tables(0).Rows(0)("M05Item_Code"))
                txtName.Text = Trim(M01.Tables(0).Rows(0)("tmpDescription"))
                Value = Trim(M01.Tables(0).Rows(0)("M05Cost"))
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("M05Retail"))
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            End If

            Sql = "select * from M12Store_Location where M12Item_Code='" & _Itemcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCell.Text = Trim(M01.Tables(0).Rows(0)("M12Cell"))
                txtRack.Text = Trim(M01.Tables(0).Rows(0)("M12Rack"))
            End If
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

    Private Sub cboCode_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.AfterCloseUp
        'Call Serch_Item()
    End Sub

    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select M04Name as [##] from M04Supplier where  M04Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 296
                '  .Rows.Band.Columns(1).Width = 160


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

    Function Search_Supplier() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Code from M04Supplier where  M04Status='A'  and M04Name='" & Trim(cboLocation.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Supplier = True
                _Supplier = Trim(M01.Tables(0).Rows(0)("M04Code"))
            End If
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

    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11ID from M11Common WHERE M11Status='LC' and M11Name='" & Trim(cboTo.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Location = True
                _Location = Trim(M01.Tables(0).Rows(0)("M11ID"))
            End If
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

    Function Load_Location()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Name as [##] from M11Common WHERE M11Status='LC'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboTo
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 283
                '  .Rows.Band.Columns(1).Width = 160


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

    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            txtCom_Invoice.Focus()
        End If
    End Sub

    Private Sub txtCom_Invoice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCom_Invoice.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtCom_Invoice.Text) <> "" Then
                cboTo.ToggleDropdown()
            End If
        End If
    End Sub

    Private Sub cboTo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboTo.KeyUp
        If e.KeyCode = 13 Then
            cboCode.ToggleDropdown()

        End If
    End Sub

    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCode.KeyUp
        If e.KeyCode = 13 Then
            If Trim(cboCode.Text) <> "" Then
                If UltraGrid2.Visible = True Then
                    UltraGrid2.Focus()
                Else
                    txtRate.Focus()
                End If
            End If
        ElseIf e.KeyCode = Keys.Escape Then
            UltraGrid2.Visible = False
        End If
    End Sub

    Private Sub txtRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtRate.Text) Then
                Value = txtRate.Text
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtSales.Focus()
        End If
    End Sub

    Private Sub txtSales_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSales.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtSales.Text) Then
                Value = txtSales.Text
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtQty.Focus()
        End If
    End Sub

    Private Sub txtRack_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRack.KeyUp
        If e.KeyCode = 13 Then
            txtCell.Focus()
        End If
    End Sub

    Function Calculating_Total()
        Dim Value As Double
        If IsNumeric(txtRate.Text) And IsNumeric(txtQty.Text) Then
            Value = CDbl(txtRate.Text) * CDbl(txtQty.Text)
            txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Function

    Private Sub txtRate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRate.TextChanged
        Call Calculating_Total()
    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        If e.KeyCode = 13 Then
            If txtQty.Text <> "" Then
                txtRack.Focus()
            End If
        End If
    End Sub

    Private Sub txtQty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQty.TextChanged
        Call Calculating_Total()
    End Sub

    Private Sub txtCell_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCell.KeyUp
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim result1 As String
        Dim T01 As DataSet

        Try
            If e.KeyCode = 13 Then
                Sql = "select * from M05Item_Master WHERE M05Status='A' and M05Item_Code='" & Trim(cboCode.Text) & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                Else
                    MsgBox("Please enter the correct Part No", MsgBoxStyle.Information, "Information .......")
                    Exit Sub
                End If
                '=============================================================================================
                If Trim(txtRate.Text) <> "" Then

                    If IsNumeric(txtRate.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the correct Cost Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtRate.Focus()
                            Exit Sub
                        End If
                    End If
                End If
                '=============================================================================================
                If Trim(txtSales.Text) <> "" Then

                    If IsNumeric(txtSales.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the correct Retail Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtSales.Focus()
                            Exit Sub
                        End If
                    End If
                End If
                '=============================================================================================
                If Trim(txtRack.Text) <> "" Then
                Else
                    MsgBox("Please enter the Rack No", MsgBoxStyle.Information, "Information ........")
                    txtRack.Focus()
                    Exit Sub
                End If
                '=============================================================================================
                If Trim(txtCell.Text) <> "" Then
                Else
                    MsgBox("Please enter the Cell No", MsgBoxStyle.Information, "Information ........")
                    txtCell.Focus()
                    Exit Sub
                End If
                '=============================================================================================
                If Trim(txtQty.Text) <> "" Then

                    If IsNumeric(txtQty.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the Correct Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtQty.Focus()
                            Exit Sub
                        End If
                    End If
                End If
                '=============================================================================================
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("#Ref.No") = _Itemcode
                newRow("#Part No") = Trim(cboCode.Text)
                newRow("Item Name") = txtName.Text
                newRow("Cost Price") = txtRate.Text
                newRow("Retail Price") = txtSales.Text
                newRow("Qty") = txtQty.Text
                ' newRow("Rec.Qty") = txtRe_Qty.Text
                newRow("#Rack No") = UCase(txtRack.Text)
                newRow("Cell No") = UCase(txtCell.Text)
                newRow("Total") = txtTotal.Text
               
                Sql = "select * from M05Item_Master where M05Item_Code='" & cboCode.Text & "' and M05Status='A' "
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    If CDbl(T01.Tables(0).Rows(0)("M05Cost")) <> CDbl(txtRate.Text) Then
                        result1 = MsgBox("Previous cost price is Rs." & CDbl(T01.Tables(0).Rows(0)("M05Cost")) & ".do you want to change new one", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Information ........")
                        If result1 = vbYes Then
                            newRow("##") = True
                            '_CostStatus = True
                        Else
                            newRow("##") = False
                        End If
                    Else
                        newRow("##") = False
                    End If
                End If

              
                c_dataCustomer1.Rows.Add(newRow)
                txtCount.Text = UltraGrid1.Rows.Count
                Call Calculation_Bill_Pay()
                Me.cboCode.Text = ""
                Me.txtName.Text = ""
                Me.txtRack.Text = ""
                Me.txtCell.Text = ""
                Me.txtRate.Text = ""
                Me.txtSales.Text = ""
                Me.txtTotal.Text = ""
                Me.txtQty.Text = ""
                _Itemcode = ""
                cboCode.ToggleDropdown()
            End If
            con.ClearAllPools()
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Sub
    Function Calculating_Discount()
        Dim Value As Double

        If IsNumeric(txtDis_Rate.Text) And IsNumeric(txtNett.Text) Then
            Value = CDbl(txtNett.Text) * CDbl(txtDis_Rate.Text)
            Value = Value / 100
            txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            Value = CDbl(txtNett.Text) - CDbl(txtDiscount.Text)
            txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Function
    Function Calculation_Bill_Pay()
        Dim i As Integer
        Dim Value As Double

        i = 0
        For Each uRow As UltraGridRow In UltraGrid1.Rows
            If IsNumeric(UltraGrid1.Rows(i).Cells(8).Text) Then
                Value = Value + CDbl(UltraGrid1.Rows(i).Cells(8).Text)
            End If
            i = i + 1
        Next
        Call Calculating_Discount()
        If txtDiscount.Text <> "" Then
            If IsNumeric(txtDiscount.Text) Then
            Else
                txtDiscount.Text = "0"
            End If
        Else
            txtDiscount.Text = "0"
        End If
        txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        Value = CDbl(txtNett.Text) - CDbl(txtDiscount.Text)
        txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
    End Function
    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='GR'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtEntry.Text = "GRN-00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtEntry.Text = "GRN-0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtEntry.Text = "GRN-" & M01.Tables(0).Rows(0)("P01No")
                End If
            End If

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

    Private Sub txtDis_Rate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDis_Rate.TextChanged
        Call Calculating_Discount()
    End Sub

    Private Sub UltraGrid1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.AfterRowsDeleted
        Call Calculation_Bill_Pay()
        txtCount.Text = UltraGrid1.Rows.Count
    End Sub

    Function Clear_text()
        Me.txtDiscount.Text = ""
        Me.txtGross.Text = ""
        Me.txtNett.Text = ""
        Me.txtNBT.Text = ""
        Me.txtVAT.Text = ""
        Me.txtPO.Text = ""
        Me.txtQty.Text = ""
        Me.txtName.Text = ""
        Me.txtCell.Text = ""
        Me.txtRate.Text = ""
        Me.txtRack.Text = ""
        Me.cboCode.Text = ""
        Me.txtCom_Invoice.Text = ""
        Me.txtCount.Text = ""
        Me.txtTotal.Text = ""
        Me.cboLocation.Text = ""
        Me.cboTo.Text = ""
        Me.txtDis_Rate.Text = ""
        _Itemcode = ""
        Me.UltraGrid2.Visible = False
        OPRUser.Visible = False
        txtUserName.Text = ""
        txtPassword.Text = ""
        _LogStaus = False
        _LogStaus = False
        Call Load_EntryNo()
        Call Load_Gride2()
        cboLocation.ToggleDropdown()

    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_text()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Search_Supplier() = True Then
        Else
            MsgBox("Please select the Supplier Name", MsgBoxStyle.Information, "Informaation ........")
            cboLocation.ToggleDropdown()
            Exit Sub
        End If

        If Search_Location() = True Then
        Else
            MsgBox("Please select the Location", MsgBoxStyle.Information, "Informaation ........")
            cboTo.ToggleDropdown()
            Exit Sub
        End If

        If Trim(txtCom_Invoice.Text) <> "" Then
        Else
            MsgBox("Please enter the Company Invoice No", MsgBoxStyle.Information, "Information .....")
            txtCom_Invoice.Focus()
            Exit Sub
        End If

        If txtPO.Text <> "" Then
        Else
            txtPO.Text = "-"
        End If
        If UltraGrid1.Rows.Count > 0 Then
        Else
            MsgBox("Please enter the transaction Details", MsgBoxStyle.Information, "Information ......")
            cboCode.ToggleDropdown()
        End If

        If txtDis_Rate.Text <> "" Then
        Else
            txtDis_Rate.Text = "0"
        End If

        If IsNumeric(txtDis_Rate.Text) Then
        Else
            MsgBox("Please enter the correct Discount Rate", MsgBoxStyle.Information, "Information .......")
            txtDis_Rate.Focus()
        End If
        If Trim(txtRemark.Text) <> "" Then
        Else
            txtRemark.Text = "-"
        End If

        If txtVAT.Text <> "" Then
        Else
            txtVAT.Text = "0"
        End If
        If txtNBT.Text <> "" Then
        Else
            txtNBT.Text = "0"
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
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim i As Integer
        Dim A As String
        Dim B As New ReportDocument
        Try
            Call Load_EntryNo()
            _GetDate = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text)

            _Get_Time = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text) & " " & Hour(Now) & ":" & Minute(Now)

            nvcFieldList1 = "select* from T01GRN_Header  where T01Ref_No='" & Trim(txtEntry.Text) & "'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                MsgBox("This entry no alrady exsist", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                Exit Function
            Else

                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='GR' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'SAVE DATA ON T01GRN_HEADER
                nvcFieldList1 = "Insert Into T01GRN_Header(T01Ref_No,T01Date,T01Time,T01Sup_Code,T01Loc_Code,T01Net_Amount,T01Dis_Rate,T01VAT,T01NBT,T01Remark,T01User,T01Status,T01Com_Invoice,T01PO,T01Tr_Type)" & _
                                                                    " values('" & Trim(txtEntry.Text) & "','" & _GetDate & "', '" & _Get_Time & "','" & _Supplier & "','" & _Location & "','" & CDbl(txtNett.Text) & "','" & Trim(txtDis_Rate.Text) & "','" & Trim(txtVAT.Text) & "','" & txtNBT.Text & "','" & Trim(txtRemark.Text) & "','" & strDisname & "','A','" & Trim(UCase(txtCom_Invoice.Text)) & "','" & txtPO.Text & "','GRN')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '===================================================================================
                'SAVE DATA ON T03SUPPLIER_TR_ACCOUNT
                nvcFieldList1 = "Insert Into T03Supplier_Tr_Account(T03Pay_No,T03Tr_Type,T03Date,T03Time,T03Sup_Code,T03Ref_No,T03Cash,T03Chq,T03Credit,T03Status)" & _
                                                                    " values('-','GRN','" & _GetDate & "', '" & _Get_Time & "','" & _Supplier & "','" & Trim(txtEntry.Text) & "','0','0','" & CDbl(txtGross.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '===================================================================================
                'SAVE DATA ON TMPTRANSACTION_LOG
                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                  " values('GRN','SAVE', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                i = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    'SAVE S01STOCK_BALANCE
                    nvcFieldList1 = "Insert Into S01Stock_Balance(S01Item_Code,S01Ref_No,S01Date,S01Time,S01Tr_Type,S01Qty,S01Status)" & _
                                                             " values('" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "' ,'" & Trim(txtEntry.Text) & "', '" & _GetDate & "','" & _Get_Time & "','GRN','" & Trim(UltraGrid1.Rows(i).Cells(5).Text) & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    '========================================================================
                    'SAVE T02GRN_FLUTTER
                    nvcFieldList1 = "Insert Into T02GRN_Flutter(T02Ref_No,T02Item_Ref,T02Part_No,T02Cost,T02Retail,T02Qty,T02Status)" & _
                                                           " values('" & Trim(txtEntry.Text) & "' ,'" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(1).Text) & "', '" & Trim(UltraGrid1.Rows(i).Cells(3).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(4).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(5).Text) & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    '========================================================================
                    'RACK NO
                    nvcFieldList1 = "select * from M12Store_Location where M12Item_Code='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "'"
                    MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(MB51) Then
                        nvcFieldList1 = "update M12Store_Location set M12Rack='" & Trim(UltraGrid1.Rows(i).Cells(6).Text) & "',M12Cell='" & Trim(UltraGrid1.Rows(i).Cells(7).Text) & "' where M12Item_Code='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "' "
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into M12Store_Location(M12Item_Code,M12Rack,M12Cell)" & _
                                                                          " values('" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(6).Text) & "', '" & Trim(UltraGrid1.Rows(i).Cells(7).Text) & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    i = i + 1
                Next
                transaction.Commit()
                connection.ClearAllPools()
                connection.Close()
                SqlClient.SqlConnection.ClearAllPools()
                A = MsgBox("Are you sure you want to print GRN Dispatch Note", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Dispatch Note .......")
                If A = vbYes Then
                    A = ConfigurationManager.AppSettings("ReportPath") + "\GRNDispatch.rpt"
                    B.Load(A.ToString)
                    B.SetDatabaseLogon("sa", "sainfinity")
                    'B.SetParameterValue("To", _To)
                    'B.SetParameterValue("From", _From)
                    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                    frmReport.CrystalReportViewer1.DisplayToolbar = True
                    frmReport.CrystalReportViewer1.SelectionFormula = "{T01GRN_Header.T01Ref_No}='" & Trim(txtEntry.Text) & "' "
                    frmReport.Refresh()
                    ' frmReport.CrystalReportViewer1.PrintReport()
                    ' B.PrintToPrinter(1, True, 0, 0)
                    frmReport.MdiParent = MDIMain
                    frmReport.Show()
                End If
                Call Load_EntryNo()
                Call Clear_text()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub txtDis_Rate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDis_Rate.ValueChanged
        If IsNumeric(txtDis_Rate.Text) Then
            Call Calculating_Discount()
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        frmSupplier_Cnt.Close()
        frmSupplier_Cnt.Show()
    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        frmItem_cnt.Close()
        frmItem_cnt.Show()
    End Sub

    Private Sub DeactivateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeactivateToolStripMenuItem.Click
        frmView_GRN_cnt.Close()
        frmView_GRN_cnt.Show()
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        txtUserName.Text = ""
        txtPassword.Text = ""
        If _LogStaus = False Then
            OPRUser.Visible = True
            txtUserName.Focus()

            Exit Sub
        Else
            If Search_Supplier() = True Then
            Else
                MsgBox("Please select the Supplier Name", MsgBoxStyle.Information, "Informaation ........")
                cboLocation.ToggleDropdown()
                Exit Sub
            End If

            If Search_Location() = True Then
            Else
                MsgBox("Please select the Location", MsgBoxStyle.Information, "Informaation ........")
                cboTo.ToggleDropdown()
                Exit Sub
            End If

            If Trim(txtCom_Invoice.Text) <> "" Then
            Else
                MsgBox("Please enter the Company Invoice No", MsgBoxStyle.Information, "Information .....")
                txtCom_Invoice.Focus()
                Exit Sub
            End If

            If txtPO.Text <> "" Then
            Else
                txtPO.Text = "-"
            End If
            If UltraGrid1.Rows.Count > 0 Then
            Else
                MsgBox("Please enter the transaction Details", MsgBoxStyle.Information, "Information ......")
                cboCode.ToggleDropdown()
            End If

            If txtDis_Rate.Text <> "" Then
            Else
                txtDis_Rate.Text = "0"
            End If

            If IsNumeric(txtDis_Rate.Text) Then
            Else
                MsgBox("Please enter the correct Discount Rate", MsgBoxStyle.Information, "Information .......")
                txtDis_Rate.Focus()
            End If
            If Trim(txtRemark.Text) <> "" Then
            Else
                txtRemark.Text = "-"
            End If

            If txtVAT.Text <> "" Then
            Else
                txtVAT.Text = "0"
            End If
            If txtNBT.Text <> "" Then
            Else
                txtNBT.Text = "0"
            End If
        End If

        Call Edit_Data()
    End Sub

    Private Sub txtUserName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.KeyUp
        If e.KeyCode = 13 Then
            txtPassword.Focus()
        End If
    End Sub

    Private Sub txtPassword_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyUp
        If e.KeyCode = 13 Then
            OK.Focus()
        End If
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim A As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet

        Dim MB51 As DataSet
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime

        Try
            If _LogStaus = False Then
                OPRUser.Visible = True
                txtUserName.Focus()
                Exit Sub
            End If

            _GetDate = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text)

            _Get_Time = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text) & " " & Hour(Now) & ":" & Minute(Now)

            A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information .........")
            If A = vbYes Then
                nvcFieldList1 = "select * from T03Supplier_Tr_Account where T03Ref_No='" & txtEntry.Text & "' and T03Status='PAID'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    MsgBox("Can't delete this Dispatch Note.", MsgBoxStyle.Information, "Information .......")
                    connection.Close()
                    Exit Sub
                End If
                '===================================================================================================
                nvcFieldList1 = "UPDATE T01GRN_Header SET T01Status='CLOSE' WHERE T01Ref_No='" & Trim(txtEntry.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T02GRN_Flutter SET T02Status='CLOSE' WHERE T02Ref_No='" & Trim(txtEntry.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T03Supplier_Tr_Account SET T03Status='CLOSE' WHERE T03Ref_No='" & Trim(txtEntry.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S01Stock_Balance SET S01Status='CLOSE' WHERE S01Ref_No='" & Trim(txtEntry.Text) & "' AND S01Tr_Type='GRN'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                              " values('GRN','DELETE', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Record deleted successfully", MsgBoxStyle.Information, "Information ........")
                transaction.Commit()
               
            End If
            connection.ClearAllPools()
            connection.Close()
            Call Clear_text()
            Call Load_EntryNo()
            _LogStaus = False
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If

        End Try
    End Sub

    Function Edit_Data()
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        SqlClient.SqlConnection.ClearAllPools()
        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim MB51 As DataSet
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim i As Integer
        Dim A As String
        Dim B As New ReportDocument
        Try
            ' Call Load_EntryNo()
            _GetDate = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text)

            _Get_Time = Month(txtDate.Text) & "/" & Microsoft.VisualBasic.Day(txtDate.Text) & "/" & Year(txtDate.Text) & " " & Hour(Now) & ":" & Minute(Now)

         
            nvcFieldList1 = "select* from T01GRN_Header  where T01Ref_No='" & Trim(txtEntry.Text) & "'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "UPDATE T01GRN_Header SET T01Date='" & _GetDate & "',T01Sup_Code='" & _Supplier & "',T01Net_Amount='" & CDbl(txtNett.Text) & "',T01Dis_Rate='" & txtDis_Rate.Text & "',T01VAT='" & txtVAT.Text & "',T01NBT='" & txtNBT.Text & "',T01Remark='" & Trim(txtRemark.Text) & "',T01Com_Invoice='" & Trim(txtCom_Invoice.Text) & "',T01PO='" & txtPO.Text & "' WHERE T01Ref_No='" & txtEntry.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T03Supplier_Tr_Account SET T03Date='" & _GetDate & "',T03Sup_Code='" & _Supplier & "',T03Credit='" & CDbl(txtGross.Text) & "' WHERE T03Ref_No='" & txtEntry.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                'SAVE DATA ON TMPTRANSACTION_LOG
                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                  " values('GRN','EDIT', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEntry.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "DELETE FROM T02GRN_Flutter WHERE T02Ref_No='" & txtEntry.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S01Stock_Balance SET S01Status='CANCEL' WHERE S01Ref_No='" & txtEntry.Text & "' AND S01Tr_Type='GRN'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    'SAVE S01STOCK_BALANCE
                    nvcFieldList1 = "Insert Into S01Stock_Balance(S01Item_Code,S01Ref_No,S01Date,S01Time,S01Tr_Type,S01Qty,S01Status)" & _
                                                             " values('" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "' ,'" & Trim(txtEntry.Text) & "', '" & _GetDate & "','" & _Get_Time & "','GRN','" & CDbl(UltraGrid1.Rows(i).Cells(5).Text) & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    '========================================================================
                    'SAVE T02GRN_FLUTTER
                    nvcFieldList1 = "Insert Into T02GRN_Flutter(T02Ref_No,T02Item_Ref,T02Part_No,T02Cost,T02Retail,T02Qty,T02Status)" & _
                                                           " values('" & Trim(txtEntry.Text) & "' ,'" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(1).Text) & "', '" & Trim(UltraGrid1.Rows(i).Cells(3).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(4).Text) & "','" & CDbl(UltraGrid1.Rows(i).Cells(5).Text) & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    '========================================================================
                    'RACK NO
                    nvcFieldList1 = "select * from M12Store_Location where M12Item_Code='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "'"
                    MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(MB51) Then
                        nvcFieldList1 = "update M12Store_Location set M12Rack='" & Trim(UltraGrid1.Rows(i).Cells(6).Text) & "',M12Cell='" & Trim(UltraGrid1.Rows(i).Cells(7).Text) & "' where M12Item_Code='" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "' "
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        nvcFieldList1 = "Insert Into M12Store_Location(M12Item_Code,M12Rack,M12Cell)" & _
                                                                          " values('" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(6).Text) & "', '" & Trim(UltraGrid1.Rows(i).Cells(7).Text) & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    i = i + 1
                Next
                MsgBox("Records change successfully", MsgBoxStyle.Information, "Information ........")
                transaction.Commit()

                Call Clear_text()
                _LogStaus = False
            Else

          
              
              
            End If
            connection.ClearAllPools()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim SQL As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        Dim M01 As DataSet

        Try
            con = DBEngin.GetConnection()
            SQL = "SELECT * FROM X02User_Loging WHERE (X02User_Name ='" & txtUserName.Text & "')and X02Password='" & txtPassword.Text & "' and X02User_Level in ('ADMIN','MANEGER','ACCOUNT') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(M01) Then
                _LogStaus = True
                _UserLevel = Trim(M01.Tables(0).Rows(0)("X02User_Level"))
                '_AthzUser = Trim(txtUserName.Text)
                OPRUser.Visible = False
            Else
                MsgBox("User name and pasword combination not found", "Information ......")
                txtUserName.Focus()
                con.ClearAllPools()
                con.CLOSE()

                Exit Sub
            End If
            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If

        End Try

    End Sub

    Private Sub cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel.Click
        OPRUser.Visible = False
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim A As String
        Dim B As New ReportDocument
        Try
            A = ConfigurationManager.AppSettings("ReportPath") + "\GRNDispatch.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "sainfinity")
            'B.SetParameterValue("To", _To)
            'B.SetParameterValue("From", _From)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{T01GRN_Header.T01Ref_No}='" & Trim(txtEntry.Text) & "' "
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.CLOSE()
            End If

        End Try
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub ItemSearchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemSearchToolStripMenuItem.Click
        strWindowName = Me.Name
        frmView_Items_cnt.Close()
        frmView_Items_cnt.Show()
    End Sub

    
    Private Sub cboCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.TextChanged
        If UltraGrid2.Visible = True Then
            Call Load_Grid_ROW()
        Else
            UltraGrid2.Visible = True
            Call Load_Grid_ROW()
        End If
    End Sub

    Private Sub UltraGrid2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid2.DoubleClickRow
        Dim _Row As Integer

        _Row = UltraGrid2.ActiveRow.Index
        _Itemcode = Trim(UltraGrid2.Rows(_Row).Cells(0).Text)
        Call Serch_Item()
        UltraGrid2.Visible = False
        txtRate.Focus()
    End Sub

    Private Sub UltraGrid2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid2.KeyUp
        Dim _Row As Integer
        If e.KeyCode = 13 Then
            _Row = UltraGrid2.ActiveRow.Index
            _Itemcode = Trim(UltraGrid2.Rows(_Row).Cells(0).Text)
            Call Serch_Item()
            UltraGrid2.Visible = False
            txtRate.Focus()
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Clear_text()
    End Sub
End Class