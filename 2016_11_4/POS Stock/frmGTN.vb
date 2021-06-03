Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmGTN
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Dim _Loccode As String
    Dim _FromLocCode As String

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [To Location] from M04Location  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 270
                '  .Rows.Band.Columns(1).Width = 160


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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet

        Dim i As Integer
        Try
            Sql = "select * from M04Location where M04Loc_Name='" & Trim(cboTo.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Loccode = Trim(M01.Tables(0).Rows(0)("M04Loc_Code"))
                Search_Location = True
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

    Function Search_LocationFrom() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet

        Dim i As Integer
        Try
            Sql = "select * from M04Location where  M04Loc_Name='" & Trim(cboLocation.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _FromLocCode = Trim(M01.Tables(0).Rows(0)("M04Loc_Code"))
                Search_LocationFrom = True
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

    Private Sub frmGTN_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("ComCode")

        'txtVAT.ReadOnly = True
        'txtMarket.ReadOnly = True
        'txtDis_Rate.ReadOnly = True
        'txtDiscount.ReadOnly = True
        'txtFree_Amount.ReadOnly = True

        'txtVAT.Text = "0.00"
        'txtMarket.Text = "0.00"
        'txtFree_Amount.Text = "0.00"
        'txtDis_Rate.Text = "0"
        'txtDiscount.Text = "0.00"

        txtDate.Text = Today
        txtRate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtDis_Rate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right

        txtSearch.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        '   txtFree.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtFree_Amount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        'txtGross.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        'txtMarket.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtRe_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtVAT.Appearance.TextHAlign = Infragistics.Win.HAlign.Right

        txtEntry.ReadOnly = True
        txtCount.ReadOnly = True
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.ReadOnly = True

        Call Load_Combo()
        Call Load_EntryNo()
        Call Load_ToLocation()
        Call Load_Item_Code()
        Call Load_Item_Name()

        Call Load_Gride2()
    End Sub


    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='IN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtEntry.Text = M01.Tables(0).Rows(0)("P01LastNo")
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

    Function Load_ToLocation()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [To Location] from M04Location  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboTo
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 270
                '  .Rows.Band.Columns(1).Width = 160


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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function


    Function Load_Item_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Code as [Item Code] from M03Item_Master where M03Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCode
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 130
                '  .Rows.Band.Columns(1).Width = 160


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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Load_Item_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Name as [Item Name] from M03Item_Master where M03Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItemName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 370
                '  .Rows.Band.Columns(1).Width = 160


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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_ItemCode()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select * from M03Item_Master where  M03Item_Name='" & Trim(cboItemName.Text) & "' and M03Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboCode.Text = M01.Tables(0).Rows(0)("M03Item_Code")
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

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

    Function Search_ItemName() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Try
            Sql = "select * from M03Item_Master where  M03Item_Code='" & Trim(cboCode.Text) & "' and M03Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_ItemName = True
                cboItemName.Text = M01.Tables(0).Rows(0)("M03Item_Name")
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

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

    Private Sub cboItemName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItemName.AfterCloseUp
        Call Search_ItemCode()
    End Sub

    Private Sub txtTotal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotal.KeyUp
        Dim result1 As String
        Dim Value As Double

        Try
            If e.KeyCode = 13 Then
                Call Calculation()
              


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



                If Trim(txtRate.Text) <> "" Then

                    If IsNumeric(txtRate.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the Correct Rate", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtRate.Focus()
                            Exit Sub
                        End If
                    End If
                Else
                    result1 = MessageBox.Show("Please enter the Correct Rate", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtRate.Focus()
                        Exit Sub
                    End If
                End If

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = Trim(cboCode.Text)
                newRow("Item Name") = cboItemName.Text
                newRow("Rate") = txtRate.Text
                newRow("Qty") = txtQty.Text
                newRow("Rec.Qty") = "0"
                newRow("Free Issue") = "0"
                newRow("Total") = txtTotal.Text

                c_dataCustomer1.Rows.Add(newRow)

                txtCount.Text = Val(txtCount.Text) + 1

                If txtNett.Text <> "" Then
                Else
                    txtNett.Text = "0"
                End If
                'Value = Double.TryParse(txtNett.Text, Value) + Double.TryParse(txtTotal.Text, Value)
                Value = CDbl(txtNett.Text) + CDbl(txtTotal.Text)

                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                'If txtFree_Amount.Text <> "" Then
                'Else
                '    txtFree_Amount.Text = "0"
                'End If

                'If txtMarket.Text <> "" Then
                'Else
                '    txtMarket.Text = "0"
                'End If

                'If txtVAT.Text <> "" Then
                'Else
                '    txtVAT.Text = "0"
                'End If

                ''If txtDiscount.Text <> "" Then
                ''Else
                ''    txtDiscount.Text = "0"
                ''End If
                ''Value = Value - Val(txtFree_Amount.Text) - Val(txtMarket.Text) - Val(txtVAT.Text)
                ''Value = Value - Val(txtDiscount.Text)

                ''txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                ''txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                cboItemName.Text = ""
                cboCode.Text = ""
                txtRate.Text = ""
                '  txtFree.Text = ""
                txtQty.Text = ""
                ' txtRe_Qty.Text = ""
                txtTotal.Text = ""
                cboCode.ToggleDropdown()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub



    Private Sub UltraGrid1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.AfterRowsDeleted
        Dim i As Integer
        Dim value As Double
        txtCount.Text = UltraGrid1.Rows.Count
        Try
            i = 0
            txtNett.Text = "0.00"
            'txtGross.Text = "0.00"
            value = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Double.TryParse(txtNett.Text, value))
                value = value + CDbl((UltraGrid1.Rows(i).Cells(6).Value))
                i = i + 1
            Next

            txtNett.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))


            value = txtNett.Text
            'value = value - (CDbl(txtFree_Amount.Text) + CDbl(txtVAT.Text) + CDbl(txtMarket.Text) + CDbl(txtDiscount.Text))
            'txtGross.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            'txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub
    Function Calculate_Gross()
        Dim Value As Double
        Try
            'Value = txtNett.Text
            'Value = Value - (CDbl(txtFree_Amount.Text) + CDbl(txtVAT.Text) + CDbl(txtMarket.Text) + CDbl(txtDiscount.Text))
            'txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            'txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Function


    Private Sub cboCode_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.AfterCloseUp
        Call Search_ItemName()
    End Sub

    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCode.KeyUp
        If e.KeyCode = Keys.Enter Then
            If cboCode.Text <> "" Then
                Call Search_ItemName()
                txtRate.Focus()
            Else
                cmdAdd.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_ItemName()
            txtRate.Focus()
        End If
    End Sub

 

    Private Sub cboTo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboTo.KeyUp
        If e.KeyCode = Keys.Enter Then
            cboCode.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboCode.ToggleDropdown()
        End If
    End Sub

    Private Sub txtRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtRate.Text) Then
                Value = txtRate.Text
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtQty.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If IsNumeric(txtRate.Text) Then
                Value = txtRate.Text
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtQty.Focus()
            End If
        End If
    End Sub

    Function Calculation()
        Dim Value As Double

        If IsNumeric(txtRate.Text) Then
            If IsNumeric(txtQty.Text) Then
                Value = Val(txtRate.Text) * Val(txtQty.Text)
                txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
        Else
            txtTotal.Text = ""
        End If
    End Function

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Dim result1 As String

        If e.KeyCode = Keys.Enter Then
            Call Calculation()
            If IsNumeric(txtQty.Text) Then
                ' txtRe_Qty.Text = txtQty.Text
            Else
                result1 = MessageBox.Show("Please enter the Correct Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtQty.Focus()
                    Exit Sub
                End If
            End If
            txtTotal.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Calculation()
            If IsNumeric(txtQty.Text) Then
                '   txtTotal.Text = txtQty.Text
            Else
                result1 = MessageBox.Show("Please enter the Correct Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtQty.Focus()
                    Exit Sub
                End If
            End If
            txtTotal.Focus()
        End If
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGRN
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function



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
        Dim M01 As DataSet
        Dim Value As Double

        Try

            'If Trim(txtDis_Rate.Text) <> "" Then
            '    If IsNumeric(txtDis_Rate.Text) Then
            '        Value = (CDbl(txtNett.Text) * CDbl(txtDis_Rate.Text)) / 100
            '        txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        Call Calculate_Gross()
            '    Else
            '        result1 = MessageBox.Show("Please enter the Discount Rate", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '        If result1 = Windows.Forms.DialogResult.OK Then
            '            txtDis_Rate.Focus()
            '            Exit Sub
            '        End If
            '    End If
            'End If

            'If Trim(txtCom_Invoice.Text) <> "" Then
            'Else
            '    result1 = MessageBox.Show("Please enter the company invoice ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    If result1 = Windows.Forms.DialogResult.OK Then
            '        txtCom_Invoice.Focus()
            '        Exit Sub
            '    End If
            'End If

            If Search_Location() = True Then
            Else
                result1 = MessageBox.Show("Please Select the To Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboTo.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Search_LocationFrom() = True Then
            Else
                result1 = MessageBox.Show("Please Select the From Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Trim(cboLocation.Text) = Trim(cboTo.Text) Then
                result1 = MessageBox.Show("Can't tranfer the same location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------------
            Call Load_EntryNo()
            If UltraGrid1.Rows.Count > 0 Then
            Else
                result1 = MessageBox.Show("Please enter the Items ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboCode.ToggleDropdown()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------------
            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='IN' and P01Com_Code='" & _Comcode & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'UPDATE T01 TRANSACTION
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                Dim _Status As String

                ' If (UltraGrid1.Rows(i).Cells(3).Value) = (UltraGrid1.Rows(i).Cells(4).Value) Then
                _Status = "OK"
                'Else
                '_Status = "BL"
                'End If
                nvcFieldList1 = "Insert Into T02Transaction_Flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Commition,T02Com_Discount,T02Qty,T02Rec_Qty,T02Free_Issue,T02Status,T02Item_Received,T02Com_Code,T02Total)" & _
                                                              " values('" & txtEntry.Text & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','0','0','0','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & (UltraGrid1.Rows(i).Cells(4).Value) & "','" & (UltraGrid1.Rows(i).Cells(5).Value) & "','A','" & _Status & "','" & _Comcode & "','" & (UltraGrid1.Rows(i).Cells(6).Value) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Free_Issue,S01Ref_No,S01Com_Code)" & _
                                                              " values('" & _Loccode & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','GT','" & Val(UltraGrid1.Rows(i).Cells(3).Value) & "','" & Val(UltraGrid1.Rows(i).Cells(4).Value) & "','" & txtEntry.Text & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Free_Issue,S01Ref_No,S01Com_Code)" & _
                                                              " values('" & _FromLocCode & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','GT','" & -Val(UltraGrid1.Rows(i).Cells(3).Value) & "','" & Val(UltraGrid1.Rows(i).Cells(4).Value) & "','" & txtEntry.Text & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = i + 1
            Next

            nvcFieldList1 = "Insert Into T01Transaction_Header(T01Trans_Type,T01Ref_No,T01Date,T01Invoice_No,T01FromLoc_Code,T01To_Loc_Code,T01Net_Amount,T01Com_Discount,T01DisRate,T01Vat,T01FreeIssue,T01Market_Return,T01Profit,T01Com_Code,T01Discount,T01User)" & _
                                                              " values('TR', '" & txtEntry.Text & "','" & txtDate.Text & "','" & txtCom_Invoice.Text & "','" & _FromLocCode & "','" & _Loccode & "','" & txtNett.Text & "','0','0','0','0','0','0','" & _Comcode & "','0','" & strDisname & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            '-------------------------------------------------------------------
        


            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0, OPR1, OPR2, OPR3)
            OPR2.Enabled = True
            OPR1.Enabled = True
            OPR0.Enabled = True
            OPR3.Enabled = True
            cmdAdd.Enabled = True
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cboLocation.ToggleDropdown()
            ' cmdSave.Enabled = False
            cmdDelete.Enabled = False
            '  Call Load_Gride()
            Call Load_Gride2()
            Call Load_EntryNo()
            Call Load_Combo()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0, OPR1, OPR2, OPR3)
        OPR2.Enabled = True
        OPR1.Enabled = True
        OPR0.Enabled = True
        OPR3.Enabled = True
        cmdAdd.Enabled = True
        cboLocation.ToggleDropdown()
        ' cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        '  Call Load_Gride()
        Call Load_Gride2()
        Call Load_EntryNo()
        Call Load_Combo()

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


        Try
            Sql = "select * from T01Transaction_Header inner join T02Transaction_Flutter  on T01Ref_No=T02Ref_No  inner join M03Item_Master on T02Item_Code=M03Item_Code inner join M04Location on M04Loc_Code=T01To_Loc_Code  where  T01Ref_No='" & Trim(txtSearch.Text) & "' and T01Trans_Type='TR'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboTo.Text = Trim(M01.Tables(0).Rows(0)("M04Loc_Name"))
                txtCom_Invoice.Text = Trim(M01.Tables(0).Rows(0)("T01Invoice_No"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("T01Date"))
                txtEntry.Text = Trim(M01.Tables(0).Rows(0)("T01Ref_No"))

                Sql = "select * from M04Location where M04Loc_Code='" & Trim(M01.Tables(0).Rows(0)("T01FromLoc_Code")) & "' "
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    cboLocation.Text = M02.Tables(0).Rows(0)("M04Loc_Name")
                End If

                Value = Trim(M01.Tables(0).Rows(0)("T01Net_Amount"))
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = Trim(M01.Tables(0).Rows(0)("T01Vat"))
                'txtVAT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtVAT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = Trim(M01.Tables(0).Rows(0)("T01Market_Return"))
                'txtMarket.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtMarket.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = Trim(M01.Tables(0).Rows(0)("T01Com_Discount"))
                'txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'txtDis_Rate.Text = Trim(M01.Tables(0).Rows(0)("T01DisRate"))

                'Value = Trim(M01.Tables(0).Rows(0)("T01FreeIssue"))
                'txtFree_Amount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtFree_Amount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = CDbl(txtNett.Text) - (CDbl(txtFree_Amount.Text) + CDbl(txtDiscount.Text) + CDbl(txtVAT.Text) + Val(txtMarket.Text))
                'txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtCount.Text = M01.Tables(0).Rows.Count

                Dim _St As String

                i = 0
                For Each DTRow2 As DataRow In M01.Tables(0).Rows

                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Item Code") = Trim(M01.Tables(0).Rows(i)("M03Item_Code"))
                    newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("M03Item_Name"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Rate") = _St
                    newRow("Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    newRow("Rec.Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    newRow("Free Issue") = Trim(M01.Tables(0).Rows(i)("T02Free_Issue"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Qty")) * Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Total") = _St

                    c_dataCustomer1.Rows.Add(newRow)


                    i = i + 1
                Next
                cmdAdd.Enabled = False
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True

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

    Private Sub txtSearch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyUp
        If e.KeyCode = 13 Then
            Call Search_RecordsUsing_Entry()
        End If
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
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
        Dim strDelete As String
        Try
            strDelete = MessageBox.Show("Are you sure you want to delete this records ", "Information ....", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If strDelete = Windows.Forms.DialogResult.Yes Then
                nvcFieldList1 = "delete from T01Transaction_Header where T01Ref_No='" & txtEntry.Text & "'  and T01Trans_Type='TR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "delete from T02Transaction_Flutter where T02Ref_No='" & txtEntry.Text & "'  "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "delete from T05Acc_Trans where T05Ref_No='" & txtEntry.Text & "' and T05Com_Code='" & _Comcode & "' and T05Acc_Type='SP'"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                Call Search_Location()
                Call Search_LocationFrom()

                nvcFieldList1 = "delete from S01Stock_Balance where S01Loc_Code='" & _Loccode & "'  and S01Ref_No='" & txtEntry.Text & "' and S01Trans_Type='TR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "delete from S01Stock_Balance where S01Loc_Code='" & _FromLocCode & "'  and S01Ref_No='" & txtEntry.Text & "' and S01Trans_Type='TR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records delete successfully", MsgBoxStyle.Information, "Information ...........")
                transaction.Commit()
            End If


            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0, OPR1, OPR2, OPR3)
            OPR2.Enabled = True
            OPR1.Enabled = True
            OPR0.Enabled = True
            OPR3.Enabled = True

            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Enabled = True
            cboLocation.ToggleDropdown()
            ' cmdSave.Enabled = False

            '  Call Load_Gride()
            Call Load_Gride2()
            Call Load_EntryNo()
            Call Load_Combo()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub


    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
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
        Dim M01 As DataSet
        Dim Value As Double

        Try

            'If Trim(txtDis_Rate.Text) <> "" Then
            '    If IsNumeric(txtDis_Rate.Text) Then
            '        Value = (CDbl(txtNett.Text) * CDbl(txtDis_Rate.Text)) / 100
            '        txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            '        txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            '        Call Calculate_Gross()
            '    Else
            '        result1 = MessageBox.Show("Please enter the Discount Rate", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '        If result1 = Windows.Forms.DialogResult.OK Then
            '            txtDis_Rate.Focus()
            '            Exit Sub
            '        End If
            '    End If
            'End If

            'If Trim(txtCom_Invoice.Text) <> "" Then
            'Else
            '    result1 = MessageBox.Show("Please enter the company invoice ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    If result1 = Windows.Forms.DialogResult.OK Then
            '        txtCom_Invoice.Focus()
            '        Exit Sub
            '    End If
            'End If

            If Search_Location() = True Then
            Else
                result1 = MessageBox.Show("Please Select To Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboTo.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Search_LocationFrom() = True Then
            Else
                result1 = MessageBox.Show("Please Select From Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Trim(cboLocation.Text) = Trim(cboTo.Text) Then
                result1 = MessageBox.Show("Can't tranfer the same location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------------
            '  Call Load_EntryNo()
            If UltraGrid1.Rows.Count > 0 Then
            Else
                result1 = MessageBox.Show("Please enter the Items ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboCode.ToggleDropdown()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------------


            'UPDATE T01 TRANSACTION
            i = 0

            nvcFieldList1 = "delete from T02Transaction_Flutter where T02Ref_No='" & txtEntry.Text & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from S01Stock_Balance where S01Ref_No='" & txtEntry.Text & "'  and S01Trans_Type='TR'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

          
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                Dim _Status As String

                ' If (UltraGrid1.Rows(i).Cells(3).Value) = (UltraGrid1.Rows(i).Cells(4).Value) Then
                _Status = "OK"
                'Else
                '_Status = "BL"
                'End If
                nvcFieldList1 = "Insert Into T02Transaction_Flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Commition,T02Com_Discount,T02Qty,T02Rec_Qty,T02Free_Issue,T02Status,T02Item_Received,T02Com_Code)" & _
                                                              " values('" & txtEntry.Text & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','0','0','0','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & (UltraGrid1.Rows(i).Cells(4).Value) & "','" & (UltraGrid1.Rows(i).Cells(5).Value) & "','A','" & _Status & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Free_Issue,S01Ref_No,S01Com_Code)" & _
                                                              " values('" & _Loccode & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','TR','" & Val(UltraGrid1.Rows(i).Cells(3).Value) & "','" & Val(UltraGrid1.Rows(i).Cells(4).Value) & "','" & txtEntry.Text & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Free_Issue,S01Ref_No,S01Com_Code)" & _
                                                             " values('" & _FromLocCode & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','TR','" & -Val(UltraGrid1.Rows(i).Cells(3).Value) & "','" & Val(UltraGrid1.Rows(i).Cells(4).Value) & "','" & txtEntry.Text & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = i + 1
            Next

            nvcFieldList1 = "update T01Transaction_Header set T01To_Loc_Code='" & _Loccode & "',T01Net_Amount='" & txtNett.Text & "',T01Com_Discount='0',T01DisRate='0',T01Vat='0',T01FreeIssue='0',T01Market_Return='0',T01FromLoc_Code='" & _FromLocCode & "' where T01Ref_No='" & txtEntry.Text & "' and T01Trans_Type='TR' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '-------------------------------------------------------------------
      
            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0, OPR1, OPR2, OPR3)
            OPR2.Enabled = True
            OPR1.Enabled = True
            OPR0.Enabled = True
            OPR3.Enabled = True

            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Enabled = True
            cboLocation.ToggleDropdown()
            ' cmdSave.Enabled = False

            '  Call Load_Gride()
            Call Load_Gride2()
            Call Load_EntryNo()
            Call Load_Combo()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub
End Class