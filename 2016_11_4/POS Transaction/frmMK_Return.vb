Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmMK_Return
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Dim _Loccode As String
    Dim _SupCode As String
    Dim _SupCode1 As String
    Dim _ItemLoc As String
    Dim _RefNo As Integer
    Dim _LogStaus As Boolean
    Dim _UserLevel As String
    Dim _AthzUser As String
    Dim _FromLoc As String


    Function Search_From() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M04Location where M04Loc_Name='" & Trim(cboLocation.Text) & "' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                _FromLoc = dsUser.Tables(0).Rows(0)("M04Loc_Code")
                Search_From = True
            Else
                Search_From = False
            End If
            '  Call Load_Location()

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function
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


    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim SQL As String
        Dim con = New SqlConnection()
        Dim M01 As DataSet

        Try
            con = DBEngin.GetConnection()
            SQL = "SELECT * FROM users WHERE (NAME ='" & txtUserName.Text & "')and Password='" & txtPassword.Text & "' and UType in ('ADMIN','SH/MANEGER','ACCOUNT') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(M01) Then
                _LogStaus = True
                _UserLevel = Trim(M01.Tables(0).Rows(0)("UType"))
                _AthzUser = Trim(txtUserName.Text)
                OPRUser.Visible = False
            Else
                MsgBox("User name and pasword combination not found", "Information ......")
                txtUserName.Focus()
                con.close()
                Exit Sub
            End If
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.CLOSE()
            End If

        End Try

    End Sub
    Function Search_Supplier() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M01Account_Master where M01Acc_Name='" & Trim(cboSupplier.Text) & "'  and M01Acc_Type='SP' and M01Com_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                _Supcode = dsUser.Tables(0).Rows(0)("M01Acc_Code")
                Search_Supplier = True
            Else
                Search_Supplier = False
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub



    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Name as [From Location] from M01Account_Master where  M01Acc_Type='SP' and M01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupplier
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

    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='MK'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    txtEntry.Text = "MK-00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    txtEntry.Text = "MK-0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtEntry.Text = "MK-" & M01.Tables(0).Rows(0)("P01LastNo")
                End If
            End If

            'Sql = "select * from M04Location"
            'M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            'If isValidDataset(M01) Then
            '    cboLocation.Text = M01.Tables(0).Rows(0)("M04Loc_Name")
            'End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_ToLocation()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [To Location] from M04Location "
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

    Function Search_Retail_Price(ByVal strCode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Try
            Sql = "select * from View_StockBalance1 where S04Item_Code='" & strCode & "' and S04Cost='" & txtRate.Text & "' and S04Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("Rate")
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function
    Function Load_Cost_Price(ByVal strCode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select CAST(S04Cost AS DECIMAL(16,2)) as [##Cost],SUM(Qty) as Qty from View_StockBalance1 where S04Item_Code='" & strCode & "' and S04Com_Code='" & _Comcode & "' group by S04Cost"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtRate
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
                .Rows.Band.Columns(1).Width = 90


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
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
            Sql = "select * from M03Item_Master where M03Item_Name='" & Trim(cboItemName.Text) & "' and M03status='A' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboCode.Text = M01.Tables(0).Rows(0)("M03Item_Code")
                Call Load_Cost_Price(cboCode.Text)
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M01.Tables(0).Rows(0)("M03Retail_Price")
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Call Search_Retail_Price(cboCode.Text)
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
            Sql = "select * from M03Item_Master where  M03Item_Code='" & Trim(cboCode.Text) & "' and M03status='A' and M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_ItemName = True
                cboItemName.Text = M01.Tables(0).Rows(0)("M03Item_Name")
                Call Load_Cost_Price(cboCode.Text)
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M01.Tables(0).Rows(0)("M03Retail_Price")
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Call Search_Retail_Price(cboCode.Text)
                If Microsoft.VisualBasic.Left(M01.Tables(0).Rows(0)("M03ExPair"), 1) = "Y" Then
                    txtEx_Date.Appearance.BackColor = Color.Gold
                End If
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

    Private Sub cboCode_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.AfterCloseUp
        Call Search_ItemName()
    End Sub

    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCode.KeyUp
        If e.KeyCode = Keys.Enter Then
            If cboCode.Text <> "" Then
                Call Search_ItemName()
                txtRate.ToggleDropdown()
            Else
                'txtVAT.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_ItemName()
            txtRate.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F1 Then
            frmViewMarkert.Show()
        ElseIf e.KeyCode = Keys.F2 Then
            OPR5.Visible = True
            txtFind.Text = ""
            Call Load_Gride_Item()
            txtFind.Focus()

            'Panel1.Visible = False
            'GRP1.Visible = False
        ElseIf e.KeyCode = Keys.Escape Then
            'OPR4.Visible = False
            OPR5.Visible = False
            'Panel1.Visible = False
            'GRP1.Visible = False
        ElseIf e.KeyCode = Keys.F5 Then
            'Call Clear_Item()
            'Panel1.Visible = True
            'GRP1.Visible = True
            'cboMain.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F3 Then

        End If
    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Dim result1 As String

        If e.KeyCode = Keys.Enter Then
            Call Calculation()
            If IsNumeric(txtQty.Text) Then
                txtEx_Date.Focus()
            Else
                'result1 = MessageBox.Show("Please enter the Correct Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'If result1 = Windows.Forms.DialogResult.OK Then
                '    txtQty.Focus()
                '    Exit Sub
                'End If
            End If
            'txtFree.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Calculation()
            If IsNumeric(txtQty.Text) Then
                'txtRe_Qty.Text = txtQty.Text
            Else
                'result1 = MessageBox.Show("Please enter the Correct Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'If result1 = Windows.Forms.DialogResult.OK Then
                '    txtQty.Focus()
                '    Exit Sub
                'End If
            End If
            'txtFree.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            frmViewMarkert.Show()
        End If
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableMK_Return
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 80
            .DisplayLayout.Bands(0).Columns(6).Width = 90
            '  .DisplayLayout.Bands(0).Columns(7).Width = 90
            ' .DisplayLayout.Bands(0).Columns(8).Width = 90

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            ' .DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub txtRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtRate.Text) Then
                Value = txtRate.Text
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtSales.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If IsNumeric(txtRate.Text) Then
                Value = txtRate.Text
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtSales.Focus()
            End If
        ElseIf e.KeyCode = Keys.F1 Then
            frmViewGRN.Show()
        End If
    End Sub

    Function Calculation()
        Dim Value As Double

        If IsNumeric(txtRate.Text) Then
            If IsNumeric(txtQty.Text) Then
                Value = CDbl(txtRate.Text) * CDbl(txtQty.Text)
                txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
        Else
            txtTotal.Text = ""
        End If
    End Function

    Private Sub txtTotal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotal.KeyUp
        Dim result1 As String
        Dim Value As Double
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim _CostStatus As Boolean
        Try
            If e.KeyCode = 13 Then
                If txtEx_Date.Appearance.BackColor = Color.Gold Then
                    If IsDate(txtEx_Date.Text) Then
                    Else
                        MsgBox("Please enter the Ex.Date", MsgBoxStyle.Information, "Information .........")
                        con.close()
                        txtEx_Date.Focus()
                        Exit Sub
                    End If
                End If
                Call Calculation()
                'If Trim(txtRe_Qty.Text) <> "" Then
                'Else
                '    If IsNumeric(txtRe_Qty.Text) Then

                '        result1 = MessageBox.Show("Please enter the Correct Rec.Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                '        If result1 = Windows.Forms.DialogResult.OK Then
                '            txtRe_Qty.Focus()
                '            Exit Sub
                '        End If
                '    End If
                'End If

                
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

                If txtSales.Text <> "" Then
                Else
                    txtSales.Text = "0"
                End If

                If Trim(txtQty.Text) <> "" Then

                    If IsNumeric(txtQty.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the Correct Retail Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtSales.Focus()
                            Exit Sub
                        End If
                    End If
                End If

                

                If txtRate.Text <> "" Then
                Else
                    txtRate.Text = "0"
                End If
                If Trim(txtRate.Text) <> "" Then

                    If IsNumeric(txtRate.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the Correct cost price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtRate.Focus()
                            Exit Sub
                        End If
                    End If
                Else
                    result1 = MessageBox.Show("Please enter the Correct cost price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtRate.Focus()

                        Exit Sub
                    End If
                End If

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = Trim(cboCode.Text)
                newRow("Item Name") = cboItemName.Text
                newRow("Cost Price") = txtRate.Text
                newRow("Retail Price") = txtSales.Text
                newRow("Qty") = txtQty.Text
                ' newRow("Rec.Qty") = txtRe_Qty.Text
                ' newRow("Free Issue") = txtFree.Text
                newRow("Total") = txtTotal.Text

                SQL = "select * from M03Item_Master where M03Item_Code='" & cboCode.Text & "' and m03Status='A' and LEFT(M03ExPair,1)='Y' and M03Com_Code='" & _Comcode & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then
                    newRow("Ex Date") = txtEx_Date.Text

                    'If CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) <> CDbl(txtRate.Text) Then
                    '    result1 = MsgBox("Previous cost price is Rs." & CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) & ".do you want to change new one", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Information ........")
                    '    If result1 = vbYes Then
                    '        newRow("##") = True
                    '        _CostStatus = True
                    '    Else
                    '        newRow("##") = False
                    '    End If
                    'End If
                End If

                SQL = "select * from M03Item_Master where M03Item_Code='" & cboCode.Text & "' and m03Status='A'  and M03Com_Code='" & _Comcode & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                Else
                    MsgBox("Please enter the correct item code", MsgBoxStyle.Information, "Information .......")
                    cboCode.Focus()
                    Exit Sub
                End If
                c_dataCustomer1.Rows.Add(newRow)

                If _CostStatus = True Then
                    Dim _lastRow As Integer

                    _lastRow = UltraGrid1.Rows.Count - 1
                    UltraGrid1.Rows(_lastRow).Cells(8).Value = True
                End If
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

                cboItemName.Text = ""
                cboCode.Text = ""
                txtRate.Text = ""
                ' txtFree.Text = ""
                txtQty.Text = ""
                txtSales.Text = ""
                txtTotal.Text = ""
                '  Me.txtRetail.Text = ""
                Me.txtEx_Date.Appearance.BackColor = Color.White
                txtEx_Date.Text = ""
                cboCode.Focus()
            ElseIf e.KeyCode = Keys.F1 Then
                frmViewGRN.Show()
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
            ' txtGross.Text = "0.00"
            value = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Double.TryParse(txtNett.Text, value))
                value = value + CDbl((UltraGrid1.Rows(i).Cells(6).Value))
                i = i + 1
            Next

            txtNett.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub

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
            Sql = "select * from M04Location where  M04Loc_Name='" & Trim(cboLocation.Text) & "' and M04Com_Code='" & _Comcode & "'"
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
        'OPRUser.Visible = False
        'Panel2.Visible = False
        _UserLevel = ""
        _LogStaus = False
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
            Sql = "select * from T01Transaction_Header inner join T02Transaction_Flutter  on T01Ref_No=T02Ref_No  inner join M03Item_Master on T02Item_Code=M03Item_Code  inner join M04Location on M04Loc_Code=T01FromLoc_Code  inner join M09Supplier on M09Code=T01To_Loc_Code where T01Grn_No='" & Trim(txtEntry.Text) & "' and T01Trans_Type='MR' and T01FromLoc_Code='" & _Comcode & "'"
            Sql = "select * from T01Transaction_Header inner join T02Transaction_Flutter  on T01Ref_No=T02Ref_No and T02Com_Code=T01Com_Code inner join M03Item_Master on T02Item_Code=M03Item_Code and M03Com_Code=T02Com_Code inner join M04Location on M04Loc_Code=T01FromLoc_Code inner join M09Supplier on M09Loc_Code=T01Com_Code where T01Com_Code='" & _Comcode & "' and T01Grn_No='" & Trim(txtEntry.Text) & "' and T01Trans_Type='MR'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboSupplier.Text = Trim(M01.Tables(0).Rows(0)("M09Name"))
                cboLocation.Text = Trim(M01.Tables(0).Rows(0)("M04Loc_Name"))
                txtCom_Invoice.Text = Trim(M01.Tables(0).Rows(0)("T01Invoice_No"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("T01Date"))
                txtEntry.Text = Trim(M01.Tables(0).Rows(0)("T01Grn_No"))
                txtRemark.Text = Trim(M01.Tables(0).Rows(0)("T01Remark"))
                _RefNo = Trim(M01.Tables(0).Rows(0)("T01Ref_No"))

                Value = Trim(M01.Tables(0).Rows(0)("T01Net_Amount"))
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = Trim(M01.Tables(0).Rows(0)("T01FreeIssue"))
                'txtFree_Amount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtFree_Amount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = (CDbl(txtNett.Text) + CDbl(txtVAT.Text)) - (CDbl(txtDiscount.Text) + Val(txtMarket.Text))
                'txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtCount.Text = M01.Tables(0).Rows.Count

                Dim _St As String
                Call Load_Gride2()

                i = 0
                For Each DTRow2 As DataRow In M01.Tables(0).Rows

                    Dim newRow As DataRow = c_dataCustomer1.NewRow
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
                    ' newRow("Free Issue") = Trim(M01.Tables(0).Rows(i)("T02Free_Issue"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Qty")) * Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Total") = _St
                    ' newRow("##") = False
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

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim B As New ReportDocument
        Dim A1 As String

        Try
            A1 = ConfigurationManager.AppSettings("ReportPath") + "\markert_Return.rpt"
            B.Load(A1.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            'B.SetParameterValue("To", _To)
            'B.SetParameterValue("From", _From)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{View_MkReturnnw.T01Ref_No} =" & _RefNo & " and {View_MkReturnnw.T01FromLoc_Code}='" & _Comcode & "' "
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(connection)
                'connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub txtRemark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyUp
        If e.KeyCode = Keys.Enter Then
            cboCode.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboCode.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F1 Then
            frmViewGRN.Show()
        End If
    End Sub

    Private Sub cboItemName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItemName.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemCode()
            txtRate.Focus()

        ElseIf e.KeyCode = Keys.F2 Then
            OPR5.Visible = True
            txtFind.Text = ""
            Call Load_Gride_Item()
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            ' OPR4.Visible = False
            OPR5.Visible = False
        ElseIf e.KeyCode = Keys.F1 Then
            frmViewGRN.Show()
        End If
    End Sub

    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            cboSupplier.Focus()
        ElseIf e.KeyCode = Keys.F1 Then

        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False

        End If
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

    Private Sub txtFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyUp
        If e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            cboCode.Focus()
        End If
    End Sub

    Private Sub txtFind_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFind.TextChanged
        Call Load_Gride_Item3()
    End Sub

    Private Sub txtSales_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSales.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtSales.Text) Then
                Value = txtSales.Text
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtQty.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If IsNumeric(txtSales.Text) Then
                Value = txtSales.Text
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtQty.Focus()
            End If
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Private Sub UltraGrid3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid3.KeyUp
        On Error Resume Next
        Dim _Rowindex As Integer
        If e.KeyCode = 13 Then

            _Rowindex = UltraGrid3.ActiveRow.Index


            cboCode.Text = UltraGrid3.Rows(_Rowindex).Cells(0).Text
            Search_ItemName()
            OPR5.Visible = False
            txtRate.Focus()
        End If
    End Sub

    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer
        _Rowindex = UltraGrid3.ActiveRow.Index
        cboCode.Text = UltraGrid3.Rows(_Rowindex).Cells(0).Text
        Search_ItemName()
        OPR5.Visible = False
        txtRate.Focus()
    End Sub

    Private Sub txtEx_Date_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEx_Date.KeyUp
        If e.KeyCode = 13 Then
            txtTotal.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtTotal.Focus()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewMarkert.Show()
        End If
    End Sub

    Private Sub frmMK_Return_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride_Item()
        ' txtEx_Date.Text = Today
        txtDate.Text = Today
        txtRate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtDis_Rate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        ' txtPO.ReadOnly = True
        'Call Load_Location()
        'txtReorder.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtCost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtRetail.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtSales.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        'txtFree.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtFree_Amount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        'txtGross.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        'txtMarket.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtRe_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        ' txtVAT.Appearance.TextHAlign = Infragistics.Win.HAlign.Right

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
            Sql = "select M03Item_Name as [Item Name] from M03Item_Master where M03status='A' "
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

    Private Sub cboSupplier_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSupplier.KeyUp
        If e.KeyCode = 13 Then
            txtRemark.Focus()
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
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
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument

        Try
            If txtPO.Text <> "" Then
            Else
                txtPO.Text = " "
            End If


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

            'If txtRemark.Text <> "" Then
            'Else
            '    txtRemark.Text = " "
            'End If

            'If Search_Location() = True Then
            'Else
            '    result1 = MessageBox.Show("Please Select the To Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    If result1 = Windows.Forms.DialogResult.OK Then
            '        cboTo.ToggleDropdown()
            '        Exit Sub
            '    End If
            'End If

            If Search_From() = True Then
            Else
                result1 = MessageBox.Show("Please Select the From Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Search_Supplier() = True Then
            Else
                result1 = MessageBox.Show("Please Select the From Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            nvcFieldList1 = "select * from P01Parameter where P01Code='IN'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = M01.Tables(0).Rows(0)("P01LastNo")
            End If
            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='IN' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='MK' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'UPDATE T01 TRANSACTION
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                Dim _Status As String
                Dim _Exdate As String

                _Exdate = " "
                'If (UltraGrid1.Rows(i).Cells(3).Value) = (UltraGrid1.Rows(i).Cells(4).Value) Then
                '    _Status = "OK"
                'Else
                _Status = "A"
                'End If

                If IsDate((UltraGrid1.Rows(i).Cells(6).Value)) Then
                    _Exdate = (UltraGrid1.Rows(i).Cells(6).Value)
                End If
                nvcFieldList1 = "Insert Into T02Transaction_Flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Commition,T02Com_Discount,T02Qty,T02Rec_Qty,T02Free_Issue,T02Status,T02Item_Received,T02Com_Code,T02Total,T02Ex_Date,T02Count)" & _
                                                              " values('" & _RefNo & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','0','0','" & (UltraGrid1.Rows(i).Cells(4).Value) & "','0','0','A','" & _Status & "','" & _Comcode & "','" & CDbl(UltraGrid1.Rows(i).Cells(6).Value) & "','" & _Exdate & "','" & i + 1 & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Ref_No,S01Com_Code,S01STATUS)" & _
                                                              " values('" & _FromLoc & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','MR','" & -(UltraGrid1.Rows(i).Cells(4).Value) & "','" & _RefNo & "','" & _Comcode & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S04Stock_Price(S04Location,S04Tr_Type,S04Date,S04Item_Code,S04Qty,S04Ref_No,S04Status,S04Rate,S04Cost,S04Com_Code)" & _
                                                               " values('" & _FromLoc & "','MR','" & txtDate.Text & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & -(UltraGrid1.Rows(i).Cells(4).Text) & "','" & _RefNo & "','A','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'EXPAIRE STOCK UPDATE
                nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "' and m03Status='A' and M03ExPair='YES'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "Insert Into S03Ex_Stock(S03Loc_Code,S03Tr_Type,S03Item_Code,S03Qty,S03Ex_Date,S03Status,S03Ref_No)" & _
                                                                " values('" & _FromLoc & "','MR', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & -(UltraGrid1.Rows(i).Cells(4).Text) & "','" & UltraGrid1.Rows(i).Cells(5).Text & "','A','" & _RefNo & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                End If

                'If (UltraGrid1.Rows(i).Cells(8).Value) = True Then
                '    nvcFieldList1 = "UPDATE M03Item_Master SET M03Cost_Price='" & (UltraGrid1.Rows(i).Cells(2).Value) & "' WHERE M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                'End If

                i = i + 1
            Next

            nvcFieldList1 = "Insert Into T01Transaction_Header(T01Trans_Type,T01Ref_No,T01Date,T01Invoice_No,T01FromLoc_Code,T01To_Loc_Code,T01Net_Amount,T01Com_Discount,T01DisRate,T01Vat,T01FreeIssue,T01Market_Return,T01Profit,T01Com_Code,T01Discount,T01User,T01Remark,T01PO_NO,T01Grn_No,T01Status)" & _
                                                              " values('MR', '" & _RefNo & "','" & txtDate.Text & "','" & txtCom_Invoice.Text & "','" & _FromLoc & "','" & _SupCode & "','" & CDbl(txtNett.Text) & "','0','0','0','0','0','0','" & _Comcode & "','0','" & strDisname & "','" & txtRemark.Text & "','" & txtPO.Text & "','" & txtEntry.Text & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            '-------------------------------------------------------------------
            ''Account
            'Dim _StrRemark As String

            '_StrRemark = "Good Received invo.No -" & txtCom_Invoice.Text
            'nvcFieldList1 = "Insert Into T05Acc_Trans(T05Acc_Type,T05Ref_No,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Com_Code,T05User,T05Status)" & _
            '                                                 " values('SP', '" & txtEntry.Text & "','" & txtDate.Text & "','" & _SupCode & "','" & _StrRemark & "','0','" & CDbl(txtGross.Text) & "','" & _Comcode & "','" & strDisname & "','GR')"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'If CDbl(txtGross.Text) > 0 Then
            '    nvcFieldList1 = "Insert Into T07Supplier_Payment(T07RefNo,T07Date,T07InvoiceAmount,T07Paid_Amount,T07Status,T07Com_Code,T07Paid_Voucher)" & _
            '                                                 " values('" & txtEntry.Text & "','" & txtDate.Text & "','" & CDbl(txtGross.Text) & "','0','N','" & _SupCode & "','0')"
            '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            'End If
            'MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            A = MsgBox("Are you sure you want to print Markert Return Note", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Markert Return Note .....")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\markert_Return.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_MkReturnnw.T01Ref_No} =" & _RefNo & " and {View_MkReturnnw.T01FromLoc_Code}='" & _Comcode & "' "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

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
            ' Call Load_Data()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

 
    Private Sub cmdEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
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

            txtUserName.Text = ""
            txtPassword.Text = ""
            If _LogStaus = False Then
                'Panel2.Visible = True
                OPRUser.Visible = True
                connection.Close()
                txtUserName.Focus()
                Exit Sub
            End If

            If txtPO.Text <> "" Then
            Else
                txtPO.Text = " "
            End If

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
                result1 = MessageBox.Show("Please Select the Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    Exit Sub
                End If
            End If
            If Search_Supplier() = True Then
            Else
                result1 = MessageBox.Show("Please Select the From Supplier ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboSupplier.ToggleDropdown()
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
            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = " "
            End If

            'UPDATE T01 TRANSACTION
            i = 0

            nvcFieldList1 = "delete from T02Transaction_Flutter where T02Ref_No='" & _RefNo & "' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from S01Stock_Balance where S01Ref_No='" & _RefNo & "'  and S01Trans_Type='MR'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from S03Ex_Stock where S03Ref_No='" & _RefNo & "'  and S03Tr_Type='MR'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from S04Stock_Price where S04Ref_No='" & _RefNo & "'  and S04Tr_Type='MR'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                Dim _Status As String
                Dim _Exdate As String

                _Exdate = " "
                'If (UltraGrid1.Rows(i).Cells(3).Value) = (UltraGrid1.Rows(i).Cells(4).Value) Then
                '    _Status = "OK"
                'Else
                _Status = "A"
                'End If

                If IsDate((UltraGrid1.Rows(i).Cells(5).Value)) Then
                    _Exdate = (UltraGrid1.Rows(i).Cells(5).Value)
                End If
                nvcFieldList1 = "Insert Into T02Transaction_Flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Commition,T02Com_Discount,T02Qty,T02Rec_Qty,T02Free_Issue,T02Status,T02Item_Received,T02Com_Code,T02Total,T02Ex_Date,T02Count)" & _
                                                              " values('" & _RefNo & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','0','0','" & (UltraGrid1.Rows(i).Cells(4).Value) & "','0','0','A','" & _Status & "','" & _Comcode & "','" & CDbl(UltraGrid1.Rows(i).Cells(6).Value) & "','" & _Exdate & "','" & i + 1 & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Ref_No,S01Com_Code,s01Status)" & _
                                                              " values('" & _Loccode & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','MR','" & -(UltraGrid1.Rows(i).Cells(4).Value) & "','" & _RefNo & "','" & _Comcode & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S04Stock_Price(S04Location,S04Tr_Type,S04Date,S04Item_Code,S04Qty,S04Ref_No,S04Status,S04Rate,S04Cost,S04Com_Code)" & _
                                                                     " values('" & _FromLoc & "','MR','" & txtDate.Text & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & -(UltraGrid1.Rows(i).Cells(4).Text) & "','" & _RefNo & "','A','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'EXPAIRE STOCK UPDATE
                nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "' and m03Status='A' and M03ExPair='YES'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "Insert Into S03Ex_Stock(S03Loc_Code,S03Tr_Type,S03Item_Code,S03Qty,S03Ex_Date,S03Status,S03Ref_No)" & _
                                                                " values('" & _Loccode & "','MR', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & -(UltraGrid1.Rows(i).Cells(4).Text) & "','" & UltraGrid1.Rows(i).Cells(5).Text & "','A','" & _RefNo & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                End If

                'If (UltraGrid1.Rows(i).Cells(8).Value) = True Then
                '    nvcFieldList1 = "UPDATE M03Item_Master SET M03Cost_Price='" & (UltraGrid1.Rows(i).Cells(2).Value) & "' WHERE M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "' and M03Location='" & _Comcode & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                'End If

                i = i + 1
            Next

            nvcFieldList1 = "update T01Transaction_Header set T01To_Loc_Code='" & _SupCode & "',T01Net_Amount='" & txtNett.Text & "',T01Com_Discount='0',T01DisRate='0',T01Vat='0',T01FreeIssue='0',T01Market_Return='0',T01Remark='" & txtRemark.Text & "',T01PO_NO='" & txtPO.Text & "',T01Invoice_No='" & txtCom_Invoice.Text & "' where T01Ref_No='" & _RefNo & "' and T01Trans_Type='MR' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '-------------------------------------------------------------------

            nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpAthz,tmpUser,tmpLog)" & _
                                                           " values('MR', 'EDIT','" & txtEntry.Text & "','" & Now & "','" & _AthzUser & "','" & strDisname & "','" & _Comcode & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'Account
            Dim _StrRemark As String

            '_StrRemark = "Good Received invo.No -" & txtCom_Invoice.Text

            'nvcFieldList1 = "update T05Acc_Trans set T05Credit='" & txtGross.Text & "' where T05Ref_No='" & txtEntry.Text & "' and T05Acc_Type='SP' and T05Acc_No='" & _Comcode & "'"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

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
            ' Call Load_Data()
            _LogStaus = False

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
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
            txtUserName.Text = ""
            txtPassword.Text = ""

            If _LogStaus = False Then
                'Panel2.Visible = True
                OPRUser.Visible = True
                connection.Close()
                txtUserName.Focus()
                Exit Sub
            End If
            strDelete = MessageBox.Show("Are you sure you want to delete this records ", "Information ....", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If strDelete = Windows.Forms.DialogResult.Yes Then
                nvcFieldList1 = "update T01Transaction_Header set T01status='I' where T01Ref_No='" & _RefNo & "'and T01Trans_Type='MR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update T02Transaction_Flutter set T02status='I' where T02Ref_No='" & _RefNo & "'  "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "Update T05Acc_Trans where T05Ref_No='" & txtEntry.Text & "'  and T05Acc_Type='SP'"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                Call Search_Location()
                nvcFieldList1 = "UPDATE S01Stock_Balance SET S01STATUS='CLOSE' where S01Loc_Code='" & _Loccode & "'  and S01Ref_No='" & _RefNo & "' and S01Trans_Type='MR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S03Ex_Stock SET S03STATUS='CLOSE' where S03Loc_Code='" & _Loccode & "'  and S03Ref_No='" & _RefNo & "' and S03Tr_Type='MR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S04Stock_Price SET S04STATUS='CLOSE' where  S04Ref_No='" & _RefNo & "' and S04Tr_Type='MR' AND S04Location='" & _Loccode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpAthz,tmpUser,tmpLog)" & _
                                                          " values('MR', 'DELETE','" & txtEntry.Text & "','" & Now & "','" & _AthzUser & "','" & strDisname & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records delete successfully", MsgBoxStyle.Information, "Information ...........")
                transaction.Commit()
            End If

            connection.Close()
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
            '  Call Load_Data()
            _LogStaus = False
            '  Call Load_Data()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    

    Private Sub txtRate_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRate.AfterCloseUp
        Call Search_Retail_Price(cboCode.Text)
    End Sub

    Private Sub txtRate_KeyUp1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Retail_Price(cboCode.Text)
            txtSales.Focus()
        End If
    End Sub
End Class