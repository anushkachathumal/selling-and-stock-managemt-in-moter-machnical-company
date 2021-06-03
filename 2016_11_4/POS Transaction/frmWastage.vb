﻿Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmWastage
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Dim _Loccode As String
    Dim _FromLocCode As String
    Dim _EntryNo As Integer
    Dim _RefNo As Integer
    Dim _LogStaus As Boolean
    Dim _AthzUser As String
    Dim _UserLevel As String

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

    Private Sub frmWastage_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCODE")
        txtDate.Text = Today
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.ReadOnly = True
        txtRate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.ReadOnly = True

        '  Call Load_Combo()
        Call Load_EntryNo()
        Call Load_Item_Name()
        Call Load_Item_Code()
        Call Load_Gride2()
        txtEntry.ReadOnly = True
        Call Load_Gride_Item()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_PO
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 230
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function



    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='WT'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    txtEntry.Text = "WT-00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    txtEntry.Text = "WT-0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtEntry.Text = "WT-" & M01.Tables(0).Rows(0)("P01LastNo")
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

    
    Function Load_Item_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Code as [Item Code] from M03Item_Master where M03Status='A' and M03Location='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCode
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 130
                '  .Rows.Band.Columns(1).Width = 160


            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Item_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Name as [Item Name] from M03Item_Master where M03Status='A' and M03Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItemName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 370
                '  .Rows.Band.Columns(1).Width = 160


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
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = M01.Tables(0).Rows(0)("M03Retail_Price")
                'txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))


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
    Private Sub txtRemark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyUp
        If e.KeyCode = Keys.Enter Then
            cboCode.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboCode.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewWastage.Show()
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

    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer
        _Rowindex = UltraGrid3.ActiveRow.Index
        cboCode.Text = UltraGrid3.Rows(_Rowindex).Cells(0).Text
        Search_ItemName()
        OPR5.Visible = False
        txtRate.Focus()
    End Sub

    
    Private Sub cboItemName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItemName.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemCode()
            txtRate.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            ' OPR4.Visible = True
        ElseIf e.KeyCode = Keys.F1 Then
            OPR5.Visible = True
            txtFind.Text = ""
            Call Load_Gride_Item()
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            ' OPR4.Visible = False
            OPR5.Visible = False
        ElseIf e.KeyCode = Keys.F2 Then
            frmViewWastage.Show()
        End If
    End Sub
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
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = M01.Tables(0).Rows(0)("M03Retail_Price")
                'txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

               
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
                txtRate.Focus()
            Else
                ' txtVAT.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_ItemName()
            txtRate.Focus()
       
        ElseIf e.KeyCode = Keys.F1 Then
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
        ElseIf e.KeyCode = Keys.F2 Then
            frmViewWastage.Show()
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
        ElseIf e.KeyCode = Keys.F2 Then
            frmViewWastage.Show()
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

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Dim result1 As String

        If e.KeyCode = Keys.Enter Then
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
            txtTotal.Focus()
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
            txtTotal.Focus()
        ElseIf e.KeyCode = Keys.F2 Then
            frmViewWastage.Show()
        End If
    End Sub



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

                'If txtRetail.Text <> "" Then
                'Else
                '    txtRetail.Text = "0"
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

           

                'If txtCost.Text <> "" Then
                'Else
                '    txtCost.Text = "0"
                'End If
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
                newRow("Rate") = txtRate.Text
                '  newRow("Retail Price") = txtSales.Text
                newRow("Qty") = txtQty.Text
                ' newRow("Rec.Qty") = txtRe_Qty.Text
                '   newRow("Free Issue") = txtFree.Text
                newRow("Total") = txtTotal.Text

                'SQL = "select * from M03Item_Master where M03Item_Code='" & cboCode.Text & "' and m03Status='A' and LEFT(M03ExPair,1)='Y'"
                'T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                'If isValidDataset(T01) Then
                '    newRow("Ex Date") = txtEx_Date.Text

                '    If CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) <> CDbl(txtRate.Text) Then
                '        result1 = MsgBox("Previous cost price is Rs." & CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) & ".do you want to change new one", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Information ........")
                '        If result1 = vbYes Then
                '            newRow("##") = True
                '            _CostStatus = True
                '        Else
                '            newRow("##") = False
                '        End If
                '    End If
                'End If
                '  newRow("##") = False
                SQL = "select * from M03Item_Master where M03Item_Code='" & cboCode.Text & "' and m03Status='A' and M03Com_Code='" & _Comcode & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                    'If CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) <> CDbl(txtRate.Text) Then
                    '    result1 = MsgBox("Previous cost price is Rs." & CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) & ".do you want to change new one", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Information ........")
                    '    If result1 = vbYes Then
                    '        newRow("##") = True
                    '    Else
                    '        newRow("##") = False
                    '    End If

                    'Else
                    '    newRow("##") = False
                    'End If
                Else
                    MsgBox("Please enter the correct item code", MsgBoxStyle.Information, "Information .......")
                    cboCode.Focus()
                    Exit Sub
                End If
                c_dataCustomer1.Rows.Add(newRow)

                'If _CostStatus = True Then
                '    Dim _lastRow As Integer

                '    _lastRow = UltraGrid1.Rows.Count - 1
                '    UltraGrid1.Rows(_lastRow).Cells(8).Value = True
                'End If
                'txtCount.Text = Val(txtCount.Text) + 1

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
                '  txtFree.Text = ""
                txtQty.Text = ""
                ' txtSales.Text = ""
                txtTotal.Text = ""
                'Me.txtRetail.Text = ""
                ' Me.txtEx_Date.Appearance.BackColor = Color.White
                'txtEx_Date.Text = ""
                cboCode.Focus()
            ElseIf e.KeyCode = Keys.F2 Then
                frmViewWastage.Show()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub

    Function Search_RecordsUsing_Entry()
        Dim result1 As String
        Dim Value As Double
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim _St As String
        Dim I As Integer
        Try
            SQL = "select * from View_Wastage_Header where T01Grn_No='" & txtEntry.Text & "' and T01Com_Code='" & _Comcode & "' and T01Status='A'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                txtRemark.Text = T01.Tables(0).Rows(0)("T01Remark")
                _RefNo = T01.Tables(0).Rows(0)("T01Ref_no")
                Value = T01.Tables(0).Rows(0)("T01Net_amount")
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If

            SQL = "select * from View_T02Transaction where T02Ref_No=" & _RefNo & ""
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            I = 0
            For Each DTRow2 As DataRow In T01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = T01.Tables(0).Rows(I)("T02Item_Code")
                newRow("Item Name") = T01.Tables(0).Rows(I)("M03Item_Name")
                Value = T01.Tables(0).Rows(I)("T02cost")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Rate") = _St
                '  newRow("Retail Price") = txtSales.Text
                newRow("Qty") = T01.Tables(0).Rows(I)("T02Qty")
                Value = T01.Tables(0).Rows(I)("T02Total")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St



                c_dataCustomer1.Rows.Add(newRow)
                I = I + 1
            Next

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Private Sub UltraGrid1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.AfterRowsDeleted
        Dim i As Integer
        Dim value As Double
        'txtCount.Text = UltraGrid1.Rows.Count
        Try
            i = 0
            txtNett.Text = "0.00"
            'txtGross.Text = "0.00"
            value = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Double.TryParse(txtNett.Text, value))
                value = value + CDbl((UltraGrid1.Rows(i).Cells(4).Value))
                i = i + 1
            Next

            txtNett.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))


            'value = txtNett.Text
            'value = value - (CDbl(txtVAT.Text) + CDbl(txtMarket.Text) + CDbl(txtDiscount.Text))
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
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument

        Try
          
            'If Trim(txtCom_Invoice.Text) <> "" Then
            'Else
            '    result1 = MessageBox.Show("Please enter the company invoice ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    If result1 = Windows.Forms.DialogResult.OK Then
            '        txtCom_Invoice.Focus()
            '        Exit Sub
            '    End If
            'End If

            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = " "
            End If

           

            '----------------------------------------------------------------------------------
            Call Load_EntryNo()
            If UltraGrid1.Rows.Count > 0 Then
            Else
                result1 = MessageBox.Show("Please enter the Items ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboCode.ToggleDropdown()
                    connection.Close()
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

            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='WT' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'UPDATE T01 TRANSACTION
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                Dim _Status As String
                Dim _Exdate As String
                Dim _retail As Double

                _Exdate = " "
                'If (UltraGrid1.Rows(i).Cells(3).Value) = (UltraGrid1.Rows(i).Cells(4).Value) Then
                '    _Status = "OK"
                'Else
                _Status = "A"
                'End If
                nvcFieldList1 = "SELECT * FROM M03Item_Master WHERE M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "' AND M03Com_Code='" & _Comcode & "' AND m03Status='A'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    _retail = M01.Tables(0).Rows(0)("M03Retail_Price")
                End If
               
                nvcFieldList1 = "Insert Into T02Transaction_Flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Commition,T02Com_Discount,T02Qty,T02Rec_Qty,T02Free_Issue,T02Status,T02Item_Received,T02Com_Code,T02Total,T02Ex_Date,T02Count)" & _
                                                              " values('" & _RefNo & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & _retail & "','0','0','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','0','0','A','" & _Status & "','" & _Comcode & "','" & CDbl(UltraGrid1.Rows(i).Cells(4).Value) & "','" & _Exdate & "','" & i + 1 & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Ref_No,S01Com_Code,S01STATUS)" & _
                                                              " values('" & _Comcode & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','WT','" & -(UltraGrid1.Rows(i).Cells(3).Value) & "','" & _RefNo & "','" & _Comcode & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S04Stock_Price(S04Location,S04Tr_Type,S04Date,S04Item_Code,S04Qty,S04Ref_No,S04Status,S04Rate,S04Com_Code)" & _
                                                               " values('" & _Comcode & "','WT','" & txtDate.Text & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & -(UltraGrid1.Rows(i).Cells(3).Text) & "','" & _RefNo & "','A','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

               
                'If (UltraGrid1.Rows(i).Cells(8).Value) = True Then
                '    nvcFieldList1 = "UPDATE M03Item_Master SET M03Cost_Price='" & (UltraGrid1.Rows(i).Cells(2).Value) & "' WHERE M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                'End If

                i = i + 1
            Next

            nvcFieldList1 = "Insert Into T01Transaction_Header(T01Trans_Type,T01Ref_No,T01Date,T01Invoice_No,T01FromLoc_Code,T01To_Loc_Code,T01Net_Amount,T01Com_Discount,T01DisRate,T01Vat,T01FreeIssue,T01Market_Return,T01Profit,T01Com_Code,T01Discount,T01User,T01Remark,T01PO_NO,T01Grn_No,T01Status)" & _
                                                              " values('WT', '" & _RefNo & "','" & txtDate.Text & "','-','" & _Comcode & "','-','" & CDbl(txtNett.Text) & "','0','0','0','0','0','0','" & _Comcode & "','0','" & strDisname & "','" & txtRemark.Text & "','-','" & txtEntry.Text & "','A')"
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
            'A = MsgBox("Are you sure you want to print Wastage Note", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Dispatch Note .....")
            'If A = vbYes Then
            '    A1 = ConfigurationManager.AppSettings("ReportPath") + "\Wastage_Note.rpt"
            '    B.Load(A1.ToString)
            '    B.SetDatabaseLogon("sa", "tommya")
            '    'B.SetParameterValue("To", _To)
            '    'B.SetParameterValue("From", _From)
            '    '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            '    frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            '    frmReport.CrystalReportViewer1.DisplayToolbar = True
            '    frmReport.CrystalReportViewer1.SelectionFormula = "{View_Wastage_Header.T01Ref_No}  =" & _RefNo & " and {View_Wastage_Header.T01FromLoc_Code} ='" & _Comcode & "'"
            '    frmReport.Refresh()
            '    ' frmReport.CrystalReportViewer1.PrintReport()
            '    ' B.PrintToPrinter(1, True, 0, 0)
            '    frmReport.MdiParent = MDIMain
            '    frmReport.Show()
            'End If

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            Call Load_EntryNo()
            Call Load_Gride2()
            txtNett.Text = "00.00"
            cboCode.Text = ""
            cboItemName.Text = ""
            txtTotal.Text = ""
            txtQty.Text = ""
            txtRemark.Text = ""
            cboCode.Focus()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim A1 As String
        Dim B As New ReportDocument

        Try

            A1 = ConfigurationManager.AppSettings("ReportPath") + "\Wastage_Note.rpt"
            B.Load(A1.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            'B.SetParameterValue("To", _To)
            'B.SetParameterValue("From", _From)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{View_Wastage_Header.T01Ref_No}  =" & _RefNo & " and {View_Wastage_Header.T01FromLoc_Code} ='" & _Comcode & "'"
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
                nvcFieldList1 = "update T01Transaction_Header set T01status='I' where T01Ref_No='" & _RefNo & "'and T01Trans_Type='WT' and T01status='A'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update T02Transaction_Flutter set T02status='I' where T02Ref_No='" & _RefNo & "' and T02status='A' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "Update T05Acc_Trans where T05Ref_No='" & txtEntry.Text & "'  and T05Acc_Type='SP'"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                ' Call Search_Location()
                nvcFieldList1 = "UPDATE S01Stock_Balance SET S01STATUS='CLOSE' where S01Loc_Code='" & _Comcode & "'  and S01Ref_No='" & _RefNo & "' and S01Trans_Type='WT' and S01STATUS='A'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S03Ex_Stock SET S03STATUS='CLOSE' where S03Loc_Code='" & _Comcode & "'  and S03Ref_No='" & _RefNo & "' and S03Tr_Type='WT' and S03STATUS='A'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S04Stock_Price SET S04STATUS='CLOSE' where  S04Ref_No='" & _RefNo & "' and S04Tr_Type='MR' AND S04Location='" & _Comcode & "' and S04STATUS='A'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpAthz,tmpUser,tmpLog)" & _
                                                          " values('WT', 'DELETE','" & txtEntry.Text & "','" & Now & "','" & _AthzUser & "','" & strDisname & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records delete successfully", MsgBoxStyle.Information, "Information ...........")
                transaction.Commit()
            End If

            connection.Close()
          
            cboCode.Text = ""
            cboItemName.Text = ""
            txtRate.Text = ""
            txtRemark.Text = ""
            txtQty.Text = ""
            txtTotal.Text = ""
            txtNett.Text = ""
            Call Load_EntryNo()
            cmdDelete.Enabled = False
            cmdAdd.Enabled = True
            ' cboLocation.ToggleDropdown()
            ' cmdSave.Enabled = False

            '  Call Load_Gride()
            Call Load_Gride2()
            Call Load_EntryNo()
            ' Call Load_Combo()
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
End Class