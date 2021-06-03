Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmPayMain_DS
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Ez_Uniq
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 70
            '.DisplayLayout.Bands(0).Columns(4).Width = 70
            '.DisplayLayout.Bands(0).Columns(5).Width = 70
            '.DisplayLayout.Bands(0).Columns(6).Width = 110


            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ''.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            ''.DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit


            '.DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride3()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_ChqPay_Uniq
        UltraGrid2.DataSource = c_dataCustomer2
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(3).Width = 110
            '.DisplayLayout.Bands(0).Columns(4).Width = 70
            '.DisplayLayout.Bands(0).Columns(5).Width = 70
            '.DisplayLayout.Bands(0).Columns(6).Width = 110


            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ''.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            '.DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            ''.DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit


            '.DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub frmPayMain_DS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCard_Amount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCash.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtPaid_Amount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtBalance.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtBill_Amount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtBalance.ReadOnly = True
        txtPaid_Amount.ReadOnly = True
        txtBill_Amount.ReadOnly = True

        Call Load_PAY_TEARMS()
        Call Load_PAY_CARDS()
        Call Load_BANK()
        Call Load_Gride2()
        Call Load_Gride3()
        txtChq_Total.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtChq_Total.ReadOnly = True
        txtTotal.ReadOnly = True
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
    End Sub

    Function Load_PAY_TEARMS()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M11Name as [##] from M11Common where M11Status='PY' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboTearms
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
                ' .Rows.Band.Columns(1).Width = 360

            End With


            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Load_PAY_CARDS()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M11Name as [##] from M11Common where M11Status='CD' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboEzy
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
                ' .Rows.Band.Columns(1).Width = 360

            End With


            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Function Load_BANK()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M11Name as [##] from M11Common where M11Status='BNK' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboBank
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
                ' .Rows.Band.Columns(1).Width = 360

            End With


            con.ClearAllPools()
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function
    Private Sub cboTearms_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboTearms.KeyUp
        If e.KeyCode = 13 Then
            If Trim(cboTearms.Text).Trim = "CASH" Then
                txtCash.Focus()
            ElseIf LTrim(cboTearms.Text).Trim = "CREDIT" Then
                cmdAdd.Focus()
                txtCash.ReadOnly = True
            ElseIf Trim(cboTearms.Text).Trim = "CREDIT CARD" Then
                cboEzy.ToggleDropdown()
                OPR2.Visible = False
                OPR1.Visible = True
                OPR1.Enabled = True
                cboEzy.Text = ""
                txtCard_no.Text = ""
                txtAmount.Text = ""
                txtCard_Amount.Text = ""
                txtTotal.Text = ""
                Call Load_Gride2()
            ElseIf Trim(cboTearms.Text).Trim = "CHQUE" Then
                cboBank.ToggleDropdown()
                cboBank.Text = ""
                Me.txtChqNo.Text = ""
                Me.txtDOR.Text = ""
                txtAmount.Text = ""
                txtTotal.Text = ""
                Call Load_Gride3()
                txtChq_Total.Text = ""
            ElseIf Trim(cboTearms.Text).Trim = "CASH/CREDIT" Then
                txtCash.Focus()
                Call Load_Gride2()
                txtChq_Total.Text = ""
                txtCash.ReadOnly = False
            End If
        End If
    End Sub

    Private Sub cboTearms_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTearms.TextChanged
        Call Calculate_Paid()
        If Trim(cboTearms.Text).Trim = "CASH" Then
            OPR1.Enabled = False
            Call Load_Gride2()
            txtTotal.Text = ""
            txtCash.Text = ""
            txtCash.ReadOnly = False
        ElseIf LTrim(cboTearms.Text).Trim = "CREDIT" Then
            OPR1.Enabled = False
            Call Load_Gride2()
            txtTotal.Text = ""
            txtCash.Text = ""
            txtCash.ReadOnly = True
        ElseIf Trim(cboTearms.Text).Trim = "CREDIT CARD" Then
            OPR1.Visible = True
            OPR1.Enabled = True
            OPR1.Text = "Ezy Payment ..."
            Call Load_Gride3()
            txtChq_Total.Text = ""
            OPR2.Visible = False
            txtCash.Text = ""
            txtCash.ReadOnly = True
            txtTotal.Text = ""
            Call Load_Gride2()
        ElseIf Trim(cboTearms.Text).Trim = "CHQUE" Then
            OPR1.Visible = False
            OPR2.Visible = True
            ' OPR1.Text = "Ezy Payment ..."
            txtCash.Text = ""
            txtCash.ReadOnly = True
            cboBank.Text = ""
            Me.txtChqNo.Text = ""
            Me.txtDOR.Text = ""
            txtAmount.Text = ""
            txtTotal.Text = ""
            Call Load_Gride3()
            txtChq_Total.Text = ""
        ElseIf Trim(cboTearms.Text).Trim = "CASH/CREDIT" Then
            OPR1.Visible = True
            OPR1.Enabled = False
            OPR2.Visible = False
            ' OPR1.Text = "Ezy Payment ..."
            txtCash.Text = ""
            txtCash.ReadOnly = False
            txtTotal.Text = ""
            Call Load_Gride2()
            txtChq_Total.Text = ""

        End If
    End Sub

    Private Sub cboBank_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBank.KeyUp
        If e.KeyCode = 13 Then
            txtChqNo.Focus()
        End If
    End Sub

    Private Sub txtChqNo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChqNo.KeyUp
        If e.KeyCode = 13 Then
            txtDOR.Focus()
        End If
    End Sub

    Private Sub txtDOR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDOR.KeyUp
        If e.KeyCode = 13 Then
            txtAmount.Focus()
        End If
    End Sub

    Private Sub txtCash_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCash.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtCash.Text) Then
                Value = txtCash.Text
                txtCash.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCash.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            cmdAdd.Focus()
        End If
    End Sub

    Private Sub txtCash_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCash.ValueChanged
        Call Calculate_Paid()
    End Sub

    Function Calculate_Paid()
        On Error Resume Next
        Dim Value As Double

        If IsNumeric(txtCash.Text) Then
            Value = txtCash.Text
        End If

        If IsNumeric(txtChq_Total.Text) Then
            Value = Value + CDbl(txtChq_Total.Text)
        End If
        If IsNumeric(txtTotal.Text) Then
            Value = Value + CDbl(txtTotal.Text)
        End If
        txtPaid_Amount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtPaid_Amount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

        If IsNumeric(txtBill_Amount.Text) And IsNumeric(txtPaid_Amount.Text) Then
            Value = CDbl(txtBill_Amount.Text) - CDbl(txtPaid_Amount.Text)
            txtBalance.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtBalance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Function

    Private Sub cboEzy_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEzy.KeyUp
        If e.KeyCode = 13 Then
            txtCard_no.Focus()
        End If
    End Sub

    Private Sub txtCard_no_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCard_no.KeyUp
        If e.KeyCode = 13 Then
            txtCard_Amount.Focus()
        End If
    End Sub

    Private Sub txtCard_Amount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCard_Amount.KeyUp

        Dim VAlue As Double
        Try
            If e.KeyCode = 13 Then
                If Trim(cboEzy.Text) <> "" Then
                Else
                    MsgBox("Please select the card name", MsgBoxStyle.Information, "Information .....")
                    Exit Sub
                End If

                If Trim(txtCard_no.Text) <> "" Then
                Else
                    txtCard_no.Text = "-"
                End If

                If IsNumeric(txtCard_Amount.Text) Then
                Else
                    MsgBox("Please enter the card amount", MsgBoxStyle.Information, "Information ........")
                    Exit Sub
                End If

                If txtCard_Amount.Text <> "" Then
                Else
                    MsgBox("Please enter the card amount", MsgBoxStyle.Information, "Information ........")
                    Exit Sub
                End If

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("#Card Name") = Trim(cboEzy.Text)
                newRow("Card No") = Trim(txtCard_no.Text)
                newRow("Amount") = txtCard_Amount.Text

                c_dataCustomer1.Rows.Add(newRow)
                Me.cboEzy.Text = ""
                Me.txtCard_Amount.Text = ""
                Me.txtCard_no.Text = ""
                cboEzy.ToggleDropdown()
                Call Calculate_Card_Total()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Function Calculate_Chq_Total()
        Dim I As Integer
        Dim Value As Double

        I = 0
        For Each uRow As UltraGridRow In UltraGrid2.Rows
            Value = Value + CDbl(UltraGrid2.Rows(I).Cells(3).Value)
            I = I + 1
        Next
        txtChq_Total.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtChq_Total.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        Call Calculate_Paid()
    End Function

    Function Calculate_Card_Total()
        Dim I As Integer
        Dim Value As Double

        I = 0
        For Each uRow As UltraGridRow In UltraGrid1.Rows
            Value = Value + CDbl(UltraGrid1.Rows(I).Cells(2).Value)
            I = I + 1
        Next
        txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        Call Calculate_Paid()
    End Function

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        Dim Value As Double
        Dim _St As String
        Try
            If e.KeyCode = 13 Then
                If Trim(cboBank.Text) <> "" Then
                Else
                    MsgBox("Please select the Bank Name", MsgBoxStyle.Information, "Information .......")
                    cboBank.ToggleDropdown()
                    Exit Sub
                End If

                If Trim(txtChqNo.Text) <> "" Then
                Else
                    MsgBox("Please enter the Chq No", MsgBoxStyle.Information, "Information .........")
                    txtChqNo.Focus()
                    Exit Sub
                End If

                If Trim(txtDOR.Text) <> "" Then
                Else
                    MsgBox("Please enter the Chq Realize date", MsgBoxStyle.Information, "Information ........")
                    txtDOR.Focus()
                    Exit Sub
                End If

                If IsDate(txtDOR.Text) Then
                Else
                    MsgBox("Please enter the correct Chq Realize date", MsgBoxStyle.Information, "Information .......")
                    txtDOR.Focus()
                    Exit Sub
                End If

                If txtAmount.Text <> "" Then
                Else
                    MsgBox("Please enter the Chq Amount", MsgBoxStyle.Information, "Information ........")
                    txtAmount.Focus()
                    Exit Sub
                End If

                If IsNumeric(txtAmount.Text) Then
                Else
                    MsgBox("Please enter the correct Chq Amount", MsgBoxStyle.Information, "Information ........")
                    txtAmount.Focus()
                    Exit Sub
                End If

                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("#Bank Name") = Trim(cboBank.Text)
                newRow("Chq No") = Trim(txtChqNo.Text)
                newRow("DOR") = txtDOR.Text
                Value = txtAmount.Text
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Chq Amount") = _St

                c_dataCustomer2.Rows.Add(newRow)

                Me.cboBank.Text = ""
                Me.txtDOR.Text = ""
                Me.txtChqNo.Text = ""
                Me.txtAmount.Text = ""
                cboBank.ToggleDropdown()
                Call Calculate_Chq_Total()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub UltraGrid2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid2.AfterRowsDeleted
        Call Calculate_Chq_Total()
    End Sub

    Private Sub UltraGrid1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.AfterRowsDeleted
        Call Calculate_Card_Total()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_text()
    End Sub

    Function Clear_text()
        Me.txtBalance.Text = ""
        Me.txtCard_Amount.Text = ""
        Me.txtPaid_Amount.Text = ""
        Me.txtCard_no.Text = ""
        Me.txtCard_no.Text = ""
        Me.txtChqNo.Text = ""
        Me.txtChq_Total.Text = ""
        Call Load_Gride2()
        Call Load_Gride3()

    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If txtCash.Text <> "" Then
        Else
            txtCash.Text = "0"
        End If

        If IsNumeric(txtCash.Text) Then
        Else
            MsgBox("Please enter the correct cash amount", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If txtTotal.Text <> "" Then
        Else
            txtTotal.Text = "0"
        End If

        If txtChq_Total.Text <> "" Then
        Else
            txtChq_Total.Text = "0"
        End If

        Call Calculate_Paid()

        If Trim(cboTearms.Text).Trim = "CASH" Then

            If CDbl(txtBalance.Text) > 0 Then
                MsgBox("Please enter the cash amount", MsgBoxStyle.Information, "Information ......")
                Exit Sub
            End If
        ElseIf Trim(cboTearms.Text).Trim = "CREDIT" Then
            'If CDbl(txtBalance.Text) > 0 Then
            '    MsgBox("Please select the correct payment terms", MsgBoxStyle.Information, "Information .......")
            '    Exit Sub
            'End If
        ElseIf Trim(cboTearms.Text).Trim = "CASH/CREDIT" Then
            'If CDbl(txtBalance.Text) < 0 Then
            '    MsgBox("Please check the cash amount", MsgBoxStyle.Information, "Information ........")
            '    Exit Sub
            'End If
        ElseIf Trim(cboTearms.Text).Trim = "CHQUE" Then
            If CDbl(txtBalance.Text) > 0 Then
                MsgBox("Please check the cheque amount", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If
        ElseIf Trim(cboTearms.Text).Trim = "CREDIT CARD" Then
            If CDbl(txtBalance.Text) > 0 Then
                MsgBox("Please check the credit card amount", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If
        End If
        frmDirect_Sales.SAVE_DATA()
        Me.Close()
    End Sub

    Private Sub txtAmount_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.ValueChanged

    End Sub
End Class