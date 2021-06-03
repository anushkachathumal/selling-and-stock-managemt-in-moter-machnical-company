Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports Microsoft.Office.Interop.Excel

Public Class frmCost
    Dim c_dataCustomer1 As DataTable
    Dim _EMP1 As String
    Dim _EMP2 As String
    Dim _EMP3 As String
    Dim _EMP4 As String

    Private Sub frmCost_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtExtra_Cast.ReadOnly = True
        txtExtra_Cast.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDaily_Salary.ReadOnly = True
        txtRef.ReadOnly = True
        txtRef.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDaily_Salary.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtGada1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtGada2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtGada3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtGada4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtTot_Gada1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTot_Gada1.ReadOnly = True
        txtTot_Gada2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTot_Gada2.ReadOnly = True
        txtTotal_Laber.ReadOnly = True
        txtTotal_Laber.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRef.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate.Text = Today

        Call Load_Parameter()
        Call Load_NAME()

        txtTotal_Gada.ReadOnly = True
        txtTotal_Gada.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_Unit.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_Cost_Gas.ReadOnly = True
        txtTotal_Cost_Gas.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        '  txtGas_Cost.ReadOnly = True
        txtGas_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtN2_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_N2_Cost.ReadOnly = True
        txtTotal_N2_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtEx_Gada_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEx_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEx_Total_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEx_Total_Cost.ReadOnly = True
        txtCov_Powder.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRed_Cement.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtCotton_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCotton_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_Cotton_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_Cotton_Cost.ReadOnly = True

        txtLeather_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLeather_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLeather_Total.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLeather_Total.ReadOnly = True

        txtSocks_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSocks_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_Socks.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_Socks.ReadOnly = True
        Call Load_Amount()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Parameter()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='CS' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtRef.Text = M01.Tables(0).Rows(0)("P01NO")
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

    Function Load_NAME()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Employee_Name as [##] from M01Employee_Master order by M01Employee_Name "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboEmp1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 240
                ' .Rows.Band.Columns(1).Width = 180
            End With

            With cboEmp2
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 240
                ' .Rows.Band.Columns(1).Width = 180
            End With

            With cboEmp3
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 240
                ' .Rows.Band.Columns(1).Width = 180
            End With

            With cboEmp4
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 240
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

    Private Sub cboEmp1_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEmp1.AfterCloseUp
        Call BASIC_SALARY()
    End Sub

    Private Sub cboEmp1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEmp1.KeyUp
        If e.KeyCode = 13 Then
            txtGada1.Focus()
        End If
    End Sub

    Private Sub txtGada1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGada1.KeyUp
        If e.KeyCode = 13 Then
            cboEmp2.ToggleDropdown()
        End If
    End Sub

    Private Sub txtGada1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGada1.TextChanged
        Dim _VALUE1 As Double
        Dim _VALUE2 As Double
        Dim _EXTRACASTING As Double
        Dim VALUE As Double

        _VALUE1 = 0
        _VALUE2 = 0
        If txtGada1.Text <> "" Then
            If IsNumeric(txtGada1.Text) Then
                _VALUE1 = txtGada1.Text
            End If
        End If

        If txtGada2.Text <> "" Then
            If IsNumeric(txtGada2.Text) Then
                _VALUE2 = txtGada2.Text
            End If
        End If

        txtTot_Gada1.Text = _VALUE1 + _VALUE2
        If txtTot_Gada1.Text <> "" Then
        Else
            txtTot_Gada1.Text = "0"
        End If

        _EXTRACASTING = 0

        If txtTot_Gada1.Text > 100 Then
            _EXTRACASTING = txtTot_Gada1.Text - 100
        End If

        If txtTot_Gada2.Text <> "" Then
        Else
            txtTot_Gada2.Text = "0"
        End If

        If txtTot_Gada2.Text > 100 Then
            _EXTRACASTING = _EXTRACASTING + (txtTot_Gada2.Text - 100)
        End If
        VALUE = _EXTRACASTING * 20

        txtExtra_Cast.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        If txtDaily_Salary.Text <> "" Then
        Else
            txtDaily_Salary.Text = "0"
        End If

        If txtExtra_Cast.Text <> "" Then
        Else
            txtExtra_Cast.Text = "0"
        End If

        VALUE = CDbl(txtDaily_Salary.Text) + CDbl(txtExtra_Cast.Text)
        txtTotal_Laber.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)

        Call CALCULATE_TOTALGADA()
    End Sub

    Private Sub cboEmp2_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEmp2.AfterCloseUp
        Call BASIC_SALARY()
    End Sub

    Private Sub cboEmp2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEmp2.KeyUp
        If e.KeyCode = 13 Then
            txtGada2.Focus()
        End If
    End Sub

    Private Sub txtGada2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGada2.KeyUp
        If e.KeyCode = 13 Then
            cboEmp3.ToggleDropdown()
        End If
    End Sub

    Private Sub txtGada2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGada2.TextChanged
        Dim _VALUE1 As Double
        Dim _VALUE2 As Double
        Dim VALUE As Double
        Dim _EXTRACASTING As Double

        _VALUE1 = 0
        _VALUE2 = 0
        If txtGada1.Text <> "" Then
            If IsNumeric(txtGada1.Text) Then
                _VALUE1 = txtGada1.Text
            End If
        End If

        If txtGada2.Text <> "" Then
            If IsNumeric(txtGada2.Text) Then
                _VALUE2 = txtGada2.Text
            End If
        End If

        txtTot_Gada1.Text = _VALUE1 + _VALUE2

        If txtTot_Gada1.Text <> "" Then
        Else
            txtTot_Gada1.Text = "0"
        End If

        _EXTRACASTING = 0

        If txtTot_Gada1.Text > 100 Then
            _EXTRACASTING = txtTot_Gada1.Text - 100
        End If

        If txtTot_Gada2.Text <> "" Then
        Else
            txtTot_Gada2.Text = "0"
        End If

        If txtTot_Gada2.Text > 100 Then
            _EXTRACASTING = _EXTRACASTING + (txtTot_Gada2.Text - 100)
        End If
        VALUE = _EXTRACASTING * 20

        txtExtra_Cast.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        If txtDaily_Salary.Text <> "" Then
        Else
            txtDaily_Salary.Text = "0"
        End If

        If txtExtra_Cast.Text <> "" Then
        Else
            txtExtra_Cast.Text = "0"
        End If

        VALUE = CDbl(txtDaily_Salary.Text) + CDbl(txtExtra_Cast.Text)
        txtTotal_Laber.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)

        Call CALCULATE_TOTALGADA()
    End Sub

    Private Sub txtGada3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGada3.KeyUp
        If e.KeyCode = 13 Then
            cboEmp4.ToggleDropdown()
        End If
    End Sub

    Private Sub txtGada3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGada3.TextChanged
        Dim _VALUE1 As Double
        Dim _VALUE2 As Double
        Dim VALUE As Double
        Dim _EXTRACASTING As Double


        _VALUE1 = 0
        _VALUE2 = 0
        If txtGada3.Text <> "" Then
            If IsNumeric(txtGada3.Text) Then
                _VALUE1 = txtGada3.Text
            End If
        End If

        If txtGada4.Text <> "" Then
            If IsNumeric(txtGada4.Text) Then
                _VALUE2 = txtGada4.Text
            End If
        End If

        txtTot_Gada2.Text = _VALUE1 + _VALUE2


        If txtTot_Gada1.Text <> "" Then
        Else
            txtTot_Gada1.Text = "0"
        End If

        _EXTRACASTING = 0

        If txtTot_Gada1.Text > 100 Then
            _EXTRACASTING = txtTot_Gada1.Text - 100
        End If

        If txtTot_Gada2.Text <> "" Then
        Else
            txtTot_Gada2.Text = "0"
        End If

        If txtTot_Gada2.Text > 100 Then
            _EXTRACASTING = _EXTRACASTING + (txtTot_Gada2.Text - 100)
        End If
        VALUE = _EXTRACASTING * 20

        txtExtra_Cast.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        If txtDaily_Salary.Text <> "" Then
        Else
            txtDaily_Salary.Text = "0"
        End If

        If txtExtra_Cast.Text <> "" Then
        Else
            txtExtra_Cast.Text = "0"
        End If

        VALUE = CDbl(txtDaily_Salary.Text) + CDbl(txtExtra_Cast.Text)
        txtTotal_Laber.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)

        Call CALCULATE_TOTALGADA()
    End Sub

    Function BASIC_SALARY()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim VALUE As Double
        Try
            VALUE = 0
            If cboEmp1.Text <> "" Then
                Sql = "select * from M01Employee_Master WHERE M01Employee_Name='" & Trim(cboEmp1.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    VALUE = M01.Tables(0).Rows(0)("M01Basic_Salary") / 25
                End If
            End If

            If cboEmp2.Text <> "" Then
                Sql = "select * from M01Employee_Master WHERE M01Employee_Name='" & Trim(cboEmp2.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    VALUE = VALUE + (M01.Tables(0).Rows(0)("M01Basic_Salary") / 25)
                End If
            End If

            If cboEmp3.Text <> "" Then
                Sql = "select * from M01Employee_Master WHERE M01Employee_Name='" & Trim(cboEmp3.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    VALUE = VALUE + (M01.Tables(0).Rows(0)("M01Basic_Salary") / 25)
                End If
            End If

            If cboEmp4.Text <> "" Then
                Sql = "select * from M01Employee_Master WHERE M01Employee_Name='" & Trim(cboEmp4.Text) & "' "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M01) Then
                    VALUE = VALUE + (M01.Tables(0).Rows(0)("M01Basic_Salary") / 25)
                End If
            End If

            txtDaily_Salary.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cboEmp3_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEmp3.AfterCloseUp
        Call BASIC_SALARY()
    End Sub

    Private Sub cboEmp4_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEmp4.AfterCloseUp
        Call BASIC_SALARY()
    End Sub

    Private Sub txtGada4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGada4.KeyUp
        If e.KeyCode = 13 Then
            txtGas_Cost.Focus()
        End If
    End Sub


    Private Sub txtGada4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGada4.TextChanged
        Dim _VALUE1 As Double
        Dim _VALUE2 As Double
        Dim VALUE As Double
        Dim _EXTRACASTING As Double


        _VALUE1 = 0
        _VALUE2 = 0
        If txtGada3.Text <> "" Then
            If IsNumeric(txtGada3.Text) Then
                _VALUE1 = txtGada3.Text
            End If
        End If

        If txtGada4.Text <> "" Then
            If IsNumeric(txtGada4.Text) Then
                _VALUE2 = txtGada4.Text
            End If
        End If

        txtTot_Gada2.Text = _VALUE1 + _VALUE2


        If txtTot_Gada1.Text <> "" Then
        Else
            txtTot_Gada1.Text = "0"
        End If

        _EXTRACASTING = 0

        If txtTot_Gada1.Text > 100 Then
            _EXTRACASTING = txtTot_Gada1.Text - 100
        End If

        If txtTot_Gada2.Text <> "" Then
        Else
            txtTot_Gada2.Text = "0"
        End If

        If txtTot_Gada2.Text > 100 Then
            _EXTRACASTING = _EXTRACASTING + (txtTot_Gada2.Text - 100)
        End If
        VALUE = _EXTRACASTING * 20

        txtExtra_Cast.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        If txtDaily_Salary.Text <> "" Then
        Else
            txtDaily_Salary.Text = "0"
        End If

        If txtExtra_Cast.Text <> "" Then
        Else
            txtExtra_Cast.Text = "0"
        End If

        VALUE = CDbl(txtDaily_Salary.Text) + CDbl(txtExtra_Cast.Text)
        txtTotal_Laber.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)

        Call CALCULATE_TOTALGADA()
    End Sub

    Private Sub cboEmp3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEmp3.KeyUp
        If e.KeyCode = 13 Then
            txtGada3.Focus()
        End If
    End Sub

    Private Sub cboEmp4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEmp4.KeyUp
        If e.KeyCode = 13 Then
            txtGada4.Focus()
        End If
    End Sub

    Function CALCULATE_TOTALGADA()
        Dim VALUE As Double
        Dim VALUE1 As Double

        VALUE = 0
        VALUE1 = 0

        If txtTot_Gada1.Text <> "" Then
            VALUE = txtTot_Gada1.Text
        Else

        End If

        If txtTot_Gada2.Text <> "" Then
            VALUE1 = txtTot_Gada2.Text
        Else

        End If

        txtTotal_Gada.Text = VALUE + VALUE1
    End Function

    Private Sub txtGas_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGas_Cost.KeyUp
        If e.KeyCode = 13 Then
            txtTotal_Unit.Focus()
        End If
    End Sub

    Private Sub txtGas_Cost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGas_Cost.TextChanged
        Dim Value As Double

        If IsNumeric(txtGas_Cost.Text) And IsNumeric(txtTotal_Gada.Text) Then
            Value = txtTotal_Gada.Text * txtGas_Cost.Text
            txtTotal_Cost_Gas.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        End If
    End Sub

    Private Sub txtN2_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtN2_Cost.KeyUp
        If e.KeyCode = 13 Then
            txtEx_Gada_Cost.Focus()
        End If
    End Sub

    Private Sub txtN2_Cost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtN2_Cost.TextChanged
        'Dim Value As Double
        'If IsNumeric(txtTotal_Gada.Text) And IsNumeric(txtN2_Cost.Text) Then
        '    Value = txtTotal_Gada.Text * txtN2_Cost.Text
        '    txtTotal_N2_Cost.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        'End If

        Call Calculation_N2()
    End Sub

    Private Sub txtTotal_Unit_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotal_Unit.KeyUp
        If e.KeyCode = 13 Then
            txtN2_Cost.Focus()
        End If
    End Sub

    Private Sub txtN2_Cost_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtN2_Cost.ValueChanged

    End Sub

    Private Sub txtEx_Gada_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEx_Gada_Cost.KeyUp
        If e.KeyCode = 13 Then
            txtEx_Qty.Focus()
        End If
    End Sub

    Private Sub txtEx_Gada_Cost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEx_Gada_Cost.TextChanged
        Call Calculation_Gada()
    End Sub

    Function Calculation_Gada()
        Dim Value As Double

        If IsNumeric(txtEx_Gada_Cost.Text) And IsNumeric(txtEx_Qty.Text) Then
            Value = txtEx_Gada_Cost.Text * txtEx_Qty.Text
            txtEx_Total_Cost.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        End If
    End Function

    Private Sub txtEx_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEx_Qty.KeyUp
        If e.KeyCode = 13 Then
            txtCov_Powder.Focus()
        End If
    End Sub

    Private Sub txtEx_Qty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEx_Qty.TextChanged
        Call Calculation_Gada()
    End Sub


    Function Calculation_N2()
        Dim Value As Double

        If IsNumeric(txtN2_Cost.Text) And IsNumeric(txtTotal_Gada.Text) Then
            Value = txtN2_Cost.Text * txtTotal_Gada.Text
            txtTotal_N2_Cost.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        End If
    End Function

    Private Sub txtCov_Powder_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCov_Powder.KeyUp
        If e.KeyCode = 13 Then
            txtRed_Cement.Focus()
        End If
    End Sub

    Function CALCULATION_COTTON()
        Dim VALUE As Double
        If IsNumeric(txtCotton_Cost.Text) And IsNumeric(txtCotton_Qty.Text) Then
            VALUE = txtCotton_Cost.Text * txtCotton_Qty.Text
            txtTotal_Cotton_Cost.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        End If
    End Function

    Function CALCULATION_LEATHER()
        Dim VALUE As Double
        If IsNumeric(txtLeather_Cost.Text) And IsNumeric(txtLeather_Qty.Text) Then
            VALUE = txtLeather_Cost.Text * txtLeather_Qty.Text
            txtLeather_Total.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        End If
    End Function

    Function CALCULATION_SOCKS()
        Dim VALUE As Double
        If IsNumeric(txtSocks_Cost.Text) And IsNumeric(txtSocks_Qty.Text) Then
            VALUE = txtSocks_Cost.Text * txtSocks_Qty.Text
            txtTotal_Socks.Text = VALUE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        End If
    End Function

    Private Sub txtCotton_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCotton_Cost.KeyUp
        If e.KeyCode = 13 Then
            txtCotton_Qty.Focus()
        End If
    End Sub

    Private Sub txtCotton_Cost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCotton_Cost.TextChanged
        Call CALCULATION_COTTON()
    End Sub

    Private Sub txtCotton_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCotton_Qty.KeyUp
        If e.KeyCode = 13 Then
            txtLeather_Cost.Focus()
        End If
    End Sub


    Private Sub txtCotton_Qty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCotton_Qty.TextChanged
        Call CALCULATION_COTTON()
    End Sub

    Private Sub txtRed_Cement_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRed_Cement.KeyUp
        If e.KeyCode = 13 Then
            txtCotton_Cost.Focus()
        End If
    End Sub

    Private Sub txtLeather_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLeather_Cost.KeyUp
        If e.KeyCode = 13 Then
            txtLeather_Qty.Focus()
        End If
    End Sub

    Private Sub txtLeather_Cost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeather_Cost.TextChanged
        Call CALCULATION_LEATHER()
    End Sub

    Private Sub txtLeather_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLeather_Qty.KeyUp
        If e.KeyCode = 13 Then
            txtSocks_Cost.Focus()
        End If
    End Sub

    Private Sub txtLeather_Qty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeather_Qty.TextChanged
        Call CALCULATION_LEATHER()
    End Sub

    Private Sub txtSocks_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSocks_Cost.KeyUp
        If e.KeyCode = 13 Then
            txtSocks_Qty.Focus()
        End If
    End Sub

    Private Sub txtSocks_Cost_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSocks_Cost.TextChanged
        Call CALCULATION_SOCKS()
    End Sub

    Private Sub txtSocks_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSocks_Qty.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        End If
    End Sub

    Private Sub txtSocks_Qty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSocks_Qty.TextChanged
        Call CALCULATION_SOCKS()
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
        Dim t01 As DataSet


        Try
            'EMPLOYEE
            If txtGada1.Text <> "" Then

            Else
                txtGada1.Text = "0"
            End If

            If txtGada2.Text <> "" Then

            Else
                txtGada2.Text = "0"
            End If

            If txtGada3.Text <> "" Then

            Else
                txtGada3.Text = "0"
            End If

            If txtGada4.Text <> "" Then

            Else
                txtGada4.Text = "0"
            End If


            If IsNumeric(txtGada1.Text) Then
            Else
                MsgBox("Please enter the correct Gada", MsgBoxStyle.Information, "Information ........")
                txtGada1.Focus()
                connection.Close()
                Exit Sub
            End If

            If IsNumeric(txtGada2.Text) Then
            Else
                MsgBox("Please enter the correct Gada", MsgBoxStyle.Information, "Information ........")
                txtGada2.Focus()
                connection.Close()
                Exit Sub
            End If

            If IsNumeric(txtGada3.Text) Then
            Else
                MsgBox("Please enter the correct Gada", MsgBoxStyle.Information, "Information ........")
                txtGada3.Focus()
                connection.Close()
                Exit Sub
            End If

            If IsNumeric(txtGada4.Text) Then
            Else
                MsgBox("Please enter the correct Gada", MsgBoxStyle.Information, "Information ........")
                txtGada4.Focus()
                connection.Close()
                Exit Sub
            End If

            Call CALCULATE_TOTALGADA()

            If cboEmp1.Text <> "" Then
                nvcFieldList1 = "SELECT * FROM M01Employee_Master WHERE M01Employee_Name='" & cboEmp1.Text & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    _EMP1 = Trim(t01.Tables(0).Rows(0)("M01Emp_No"))
                End If
            End If


            If cboEmp2.Text <> "" Then
                nvcFieldList1 = "SELECT * FROM M01Employee_Master WHERE M01Employee_Name='" & cboEmp2.Text & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    _EMP2 = Trim(t01.Tables(0).Rows(0)("M01Emp_No"))
                End If
            End If

            If cboEmp3.Text <> "" Then
                nvcFieldList1 = "SELECT * FROM M01Employee_Master WHERE M01Employee_Name='" & cboEmp3.Text & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    _EMP3 = Trim(t01.Tables(0).Rows(0)("M01Emp_No"))
                End If
            End If

            If cboEmp4.Text <> "" Then
                nvcFieldList1 = "SELECT * FROM M01Employee_Master WHERE M01Employee_Name='" & cboEmp4.Text & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    _EMP4 = Trim(t01.Tables(0).Rows(0)("M01Emp_No"))
                End If
            End If
            '==========================================================================================================
            'GAS
            If txtTotal_Unit.Text <> "" Then
            Else
                txtTotal_Unit.Text = "0"
            End If

            If txtGas_Cost.Text <> "" Then
            Else
                txtGas_Cost.Text = "0"
            End If

            If IsNumeric(txtTotal_Unit.Text) Then
            Else
                MsgBox("Please enter the correct gas unit", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtTotal_Unit.Focus()
                Exit Sub
            End If

            If IsNumeric(txtGas_Cost.Text) Then
            Else
                MsgBox("Please enter the correct gas cost", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtGas_Cost.Focus()
                Exit Sub
            End If
            '--------------------------------------------------------------------------------------------------
            'N2
            Call Calculation_N2()
            If txtN2_Cost.Text <> "" Then
            Else
                txtN2_Cost.Text = "0"
            End If


            If IsNumeric(txtN2_Cost.Text) Then
            Else
                MsgBox("Please enter the correct n2 cost", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtN2_Cost.Focus()
                Exit Sub
            End If
            '---------------------------------------------------------------------------------------------------
            'GADA
            Call Calculation_Gada()
            If txtEx_Gada_Cost.Text <> "" Then
            Else
                txtEx_Gada_Cost.Text = "0"
            End If

            If txtEx_Qty.Text <> "" Then
            Else
                txtEx_Qty.Text = "0"
            End If

            If IsNumeric(txtEx_Gada_Cost.Text) Then
            Else
                MsgBox("Please enter the correct Gada cost", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtEx_Gada_Cost.Focus()
                Exit Sub
            End If

            If IsNumeric(txtEx_Qty.Text) Then
            Else
                MsgBox("Please enter the correct Gada Qty", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtEx_Qty.Focus()
                Exit Sub
            End If
            '----------------------------------------------------------------------------------------------------
            If txtCov_Powder.Text <> "" Then
            Else
                txtCov_Powder.Text = "0"
            End If

            If IsNumeric(txtCov_Powder.Text) Then
            Else
                MsgBox("Please enter the correct Cost for C.Powder", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtCov_Powder.Focus()
                Exit Sub
            End If
            '----------------------------------------------------------------------------------------------------
            If txtRed_Cement.Text <> "" Then
            Else
                txtRed_Cement.Text = "0"
            End If

            If IsNumeric(txtRed_Cement.Text) Then
            Else
                MsgBox("Please enter the correct Cost of Red Cement", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtRed_Cement.Focus()
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------
            'COTTON CLOVE
            Call CALCULATION_COTTON()
            If txtCotton_Cost.Text <> "" Then
            Else
                txtCotton_Cost.Text = "0"
            End If

            If IsNumeric(txtCotton_Cost.Text) Then
            Else
                MsgBox("Please enter the correct Cost of Cotton/G", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtCotton_Cost.Focus()
                Exit Sub
            End If

            If txtCotton_Qty.Text <> "" Then
            Else
                txtCotton_Qty.Text = "0"
            End If

            If IsNumeric(txtCotton_Qty.Text) Then
            Else
                MsgBox("Please enter the correct Qty of Cotton/G", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtCotton_Qty.Focus()
                Exit Sub
            End If
            '---------------------------------------------------------------------------------------------------------
            'LETHER CLOVE
            Call CALCULATION_LEATHER()
            If txtLeather_Cost.Text <> "" Then
            Else
                txtLeather_Cost.Text = "0"
            End If

            If IsNumeric(txtLeather_Cost.Text) Then
            Else
                MsgBox("Please enter the correct Cost of Leather/G", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtLeather_Cost.Focus()
                Exit Sub
            End If

            If txtLeather_Qty.Text <> "" Then
            Else
                txtLeather_Qty.Text = "0"
            End If

            If IsNumeric(txtLeather_Qty.Text) Then
            Else
                MsgBox("Please enter the correct Qty of Leather/G", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtLeather_Qty.Focus()
                Exit Sub
            End If
            '---------------------------------------------------------------------------------------------------------
            'SOCKES
            Call CALCULATION_SOCKS()
            If txtSocks_Cost.Text <> "" Then
            Else
                txtSocks_Cost.Text = "0"
            End If

            If IsNumeric(txtSocks_Cost.Text) Then
            Else
                MsgBox("Please enter the correct Cost of Socks", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtLeather_Cost.Focus()
                Exit Sub
            End If

            If txtSocks_Qty.Text <> "" Then
            Else
                txtSocks_Qty.Text = "0"
            End If

            If IsNumeric(txtSocks_Qty.Text) Then
            Else
                MsgBox("Please enter the correct Qty of Socks", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtSocks_Qty.Focus()
                Exit Sub
            End If
            '------------------------------------------------------------------------------------------------------
            'UPDATE DATA
            Call Load_Parameter()
            nvcFieldList1 = "SELECT * FROM T01Casting_Emp WHERE T01Ref='" & txtRef.Text & "'"
            t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(t01) Then
                nvcFieldList1 = "UPDATE T01Casting_Emp SET T01Date='" & txtDate.Text & "',T01Emp1='" & _EMP1 & "',T01Emp2='" & _EMP2 & "',T01Emp3='" & _EMP3 & "',T01Emp4='" & _EMP4 & "',T01Qty1='" & txtGada1.Text & "',T01Qty1='" & txtGada2.Text & "',T01Qty3='" & txtGada3.Text & "',T01Qty4='" & txtGada4.Text & "' WHERE T01Ref='" & txtRef.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T02Casting_Details SET T02Date='" & txtDate.Text & "',T02Daily_Salary='" & txtDaily_Salary.Text & "',T02Ex_Casting='" & txtExtra_Cast.Text & "',T02Total_Gada='" & txtTotal_Gada.Text & "',T02Cost_Gas='" & txtGas_Cost.Text & "',T02Unit='" & txtTotal_Unit.Text & "',T02N2_Cost='" & txtN2_Cost.Text & "',T02Gada='" & txtEx_Gada_Cost.Text & "',T02Gada_Qty='" & txtEx_Qty.Text & "',T02C_Powder='" & txtCov_Powder.Text & "',T02Red_Cement='" & txtRed_Cement.Text & "',T02Cotton_Cost='" & txtCotton_Cost.Text & "',T02Cotton_Qty='" & txtCotton_Qty.Text & "',T02L_Cost='" & txtLeather_Cost.Text & "',T02L_Qty='" & txtLeather_Qty.Text & "',T02Sock_Cost='" & txtSocks_Cost.Text & "',T02Sock_Qty='" & txtSocks_Qty.Text & "' WHERE T02Ref='" & txtRef.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                Call Load_Parameter()

                nvcFieldList1 = "UPDATE P01PARAMETER SET P01NO=P01NO +" & 1 & " WHERE P01CODE='CS'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T01Casting_Emp(T01Ref,T01Date,T01Emp1,T01Emp2,T01Emp3,T01Emp4,T01Qty1,T01Qty2,T01Qty3,T01Qty4,T01User)" & _
                                                               " values('" & txtRef.Text & "', '" & Trim(txtDate.Text) & "','" & _EMP1 & "','" & _EMP2 & "','" & _EMP3 & "','" & _EMP4 & "','" & txtGada1.Text & "','" & txtGada2.Text & "','" & txtGada3.Text & "','" & txtGada4.Text & "','" & strDisname & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T02Casting_Details(T02Ref,T02Date,T02Daily_Salary,T02Ex_Casting,T02Total_Gada,T02Cost_Gas,T02Unit,T02N2_Cost,T02Gada,T02Gada_Qty,T02C_Powder,T02Red_Cement,T02Cotton_Cost,T02Cotton_Qty,T02L_Cost,T02L_Qty,T02Sock_Cost,T02Sock_Qty,T02User)" & _
                                                              " values('" & txtRef.Text & "', '" & Trim(txtDate.Text) & "','" & txtDaily_Salary.Text & "','" & txtExtra_Cast.Text & "','" & txtTotal_Gada.Text & "','" & txtGas_Cost.Text & "','" & txtTotal_Unit.Text & "','" & txtN2_Cost.Text & "','" & txtEx_Gada_Cost.Text & "','" & txtEx_Qty.Text & "','" & txtCov_Powder.Text & "','" & txtRed_Cement.Text & "','" & txtCotton_Cost.Text & "','" & txtCotton_Qty.Text & "','" & txtLeather_Cost.Text & "','" & txtLeather_Qty.Text & "','" & txtSocks_Cost.Text & "','" & txtSocks_Qty.Text & "','" & strDisname & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information ......")
            transaction.Commit()
            connection.Close()
            Call CLEAR_TEXT()
            Call Load_Parameter()
            Call Load_Amount()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try

    End Sub


    Function CLEAR_TEXT()
        Me.cboEmp1.Text = ""
        Me.cboEmp2.Text = ""
        Me.cboEmp3.Text = ""
        Me.cboEmp4.Text = ""
        Me.txtGada1.Text = ""
        Me.txtGada2.Text = ""
        Me.txtGada3.Text = ""
        Me.txtGada4.Text = ""
        Me.txtTotal_Unit.Text = ""
        Me.txtTot_Gada1.Text = ""
        Me.txtTot_Gada2.Text = ""
        Me.txtDaily_Salary.Text = ""
        Me.txtExtra_Cast.Text = ""
        Me.txtTotal_Laber.Text = ""
        Me.txtGas_Cost.Text = ""
        Me.txtTotal_Unit.Text = ""
        Me.txtTotal_Gada.Text = ""
        Me.txtTotal_Cost_Gas.Text = ""
        Me.txtN2_Cost.Text = ""
        Me.txtTotal_N2_Cost.Text = ""
        Me.txtEx_Qty.Text = ""
        Me.txtEx_Total_Cost.Text = ""
        Me.txtEx_Gada_Cost.Text = ""
        Me.txtCov_Powder.Text = ""
        Me.txtRed_Cement.Text = ""
        Me.txtCotton_Cost.Text = ""
        Me.txtCotton_Qty.Text = ""
        Me.txtTotal_Cotton_Cost.Text = ""
        Me.txtLeather_Cost.Text = ""
        Me.txtLeather_Qty.Text = ""
        Me.txtLeather_Total.Text = ""
        Me.txtSocks_Cost.Text = ""
        Me.txtSocks_Qty.Text = ""
        Me.txtTotal_Socks.Text = ""
        Me.lblTotal.Text = "00.00"
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call CLEAR_TEXT()
    End Sub

   
    Function Load_Amount()
        Dim Value As Double
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim _St As String
        Try
            Sql = "select sum(T02Cost_Gas*T02Total_Gada) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(21).SharedProps.Caption = "GAS " & _St
            End If


            Sql = "select sum(T02N2_Cost*T02Total_Gada) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(22).SharedProps.Caption = "N2 " & _St
            End If

            Sql = "select sum(T02Sock_Cost*T02Sock_Qty) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(23).SharedProps.Caption = "Sockes " & _St
            End If

            Sql = "select sum(T02Red_Cement) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(24).SharedProps.Caption = "Red Cement " & _St
            End If

            Sql = "select sum(T02Red_Cement) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(25).SharedProps.Caption = "Gada " & _St
            End If

            Sql = "select sum(T02Gada_Qty*T02Gada) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(25).SharedProps.Caption = "Gada " & _St
            End If

            Sql = "select sum(T02L_Cost*T02L_Qty) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(26).SharedProps.Caption = "Leather/G " & _St
            End If

            Sql = "select sum(T02Cotton_Qty*T02Cotton_Cost) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(27).SharedProps.Caption = "Cotton/G " & _St
            End If


            Sql = "select sum(T02C_Powder) as Qty from T02Casting_Details where month(T02Date)='" & Month(Today) & "' and year(T02Date)='" & Year(Today) & "' group by month(T02Date),year(T02Date)"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(28).SharedProps.Caption = "Coveall Powder " & _St
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try

    End Function

    Function Create_File()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim T02 As DataSet
        Dim n_Date As Date
        Dim exc As New Application
        Dim _from As Date
        Dim workbooks As Workbooks = exc.Workbooks
        Dim workbook As _Workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Sheets = workbook.Worksheets
        Dim worksheet1 As _Worksheet = CType(sheets.Item(1), _Worksheet)
        Dim _rowcount As Integer
        Dim range1 As Range
        Dim _Emp As String

        exc.Visible = True

        worksheet1.Rows(2).Font.size = 18
        worksheet1.Rows(2).Font.name = "Tahoma"
        worksheet1.Rows(2).rowheight = 24.25
        worksheet1.Rows(6).rowheight = 15.25
        worksheet1.Rows(7).rowheight = 15.25
        worksheet1.Rows(8).rowheight = 20.25

       
        worksheet1.Columns("C").ColumnWidth = 40
        worksheet1.Columns("E").ColumnWidth = 40
        worksheet1.Columns("F").ColumnWidth = 15
        worksheet1.Columns("G").ColumnWidth = 2
        worksheet1.Columns("H").ColumnWidth = 15
        worksheet1.Columns("I").ColumnWidth = 15
        worksheet1.Columns("J").ColumnWidth = 15
        worksheet1.Columns("K").ColumnWidth = 15
        worksheet1.Columns("L").ColumnWidth = 8
        worksheet1.Columns("M").ColumnWidth = 15
        worksheet1.Columns("N").ColumnWidth = 8
        worksheet1.Columns("O").ColumnWidth = 15
        worksheet1.Columns("P").ColumnWidth = 8
        worksheet1.Columns("Q").ColumnWidth = 15
        worksheet1.Columns("R").ColumnWidth = 15
        worksheet1.Columns("S").ColumnWidth = 15



        worksheet1.Rows(6).Font.size = 9
        worksheet1.Rows(6).Font.name = "Tahoma"
        worksheet1.Rows(6).Font.Bold = True

        worksheet1.Rows(7).Font.size = 9
        worksheet1.Rows(7).Font.name = "Tahoma"
        worksheet1.Rows(7).Font.Bold = True

        worksheet1.Rows(8).Font.size = 9
        worksheet1.Rows(8).Font.name = "Tahoma"
        worksheet1.Rows(8).Font.Bold = True


        worksheet1.Cells(6, 1) = "Date"
        worksheet1.Range("a6:A8").MergeCells = True
        worksheet1.Range("A6:A6").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 1).WrapText = True
        worksheet1.Cells(6, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 2) = "A"
        worksheet1.Range("B6:C6").MergeCells = True
        worksheet1.Range("B6:C6").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 2).WrapText = True
        worksheet1.Cells(6, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet1.Cells(6, 4) = "B"

        worksheet1.Range("D6:E6").MergeCells = True
        worksheet1.Range("D6:E6").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 4).WrapText = True
        worksheet1.Cells(6, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 2) = "Gada"
        worksheet1.Range("B7:B8").MergeCells = True
        worksheet1.Range("B7:B8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 2).WrapText = True
        worksheet1.Cells(7, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 3) = "Workers"
        worksheet1.Range("C7:C8").MergeCells = True
        worksheet1.Range("C7:C8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 3).WrapText = True
        worksheet1.Cells(7, 3).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 4) = "Gada"
        worksheet1.Range("D7:D8").MergeCells = True
        worksheet1.Range("D7:D8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 4).WrapText = True
        worksheet1.Cells(7, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter


        worksheet1.Cells(7, 5) = "Workers"
        worksheet1.Range("E7:E8").MergeCells = True
        worksheet1.Range("E7:E8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 5).WrapText = True
        worksheet1.Cells(7, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(6, 6) = "Total Gada"
        worksheet1.Range("f6:f8").MergeCells = True
        worksheet1.Range("f6:f6").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 6).WrapText = True
        worksheet1.Cells(6, 6).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(6, 8) = "COST"
        worksheet1.Range("H6:S6").MergeCells = True
        worksheet1.Range("H6:S6").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(6, 8).WrapText = True
        worksheet1.Cells(6, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 8) = "Gas"
        worksheet1.Range("h7:h8").MergeCells = True
        worksheet1.Range("h7:h8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 8).WrapText = True
        worksheet1.Cells(7, 8).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 9) = "N2"
        worksheet1.Range("i7:i8").MergeCells = True
        worksheet1.Range("i7:i8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 9).WrapText = True
        worksheet1.Cells(7, 9).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 10) = "Gada"
        worksheet1.Range("J7:J8").MergeCells = True
        worksheet1.Range("J7:J8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 10).WrapText = True
        worksheet1.Cells(7, 10).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 11) = "Coveall Powder"
        worksheet1.Range("K7:K8").MergeCells = True
        worksheet1.Range("K7:K8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 11).WrapText = True
        worksheet1.Cells(7, 11).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 12) = "Cotton/G"
        worksheet1.Range("L7:M7").MergeCells = True
        worksheet1.Range("L7:M7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 12).WrapText = True
        worksheet1.Cells(7, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(8, 12) = "Qty"
        worksheet1.Range("L7:L7").MergeCells = True
        worksheet1.Range("L7:L7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(8, 12).WrapText = True
        worksheet1.Cells(8, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(8, 13) = "Price"
        worksheet1.Range("M7:M7").MergeCells = True
        worksheet1.Range("M7:M7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(8, 13).WrapText = True
        worksheet1.Cells(8, 13).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 14) = "Leather/G"
        worksheet1.Range("N7:O7").MergeCells = True
        worksheet1.Range("N7:O7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 14).WrapText = True
        worksheet1.Cells(7, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(8, 14) = "Qty"
        worksheet1.Range("N7:N7").MergeCells = True
        worksheet1.Range("N7:N7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(8, 14).WrapText = True
        worksheet1.Cells(8, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(8, 15) = "Price"
        worksheet1.Range("O7:O7").MergeCells = True
        worksheet1.Range("O7:O7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(8, 15).WrapText = True
        worksheet1.Cells(8, 15).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 16) = "Socks"
        worksheet1.Range("P7:Q7").MergeCells = True
        worksheet1.Range("P7:Q7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 16).WrapText = True
        worksheet1.Cells(7, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(8, 16) = "Qty"
        worksheet1.Range("P7:P7").MergeCells = True
        worksheet1.Range("P7:P7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(8, 16).WrapText = True
        worksheet1.Cells(8, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(8, 17) = "Price"
        worksheet1.Range("Q7:Q7").MergeCells = True
        worksheet1.Range("Q7:Q7").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(8, 17).WrapText = True
        worksheet1.Cells(8, 17).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 18) = "Red Cement"
        worksheet1.Range("R7:R8").MergeCells = True
        worksheet1.Range("R7:R8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 18).WrapText = True
        worksheet1.Cells(7, 18).HorizontalAlignment = XlHAlign.xlHAlignCenter

        worksheet1.Cells(7, 19) = "Total Cost"
        worksheet1.Range("S7:S8").MergeCells = True
        worksheet1.Range("S7:S8").VerticalAlignment = XlVAlign.xlVAlignCenter
        worksheet1.Cells(7, 19).WrapText = True
        worksheet1.Cells(7, 19).HorizontalAlignment = XlHAlign.xlHAlignCenter


        worksheet1.Range("G1:G60").MergeCells = True
        worksheet1.Range("G1:G60").VerticalAlignment = XlVAlign.xlVAlignCenter

        Dim _Chart As Integer
        Dim X As Integer
        Dim I As Integer
        X = 6
        For X = 6 To 8
            _Chart = 97
            For I = 1 To 19
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Interior.Color = RGB(0, 32, 96)
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Font.Color = RGB(255, 255, 255)
                'worksheet1.Cells(X, i).WrapText = True
                'worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).MergeCells = True
                'worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).VerticalAlignment = XlVAlign.xlVAlignCenter
                'worksheet1.Cells(X, i).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
                worksheet1.Range(ChrW(_Chart) & X, ChrW(_Chart) & X).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

                _Chart = _Chart + 1
            Next
        Next
        'X = X + 1
        _from = Month(Today) & "/1/" & Year(Today)
        Dim daysInFeb As Integer = System.DateTime.DaysInMonth(Microsoft.VisualBasic.Year(Today), Microsoft.VisualBasic.Month(Today))
        n_Date = Month(Today) & "/" & daysInFeb.ToString & "/" & Year(Today)
        ' n_Date = CDate(n_Date).AddDays(+1)
        Dim TM As TimeSpan
        Dim z As Integer
        TM = n_Date.Subtract(_from)
        X = 0
        _rowcount = 9
        For I = 1 To TM.Days + 1
            worksheet1.Cells(_rowcount, 1) = _from
            worksheet1.Cells(_rowcount, 1).EntireColumn.NumberFormat = "dd-MMM"
            worksheet1.Rows(_rowcount).Font.size = 10
            worksheet1.Cells(_rowcount, 1).WrapText = True
            worksheet1.Cells(_rowcount, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

            SQL = "select sum(T01Qty1+T01Qty2) as SecA,sum(T01Qty3+T01Qty4) as SecB from T01Casting_Emp where T01Date='" & _from & "' group by T01Date"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                worksheet1.Cells(_rowcount, 2) = T01.Tables(0).Rows(0)("SecA")
                worksheet1.Cells(_rowcount, 2).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 2)
                range1.NumberFormat = "0.00"

                worksheet1.Cells(_rowcount, 4) = T01.Tables(0).Rows(0)("SecB")
                worksheet1.Cells(_rowcount, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 4)
                range1.NumberFormat = "0.00"
            End If

            _Emp = ""
            z = 0

            SQL = "select * from T01Casting_Emp where T01Date='" & _from & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                SQL = "select * from M01Employee_Master where M01Emp_No='" & Trim(T01.Tables(0).Rows(0)("T01Emp1")) & "'"
                T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T02) Then
                    If _Emp <> "" Then
                        _Emp = _Emp & "," & T02.Tables(0).Rows(0)("M01Employee_Name")
                    Else
                        _Emp = T02.Tables(0).Rows(0)("M01Employee_Name")
                    End If
                End If

                SQL = "select * from M01Employee_Master where M01Emp_No='" & Trim(T01.Tables(0).Rows(0)("T01Emp2")) & "'"
                T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T02) Then
                    If _Emp <> "" Then
                        _Emp = _Emp & "," & T02.Tables(0).Rows(0)("M01Employee_Name")
                    Else
                        _Emp = T02.Tables(0).Rows(0)("M01Employee_Name")
                    End If
                End If
            End If

            worksheet1.Cells(_rowcount, 3) = _Emp
            _Emp = ""

            SQL = "select * from T01Casting_Emp where T01Date='" & _from & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                SQL = "select * from M01Employee_Master where M01Emp_No='" & Trim(T01.Tables(0).Rows(0)("T01Emp3")) & "'"
                T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T02) Then
                    If _Emp <> "" Then
                        _Emp = _Emp & "," & T02.Tables(0).Rows(0)("M01Employee_Name")
                    Else
                        _Emp = T02.Tables(0).Rows(0)("M01Employee_Name")
                    End If
                End If

                SQL = "select * from M01Employee_Master where M01Emp_No='" & Trim(T01.Tables(0).Rows(0)("T01Emp4")) & "'"
                T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T02) Then
                    If _Emp <> "" Then
                        _Emp = _Emp & "," & T02.Tables(0).Rows(0)("M01Employee_Name")
                    Else
                        _Emp = T02.Tables(0).Rows(0)("M01Employee_Name")
                    End If
                End If
            End If
            worksheet1.Cells(_rowcount, 5) = _Emp

            worksheet1.Range("F" & _rowcount).Formula = "=SUM(B" & _rowcount & ":D" & _rowcount & ") "
            ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
            range1 = worksheet1.Cells(_rowcount, 6)
            range1.NumberFormat = "0.00"
            '-------------------------------------------------------------------------------------
            'Gada
            SQL = "select * from T02Casting_Details where T02Date='" & _from & "'"
            T02 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T02) Then
                ' worksheet1.Range("H" & _rowcount).Formula = "=F" & _rowcount & "*" & Trim(T02.Tables(0).Rows(0)("T02Cost_Gas"))  "
                worksheet1.Cells(_rowcount, 8) = "=F" & _rowcount & "*" & Trim(T02.Tables(0).Rows(0)("T02Cost_Gas"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 8)
                range1.NumberFormat = "0.00"

                worksheet1.Cells(_rowcount, 9) = "=F" & _rowcount & "*" & Trim(T02.Tables(0).Rows(0)("T02N2_Cost"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 9)
                range1.NumberFormat = "0.00"

                worksheet1.Cells(_rowcount, 10) = "=F" & _rowcount & "*" & Trim(T02.Tables(0).Rows(0)("T02Gada"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 10)
                range1.NumberFormat = "0.00"

                worksheet1.Cells(_rowcount, 11) = Trim(T02.Tables(0).Rows(0)("T02C_Powder"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 11)
                range1.NumberFormat = "0.00"

                worksheet1.Cells(_rowcount, 12) = Trim(T02.Tables(0).Rows(0)("T02Cotton_Qty"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 12)
                range1.NumberFormat = "0"
                worksheet1.Cells(_rowcount, 12).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(_rowcount, 13) = "=L" & _rowcount & "*" & Trim(T02.Tables(0).Rows(0)("T02Cotton_Cost"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 13)
                range1.NumberFormat = "0.00"

                worksheet1.Cells(_rowcount, 14) = Trim(T02.Tables(0).Rows(0)("T02L_Qty"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 14)
                range1.NumberFormat = "0"
                worksheet1.Cells(_rowcount, 14).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(_rowcount, 15) = "=N" & _rowcount & "*" & Trim(T02.Tables(0).Rows(0)("T02L_Cost"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 15)
                range1.NumberFormat = "0.00"

                worksheet1.Cells(_rowcount, 16) = Trim(T02.Tables(0).Rows(0)("T02Sock_Qty"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 16)
                range1.NumberFormat = "0"
                worksheet1.Cells(_rowcount, 16).HorizontalAlignment = XlHAlign.xlHAlignCenter

                worksheet1.Cells(_rowcount, 17) = "=P" & _rowcount & "*" & Trim(T02.Tables(0).Rows(0)("T02Sock_Cost"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 17)
                range1.NumberFormat = "0.00"


                worksheet1.Cells(_rowcount, 18) = Trim(T02.Tables(0).Rows(0)("T02Red_Cement"))
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 18)
                range1.NumberFormat = "0.00"


                ' worksheet1.Range("S" & _rowcount).Formula = "=SUM(H" & _rowcount & ":I" & _rowcount & ") "
                worksheet1.Cells(_rowcount, 19) = "=H" & _rowcount & "+I" & _rowcount & "+J" & _rowcount & "+K" & _rowcount & "+M" & _rowcount & "+O" & _rowcount & "+Q" & _rowcount & "+R" & _rowcount
                ' worksheet1.Cells(_rowcount, 5).HorizontalAlignment = XlHAlign.xlHAlignCenter
                range1 = worksheet1.Cells(_rowcount, 19)
                range1.NumberFormat = "0.00"
            End If
            _from = _from.AddDays(+1)
            _rowcount = _rowcount + 1
        Next


        '=========================================================================================
        'TOTAL GADA
        Dim chartPage As Microsoft.Office.Interop.Excel.Chart
        Dim xlCharts As Microsoft.Office.Interop.Excel.ChartObjects
        Dim myChart As Microsoft.Office.Interop.Excel.ChartObject
        Dim chartRange As Microsoft.Office.Interop.Excel.Range
        Dim chartRange1 As Microsoft.Office.Interop.Excel.Range
        Dim chartRange2 As Microsoft.Office.Interop.Excel.Range


        Dim t_SerCol As Microsoft.Office.Interop.Excel.SeriesCollection
        Dim t_Series As Microsoft.Office.Interop.Excel.Series
        Dim z1 As Integer
        Dim sh As Worksheet
        Dim RH As Double

        xlCharts = worksheet1.ChartObjects
        myChart = xlCharts.Add(7, 550, 505, 300)
        chartPage = myChart.Chart
        chartRange = worksheet1.Range("F9", "F" & (_rowcount - 1))
        chartRange1 = worksheet1.Range("A9", "A" & (_rowcount - 1))
        'chartRange = worksheet1.Range("H8", "K" & (X - 1))
        'chartRange = worksheet1.Range("H8:K39", "A9:A39")
        ' chartPage.SetSourceData(Source:=chartRange)
        t_SerCol = chartPage.SeriesCollection
        t_Series = t_SerCol.NewSeries
        With t_Series
            .Name = "TOTAL GADA"
            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

        End With
        t_Series.Border.Color = RGB(255, 0, 0)
        chartPage.Refresh()
        chartPage.SeriesCollection(1).Interior.Color = RGB(255, 215, 0)
        chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlLineMarkers
        chartPage.HasTitle = True
        chartPage.ChartTitle.Text = ("Total Gada")


        xlCharts = worksheet1.ChartObjects
        myChart = xlCharts.Add(520, 550, 505, 300)
        chartPage = myChart.Chart
        chartRange = worksheet1.Range("H9", "H" & (_rowcount - 1))
        chartRange1 = worksheet1.Range("A9", "A" & (_rowcount - 1))
        'chartRange = worksheet1.Range("H8", "K" & (X - 1))
        'chartRange = worksheet1.Range("H8:K39", "A9:A39")
        ' chartPage.SetSourceData(Source:=chartRange)
        t_SerCol = chartPage.SeriesCollection
        t_Series = t_SerCol.NewSeries
        With t_Series
            .Name = "Gas Cost"
            t_Series.XValues = chartRange1 '("=Friction!R11C1:R17C1") 'Reference a valid RANGE
            t_Series.Values = chartRange '("=Friction!R11C2:R17C2") 'Reference a valid RANGE

        End With
        t_Series.Border.Color = RGB(255, 0, 0)
        chartPage.Refresh()
        chartPage.SeriesCollection(1).Interior.Color = RGB(255, 215, 0)
        chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlLineMarkers
        chartPage.HasTitle = True
        chartPage.ChartTitle.Text = ("Gas Cost")
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Create_File()
    End Sub

    Private Sub cboEmp1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboEmp1.InitializeLayout

    End Sub
End Class