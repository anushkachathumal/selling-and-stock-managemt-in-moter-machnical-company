Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmAvg_Price
    Dim c_dataCustomer1 As DataTable
    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Private Sub frmAvg_Price_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtYear1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYear2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYear3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYear4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYear5.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtYear1.Text = Year(Today)
        txtYear2.Text = Year(Today)
        txtYear3.Text = Year(Today)
        txtYear4.Text = Year(Today)
        txtYear5.Text = Year(Today)

        txtAmount1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtAmount2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtAmount3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtAmount4.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtAmount5.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Months()

    End Sub

    Function Load_Months()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M07Name as [##] from M07Month order by M07Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboMonth1
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
                ' .Rows.Band.Columns(1).Width = 180


            End With

            With cboMonth2
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
                ' .Rows.Band.Columns(1).Width = 180


            End With

            With cboMonth3
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
                ' .Rows.Band.Columns(1).Width = 180


            End With

            With cboMonth4
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
                ' .Rows.Band.Columns(1).Width = 180


            End With

            With cboMonth5
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
                ' .Rows.Band.Columns(1).Width = 180


            End With

            cboMonth1.Text = MonthName(Month(Today))
            cboMonth2.Text = MonthName(Month(Today))
            cboMonth3.Text = MonthName(Month(Today))
            cboMonth4.Text = MonthName(Month(Today))
            cboMonth5.Text = MonthName(Month(Today))

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

    Private Sub cboMonth1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboMonth1.InitializeLayout

    End Sub

    Private Sub cboMonth1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMonth1.KeyUp
        If e.KeyCode = 13 Then
            txtYear1.Focus()
        End If
    End Sub

    Private Sub txtYear1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYear1.KeyUp
        If e.KeyCode = 13 Then
            txtAmount1.Focus()
        End If
    End Sub

    Private Sub txtAmount1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount1.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtAmount1.Text <> "" Then
                If IsNumeric(txtAmount1.Text) Then
                    Value = txtAmount1.Text
                    txtAmount1.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                End If
            End If
            cboMonth2.ToggleDropdown()
        End If
    End Sub

    Private Sub cboMonth2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMonth2.KeyUp
        If e.KeyCode = 13 Then
            txtYear2.Focus()
        End If
    End Sub

    Private Sub txtYear2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYear2.KeyUp
        If e.KeyCode = 13 Then
            txtAmount2.Focus()
        End If
    End Sub

    Private Sub txtAmount2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount2.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtAmount2.Text <> "" Then
                If IsNumeric(txtAmount2.Text) Then
                    Value = txtAmount2.Text
                    txtAmount2.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                End If
            End If
            cboMonth3.ToggleDropdown()
        End If
    End Sub

    Private Sub cboMonth3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMonth3.KeyUp
        If e.KeyCode = 13 Then
            txtYear3.Focus()
        End If
    End Sub

    Private Sub txtYear3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYear3.KeyUp
        If e.KeyCode = 13 Then
            txtAmount3.Focus()
        End If
    End Sub

    Private Sub txtAmount3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount3.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtAmount3.Text <> "" Then
                If IsNumeric(txtAmount3.Text) Then
                    Value = txtAmount3.Text
                    txtAmount3.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                End If
            End If
            cboMonth4.ToggleDropdown()
        End If
    End Sub

    Private Sub cboMonth4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMonth4.KeyUp
        If e.KeyCode = 13 Then
            txtYear4.Focus()
        End If
    End Sub

    Private Sub txtYear4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYear4.KeyUp
        If e.KeyCode = 13 Then
            txtAmount4.Focus()
        End If
    End Sub

    Private Sub txtAmount4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount4.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtAmount4.Text <> "" Then
                If IsNumeric(txtAmount4.Text) Then
                    Value = txtAmount4.Text
                    txtAmount4.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                End If
            End If
            cboMonth5.ToggleDropdown()
        End If
    End Sub

    Private Sub cboMonth5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMonth5.KeyUp
        If e.KeyCode = 13 Then
            txtYear5.Focus()
        End If
    End Sub

    Private Sub txtYear5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtYear5.KeyUp
        If e.KeyCode = 13 Then
            txtAmount5.Focus()
        End If
    End Sub

    Private Sub txtAmount5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount5.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If txtAmount5.Text <> "" Then
                If IsNumeric(txtAmount5.Text) Then
                    Value = txtAmount5.Text
                    txtAmount5.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                End If
            End If
            cmdSave.Focus()
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
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
            If cboMonth1.Text <> "" And txtYear1.Text <> "" And txtAmount1.Text <> "" Then
                If IsNumeric(txtYear1.Text) Then

                Else
                    MsgBox("Please enter the correct Gas Cost year", MsgBoxStyle.Information, "Information ....")
                    txtYear1.Focus()
                    connection.Close()
                    Exit Sub
                End If

                If IsNumeric(txtAmount1.Text) Then

                Else
                    MsgBox("Please enter the correct Gas Cost", MsgBoxStyle.Information, "Information ....")
                    txtAmount1.Focus()
                    connection.Close()
                    Exit Sub
                End If

                nvcFieldList1 = "SELECT * FROM M04Gas_Avg_Cost WHERE M04Month='" & Trim(cboMonth1.Text) & "' AND M04year='" & Trim(txtYear1.Text) & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    nvcFieldList1 = "UPDATE M04Gas_Avg_Cost SET M05Amount='" & txtYear1.Text & "' WHERE  M04Month='" & Trim(cboMonth1.Text) & "' AND M04year='" & Trim(txtYear1.Text) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M04Gas_Avg_Cost(M04Month,M04year,M05Amount)" & _
                                                               " values('" & Trim(cboMonth1.Text) & "', '" & Trim(txtYear1.Text) & "','" & Trim(txtAmount1.Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            End If


            If cboMonth2.Text <> "" And txtYear2.Text <> "" And txtAmount2.Text <> "" Then
                If IsNumeric(txtYear2.Text) Then

                Else
                    MsgBox("Please enter the correct Leather Gloves Cost year", MsgBoxStyle.Information, "Information ....")
                    txtYear2.Focus()
                    connection.Close()
                    Exit Sub
                End If

                If IsNumeric(txtAmount2.Text) Then

                Else
                    MsgBox("Please enter the correct Leather Gloves Cost", MsgBoxStyle.Information, "Information ....")
                    txtAmount2.Focus()
                    connection.Close()
                    Exit Sub
                End If

                nvcFieldList1 = "SELECT * FROM M05Leather_Gloves_Avg_Cost WHERE M05Month='" & Trim(cboMonth2.Text) & "' AND M05year='" & Trim(txtYear2.Text) & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    nvcFieldList1 = "UPDATE M05Leather_Gloves_Avg_Cost SET M05Amount='" & txtYear2.Text & "' WHERE  M05Month='" & Trim(cboMonth2.Text) & "' AND M05year='" & Trim(txtYear2.Text) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M05Leather_Gloves_Avg_Cost(M05Month,M05year,M05Amount)" & _
                                                               " values('" & Trim(cboMonth2.Text) & "', '" & Trim(txtYear2.Text) & "','" & Trim(txtAmount2.Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            End If
            '============================================================================================
            If cboMonth3.Text <> "" And txtYear3.Text <> "" And txtAmount3.Text <> "" Then
                If IsNumeric(txtYear3.Text) Then

                Else
                    MsgBox("Please enter the correct N02 Cost year", MsgBoxStyle.Information, "Information ....")
                    txtYear3.Focus()
                    connection.Close()
                    Exit Sub
                End If

                If IsNumeric(txtAmount3.Text) Then

                Else
                    MsgBox("Please enter the correct N2 Cost", MsgBoxStyle.Information, "Information ....")
                    txtAmount3.Focus()
                    connection.Close()
                    Exit Sub
                End If

                nvcFieldList1 = "SELECT * FROM M06Nitragon_Avg_Cost WHERE M06Month='" & Trim(cboMonth3.Text) & "' AND M06year='" & Trim(txtYear3.Text) & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    nvcFieldList1 = "UPDATE M06Nitragon_Avg_Cost SET M06Amount='" & txtYear3.Text & "' WHERE  M06Month='" & Trim(cboMonth3.Text) & "' AND M06year='" & Trim(txtYear3.Text) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M06Nitragon_Avg_Cost(M06Month,M06year,M06Amount)" & _
                                                               " values('" & Trim(cboMonth3.Text) & "', '" & Trim(txtYear3.Text) & "','" & Trim(txtAmount3.Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            End If
            '======================================================================================
            If cboMonth4.Text <> "" And txtYear4.Text <> "" And txtAmount4.Text <> "" Then
                If IsNumeric(txtYear4.Text) Then

                Else
                    MsgBox("Please enter the correct Cotton Gloves Cost year", MsgBoxStyle.Information, "Information ....")
                    txtYear4.Focus()
                    connection.Close()
                    Exit Sub
                End If

                If IsNumeric(txtAmount4.Text) Then

                Else
                    MsgBox("Please enter the correct Cotton Gloves Cost", MsgBoxStyle.Information, "Information ....")
                    txtAmount4.Focus()
                    connection.Close()
                    Exit Sub
                End If

                nvcFieldList1 = "SELECT * FROM M07Cotton_Glove_Avg_Cost WHERE M07Month='" & Trim(cboMonth4.Text) & "' AND M07year='" & Trim(txtYear4.Text) & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    nvcFieldList1 = "UPDATE M07Cotton_Glove_Avg_Cost SET M07Amount='" & txtYear4.Text & "' WHERE  M07Month='" & Trim(cboMonth4.Text) & "' AND M07year='" & Trim(txtYear4.Text) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M07Cotton_Glove_Avg_Cost(M07Month,M07year,M07Amount)" & _
                                                               " values('" & Trim(cboMonth4.Text) & "', '" & Trim(txtYear4.Text) & "','" & Trim(txtAmount4.Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            End If
            '==================================================================================
            If cboMonth5.Text <> "" And txtYear5.Text <> "" And txtAmount5.Text <> "" Then
                If IsNumeric(txtYear5.Text) Then

                Else
                    MsgBox("Please enter the correct Cotton Gloves Cost year", MsgBoxStyle.Information, "Information ....")
                    txtYear5.Focus()
                    connection.Close()
                    Exit Sub
                End If

                If IsNumeric(txtAmount5.Text) Then

                Else
                    MsgBox("Please enter the correct Cotton Gloves Cost", MsgBoxStyle.Information, "Information ....")
                    txtAmount5.Focus()
                    connection.Close()
                    Exit Sub
                End If

                nvcFieldList1 = "SELECT * FROM M08Gada_Avg_Cost WHERE M08Month='" & Trim(cboMonth5.Text) & "' AND M08year='" & Trim(txtYear5.Text) & "'"
                t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(t01) Then
                    nvcFieldList1 = "UPDATE M08Gada_Avg_Cost SET M08Amount='" & txtYear5.Text & "' WHERE  M08Month='" & Trim(cboMonth5.Text) & "' AND M08year='" & Trim(txtYear5.Text) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M08Gada_Avg_Cost(M08Month,M08year,M08Amount)" & _
                                                               " values('" & Trim(cboMonth5.Text) & "', '" & Trim(txtYear5.Text) & "','" & Trim(txtAmount5.Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            End If


            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            Call Clear_Text()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Function Clear_Text()
        Me.cboMonth1.Text = ""
        Me.cboMonth2.Text = ""
        Me.cboMonth3.Text = ""
        Me.cboMonth4.Text = ""
        Me.txtYear1.Text = ""
        Me.txtYear2.Text = ""
        Me.txtYear3.Text = ""
        Me.txtYear4.Text = ""
        Me.txtYear5.Text = ""
        Me.cboMonth5.Text = ""
        Me.txtAmount1.Text = ""
        Me.txtAmount2.Text = ""
        Me.txtAmount3.Text = ""
        Me.txtAmount5.Text = ""
        cboMonth1.ToggleDropdown()

        txtYear1.Text = Year(Today)
        txtYear2.Text = Year(Today)
        txtYear3.Text = Year(Today)
        txtYear4.Text = Year(Today)
        txtYear5.Text = Year(Today)

    End Function

    Function Last_Month()
        Dim Value As Double
        Dim m01 As DataSet

        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        ' Dim M01 As DataSet

        'GSA
        Sql = "select * from M04Gas_Avg_Cost order by M04Month,M04year DESC "
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        If isValidDataset(M01) Then
            cboMonth1.Text = Trim(m01.Tables(0).Rows(0)("M04Month"))
            txtYear1.Text = Trim(m01.Tables(0).Rows(0)("M04year"))
            Value = Trim(m01.Tables(0).Rows(0)("M05Amount"))
            txtAmount1.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
        End If

    End Function

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        Call Last_Month()
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Clear_Text()
    End Sub
End Class