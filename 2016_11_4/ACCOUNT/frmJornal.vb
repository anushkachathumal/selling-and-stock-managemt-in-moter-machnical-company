Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Configuration
Public Class frmJornal
    Dim _Acc_Type As String
    Dim _Comcode As String
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Code as [##],M01Acc_Name as [Acc Name] from M01Account_Master where m01ACC='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboMain
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 120
                .Rows.Band.Columns(1).Width = 260


            End With

            Sql = "select M01Acc_Name as [Acc Name] from M01Account_Master where m01ACC='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDis
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 320
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
            End If
        End Try
    End Function

    Private Sub frmJornal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Combo()
        _Comcode = ConfigurationManager.AppSettings("ComCode")
        txtCr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTo.Text = Today
        Call Load_Gride2()
        txtTot_Cr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTot_Dr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTot_Cr.ReadOnly = True
        txtTot_Dr.ReadOnly = True

    End Sub

    Private Sub cmdExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Journal
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Search_AccNo() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master where M01Acc_Code='" & cboMain.Text & "' and M01Com_Code='" & _Comcode & "' and m01ACC='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_AccNo = True
                cboDis.Text = M01.Tables(0).Rows(0)("M01Acc_Name")
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
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


    Function Search_AccName() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master where M01Acc_Name='" & cboDis.Text & "' and M01Com_Code='" & _Comcode & "' and m01ACC='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_AccName = True
                cboMain.Text = M01.Tables(0).Rows(0)("M01Acc_Code")
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
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

    Private Sub cboMain_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMain.AfterCloseUp
        Call Search_AccNo()
    End Sub

    Private Sub cboMain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMain.KeyUp
        If e.KeyCode = 13 Then
            Call Search_AccNo()
            txtDr.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_AccNo()
            txtDr.Focus()
        End If
    End Sub

    Private Sub cboDis_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDis.AfterCloseUp
        Call Search_AccName()
    End Sub


    Private Sub cboDis_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDis.KeyUp
        If e.KeyCode = 13 Then
            Call Search_AccName()
            txtDr.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_AccName()
            txtDr.Focus()
        End If
    End Sub

    Private Sub txtDr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDr.KeyUp
        Dim Value As Double

        If e.KeyCode = 13 Then
            If txtDr.Text <> "" Then
                If IsNumeric(txtDr.Text) Then
                    Value = txtDr.Text
                    txtDr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtDr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    txtCr.Focus()
                Else
                    MsgBox("Please enter the Correct Amount", MsgBoxStyle.Information, "Information ......")
                    Exit Sub
                End If
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If txtDr.Text <> "" Then
                If IsNumeric(txtDr.Text) Then
                    Value = txtDr.Text
                    txtDr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtDr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    txtCr.Focus()
                Else
                    MsgBox("Please enter the Correct Amount", MsgBoxStyle.Information, "Information ......")
                    Exit Sub
                End If
            End If
        End If
    End Sub

  
    Private Sub txtCr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCr.KeyUp
        Dim Value As Double

        If e.KeyCode = 13 Then
            If Search_AccNo() = True Then
            Else
                MsgBox("Please select the Account code", MsgBoxStyle.Information, "Information ......")
                cboMain.ToggleDropdown()
                Exit Sub
            End If
            If UltraGrid1.Rows.Count >= 3 Then
                MsgBox("Can't Add Account", MsgBoxStyle.Information, "Information .....")
                Exit Sub
            End If
            If txtDr.Text <> "" Then
                If IsNumeric(txtDr.Text) Then

                Else
                    MsgBox("Please enter the correct Amount", MsgBoxStyle.Information, "Information .....")
                    txtDr.Focus()
                    Exit Sub
                End If
            End If
            If txtCr.Text <> "" Then
                If IsNumeric(txtCr.Text) Then

                Else
                    MsgBox("Please enter the correct Amount", MsgBoxStyle.Information, "Information .....")
                    txtCr.Focus()
                    Exit Sub
                End If
            End If
            'If UltraGrid1.Rows.Count = 1 Then
            '    If IsNumeric(UltraGrid1.Rows(0).Cells(2).Value) Then
            '        If IsNumeric(txtDr.Text) Then
            '            MsgBox("Debit amount alrady exist", MsgBoxStyle.Information, "Information ......")
            '            txtDr.Focus()
            '            Exit Sub
            '        Else

            '        End If
            '    End If

            '    If UltraGrid1.Rows(0).Cells(3).Text <> "" Then
            '        If IsNumeric(txtCr.Text) Then
            '            MsgBox("Credit amount alrady exist", MsgBoxStyle.Information, "Information ......")
            '            txtCr.Focus()
            '            Exit Sub
            '        Else

            '        End If
            '    End If
            ' End If
            If UltraGrid1.Rows.Count = 2 Then
                If UltraGrid1.Rows(2).Cells(3).Text <> "" Then
                    If IsNumeric(txtDr.Text) Then
                        MsgBox("Debit amount alrady exist", MsgBoxStyle.Information, "Information ......")
                        txtDr.Focus()
                        Exit Sub
                    Else

                    End If
                End If

                If UltraGrid1.Rows(2).Cells(4).Text <> "" Then
                    If IsNumeric(txtCr.Text) Then
                        MsgBox("Debit amount alrady exist", MsgBoxStyle.Information, "Information ......")
                        txtCr.Focus()
                        Exit Sub
                    Else

                    End If
                End If
            End If

            If IsNumeric(txtDr.Text) And IsNumeric(txtCr.Text) Then
                MsgBox("Please enter the credit or debit amount", MsgBoxStyle.Information, "Information ......")
                Exit Sub
            Else
                If IsNumeric(txtCr.Text) Then
                    Value = txtCr.Text
                    txtCr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtCr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Acc Code") = cboMain.Text
                newRow("Acc Name") = cboDis.Text
                newRow("Debit") = txtDr.Text
                newRow("Credit") = txtCr.Text
                c_dataCustomer1.Rows.Add(newRow)
                If IsNumeric(txtCr.Text) Then
                    If IsNumeric(txtTot_Cr.Text) Then
                    Else
                        txtTot_Cr.Text = "0"
                    End If
                    Value = CDbl(txtCr.Text) + CDbl(txtTot_Cr.Text)
                    txtTot_Cr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtTot_Cr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If

                If IsNumeric(txtDr.Text) Then
                    If IsNumeric(txtTot_Dr.Text) Then
                    Else
                        txtTot_Dr.Text = "0"
                    End If
                    Value = CDbl(txtDr.Text) + CDbl(txtTot_Dr.Text)
                    txtTot_Dr.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtTot_Dr.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If


                txtCr.Text = ""
                txtDr.Text = ""
                cboDis.Text = ""
                cboMain.Text = ""
                cboMain.ToggleDropdown()


            End If
        End If
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
        Dim _RefNo As Integer
        Dim M01 As DataSet

        Dim i As Integer
        i = 0
        Try
            nvcFieldList1 = "select * from P01Parameter where P01Code='AC' and P01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = M01.Tables(0).Rows(0)("P01LastNo")
            End If

            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo +" & 1 & " where P01Code='AC' and P01Com_Code='" & _Comcode & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            For Each uRow As UltraGridRow In UltraGrid1.Rows
                nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Com_Code,T05User,T05Status)" & _
                                                               " values('" & _RefNo & "','JE','" & txtTo.Text & "', '" & Trim(UltraGrid1.Rows(i).Cells(0).Text) & "','Journal Entry','" & Trim(UltraGrid1.Rows(i).Cells(3).Text) & "','" & Trim(UltraGrid1.Rows(i).Cells(2).Text) & "','" & _Comcode & "','" & strDisname & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                i = i + 1
            Next

            transaction.Commit()
            Call Load_Gride2()
            txtCr.Text = ""
            txtDr.Text = ""
            cboDis.Text = ""
            cboMain.Text = ""
            txtTot_Cr.Text = ""
            txtTot_Dr.Text = ""
            connection.Close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Load_Gride2()
        txtCr.Text = ""
        txtDr.Text = ""
        cboDis.Text = ""
        cboMain.Text = ""
        txtTot_Cr.Text = ""
        txtTot_Dr.Text = ""
    End Sub

    Private Sub txtCr_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCr.ValueChanged

    End Sub
End Class