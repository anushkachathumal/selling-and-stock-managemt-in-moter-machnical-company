Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmDiscount_2
    Dim _From As Date
    Dim _Comcode As String
    Dim _Disname As String
    Dim c_dataCustomer1 As DataTable
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmDiscount_2_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmView_Customer.Close()
    End Sub

    Private Sub frmDiscount_2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCash.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_Discount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Customer()
        Call Load_Cus_Code()
        Call Load_Gride2()
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableCom_Discount
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 120
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False



            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit

        End With
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        txtAmount.Text = ""
        txtCash.Text = ""
        txtTotal_Discount.Text = ""
        txtDiscount.Text = ""
        Call Load_Gride2()
        cboCustomer.Text = ""
        cboCus_Code.Text = ""
        Call Load_Customer()
    End Sub

    Function Load_Customer()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M05Shop_Name  as [Shop Name] from M05New_Customer  where M05Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 265
                '  .Rows.Band.Columns(1).Width = 160


            End With
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Cus_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M05Cus_Code as [##] from M05New_Customer  where M05Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCus_Code
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 114
                '  .Rows.Band.Columns(1).Width = 160


            End With
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Cus_Code(ByVal strCode As String) As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M05New_Customer  where M05Status='A' and M05Cus_Code='" & strCode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboCustomer.Text = Trim(M01.Tables(0).Rows(0)("M05Shop_Name"))
            End If

            Call Search_Records()
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Cus_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M05New_Customer  where M05Status='A' and M05Shop_Name='" & Trim(cboCustomer.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboCus_Code.Text = Trim(M01.Tables(0).Rows(0)("M05Cus_Code"))
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
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
        Dim _ST As String
        Try
            Sql = "select * from M18Advance_Discount_Customer_Specific_Header where M18Cus_Code='" & Trim(cboCus_Code.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = Trim(M01.Tables(0).Rows(0)("M18Cash_Discount"))
                txtCash.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCash.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("M18Total_Discount"))
                txtTotal_Discount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtTotal_Discount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            End If
            Call Load_Gride2()
            Sql = "select * from M19Advance_Discount_Customer_Fluter where M19Cus_Code='" & Trim(cboCus_Code.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                Value = Trim(M01.Tables(0).Rows(i)("M19Amount"))
                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Invo.Amount") = _ST
                Value = Trim(M01.Tables(0).Rows(i)("M19Discount"))
                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount(%)") = _ST
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboCus_Code_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCus_Code.AfterCloseUp
        Call Search_Cus_Code(Trim(cboCus_Code.Text))
    End Sub

    Private Sub cboCus_Code_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCus_Code.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Cus_Code(Trim(cboCus_Code.Text))
            If Trim(cboCus_Code.Text) <> "" Then
                txtCash.Focus()
            Else
                cboCustomer.ToggleDropdown()
            End If
        ElseIf e.KeyCode = Keys.F1 Then
            strWindowName = Me.Name.ToString
            frmView_Customer.Close()
            frmView_Customer.Show()
        End If
    End Sub

    Private Sub cboCustomer_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustomer.AfterCloseUp
        Call Search_Cus_Name()
    End Sub

    Private Sub cboCustomer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCustomer.KeyUp
        If e.KeyCode = 13 Then
            txtCash.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            strWindowName = Me.Name.ToString
            frmView_Customer.Close()
            frmView_Customer.Show()
        End If
    End Sub

    Private Sub txtCash_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCash.KeyUp
        If e.KeyCode = 13 Then
            txtTotal_Discount.Focus()
        End If
    End Sub

    Private Sub txtTotal_Discount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotal_Discount.KeyUp
        If e.KeyCode = 13 Then
            txtAmount.Focus()
        End If
    End Sub

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtAmount.Text) Then
                Value = txtAmount.Text
                txtAmount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAmount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            txtDiscount.Focus()
        End If
    End Sub

    Private Sub txtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDiscount.KeyUp
        Try
            Dim Value As Double
            If e.KeyCode = 13 Then


                If txtAmount.Text <> "" Then
                Else
                    MsgBox("Please enter the Monthly Invoice Amount", MsgBoxStyle.Information, "Information ........")
                    Exit Sub
                End If

                If IsNumeric(txtAmount.Text) Then
                Else
                    MsgBox("Please enter the correct Monthly Invoice Amount", MsgBoxStyle.Information, "Information ......")
                    Exit Sub
                End If


                If txtDiscount.Text <> "" Then
                Else
                    MsgBox("Please enter the Discount", MsgBoxStyle.Information, "Information ........")
                    Exit Sub
                End If

                If IsNumeric(txtDiscount.Text) Then
                Else
                    MsgBox("Please enter the correct Discount", MsgBoxStyle.Information, "Information ......")
                    Exit Sub
                End If
                '==============================================================================================
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                Value = txtAmount.Text
                txtAmount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAmount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Invo.Amount") = Trim(txtAmount.Text)
                newRow("Discount(%)") = Trim(txtDiscount.Text)
                c_dataCustomer1.Rows.Add(newRow)
                txtAmount.Text = ""
                txtDiscount.Text = ""
                txtAmount.Focus()

            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
       
        If txtCash.Text <> "" Then
        Else
            txtCash.Text = "0"
        End If

        If IsNumeric(txtCash.Text) = True Then
        Else
            MsgBox("Please enter the correct cash amount", MsgBoxStyle.Information, "Information ........")
            Exit Sub
        End If

        If txtTotal_Discount.Text <> "" Then
        Else
            txtTotal_Discount.Text = "0"
        End If

        If IsNumeric(txtTotal_Discount.Text) = True Then
        Else
            MsgBox("Please enter the correct Total Discount ", MsgBoxStyle.Information, "Information ........")
            Exit Sub
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
        Dim I As Integer
        Try
            nvcFieldList1 = "select * from M05New_Customer where M05Cus_Code='" & Trim(cboCus_Code.Text) & "' and M05Status='A'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
            Else
                MsgBox("Please select the customer code", MsgBoxStyle.Information, "Information .....")
                connection.Close()
                Exit Function
            End If
            nvcFieldList1 = "SELECT * FROM M18Advance_Discount_Customer_Specific_Header WHERE M18Cus_Code='" & Trim(cboCus_Code.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M18Advance_Discount_Customer_Specific_Header set M18Cash_Discount='" & Trim(txtCash.Text) & "',M18Total_Discount='" & Trim(txtTotal_Discount.Text) & "' where M18Cus_Code='" & Trim(cboCus_Code.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                '                                               " values('NEW_ROOTS','EDIT', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "Insert Into M18Advance_Discount_Customer_Specific_Header(M18Cus_Code,M18Cash_Discount,M18Total_Discount)" & _
                                                                  " values('" & Trim(cboCus_Code.Text) & "','" & Trim(txtCash.Text) & "','" & Trim(txtTotal_Discount.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            nvcFieldList1 = "DELETE FROM M19Advance_Discount_Customer_Fluter where M19Cus_Code='" & Trim(cboCus_Code.Text) & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            I = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                nvcFieldList1 = "Insert Into M19Advance_Discount_Customer_Fluter(M19Cus_Code,M19Amount,M19Discount)" & _
                                                                   " values('" & Trim(cboCus_Code.Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(1).Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                I = I + 1
            Next
            MsgBox("Customer Discount Update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()

            connection.Close()
            Me.cboCus_Code.Text = ""
            Me.cboCustomer.Text = ""
            Me.txtCash.Text = ""
            Me.txtTotal_Discount.Text = ""
            Me.txtDiscount.Text = ""
            Me.txtAmount.Text = ""
            Call Load_Gride2()
            cboCus_Code.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim A As String
        A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information ........")
        If A = vbYes Then
            Call Delete_Data()
        End If
    End Sub

    Function Delete_Data()
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
        Dim I As Integer
        Try

            nvcFieldList1 = "DELETE FROM M18Advance_Discount_Customer_Specific_Header WHERE M18Cus_Code='" & Trim(cboCus_Code.Text) & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
           

            nvcFieldList1 = "DELETE FROM M19Advance_Discount_Customer_Fluter where M19Cus_Code='" & Trim(cboCus_Code.Text) & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
           
            MsgBox("Customer Discount Deleted successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()

            connection.Close()
            Me.cboCus_Code.Text = ""
            Me.cboCustomer.Text = ""
            Me.txtCash.Text = ""
            Me.txtTotal_Discount.Text = ""
            Me.txtDiscount.Text = ""
            Me.txtAmount.Text = ""
            Call Load_Gride2()
            cboCus_Code.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

End Class