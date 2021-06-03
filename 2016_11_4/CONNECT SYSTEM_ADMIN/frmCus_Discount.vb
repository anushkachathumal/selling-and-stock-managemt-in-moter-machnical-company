Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmCus_Discount
    Dim _From As Date
    Dim _Comcode As String
    Dim _Disname As String
    Dim c_dataCustomer1 As DataTable

    Private Sub frmCus_Discount_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmDiscount_1.Close()
        frmDiscount_2.Close()
    End Sub
    
    Private Sub frmCus_Discount_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCash.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDays.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal_Discount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Gride2()
        Call Search_Compay_Discount()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub txtCash_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCash.KeyUp
        If e.KeyCode = 13 Then
            txtTotal_Discount.Focus()
        End If
    End Sub

    Private Sub txtCash_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCash.ValueChanged

    End Sub

    Private Sub txtDays_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDays.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtDays.Text) <> "" Then
                txtDiscount.Focus()
            End If
        End If
    End Sub

    Private Sub txtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDiscount.KeyUp
        Try
            If e.KeyCode = 13 Then
                'If txtCash.Text <> "" Then
                'Else
                '    txtCash.Text = ""
                'End If

                'If IsNumeric(txtCash.Text) Then
                'Else
                '    MsgBox("Please enter the correct Cash Discount", MsgBoxStyle.Information, "Information ......")
                '    Exit Sub
                'End If

                If txtDays.Text <> "" Then
                Else
                    MsgBox("Please enter the Chq Days", MsgBoxStyle.Information, "Information ........")
                    Exit Sub
                End If

                If IsNumeric(txtDays.Text) Then
                Else
                    MsgBox("Please enter the correct Chq Days", MsgBoxStyle.Information, "Information ......")
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
                newRow("Chq Days") = Trim(txtDays.Text)
                newRow("Discount(%)") = Trim(txtDiscount.Text)
                c_dataCustomer1.Rows.Add(newRow)
                txtDays.Text = ""
                txtDiscount.Text = ""
                txtDays.Focus()

            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableCus_Discount
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

          

            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
          
        End With
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Me.txtDays.Text = ""
        Me.txtCash.Text = ""
        Me.txtDiscount.Text = ""
        Call Load_Gride2()
    End Sub



    Private Sub txtTotal_Discount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotal_Discount.KeyUp
        If e.KeyCode = 13 Then
            If txtTotal_Discount.Text <> "" Then
                txtDays.Focus()
            End If
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If txtCash.Text <> "" Then
            If IsNumeric(txtCash.Text) Then
            Else
                MsgBox("Please enter the correct Cash Discount%", MsgBoxStyle.Information, "Information ...........")
                Exit Sub
            End If
        Else
            txtCash.Text = "0"
        End If

        If txtTotal_Discount.Text <> "" Then
            If IsNumeric(txtTotal_Discount.Text) Then
            Else
                MsgBox("Please enter the correct Total Invoice Discount%", MsgBoxStyle.Information, "Information ...........")
                Exit Sub
            End If
        Else
            txtTotal_Discount.Text = "0"

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

            nvcFieldList1 = "SELECT * FROM M01Setup_Company WHERE M01Ref_No='CE' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M01Setup_Company set M01Sales_Discount='" & Trim(txtCash.Text) & "',M01Add_Discount='" & Trim(txtTotal_Discount.Text) & "' where M01Ref_No='CE'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                '                                               " values('NEW_ROOTS','EDIT', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)
           
            End If

            nvcFieldList1 = "DELETE FROM M16Comapy_Discount"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            I = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                nvcFieldList1 = "Insert Into M16Comapy_Discount(M16Days,M16Discount)" & _
                                                                   " values('" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(1).Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                I = I + 1
            Next
            MsgBox("Customer Discount Update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
          
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Function Search_Compay_Discount()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        Try
            Sql = "select * from M01Setup_Company where  M01Ref_No='CE' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCash.Text = Trim(M01.Tables(0).Rows(0)("M01Sales_Discount"))
                txtTotal_Discount.Text = Trim(M01.Tables(0).Rows(0)("M01Add_Discount"))
            End If
            Call Load_Gride2()
            Sql = "select * from M16Comapy_Discount order by M16ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Chq Days") = Trim(M01.Tables(0).Rows(i)("M16Days"))
                newRow("Discount(%)") = Trim(M01.Tables(0).Rows(i)("M16Discount"))
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

    Private Sub CompanyDiscountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompanyDiscountToolStripMenuItem.Click
        frmDiscount_1.Close()
        frmDiscount_1.Show()
    End Sub

    Private Sub CustomerSpecificDiscountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerSpecificDiscountToolStripMenuItem.Click
        frmDiscount_2.Close()
        frmDiscount_2.Show()
        frmDiscount_2.cboCus_Code.ToggleDropdown()
    End Sub
End Class