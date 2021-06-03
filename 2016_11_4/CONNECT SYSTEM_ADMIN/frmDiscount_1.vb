Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmDiscount_1
    Dim _From As Date
    Dim _Comcode As String
    Dim _Disname As String
    Dim c_dataCustomer1 As DataTable

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmDiscount_1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Gride2()
        Call Search_Compay_Discount()
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

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        On Error Resume Next
        Dim Value As Double
        If e.KeyCode = 13 Then
            If Trim(txtAmount.Text) <> "" Then
                If IsNumeric(txtAmount.Text) Then
                    Value = txtAmount.Text
                    txtAmount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtAmount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                End If
            End If
            txtDiscount.Focus()
        End If
    End Sub

 

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        txtAmount.Text = ""
        txtDiscount.Text = ""
    End Sub

    Private Sub txtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDiscount.KeyUp
        Try
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
        If UltraGrid2.Rows.Count >= 1 Then
        Else
            MsgBox("Please enter the Discounts", MsgBoxStyle.Information, "Information ........")
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



            nvcFieldList1 = "DELETE FROM M17Company_Advance_Discount"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            I = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                nvcFieldList1 = "Insert Into M17Company_Advance_Discount(M17Amount,M17Discount)" & _
                                                                   " values('" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(1).Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                I = I + 1
            Next
            MsgBox("Advance Discount Update successfully", MsgBoxStyle.Information, "Information ..........")
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
        Dim Value As Double
        Dim _St As String

        Try
         
            Call Load_Gride2()
            Sql = "select * from M17Company_Advance_Discount order by M17ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                Value = Trim(M01.Tables(0).Rows(i)("M17Amount"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Invo.Amount") = _St
                Value = Trim(M01.Tables(0).Rows(i)("M17Discount"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Discount(%)") = _St
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

End Class