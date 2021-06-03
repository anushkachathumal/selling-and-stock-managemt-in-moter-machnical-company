Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmCompany_Cnt
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Private Sub frmCompany_Cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtRemainder.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSales_Dis.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtAdd_Dis.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCom_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCom_2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCom_3.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Search_Records()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Me.txtAdd_Dis.Text = ""
        Me.txtAddress.Text = ""
        Me.txtCompany.Text = ""
        Me.txtSales_Dis.Text = ""
        Me.txtRemainder.Text = ""
        Me.txtCom_3.Text = ""
        Me.txtCom_2.Text = ""
        Me.txtCom_1.Text = ""
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select * from M01Setup_Company where M01Ref_No='CE' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCompany.Text = Trim(M01.Tables(0).Rows(0)("M01Com_Name"))
                txtAddress.Text = Trim(M01.Tables(0).Rows(0)("M01Address"))
                txtVAT.Text = Trim(M01.Tables(0).Rows(0)("M01VAT_No"))
                Value = Trim(M01.Tables(0).Rows(0)("M01Sales_Discount"))
                txtSales_Dis.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales_Dis.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("M01Add_Discount"))
                txtAdd_Dis.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAdd_Dis.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("M01Sales_Comm"))
                txtCom_1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCom_1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("M01Cash_Collection"))
                txtCom_2.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCom_2.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))


                Value = Trim(M01.Tables(0).Rows(0)("M02Chq_Collection"))
                txtCom_3.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCom_3.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtRemainder.Text = Trim(M01.Tables(0).Rows(0)("M01Chq_Reminder"))
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Trim(txtCompany.Text) <> "" Then
        Else
            MsgBox("Please enter the company name", MsgBoxStyle.Information, "Information ........")
            Exit Sub
        End If

        If Trim(txtAddress.Text) <> "" Then
        Else
            txtAddress.Text = "-"
        End If

        If Trim(txtVAT.Text) <> "" Then
        Else
            txtVAT.Text = "-"
        End If

        If IsNumeric(txtSales_Dis.Text) Then
        Else
            MsgBox("Please enter the correct Sales Discount(%)", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If IsNumeric(txtAdd_Dis.Text) Then
        Else
            MsgBox("Please enter the correct Additional Discount(%)", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If IsNumeric(txtRemainder.Text) Then
        Else
            MsgBox("Please enter the correct Chq Reminder Days", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If IsNumeric(txtCom_1.Text) Then
        Else
            MsgBox("Please enter the correct Sales Commision", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If IsNumeric(txtCom_2.Text) Then
        Else
            MsgBox("Please enter the correct Collection Commision-Cash", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If IsNumeric(txtCom_3.Text) Then
        Else
            MsgBox("Please enter the correct Collection Commision-Chque", MsgBoxStyle.Information, "Information ......")
            Exit Sub
        End If

        If txtSales_Dis.Text <> "" Then
        Else
            txtSales_Dis.Text = "0"

        End If

        If txtAdd_Dis.Text <> "" Then
        Else
            txtAdd_Dis.Text = "0"

        End If

        If txtRemainder.Text <> "" Then
        Else
            txtRemainder.Text = "0"

        End If

        If txtCom_1.Text <> "" Then
        Else
            txtCom_1.Text = "0"
        End If

        If txtCom_2.Text <> "" Then
        Else
            txtCom_2.Text = "0"
        End If

        If txtCom_3.Text <> "" Then
        Else
            txtCom_3.Text = "0"
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
        Try
          
            nvcFieldList1 = "SELECT * FROM M01Setup_Company WHERE M01Ref_No='CE' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "UPDATE M01Setup_Company SET M01Com_Name='" & Trim(txtCompany.Text) & "',M01Address='" & Trim(txtAddress.Text) & "',M01Sales_Discount='" & txtSales_Dis.Text & "',M01Add_Discount='" & txtAdd_Dis.Text & "',M01Chq_Reminder='" & txtRemainder.Text & "',M01Sales_Comm='" & txtCom_1.Text & "',M01Cash_Collection='" & txtCom_2.Text & "',M02Chq_Collection='" & txtCom_3.Text & "' WHERE M01Ref_No='CE'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "Insert Into M01Setup_Company(M01Ref_No,M01Com_Name,M01Address,M01VAT_No,M01Sales_Discount,M01Add_Discount,M01Chq_Reminder,M01Sales_Comm,M01Cash_Collection,M02Chq_Collection)" & _
                                                                  " values('CE','" & Trim(txtCompany.Text) & "','" & Trim(txtAddress.Text) & "','" & Trim(txtVAT.Text) & "','" & txtSales_Dis.Text & "','" & txtAdd_Dis.Text & "','" & txtRemainder.Text & "','" & txtCom_1.Text & "','" & txtCom_2.Text & "','" & txtCom_3.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
           
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try

    End Function
End Class