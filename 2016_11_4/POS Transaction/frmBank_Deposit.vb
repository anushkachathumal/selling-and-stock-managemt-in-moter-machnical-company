
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmBank_Deposit
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable

    Dim _EmpNo As String
    Dim _CusType As String

    Dim _Category As String
    Dim _Comcode As String
    Dim _Supcode As String
    Dim _Cuscode As String
    Dim _RootNo As String
    Dim _ACCCODE As String


    Function Load_Gride3()
        'Dim CustomerDataClass As New DAL_InterLocation()
        'c_dataCustomer3 = CustomerDataClass.MakeDataTableChq
        'UltraGrid3.DataSource = c_dataCustomer3
        'With UltraGrid3
        '    .DisplayLayout.Bands(0).Columns(0).Width = 70
        '    .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
        '    .DisplayLayout.Bands(0).Columns(1).Width = 110
        '    .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

        '    .DisplayLayout.Bands(0).Columns(2).Width = 110
        '    .DisplayLayout.Bands(0).Columns(3).Width = 70
        '    ' .DisplayLayout.Bands(0).Columns(4).Width = 70
        '    ' .DisplayLayout.Bands(0).Columns(4).Width = 70
        '    '.DisplayLayout.Bands(0).Columns(6).Width = 60

        '    .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '    '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '    '.DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '    '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        '    '  .DisplayLayout.Bands(0).Columns(6).Width = 90
        '    ' .DisplayLayout.Bands(0).Columns(7).Width = 90

        '    ' .DisplayLayout.Bands(0).Columns(3).Width = 300
        '    '.DisplayLayout.Bands(0).Columns(4).Width = 300
        'End With
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmBank_Deposit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("ComCode")
        Call Load_Gride3()
        txtDate.Text = Today
        Call Load_Combo()
        Call Load_Bank()
        Call Load_Parameter()
        txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtChq_Tot.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtCash.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSearch.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtEntry.ReadOnly = True
        txtDue.Text = Today

    End Sub

    Function Load_Bank()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M42Name as [##] from M42Bank "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboBank
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

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Name as [##] from M01Account_Master_New where M01Com_Code='" & _Comcode & "' AND M01Acc_Type='3'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCode
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

    Function Load_Parameter()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where P01Code='BD' and P01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtEntry.Text = Trim(M01.Tables(0).Rows(0)("P01LastNo"))
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
        Call Load_Parameter()
        txtAmount.Text = ""
        txtChq.Text = ""
        txtChq_Tot.Text = ""
        txtCash.Text = ""
        txtPay.Text = ""
        txtRemark.Text = ""
        Call Load_Gride3()

        cboBank.Text = ""
        cboCode.Text = ""

    End Sub

    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCode.KeyUp
        If e.KeyCode = 13 Then
            txtPay.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtPay.Focus()
        End If
    End Sub

    Private Sub txtPay_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPay.KeyUp
        If e.KeyCode = 13 Then
            txtCash.Focus()
        End If
    End Sub

    Private Sub txtCash_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCash.KeyUp
        Dim VALUE As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtCash.Text) Then
                VALUE = txtCash.Text
                txtCash.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCash.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
            End If
            txtRemark.Focus()
        End If
    End Sub

    Private Sub cboBank_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBank.KeyUp
        If e.KeyCode = 13 Then
            txtChq.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtChq.Focus()
        End If
    End Sub

    Private Sub txtChq_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChq.KeyUp
        If e.KeyCode = 13 Then
            txtAmount.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtAmount.Focus()
        End If
    End Sub

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        Dim Value As Double
        Try
            If e.KeyCode = 13 Then
                If cboBank.Text <> "" Then
                Else
                    MsgBox("Please select the Bank Name", MsgBoxStyle.Information, "Information ......")
                    cboBank.ToggleDropdown()
                    Exit Sub

                End If

                If txtChq.Text <> "" Then
                Else
                    MsgBox("Please select the Chq Number", MsgBoxStyle.Information, "Information ......")
                    txtChq.Focus()
                    Exit Sub

                End If

                If cboBank.Text <> "" Then
                    If IsNumeric(txtAmount.Text) Then
                    Else
                        MsgBox("Please select the Amount", MsgBoxStyle.Information, "Information ......")
                        txtAmount.Focus()
                        Exit Sub
                    End If
                Else
                    MsgBox("Please select the Amount", MsgBoxStyle.Information, "Information ......")
                    txtAmount.Focus()
                    Exit Sub

                End If
                Dim _St As String

                Value = txtAmount.Text
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Dim newRow As DataRow = c_dataCustomer3.NewRow
                newRow("Bank Name") = Trim(cboBank.Text)
                newRow("Due Date") = txtDue.Text
                newRow("Chque No") = txtChq.Text
                newRow("Amount") = _St


                c_dataCustomer3.Rows.Add(newRow)

                If txtChq_Tot.Text <> "" Then
                Else
                    txtChq_Tot.Text = "0"
                End If

                Value = CDbl(txtChq_Tot.Text) + CDbl(txtAmount.Text)
                txtChq_Tot.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtChq_Tot.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                'Value = CDbl(txtBill.Text) - CDbl(txtChq_Tot.Text)
                'txtBalance.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtBalance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtChq.Text = ""
                txtAmount.Text = ""
                cboBank.ToggleDropdown()

            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Function Search_Account() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master_New where M01Com_Code='" & _Comcode & "' AND M01Acc_Type='3' and M01Acc_Name='" & cboCode.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Account = True
                _ACCCODE = M01.Tables(0).Rows(0)("M01Acc_Code")
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
        Dim M01 As DataSet
        Dim _REMARK As String
        Dim _COUNTNO As Integer
        Dim I As Integer

        Try
            If txtRemark.Text <> "" Then

            Else
                txtRemark.Text = " "
            End If

            If Search_Account() = True Then
            Else
                MsgBox("Please select the Account", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub
            End If

            If txtCash.Text <> "" Then
                If IsNumeric(txtCash.Text) Then
                Else
                    MsgBox("Please enter the correct cash amount", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    Exit Sub
                End If
            End If

            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo + " & 1 & " WHERE P01Code='BD' AND P01Com_Code='" & _Comcode & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            _REMARK = ""

            If txtCash.Text <> "" Then

                If txtPay.Text <> "" Then
                    _REMARK = "CASH DEPOSIT (" & txtPay.Text & ")"
                Else
                    _REMARK = "CASH DEPOSIT"
                End If
                nvcFieldList1 = "Insert Into tmpCash_Book(tmpDate,tmpType,tmpRef_No,tmpDiscription,tmpChqNo,tmpCrdit,tmpDebit,tmpCom_Code)" & _
                                                                                   " values('" & txtDate.Text & "', 'BD','" & txtEntry.Text & "','" & _REMARK & "','','0','" & txtCash.Text & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "SELECT * FROM P01Parameter WHERE P01Code='CON' AND P01Com_Code='" & _Comcode & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    _COUNTNO = M01.Tables(0).Rows(0)("P01LastNo")
                End If
                nvcFieldList1 = "Insert Into tmpBank_Rec(tmpRef_No,tmpDate,tmpMonth,tmpYear,tmpAcc_No,tmpDis,tmpCredit,tmpDebit,tmpCR_Status,tmpDR_Status,tmpCom_Code,tmpRemark,tmpStatus,tmpChq_No,tmpBank,tmpCount_No,tmpSlipNo)" & _
                                                                                   " values('" & txtEntry.Text & "', '" & txtDate.Text & "','" & Month(txtDate.Text) & "','" & Year(txtDate.Text) & "','" & _ACCCODE & "','" & txtPay.Text & "','" & txtCash.Text & "','0','N','N','" & _Comcode & "','" & _REMARK & "','DEPOSIT',' ',' ','" & _COUNTNO & "','" & txtPay.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo + " & 1 & " WHERE P01Code='CON' AND P01Com_Code='" & _Comcode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Com_Code,T05User,T05Status)" & _
                                                                 " values('" & txtEntry.Text & "','BD','" & txtDate.Text & "', '" & _ACCCODE & "','" & _REMARK & "','" & txtCash.Text & "','0','" & _Comcode & "','" & strDisname & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If

            If UltraGrid3.Rows.Count > 0 Then
                I = 0
                For Each uRow As UltraGridRow In UltraGrid3.Rows
                    _REMARK = "CHQ DEPOSIT (" & UltraGrid3.Rows(I).Cells(0).Value & ")"

                    nvcFieldList1 = "SELECT * FROM P01Parameter WHERE P01Code='CON' AND P01Com_Code='" & _Comcode & "'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                        _COUNTNO = M01.Tables(0).Rows(0)("P01LastNo")
                    End If

                    nvcFieldList1 = "Insert Into tmpCash_Book(tmpDate,tmpType,tmpRef_No,tmpDiscription,tmpChqNo,tmpCrdit,tmpDebit,tmpCom_Code)" & _
                                                                                   " values('" & txtDate.Text & "', 'BD','" & txtEntry.Text & "','" & _REMARK & "','" & UltraGrid3.Rows(I).Cells(0).Value & "','0','" & UltraGrid3.Rows(I).Cells(4).Value & "','" & _Comcode & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into tmpBank_Rec(tmpRef_No,tmpDate,tmpMonth,tmpYear,tmpAcc_No,tmpDis,tmpCredit,tmpDebit,tmpCR_Status,tmpDR_Status,tmpCom_Code,tmpRemark,tmpStatus,tmpChq_No,tmpBank,tmpCount_No,tmpDue_Date,tmpSlipNo)" & _
                                                                                   " values('" & txtEntry.Text & "', '" & txtDate.Text & "','" & Month(UltraGrid3.Rows(I).Cells(3).Value) & "','" & Year(UltraGrid3.Rows(I).Cells(3).Value) & "','" & _ACCCODE & "','" & txtPay.Text & "','" & UltraGrid3.Rows(I).Cells(4).Value & "','0','N','N','" & _Comcode & "','" & txtRemark.Text & "','DEPOSIT','" & UltraGrid3.Rows(I).Cells(0).Value & "','" & UltraGrid3.Rows(I).Cells(1).Value & "','" & _COUNTNO & "','" & UltraGrid3.Rows(I).Cells(2).Value & "','" & txtPay.Text & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo + " & 1 & " WHERE P01Code='CON' AND P01Com_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Com_Code,T05User,T05Status)" & _
                                                                 " values('" & txtEntry.Text & "','BD','" & txtDate.Text & "', '" & _ACCCODE & "','" & txtRemark.Text & "','" & UltraGrid3.Rows(I).Cells(4).Value & "','0','" & _Comcode & "','" & strDisname & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                    nvcFieldList1 = "UPDATE T04Chq_Trans SET T04Status='D' WHERE T04Chq_no='" & UltraGrid3.Rows(I).Cells(0).Value & "' AND T04Bank_Name='" & Trim(UltraGrid3.Rows(I).Cells(1).Value) & "' AND T04Status='ON' AND T04Com_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    I = I + 1
                Next
            End If
            MsgBox("Record update sucessfully", MsgBoxStyle.Information, "Information ......")
            transaction.Commit()
            connection.Close()
            Call Load_Parameter()
            txtAmount.Text = ""
            txtChq.Text = ""
            txtChq_Tot.Text = ""
            txtCash.Text = ""
            txtPay.Text = ""
            txtRemark.Text = ""
            Call Load_Gride3()
            cboBank.Text = ""
            cboCode.Text = ""
            cboCode.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try

    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim value As Double
        Dim I As Integer

        Try
            Sql = "select * from tmpBank_Rec inner join M01Account_Master_New on M01Acc_Code=tmpAcc_No where tmpRef_No='" & txtSearch.Text & "' and tmpCom_Code='" & _Comcode & "' and tmpChq_No=''"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                value = M01.Tables(0).Rows(0)("tmpCredit")
                txtCash.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCash.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                cboCode.Text = M01.Tables(0).Rows(0)("M01Acc_Name")
                txtDate.Text = M01.Tables(0).Rows(0)("tmpDate")
                txtRemark.Text = M01.Tables(0).Rows(0)("tmpRemark")
                txtPay.Text = M01.Tables(0).Rows(0)("tmpSlipNo")

                cmdAdd.Enabled = False
                cmdDelete.Enabled = True
            End If

            Call Load_Gride3()
            Sql = "SELECT * FROM tmpBank_Rec WHERE tmpRef_No='" & txtSearch.Text & "' AND tmpCom_Code='" & _Comcode & "' AND tmpChq_No<>''"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            I = 0
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim _St As String
                txtDate.Text = M01.Tables(0).Rows(I)("tmpDate")
                txtPay.Text = M01.Tables(0).Rows(I)("tmpSlipNo")
                txtRemark.Text = M01.Tables(0).Rows(I)("tmpDis")

                value = M01.Tables(0).Rows(I)("tmpCredit")
                _St = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                Dim newRow As DataRow = c_dataCustomer3.NewRow
                newRow("Bank Name") = M01.Tables(0).Rows(I)("tmpBank")
                newRow("Due Date") = Microsoft.VisualBasic.Day(M01.Tables(0).Rows(I)("tmpDue_Date")) & "/" & Month(M01.Tables(0).Rows(I)("tmpDue_Date")) & "/" & Year(M01.Tables(0).Rows(I)("tmpDue_Date"))
                newRow("Chque No") = M01.Tables(0).Rows(I)("tmpChq_No")
                newRow("Amount") = _St


                c_dataCustomer3.Rows.Add(newRow)

                If txtChq_Tot.Text <> "" Then
                Else
                    txtChq_Tot.Text = "0"
                End If

                value = CDbl(txtChq_Tot.Text) + value
                txtChq_Tot.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtChq_Tot.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))

                I = I + 1
            Next
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

    Private Sub txtSearch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Records()
        End If
    End Sub

    Private Sub txtSearch_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.ValueChanged

    End Sub
End Class