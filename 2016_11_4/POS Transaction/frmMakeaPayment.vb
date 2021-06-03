Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine

Public Class frmMakeaPayment
    Dim _Acctype As String
    Dim _Comcode As String
    Dim c_dataCustomer1 As DataTable
    Dim _Bank_Code As String

    Function Load_Status()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select M06Name as [##] from M06Status"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboStatus
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 65
            End With
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Account_no()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select M01Acc_Code as [Account No] from M01Account_Master where  M01Acc_Type<>'BN'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboAccno
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 130
            End With

            SQL = "select M01Acc_Code as [Account No] from M01Account_Master where   M01Acc_Type='BN'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboBank
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 130
            End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTablePayment
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 150
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            '.DisplayLayout.Bands(1).Columns(1).Width = 130
            '.DisplayLayout.Bands(1).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(2).Columns(1).Width = 90
            '.DisplayLayout.Bands(2).Columns(1).AutoEdit = False
            '.DisplayLayout.Bands(3).Columns(1).Width = 90
            '.DisplayLayout.Bands(3).Columns(1).AutoEdit = False

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Account_name()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select M01Acc_Name as [Account Name] from M01Account_Master "
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboName
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 230
            End With
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Function Load_Bank_Name()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select M01Acc_Name as [Bank Name] from M01Account_Master where  M01Acc_Type='BN'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            With cboBank
                .DataSource = T01
                .Rows.Band.Columns(0).Width = 230
            End With
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmMakeaPayment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("ComCode")
        Call Load_Status()
        Call Load_Account_no()
        Call Load_Account_name()
        ' cboBank.ReadOnly = True
        '  Call Load_Bank_Name()

        txtCash.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtPayment.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtPayment.ReadOnly = True

        txtDate.Text = Today
        txtDue.Text = Today
        Call Load_Gride2()

    End Sub

    Private Sub cmdExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Search_Records() As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select * from M01Account_Master where  M01Acc_Code='" & Trim(cboAccno.Text) & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                cboStatus.Text = T01.Tables(0).Rows(0)("M01Status")
                _Acctype = Trim(T01.Tables(0).Rows(0)("M01Acc_Type"))
                cboName.Text = T01.Tables(0).Rows(0)("M01Acc_Name")
                txtAddress.Text = T01.Tables(0).Rows(0)("M01Address")
                txtTp.Text = T01.Tables(0).Rows(0)("M01TP")
                Search_Records = True

                Call Load_Gridewith_Detail()
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Function Search_Bank_Code() As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select * from M01Account_Master where  M01Acc_Name='" & Trim(cboBank.Text) & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                _Bank_Code = T01.Tables(0).Rows(0)("M01Acc_Code")
             
                Search_Bank_Code = True


            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Records1() As Boolean
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Try
            SQL = "select * from M01Account_Master where  M01Acc_Name='" & Trim(cboName.Text) & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(T01) Then
                cboStatus.Text = T01.Tables(0).Rows(0)("M01Status")
                _Acctype = Trim(T01.Tables(0).Rows(0)("M01Acc_Type"))
                cboAccno.Text = T01.Tables(0).Rows(0)("M01Acc_Code")
                txtAddress.Text = T01.Tables(0).Rows(0)("M01Address")
                txtTp.Text = T01.Tables(0).Rows(0)("M01TP")
                Search_Records1 = True

                Call Load_Gridewith_Detail()
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboAccno_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccno.AfterCloseUp
        Call Search_Records()
    End Sub

    Private Sub cboAccno_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboAccno.InitializeLayout

    End Sub

    Private Sub cboAccno_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboAccno.KeyUp
        If e.KeyCode = Keys.Enter Then
            Call Search_Records()
            txtRef.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_Records()
            txtRef.Focus()
        End If
    End Sub

    Private Sub txtRef_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRef.KeyUp
        If e.KeyCode = Keys.Enter Then
            txtName.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtName.Focus()
        End If
    End Sub

    Private Sub txtRef_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRef.ValueChanged

    End Sub

    Private Sub txtName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyUp
        If e.KeyCode = Keys.Enter Then
            txtDetails.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDetails.Focus()
        End If
    End Sub

    Private Sub txtName_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.ValueChanged

    End Sub

    Private Sub txtDetails_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDetails.KeyUp
        If e.KeyCode = Keys.Enter Then
            txtCash.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtCash.Focus()
        End If
    End Sub

    Private Sub txtDetails_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDetails.ValueChanged

    End Sub

    Private Sub txtCash_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCash.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtCash.Text) Then
                Value = txtCash.Text
                txtCash.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCash.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                Call Calcution()
                cboBank.ToggleDropdown()
            Else
                Call Calcution()
                cboBank.ToggleDropdown()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If IsNumeric(txtCash.Text) Then
                Value = txtCash.Text
                txtCash.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCash.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                Call Calcution()
                cboBank.ToggleDropdown()
            Else
                Call Calcution()
                cboBank.ToggleDropdown()
            End If
            End If
    End Sub

    Function Calcution()
        Dim Value As Double
        If Trim(txtCash.Text) <> "" Then
            If IsNumeric(txtCash.Text) Then
            Else
                Exit Function
            End If
        Else
            txtCash.Text = "0"
        End If

        If Trim(txtAmount.Text) <> "" Then
            If IsNumeric(txtAmount.Text) Then
            Else
                Exit Function
            End If
        Else
            txtAmount.Text = "0"
        End If

        If Trim(txtCash.Text) <> "" And Trim(txtAmount.Text) <> "" Then
            Value = CDbl(txtCash.Text) + CDbl(txtAmount.Text)
            txtPayment.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtPayment.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Function

    Function Load_Gridewith_Detail()
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim i As Integer
        Dim _ST As String
        Dim Value As Double

        Try
            Call Load_Gride2()
            SQL = "select * from T05Acc_Trans where  T05Acc_No='" & Trim(cboAccno.Text) & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            i = 0
            For Each DTRow2 As DataRow In T01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Date") = T01.Tables(0).Rows(i)("T05Date")
                newRow("Description") = T01.Tables(0).Rows(i)("T05Remark")
                Value = T01.Tables(0).Rows(i)("T05Credit")
                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Payment") = _ST
                Value = T01.Tables(0).Rows(i)("T05Debit")
                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                newRow("W.Drawals") = _ST
                

                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub txtChq_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChq.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtChq.Text) <> "" Then
                txtDue.Focus()
            Else
                cmdAdd.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If Trim(txtChq.Text) <> "" Then
                txtDue.Focus()
            Else
                cmdAdd.Focus()
            End If
        End If
    End Sub

    Private Sub txtChq_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChq.ValueChanged

    End Sub

    Private Sub txtDue_BeforeDropDown(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDue.BeforeDropDown

    End Sub

    Private Sub txtDue_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDue.KeyUp
        If e.KeyCode = 13 Then
            ' Call Calcution()
            txtAmount.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            'Call Calcution()
            txtAmount.Focus()

        End If
    End Sub

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtAmount.Text) Then
                Value = txtAmount.Text
                txtAmount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAmount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                Call Calcution()
                cmdAdd.Focus()
            Else
                Call Calcution()
                cmdAdd.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If IsNumeric(txtAmount.Text) Then
                Value = txtAmount.Text
                txtAmount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtAmount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                Call Calcution()
                cmdAdd.Focus()
            Else
                Call Calcution()
                cmdAdd.Focus()
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
        Dim MB51 As DataSet
        Dim i As Integer
        Dim result1 As String
        Dim M01 As DataSet
        Dim Value As Double
        Dim _PROFIT As Double
        Dim _PROFITCOM As Double
        Dim T01 As DataSet
        Dim _RefNo As Integer
        Dim B As New ReportDocument
        Try
            Call Calcution()
            If _Acctype = "BN" Then
            Else
                If Search_Records() = True Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Account ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        cboAccno.ToggleDropdown()
                        connection.Close()
                        Exit Sub
                    End If
                End If
            End If
            '-----------------------------------------------------------------
            If Trim(txtCash.Text) <> "" Then
            Else
                If IsNumeric(txtCash.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Cash Amount ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtCash.Focus()
                        connection.Close()
                        Exit Sub
                    End If
                End If
            End If

            If txtAmount.Text = "0" Then
                txtAmount.Text = ""
            End If
            If Trim(txtAmount.Text) <> "" Then
                If Trim(txtChq.Text) <> "" Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Chque No ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtChq.Focus()
                        connection.Close()
                        Exit Sub
                    End If
                End If

                If IsNumeric(txtAmount.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Chque Amount ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtChq.Focus()
                        connection.Close()
                        Exit Sub
                    End If
                End If

                'nvcFieldList1 = "select * from M01Account_Master where M01Acc_Code='" & Trim(cboBank.Text) & "' and M01Acc_Type='BN' and M01Com_Code='" & _Comcode & "'"
                'T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(T01) Then
                'Else
                '    result1 = MessageBox.Show("Please enter the Correct Account No ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                '    If result1 = Windows.Forms.DialogResult.OK Then
                '        cboBank.ToggleDropdown()
                '        Exit Sub
                '    End If
                'End If
                If cboBank.Text <> "" Then
                    'If Search_Bank_Code() = True Then
                    'Else

                    '    result1 = MessageBox.Show("Please enter the Correct Bank Name ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    '    If result1 = Windows.Forms.DialogResult.OK Then
                    '        cboBank.ToggleDropdown()
                    '        connection.Close()
                    '        Exit Sub
                    '    End If
                    'End If
                Else
                    result1 = MessageBox.Show("Please enter the Correct Bank Name ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        cboBank.ToggleDropdown()
                        connection.Close()
                        Exit Sub
                    End If
                End If
            End If

            ' Call Calcution()

            nvcFieldList1 = "select * from P01Parameter where  P01Code='IN'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = Trim(M01.Tables(0).Rows(0)("P01LastNo"))
            End If

            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='IN' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Dim _Remark As String

            _Remark = ""
            If txtRef.Text <> "" And txtDetails.Text <> "" Then
                _Remark = "Make a Payment for Invoice No- " & txtRef.Text & " - " & txtDetails.Text
            ElseIf txtRef.Text <> "" Then
                _Remark = "Make a Payment for Invoice No- " & txtRef.Text
            ElseIf txtDetails.Text <> "" Then
                _Remark = "Make a Payment ( " & txtDetails.Text & " )"
            Else
                _Remark = "Make a Payment on Voucher No " & _RefNo
            End If

            nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Invo,T05Name,T05Com_Code,T05User,T05Status)" & _
                                                                             " values('" & _RefNo & "', '" & _Acctype & "','" & txtDate.Text & "','" & Trim(cboAccno.Text) & "','" & _Remark & "','0','" & txtPayment.Text & "','" & txtRef.Text & "','" & txtName.Text & "','" & _Comcode & "','" & strDisname & "','MP')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'Need to Delete KHills Distributors system
            'If CDbl(txtCash.Text) > 0 Then
            '    nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Invo,T05Name,T05Com_Code,T05User,T05Status)" & _
            '                                                                                 " values('" & _RefNo & "', 'SP','" & txtDate.Text & "','" & _Comcode & "','" & _Remark & "','" & txtCash.Text & "','0','" & txtRef.Text & "','" & txtName.Text & "','" & _Comcode & "','" & strDisname & "','MP')"
            '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'End If
            If _Acctype = "BN" Then

            Else
                If Val(txtAmount.Text) > 0 And txtChq.Text <> "" And cboBank.Text <> "" Then
                    nvcFieldList1 = "SELECT * FROM M01Account_Master WHERE M01Acc_Code='" & cboBank.Text & "'  AND M01Acc_Type='BN'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                        nvcFieldList1 = "Insert Into T04Chq_Trans(T04Ref_No,T04Acc_Type,T04Chq_no,T04Amount,T04DOR,T04ACC_No,T04Status,T04Com_Code)" & _
                                                                                    " values('" & _RefNo & "', '" & _Acctype & "','" & txtChq.Text & "','" & Trim(txtAmount.Text) & "','" & txtDue.Text & "','" & cboBank.Text & "','MP','" & _Comcode & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Invo,T05Name,T05Com_Code,T05User,T05Status)" & _
                                                                                               " values('" & _RefNo & "', 'BN','" & txtDate.Text & "','" & cboBank.Text & "','" & _Remark & "','" & txtAmount.Text & "','0','" & txtRef.Text & "','" & txtName.Text & "','" & _Comcode & "','" & strDisname & "','MP')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Else
                        MsgBox("Please enter the correct Bank Account", MsgBoxStyle.Information, "Technova .........")
                        cboBank.ToggleDropdown()
                        Exit Sub
                    End If
                End If
            End If

            Dim _Balance As Double
            Dim _PayVoucher As String
            _Balance = txtPayment.Text

            nvcFieldList1 = "select (T07InvoiceAmount-T07Paid_Amount) as Amount,T07RefNo,T07Paid_Voucher from T07Supplier_Payment where  T07Status='N' order by T07RefNo"
            T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            i = 0
            For Each DTRow2 As DataRow In T01.Tables(0).Rows


                If T01.Tables(0).Rows(i)("T07Paid_Voucher") = "0" Then
                    _PayVoucher = _RefNo
                Else
                    _PayVoucher = T01.Tables(0).Rows(i)("T07Paid_Voucher") & "," & _RefNo
                End If

                If T01.Tables(0).Rows(i)("AMOUNT") <= _Balance Then
                    nvcFieldList1 = "update T07Supplier_Payment set T07Paid_Amount=T07Paid_Amount+" & T01.Tables(0).Rows(i)("Amount") & ",T07Paid_Voucher='" & _PayVoucher & "',T07Status='Y' where T07RefNo='" & T01.Tables(0).Rows(i)("T07RefNo") & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    _Balance = _Balance - T01.Tables(0).Rows(i)("AMOUNT")
                Else
                    nvcFieldList1 = "update T07Supplier_Payment set T07Paid_Amount=T07Paid_Amount+" & _Balance & ",T07Paid_Voucher='" & _PayVoucher & "' where T07RefNo='" & T01.Tables(0).Rows(i)("T07RefNo") & "' "
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    Exit For
                End If
                i = i + 1
            Next
            transaction.Commit()

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            'transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""

            result1 = MsgBox("Do you want to Print Payment Voucher", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Payment Voucher ......")
            If result1 = vbYes Then
                'Dim B As New ReportDocument
                Dim A As String


                A = ConfigurationManager.AppSettings("ReportPath") + "\PayVoucher.rpt"
                B.Load(A.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                B.SetParameterValue("Chq", txtChq.Value)
                B.SetParameterValue("Due", txtDue.Value)
                B.SetParameterValue("Amount", txtAmount.Value)

                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T05Acc_Trans.T05Ref_No}=" & _RefNo & "  and {T05Acc_Trans.T05Com_Code} ='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If
            common.ClearAll(OPR2, OPR3, OPR4, OPR5)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            OPR2.Enabled = True
            OPR3.Enabled = True
            OPR4.Enabled = True
            OPR5.Enabled = True

            cmdAdd.Enabled = True
            cboAccno.ToggleDropdown()
            ' cmdSave.Enabled = False

            Call Load_Gride2()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub txtCash_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCash.ValueChanged

    End Sub

    Private Sub cboBank_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboBank.InitializeLayout

    End Sub

    Private Sub cboBank_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBank.KeyUp
        If e.KeyCode = 13 Then
            txtChq.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtChq.Focus()

        End If
    End Sub

    Private Sub UltraGroupBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPR4.Click

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR2, OPR3, OPR4, OPR5)
        ' OPR2.Enabled = False
        'OPR1.Enabled = False
        OPR2.Enabled = True
        OPR3.Enabled = True
        OPR4.Enabled = True
        OPR5.Enabled = True

        cmdAdd.Enabled = True
        cboAccno.ToggleDropdown()
        ' cmdSave.Enabled = False

        Call Load_Gride2()
    End Sub

    Private Sub cboName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboName.AfterCloseUp
        Call Search_Records1()
    End Sub

    Private Sub cboName_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboName.InitializeLayout

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim A As String
        Dim B As New ReportDocument
        Dim Sql As String
        Dim M01 As DataSet
        Dim M02 As DataSet

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Sql = "select * from T05Acc_Trans where T05Status='MP' and T05Com_Code='" & _Comcode & "' order by T05Ref_No DESC "
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        If isValidDataset(M01) Then
            A = ConfigurationManager.AppSettings("ReportPath") + "\PayVoucher.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            Sql = "select * from T04Chq_Trans where T04Ref_No='" & M01.Tables(0).Rows(0)("T05Ref_No") & "' and T04Com_Code='" & _Comcode & "'"
            M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M02) Then
                B.SetParameterValue("Chq", M02.Tables(0).Rows(0)("T04Chq_no"))
            Else
                B.SetParameterValue("Chq", " ")
                B.SetParameterValue("Due", " ")
                B.SetParameterValue("Amount", " ")
            End If
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{T05Acc_Trans.T05Ref_No}=" & M01.Tables(0).Rows(0)("T05Ref_No") & "  and {T05Acc_Trans.T05Com_Code} ='" & _Comcode & "'"
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()
        End If
    End Sub
End Class