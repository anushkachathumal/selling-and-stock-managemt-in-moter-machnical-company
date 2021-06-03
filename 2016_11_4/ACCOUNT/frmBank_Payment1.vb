Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.DAL_frmWinner
Imports DBLotVbnet.common
Imports DBLotVbnet.MDIMain
Imports System.Net.NetworkInformation
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports System.Configuration
Imports DBLotVbnet.modlVar1
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmBank_Payment1
    Dim _Acc_Type As String
    Dim _Comcode As String
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim c_dataCustomer3 As DataTable
    Dim _Acc_Code As String

    Dim c_dataCustomer4 As DataTable
    Dim _From As Date
    Dim _To As Date

    Dim _Suppcode As String
    Dim _Search_Status As String

    Function Load_Data()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim _Qty As Integer
        Dim Value As Double
        Dim _Rowcount As Integer
        Dim _St As String
        Dim _Total As Double
        Dim _LastRow As Integer
        Dim _Paid As Double

        Try
            ' Me.Cursor = Cursors.WaitCursor
            Call Load_Gride2()
            Sql = "select *  from View_Account_Detailes where T05com_code='" & _Comcode & "' and T05Acc_type='BANK_PAY' order by T05Ref_No DESC"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            _Total = 0
            _Paid = 0
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                _Paid = 0
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Voucher No") = M01.Tables(0).Rows(i)("t05INVO")
                newRow("Date") = Month(M01.Tables(0).Rows(i)("T05Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T05Date")) & "/" & Year(M01.Tables(0).Rows(i)("T05Date"))
                'newRow("Com Invoice") = M01.Tables(0).Rows(i)("T01Invoice_No")

                newRow("Acc_Code") = M01.Tables(0).Rows(i)("T05aCC_NO")
                newRow("Acc_Name") = M01.Tables(0).Rows(i)("M01ACC_NAME")
                Value = M01.Tables(0).Rows(i)("T05CREDIT")
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                _Paid = Value
                newRow("Amount") = _St

              

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next

          
            con.close()
            ' Me.Cursor = Cursors.Arrow
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                Me.Cursor = Cursors.Arrow
                con.close()
            End If
        End Try
    End Function

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Bank_PAY
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 140
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False
          

            .DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(4).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
           

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Search_AccName() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master where M01Acc_Name='" & Trim(txtP_Name.Text) & "' and M01Status='A' and M01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Acc_Code = M01.Tables(0).Rows(0)("M01Acc_Code")
                Search_AccName = True
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

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        Try
            Dim Value As Double
            Dim _ST As String
            If e.KeyCode = 13 Then
                If Search_AccName() = True Then
                Else
                    MsgBox("Please enter the correct Pay Account", MsgBoxStyle.Information, "Information .....")
                    txtP_Name.ToggleDropdown()
                    Exit Sub
                End If

                If IsNumeric(txtAmount.Text) Then
                    If txtAmount.Text > 0 Then
                    Else
                        MsgBox("Please enter the correct Amount", MsgBoxStyle.Information, "Information .....")
                        txtAmount.Focus()
                        Exit Sub
                    End If
                Else
                    MsgBox("Please enter the correct Amount", MsgBoxStyle.Information, "Information .....")
                    txtAmount.Focus()
                    Exit Sub
                End If

                If txtAmount.Text <> "" Then
                Else
                    MsgBox("Please enter the correct Amount", MsgBoxStyle.Information, "Information .....")
                    txtAmount.Focus()
                    Exit Sub
                End If

                Value = CDbl(txtAmount.Text)
                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Dim newRow As DataRow = c_dataCustomer2.NewRow
                newRow("Pay To") = _Acc_Code
                newRow("Description") = txtP_Name.Text
                newRow("Amount") = _ST
                c_dataCustomer2.Rows.Add(newRow)

                If txtTotal.Text <> "" Then
                Else
                    txtTotal.Text = "0"
                End If
                Value = CDbl(txtTotal.Text) + CDbl(txtAmount.Text)
                txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtP_Name.Text = ""
                txtAmount.Text = ""
                txtP_Name.Text = ""
                txtP_Name.ToggleDropdown()
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub


    Private Sub frmBank_Payment1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride2()
        Call Load_Data()
        '  Call Load_Gride_Data()
        ' Call Load_Supplier()
        txtRef.ReadOnly = True
        txtRef.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtName.ReadOnly = True
        txtTotal.ReadOnly = True
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Voucher()
        Call Load_Gride3()
        '  txtPay_C.ReadOnly = True
        txtPay_Amount.ReadOnly = True
        txtPay_Dis.ReadOnly = True
        Call Load_Gride_4()

        Call Load_Bank()
        Call Load_Payee()
        txtDate.Text = Today

        txtVoucher_1.ReadOnly = True
        txtVoucher_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate_1.ReadOnly = True
        txtDate_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDue_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDue_1.ReadOnly = True

        txtB_Code1.ReadOnly = True
        txtB_Name1.ReadOnly = True
        txtChq_1.ReadOnly = True
        txtPay_1.ReadOnly = True
        txtnet_1.ReadOnly = True
        txtnet_1.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
    End Sub
    Private Sub txtBank_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBank.AfterCloseUp
        Call Search_Bank()
    End Sub




    Private Sub txtBank_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBank.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Bank()
            txtChq.Focus()
        End If
    End Sub

    Private Sub txtChq_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChq.KeyUp
        If e.KeyCode = 13 Then
            txtDue.Focus()
        End If
    End Sub

    Private Sub txtDue_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDue.KeyUp
        If e.KeyCode = 13 Then
            txtP_Name.ToggleDropdown()
        End If
    End Sub

    Private Sub txtP_Name_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtP_Name.KeyUp
        If e.KeyCode = 13 Then
            txtAmount.Focus()
        End If
    End Sub
    Function Load_Gride_4()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer4 = CustomerDataClass.MakeDataTable_BankPay
        UltraGrid3.DataSource = c_dataCustomer4
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 240
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).Width = 110
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function


    Function Load_Gride3()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTable_BankPay
        UltraGrid2.DataSource = c_dataCustomer2
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 260
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).Width = 110
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    

    Function Load_Bank()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Code as [##] from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Status='A' and M01Acc_Type='3'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtBank
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 130
                ' .Rows.Band.Columns(1).Width = 180
            End With

            With cboAccount
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 130
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Function Load_Voucher()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where P01Code='VOU' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    txtRef.Text = _Comcode & "/VU00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    txtRef.Text = _Comcode & "/VU0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtRef.Text = _Comcode & "/VU" & M01.Tables(0).Rows(0)("P01LastNo")

                End If
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_Bank() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Status='A' and M01Acc_Type='3' and M01Acc_Code='" & Trim(txtBank.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Bank = True
                txtName.Text = Trim(M01.Tables(0).Rows(0)("M01Acc_Name"))
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try


    End Function


    Function Load_Payee()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Name as [##] from M01Account_Master where M01Com_Code='" & _Comcode & "' and M01Status='A' and M01Acc_Type<>'SP'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtP_Name
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 290
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub NewVoucherToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewVoucherToolStripMenuItem.Click
        OPR3.Visible = True
    End Sub

    Private Sub txtP_Name_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles txtP_Name.InitializeLayout

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
        Dim A As String
        Dim M01 As DataSet
        Dim _STRDIS As String
        Dim A1 As String
        Dim B As New ReportDocument
        Dim _Voucher As String
        Dim _RefNo As Integer

        Try
            If Search_Bank() = True Then
            Else
                MsgBox("Please select the Bank Account", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                Exit Sub
            End If

            If UltraGrid2.Rows.Count > 0 Then
            Else
                MsgBox("Please select the Payee Account", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                Exit Sub
            End If

            

            If txtChq.Text <> "" Then
            Else
                MsgBox("Please enter the cheque no", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub
            End If

            If txtTotal.Text <> "" Then
            Else
                txtTotal.Text = "0"
            End If
            _STRDIS = Num2String(CDbl(txtTotal.Text))
            If IsDate(txtDue.Text) Then
            Else
                MsgBox("Please enter the correct due date", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub

            End If


            nvcFieldList1 = "select * from P01Parameter where P01Code='IN'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                _RefNo = MB51.Tables(0).Rows(0)("P01LastNo")
            End If

            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo+" & 1 & " WHERE P01Code='VOU' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo+" & 1 & " WHERE P01Code='IN' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)




            nvcFieldList1 = "Insert Into T06Acc_Tran_Main(T06Ref_No,T06Acc_No,T06Date,T06Amount,T06Remark,T06Pay_To,T06User,T06Com_Code,T06Description,T06Chq_No,T06Type,t06DUE,T06Status)" & _
                                                                " values('" & _RefNo & "', '" & txtBank.Text & "','" & txtDate.Text & "','" & CDbl(txtTotal.Text) & "','" & txtRemark.Text & "','-','" & strDisname & "','" & _Comcode & "','" & _STRDIS & "','" & txtChq.Text & "','BANK_PAY','" & txtDue.Text & "','A')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'TRANSACTION HEADER
            nvcFieldList1 = "Insert Into T15Bank_Transaction(T15Ref,T15Date,T15TR_Type,T15Bank_Code,T15Remark,T15Tr_Status,T15Cr,T15Dr,T15Status,T15Com_Code,T15User,T15Pay_No)" & _
                                                                      " values(" & _RefNo & ", '" & txtDate.Text & "','BANK_PAY','" & Trim(txtBank.Text) & "','" & Trim(txtRemark.Text) & "','NO','0','" & CDbl(txtTotal.Text) & "','A','" & _Comcode & "','" & strDisname & "','" & txtRef.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'CHQ HEADER
            nvcFieldList1 = "Insert Into T04Chq_Trans(T04Ref_No,T04Acc_Type,T04Chq_no,T04ACC_No,T04Amount,T04DOR,T04Status,T04Com_Code)" & _
                                                                     " values(" & _RefNo & ",'BANK_PAY','" & Trim(txtChq.Text) & "','" & Trim(txtBank.Text) & "','" & CDbl(txtTotal.Text) & "','" & txtDue.Text & "','A','" & _Comcode & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'TRANSACTION LOG
            nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpUser,tmpLog)" & _
                                                                     " values('BANK_PAY','SAVE','" & Trim(txtRef.Text) & "','" & Now & "','" & strDisname & "','" & _Comcode & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                'nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Credit,T05Debit,T05Invo,T05Name,T05Com_Code,T05User,T05Status)" & _
                '                                               " values('" & _RefNo & "', 'BANK_PAY','" & txtDate.Text & "','" & txtBank.Text & "','" & CDbl(UltraGrid1.Rows(i).Cells(2).Value) & "','0',' ','" & txtRef.Text & "','" & _Comcode & "','" & strDisname & "','A')"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Credit,T05Debit,T05Invo,T05Name,T05Com_Code,T05User,T05Status)" & _
                                                              " values('" & _RefNo & "', 'BANK_PAY','" & txtDate.Text & "','" & UltraGrid2.Rows(i).Cells(0).Value & "','" & CDbl(UltraGrid2.Rows(i).Cells(2).Value) & "','0','" & txtRef.Text & "','-','" & _Comcode & "','" & strDisname & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                i = i + 1
            Next
            transaction.Commit()
          
            connection.Close()

            A = MsgBox("Are you sure you want to Print Voucher", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Voucher ...")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\pay_Voucher1.rpt.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                ' B.SetParameterValue("Amount", txtTotal.Text)
                'B.SetParameterValue("Dis", txtName.Text)
                'B.SetParameterValue("Voucher", _Voucher)
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_Account_Detailes.T05Invo} =" & txtRef.Text & " and {View_Account_Detailes.T05Com_Code}='" & _Comcode & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            Call Load_Voucher()
            ' Me.txtSearch.Text = ""
            Me.txtBank.Text = ""
            Me.txtChq.Text = ""
            Me.txtTotal.Text = ""
            Me.txtAmount.Text = ""
            Me.txtP_Name.Text = ""
            Me.txtP_Name.Text = ""
            ' Me.txtPayee.Text = ""
            Me.txtRemark.Text = ""
            Me.txtName.Text = ""
            Call Load_Gride3()
            Call Load_Data()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        OPR3.Visible = False
    End Sub

    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        OPR3.Visible = False
    End Sub
End Class