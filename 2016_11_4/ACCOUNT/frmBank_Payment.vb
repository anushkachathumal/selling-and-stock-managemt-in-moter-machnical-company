
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


Public Class frmBank_Payment
    Dim _Acc_Type As String
    Dim _Comcode As String
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub


    Private Sub frmBank_Payment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("ComCode")

        txtDate.Text = Today
        txtRef.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSearch.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtAmount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Right

        txtRef.ReadOnly = True
        txtTotal.ReadOnly = True

        Call Load_Combo()
        Call Load_Gride2()
        Call Load_PayTo()
        Call Load_Parameter()
        Call Load_PayName()

    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_BankPay
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 340
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(2).Width = 110
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select m01Acc_code as [Acc Code],m01Acc_Name as [Acc Name]  from View_Main_Acc where left(m01Acc_Code,6)='130206' and m01Acc_Code not in ('130206','130206004') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtBank
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 120
                .Rows.Band.Columns(1).Width = 260


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

  
    Function Load_PayTo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select m01Acc_code as [Acc Code],m01Acc_Name as [Acc Name]  from View_Main_Acc  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPay
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 120
                .Rows.Band.Columns(1).Width = 260


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


    Function Load_PayName()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Name as [##]  from View_Main_Acc  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With txtP_Name
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 270
                ' .Rows.Band.Columns(1).Width = 260


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

    Private Sub txtBank_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBank.AfterCloseUp
        Call Search_BankName()
    End Sub


    Private Sub txtBank_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBank.KeyUp
        If e.KeyCode = 13 Then
            Call Search_BankName()
            txtChq.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_BankName()
            txtChq.Focus()
        End If
    End Sub

    Private Sub cmdExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Parameter()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where P01Code='VU' and P01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtRef.Text = M01.Tables(0).Rows(0)("P01LastNo")
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


    Function Search_BankName() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master where M01Acc_Code='" & Trim(txtBank.Text) & "' and  left(m01Acc_Code,6)='130206' and m01Acc_Code not in ('130206','130206004')"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtName.Text = M01.Tables(0).Rows(0)("M01Acc_Name")
                Search_BankName = True
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

    Function Search_AccName() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M01Account_Master where M01Acc_Code='" & Trim(cboPay.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtP_Name.Text = M01.Tables(0).Rows(0)("M01Acc_Name")
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


    Private Sub txtBank_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBank.LostFocus
        Call Search_BankName()
    End Sub

    Private Sub txtChq_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChq.KeyUp
        If e.KeyCode = 13 Then
            cboPay.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboPay.ToggleDropdown()
        End If
    End Sub



    Private Sub cboPay_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPay.KeyUp
        If e.KeyCode = 13 Then
            Call Search_AccName()
            If cboPay.Text <> "" Then
                txtAmount.Focus()
            Else
                txtPayee.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_AccName()
            txtAmount.Focus()
        End If
    End Sub

    Private Sub txtAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyUp
        Try
            Dim Value As Double
            Dim _ST As String
            If e.KeyCode = 13 Then
                If Search_AccName() = True Then
                Else
                    MsgBox("Please enter the correct Pay Account", MsgBoxStyle.Information, "Information .....")
                    cboPay.ToggleDropdown()
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

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Pay To") = cboPay.Text
                newRow("Description") = txtP_Name.Text
                newRow("Amount") = _ST
                c_dataCustomer1.Rows.Add(newRow)

                If txtTotal.Text <> "" Then
                Else
                    txtTotal.Text = "0"
                End If
                Value = CDbl(txtTotal.Text) + CDbl(txtAmount.Text)
                txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                cboPay.Text = ""
                txtAmount.Text = ""
                txtP_Name.Text = ""
                cboPay.ToggleDropdown()
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub



    Private Sub UltraGrid1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.AfterRowsDeleted
        On Error Resume Next
        Dim I As Integer
        Dim Value As Double
        I = 0
        For Each uRow As UltraGridRow In UltraGrid1.Rows
            Value = Value + CDbl((UltraGrid1.Rows(I).Cells(2).Value))
            I = I + 1
        Next
        txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))


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

        Try
            If Search_BankName() = True Then
            Else
                MsgBox("Please select the Bank Account", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                Exit Sub
            End If

            If UltraGrid1.Rows.Count > 0 Then
            Else
                MsgBox("Please select the Payee Account", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                Exit Sub
            End If

            If txtPayee.Text <> "" Then
            Else
                MsgBox("Please select the Payee name", MsgBoxStyle.Information, "Information ......")
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

            nvcFieldList1 = "SELECT * FROM T06Acc_Tran_Main WHERE T06Ref_No=" & txtRef.Text & ""
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then

            Else

                nvcFieldList1 = "UPDATE P01Parameter SET P01LastNo=P01LastNo+" & 1 & " WHERE P01Code='VU' AND P01Com_Code='" & _Comcode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T06Acc_Tran_Main(T06Ref_No,T06Acc_No,T06Date,T06Amount,T06Remark,T06Pay_To,T06User,T06Com_Code,T06Description,T06Chq_No,T06Type)" & _
                                                                    " values('" & txtRef.Text & "', '" & txtBank.Text & "','" & txtDate.Text & "','" & CDbl(txtTotal.Text) & "','" & txtRemark.Text & "','" & txtPayee.Text & "','" & strDisname & "','" & _Comcode & "','" & _STRDIS & "','" & txtChq.Text & "','AP')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                i = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Credit,T05Debit,T05Invo,T05Name,T05Com_Code,T05User,T05Status)" & _
                                                                   " values('" & txtRef.Text & "', 'AP','" & txtDate.Text & "','" & txtBank.Text & "','" & CDbl(UltraGrid1.Rows(i).Cells(2).Value) & "','0',' ',' ','" & _Comcode & "','" & strDisname & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    nvcFieldList1 = "Insert Into T05Acc_Trans(T05Ref_No,T05Acc_Type,T05Date,T05Acc_No,T05Credit,T05Debit,T05Invo,T05Name,T05Com_Code,T05User,T05Status)" & _
                                                                  " values('" & txtRef.Text & "', 'AP','" & txtDate.Text & "','" & UltraGrid1.Rows(i).Cells(0).Value & "','0','" & CDbl(UltraGrid1.Rows(i).Cells(2).Value) & "',' ',' ','" & _Comcode & "','" & strDisname & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    i = i + 1
                Next
                transaction.Commit()
            End If

            connection.Close()
            If Len(txtRef.Text) >= 100 Then
                _Voucher = "0" & txtRef.Text
            ElseIf Len(txtRef.Text) >= 10 Then
                _Voucher = "00" & txtRef.Text
            Else
                _Voucher = "000" & txtRef.Text
            End If
            A = MsgBox("Are you sure you want to Print Voucher", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Voucher ...")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\Pay_Voucher.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "Admin@123")
                B.SetParameterValue("Amount", txtTotal.Text)
                B.SetParameterValue("Dis", txtName.Text)
                B.SetParameterValue("Voucher", _Voucher)
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{T06Acc_Tran_Main.T06Ref_No} =" & txtRef.Text & " and {T05Acc_Trans.T05Acc_No} <>  '" & txtBank.Text & "'"
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            Call Load_Parameter()
            Me.txtSearch.Text = ""
            Me.txtBank.Text = ""
            Me.txtChq.Text = ""
            Me.txtTotal.Text = ""
            Me.txtAmount.Text = ""
            Me.cboPay.Text = ""
            Me.txtP_Name.Text = ""
            Me.txtPayee.Text = ""
            Me.txtRemark.Text = ""
            Me.txtName.Text = ""
            Call Load_Gride2()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub txtPayee_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPayee.KeyUp
        If e.KeyCode = 13 Then
            txtRemark.Focus()
        End If
    End Sub



    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Load_Parameter()
        Me.txtSearch.Text = ""
        Me.txtBank.Text = ""
        Me.txtChq.Text = ""
        Me.txtTotal.Text = ""
        Me.txtAmount.Text = ""
        Me.cboPay.Text = ""
        Me.txtP_Name.Text = ""
        Me.txtPayee.Text = ""
        Me.txtRemark.Text = ""
        Me.txtName.Text = ""
        Call Load_Gride2()
    End Sub



    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Dim Value As Double
        Dim i As Integer
        Dim _ST As String

        Try
            Sql = "select * from T06Acc_Tran_Main inner join M01Account_Master on M01Acc_Code=T06Acc_No where T06Ref_No='" & txtSearch.Text & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                txtBank.Text = dsUser.Tables(0).Rows(0)("M01Acc_Code")
                txtName.Text = dsUser.Tables(0).Rows(0)("M01Acc_Name")
                txtChq.Text = dsUser.Tables(0).Rows(0)("T06Chq_No")
                txtDate.Text = dsUser.Tables(0).Rows(0)("M01DOC")
                Value = dsUser.Tables(0).Rows(0)("T06Amount")
                txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

            End If

            Call Load_Gride2()
            Sql = "select * from T05Acc_Trans inner join M01Account_Master on T05Acc_No=M01Acc_Code where T05Ref_No='" & txtSearch.Text & "' and T05Acc_Type='AP' and T05Debit>0"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In dsUser.Tables(0).Rows
                Value = dsUser.Tables(0).Rows(i)("T05Debit")
                _ST = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _ST = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Pay To") = dsUser.Tables(0).Rows(i)("M01Acc_Code")
                newRow("Description") = dsUser.Tables(0).Rows(i)("M01Acc_Name")
                newRow("Amount") = _ST
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim _Voucher As String
        Dim A1 As String
        Dim B As New reportdocument


        Try
            If Len(txtSearch.Text) >= 100 Then
                _Voucher = "0" & txtSearch.Text
            ElseIf Len(txtSearch.Text) >= 10 Then
                _Voucher = "00" & txtSearch.Text
            Else
                _Voucher = "000" & txtSearch.Text
            End If

            A1 = ConfigurationManager.AppSettings("ReportPath") + "\Pay_Voucher.rpt"
            B.Load(A1.ToString)
            B.SetDatabaseLogon("sa", "Admin@123")
            B.SetParameterValue("Amount", txtTotal.Text)
            B.SetParameterValue("Dis", txtName.Text)
            B.SetParameterValue("Voucher", _Voucher)
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{T06Acc_Tran_Main.T06Ref_No} =" & txtSearch.Text & " and {T05Acc_Trans.T05Acc_No} <>  '" & txtBank.Text & "'"
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub txtSearch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Records()
        End If
    End Sub

    Function Search_AccName1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select * from M01Account_Master where M01Acc_Name='" & txtP_Name.Text & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                cboPay.Text = dsUser.Tables(0).Rows(0)("M01Acc_Code")
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtP_Name_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP_Name.AfterCloseUp
        Call Search_AccName1()
    End Sub

    Private Sub txtP_Name_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles txtP_Name.InitializeLayout

    End Sub
End Class