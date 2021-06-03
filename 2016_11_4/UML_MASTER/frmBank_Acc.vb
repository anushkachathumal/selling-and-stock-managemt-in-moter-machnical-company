Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmBank_Acc

    Private Sub frmBank_Acc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtOBalance.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Status()
        Call Load_Grid1()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Function Clear_Text()
        Me.txtCode.Text = ""
        Me.cboBranch.Text = ""
        Me.cboName.Text = ""
        Me.cboStatus.Text = ""
        Me.txtOBalance.Text = ""
        Me.txtTp.Text = ""
        OPR4.Visible = False
        txtCode.Focus()
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Clear_Text()
    End Sub
    Function Load_Grid1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M15Acc_Type as [Account Type],M15Acc_no as [Account No],M15Bank_Name as [Bank Name] from M15Bank_Account where M15Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 110
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 270
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            End With
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Function Load_Status()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M14Dis as [##] from M14Acc_Type "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboStatus
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

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = 13 Then
            If txtCode.Text <> "" Then
                Call Search_Records()
                cboStatus.ToggleDropdown()
            End If
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            Call Load_Grid1()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR4.Visible = False
        End If
    End Sub

    Private Sub cboStatus_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboStatus.KeyUp
        If e.KeyCode = 13 Then
            cboName.ToggleDropdown()
        End If
    End Sub

    Private Sub cboName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboName.KeyUp
        If e.KeyCode = 13 Then
            cboBranch.ToggleDropdown()
        End If
    End Sub

    Private Sub cboBranch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBranch.KeyUp
        If e.KeyCode = 13 Then
            txtTp.Focus()
        End If
    End Sub

    Private Sub txtTp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTp.KeyUp
        If e.KeyCode = 13 Then
            txtOBalance.Focus()
        End If
    End Sub

    Private Sub txtOBalance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOBalance.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtOBalance.Text) Then
                Value = txtOBalance.Text
                txtOBalance.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
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

        '   Dim P01Parameter As Integer
        Dim M01 As DataSet
        Try
            If txtCode.Text <> "" Then
            Else
                MsgBox("Please enter the Bank Account", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtCode.Focus()
                Exit Sub
            End If

            If cboName.Text <> "" Then
            Else
                MsgBox("Please enter the Bank Name", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                cboName.ToggleDropdown()
                Exit Sub
            End If

            If cboBranch.Text <> "" Then
            Else
                MsgBox("Please enter the Branch Name", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                cboBranch.ToggleDropdown()
                Exit Sub
            End If

            If txtTp.Text <> "" Then
            Else
                txtTp.Text = " "
            End If

            If txtOBalance.Text <> "" Then
            Else
                txtOBalance.Text = "0"
            End If

            If IsNumeric(txtOBalance.Text) Then

            Else
                MsgBox("Please enter the correct Operning Balance", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtOBalance.Focus()
                Exit Sub
            End If

            nvcFieldList1 = "SELECT * FROM M15Bank_Account WHERE M15Acc_no='" & txtCode.Text & "' "
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                nvcFieldList1 = "UPDATE M15Bank_Account SET M15Status='A',M15Bank_Name='" & cboName.Text & "',M15Branch='" & cboBranch.Text & "',M15Tp='" & txtTp.Text & "',M15O_Balance='" & txtOBalance.Text & "',M15Acc_Type='" & cboStatus.Text & "' WHERE M15Acc_no='" & txtCode.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE M01Account_Master SET M01Status='A',M01Acc_Name='" & cboName.Text & "',M01Address='" & cboBranch.Text & "',M01TP='" & txtTp.Text & "',M01OB_Chq='" & txtOBalance.Text & "' WHERE M01Acc_Code='" & txtCode.Text & "' AND M01Acc_Type='BN'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else

                nvcFieldList1 = "Insert Into M15Bank_Account(M15Acc_no,M15Date,M15Bank_Name,M15Branch,M15Tp,M15O_Balance,M15Status,M15Acc_Type)" & _
                                                                    " values('" & Trim(txtCode.Text) & "', '" & Today & "','" & Trim(cboName.Text) & "','" & cboBranch.Text & "','" & txtTp.Text & "','" & txtOBalance.Text & "','A','" & cboStatus.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into M01Account_Master(M01Acc_Type,M01Acc_Code,M01Acc_Name,M01Address,M01TP,M01DOC,M01Status,M01User,M01year,M01ACC_OF,M01OB_Chq)" & _
                                                                  " values('BN', '" & txtCode.Text & "','" & Trim(cboName.Text) & "','" & cboBranch.Text & "','" & txtTp.Text & "','" & Today & "','A','" & strDisname & "','" & Year(Today) & "','MS','" & txtOBalance.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into T01Bank_Transaction(T01Tr_Type,T01Acc_Code,T01Date,T01Cr,T01Dr,T01Chq_No,T01Naretion,T01Status)" & _
                                                                 " values('OB', '" & txtCode.Text & "','" & Today & "','" & txtOBalance.Text & "','0',' ','Operning Bank Balance','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            Call Clear_Text()
            Call Load_Grid1()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub


    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Try
            Sql = "select * from M15Bank_Account where  M15Acc_no='" & txtCode.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01

                    cboStatus.Text = .Tables(0).Rows(0)("M15Acc_type")
                    cboName.Text = .Tables(0).Rows(0)("M15Bank_Name")
                    ' cboType.Text = .Tables(0).Rows(0)("M09Type")
                    cboBranch.Text = .Tables(0).Rows(0)("M15Branch")
                    txtTp.Text = .Tables(0).Rows(0)("M15Tp")
                    Value = .Tables(0).Rows(0)("M15O_Balance")
                    txtOBalance.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                End With
            End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub txtCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.ValueChanged

    End Sub

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _rowcount As Integer

        _rowcount = UltraGrid1.ActiveRow.Index
        txtCode.Text = UltraGrid1.Rows(_rowcount).Cells(1).Text
        Call Search_Records()
        OPR4.Visible = False
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim A As String
        Try
            A = MsgBox("Are you sure you want to delete this Account", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete Records .....")
            If A = vbYes Then
                nvcFieldList1 = "update M15Bank_Account set M15Status='I' where M15Acc_no='" & txtCode.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update M01Account_Master set M01Status='I' where M01Acc_Code='" & txtCode.Text & "' and M01Acc_Type='BN'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update T01Bank_Transaction set T01Status='I' where T01Acc_Code='" & txtCode.Text & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Account Deleted successfully", MsgBoxStyle.Information, "Information .......")
                transaction.Commit()

            End If
            connection.Close()
            Call Clear_Text()
            Call Load_Grid1()
            txtCode.Focus()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub
End Class