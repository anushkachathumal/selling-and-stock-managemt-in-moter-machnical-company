Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmSub_Manifacter

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Private Sub frmSub_Manifacter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCode.ReadOnly = True
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtOBalance.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Parameter()
        Call Load_Combo_Name()
        Call Load_Grid1()
    End Sub

    Function Load_Parameter()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='SM'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01NO") <= 10 Then
                    txtCode.Text = "SM00" & M01.Tables(0).Rows(0)("P01NO")
                ElseIf M01.Tables(0).Rows(0)("P01NO") > 10 And M01.Tables(0).Rows(0)("P01NO") <= 100 Then
                    txtCode.Text = "SM0" & M01.Tables(0).Rows(0)("P01NO")
                Else
                    txtCode.Text = "SM" & M01.Tables(0).Rows(0)("P01NO")
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

    Function Clear_Text()
        Me.txtCode.Text = ""
        Me.cboName.Text = ""
        Me.txtContact.Text = ""
        Me.txtTp.Text = ""
        Me.txtAddress.Text = ""
        Me.txtAddress1.Text = ""
        Me.txtfax.Text = ""
        Me.txtEmail.Text = ""
        Me.txtOBalance.Text = ""
        OPR4.Visible = False
        Call Load_Parameter()
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Load_Parameter()
        cboName.ToggleDropdown()
    End Sub

    Function Load_Combo_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M18Name as [##] from M18Sub_Manufacture where M18Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 370
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

    Private Sub cboName_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboName.InitializeLayout

    End Sub

    Private Sub cboName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboName.KeyUp
        If e.KeyCode = 13 Then
            txtAddress.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
        ElseIf e.KeyCode = Keys.Escape Then
            OPR4.Visible = False
        End If
    End Sub

    Private Sub txtAddress_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyUp
        If e.KeyCode = 13 Then
            txtAddress1.Focus()
        End If
    End Sub

    Private Sub txtAddress1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress1.KeyUp
        If e.KeyCode = 13 Then
            txtTp.Focus()
        End If
    End Sub

    Private Sub txtTp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTp.KeyUp
        If e.KeyCode = 13 Then
            txtfax.Focus()
        End If
    End Sub

    Private Sub txtfax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtfax.KeyUp
        If e.KeyCode = 13 Then
            txtContact.Focus()
        End If
    End Sub

    Private Sub txtContact_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtContact.KeyUp
        If e.KeyCode = 13 Then
            txtEmail.Focus()
        End If
    End Sub

    Private Sub txtEmail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEmail.KeyUp
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

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim T01 As DataSet
        Try
            If cboName.Text <> "" Then
            Else
                MsgBox("Please enter the Company Name", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                cboName.ToggleDropdown()
                Exit Sub
            End If

            If txtAddress.Text <> "" Then
            Else
                txtAddress.Text = " "
            End If

            If txtAddress1.Text <> "" Then
            Else
                txtAddress1.Text = " "
            End If

            If txtTp.Text <> "" Then
            Else
                txtTp.Text = " "
            End If

            If txtfax.Text <> "" Then
            Else
                txtfax.Text = " "
            End If

            If txtContact.Text <> "" Then
            Else
                txtContact.Text = " "
            End If

            If txtEmail.Text <> "" Then
            Else
                txtEmail.Text = " "
            End If

            If txtOBalance.Text <> "" Then
            Else
                txtOBalance.Text = "0"
            End If

            If IsNumeric(txtOBalance.Text) Then
            Else
                MsgBox("Please enter the correct Balance Amount", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtOBalance.Focus()
                Exit Sub
            End If

            nvcFieldList1 = "SELECT * FROM M18Sub_Manufacture WHERE M18Code='" & txtCode.Text & "'"
            T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(T01) Then
                nvcFieldList1 = "UPDATE M01Account_Master SET M01Status='A',M01Acc_Name='" & cboName.Text & "',M01Address='" & txtAddress.Text & "',M01TP='" & txtTp.Text & "',M01OB_Chq='" & txtOBalance.Text & "' WHERE M01Acc_Code='" & txtCode.Text & "' AND M01Acc_Type='SM'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE M18Sub_Manufacture SET M18Name='" & cboName.Text & "',M18Address='" & txtAddress.Text & "',M18Address1='" & txtAddress1.Text & "',M18Tp='" & txtTp.Text & "',M18Fax='" & txtfax.Text & "',M18Email='" & txtEmail.Text & "',M18Contact='" & txtContact.Text & "',M18Balance='" & txtOBalance.Text & "',M18Status='A' WHERE M18Code='" & txtCode.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                Call Load_Parameter()
                nvcFieldList1 = "UPDATE P01PARAMETER SET P01NO=P01NO + " & 1 & " WHERE P01CODE='SM'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M18Sub_Manufacture(M18Code,M18Date,M18Name,M18Address,M18Address1,M18Tp,M18Fax,M18Email,M18Contact,M18Balance,M18Status,M18User)" & _
                                                                      " values('" & Trim(txtCode.Text) & "', '" & Today & "','" & Trim(cboName.Text) & "','" & txtAddress.Text & "','" & txtAddress1.Text & "','" & txtTp.Text & "','" & txtfax.Text & "','" & txtEmail.Text & "','" & txtContact.Text & "','" & txtOBalance.Text & "','A','" & strDisname & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M01Account_Master(M01Acc_Type,M01Acc_Code,M01Acc_Name,M01Address,M01TP,M01DOC,M01Status,M01User,M01year,M01ACC_OF,M01OB_Chq)" & _
                                                                      " values('SM', '" & txtCode.Text & "','" & Trim(cboName.Text) & "','" & txtAddress.Text & "','" & txtTp.Text & "','" & Today & "','A','" & strDisname & "','" & Year(Today) & "','MS','" & txtOBalance.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into T02Sub_Manufacture(T02Tr_Type,T02Acc_No,T02Date,T02Cr,T02Dr,T02Remark,T02Status,T02User)" & _
                                                                 " values('OB', '" & txtCode.Text & "','" & Today & "','" & txtOBalance.Text & "','0','Operning Balance','A','" & strDisname & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If

            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            Call Clear_Text()
            Call Load_Parameter()
            Call Load_Combo_Name()
            Call Load_Grid1()
            cboName.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Function Load_Grid1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M18Code as [Company Code],M18Name as [Company Name] from M18Sub_Manufacture where M18Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 110
                .DisplayLayout.Bands(0).Columns(0).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 270
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
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
        If e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
        ElseIf e.KeyCode = Keys.Escape Then
            OPR4.Visible = False
        End If
    End Sub

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _rowindex As Integer

        _rowindex = UltraGrid1.ActiveRow.Index
        txtCode.Text = UltraGrid1.Rows(_rowindex).Cells(0).Text
        Call Search_Records()
        OPR4.Visible = False
    End Sub
    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Try
            Sql = "select * from M18Sub_Manufacture where  M18Code='" & txtCode.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01

                    txtAddress.Text = .Tables(0).Rows(0)("M18Address")
                    cboName.Text = .Tables(0).Rows(0)("M18Name")
                    txtfax.Text = .Tables(0).Rows(0)("M18Fax")
                    txtEmail.Text = .Tables(0).Rows(0)("M18Email")
                    txtContact.Text = .Tables(0).Rows(0)("M18Contact")
                    txtAddress1.Text = .Tables(0).Rows(0)("M18Address1")
                    txtTp.Text = .Tables(0).Rows(0)("M18Tp")
                    Value = .Tables(0).Rows(0)("M18Balance")
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
                nvcFieldList1 = "update M18Sub_Manufacture set M18Status='I' where M18Code='" & txtCode.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update M01Account_Master set M01Status='I' where M01Acc_Code='" & txtCode.Text & "' and M01Acc_Type='SM'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update T02Sub_Manufacture set T02Status='I' where T02Acc_No='" & txtCode.Text & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Account Deleted successfully", MsgBoxStyle.Information, "Information .......")
                transaction.Commit()

            End If
            connection.Close()
            Call Clear_Text()
            Call Load_Grid1()
            Call Load_Combo_Name()
            cboName.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub
End Class