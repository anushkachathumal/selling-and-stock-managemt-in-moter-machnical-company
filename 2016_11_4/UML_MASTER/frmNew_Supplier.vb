Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmNew_Supplier
    Dim c_dataCustomer1 As DataTable
    Dim _ComCode As String

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Function Claer_Text()
        Me.txtCode.Text = ""
        Me.txtContact.Text = ""
        Me.txtAdd1.Text = ""
        Me.txtAddress.Text = ""
        Me.txtVAT.Text = ""
        Me.txtFax.Text = ""
        Me.txtTp.Text = ""
        Me.cboName.Text = ""
        Me.cboStatus.Text = ""
        Me.cboType.Text = ""
        Call Load_Grid()
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Claer_Text()
        Call Load_Supp_Code()
        cboStatus.ToggleDropdown()
    End Sub

    Function Load_Combo_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Name as [##] from M09Supplier where M09Active='A' and M09Loc_Code='" & _ComCode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 210
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

    Function Load_Combo_Status()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M10Dis as [##] from M10Status  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboStatus
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 90
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

    Function Load_Combo_Type()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Dis as [##] from M11Supp_Type  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboType
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 120
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

    Private Sub frmNew_Supplier_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        strWindowName = ""
    End Sub

    Private Sub frmNew_Supplier_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _ComCode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Combo_Name()
        Call Load_Combo_Status()
        Call Load_Combo_Type()
        Call Load_Supp_Code()
        Call Load_Grid()
        txtCode.ReadOnly = True

    End Sub

    Function Load_Supp_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='SP' and P01Com_Code='" & _ComCode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") <= 10 Then
                    txtCode.Text = "SP00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") > 10 And M01.Tables(0).Rows(0)("P01LastNo") <= 100 Then
                    txtCode.Text = "SP0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtCode.Text = "SP" & M01.Tables(0).Rows(0)("P01LastNo")
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

  

    Private Sub cboStatus_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboStatus.KeyUp
        If e.KeyCode = 13 Then
            If cboStatus.Text <> "" Then
                cboName.ToggleDropdown()
            End If
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
        End If
    End Sub

    Private Sub cboName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboName.KeyUp
        If e.KeyCode = 13 Then
            txtAddress.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
        End If
    End Sub

    Private Sub txtAddress_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyUp
        If e.KeyCode = 13 Then
            txtAdd1.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
        End If
    End Sub

    Private Sub txtAdd1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAdd1.KeyUp
        If e.KeyCode = 13 Then
            txtTp.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
        End If
    End Sub

    Private Sub txtTp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTp.KeyUp
        If e.KeyCode = 13 Then
            txtFax.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
        End If
    End Sub

    Private Sub txtFax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFax.KeyUp
        If e.KeyCode = 13 Then
            txtContact.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
        End If
    End Sub

    Private Sub txtContact_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtContact.KeyUp
        If e.KeyCode = 13 Then
            txtVAT.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
        End If
    End Sub

    Private Sub txtVAT_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVAT.KeyUp
        If e.KeyCode = 13 Then
            cboType.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
        End If
    End Sub

    Private Sub cboType_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboType.KeyUp
        If e.KeyCode = 13 Then
            cmdSave.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            strWindowName = "frmNew_Supplier"
            txtFind.Focus()
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
        Dim t01 As DataSet

        Try
            If cboName.Text <> "" Then
            Else
                MsgBox("Please enter the Supplier Name", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                cboName.ToggleDropdown()
                Exit Sub
            End If

            If cboStatus.Text <> "" Then
            Else
                MsgBox("Please enter the Status", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                cboStatus.ToggleDropdown()
                Exit Sub
            End If

            If cboType.Text <> "" Then
            Else
                MsgBox("Please enter the Type", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                cboType.ToggleDropdown()
                Exit Sub
            End If


            If txtAdd1.Text <> "" Then
            Else
                txtAdd1.Text = " "
            End If


            If txtAddress.Text <> "" Then
            Else
                txtAddress.Text = " "
            End If

            If txtVAT.Text <> "" Then
            Else
                txtVAT.Text = " "
            End If

            If txtContact.Text <> "" Then
            Else
                txtContact.Text = " "
            End If

            If txtTp.Text <> "" Then
            Else
                txtTp.Text = " "
            End If

            If txtFax.Text <> "" Then
            Else
                txtFax.Text = " "
            End If

            nvcFieldList1 = "SELECT * FROM M09Supplier where M09Code='" & txtCode.Text & "' and M09Loc_Code='" & _ComCode & "'"
            t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(t01) Then

                nvcFieldList1 = "UPDATE M09Supplier SET M09Status='" & cboStatus.Text & "',M09Name='" & cboName.Text & "',M09Address='" & txtAddress.Text & "',M09Address1='" & txtAdd1.Text & "',M09TP='" & txtTp.Text & "',M09VAT='" & txtVAT.Text & "',M09Fax='" & txtFax.Text & "',M09Contact_On='" & txtContact.Text & "',M09Type='" & cboType.Text & "',M09Active='A' WHERE M09Code='" & txtCode.Text & "' and M09Loc_Code='" & _ComCode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE M01Account_Master SET M01Acc_Name='" & cboName.Text & "',M01Address='" & txtAddress.Text & "',M01Address2='" & txtAdd1.Text & "',M01TP='" & txtTp.Text & "',M01Status='A' WHERE M01Acc_Code='" & txtCode.Text & "' AND M01Acc_Type='SP'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            Else
                Call Load_Supp_Code()

                nvcFieldList1 = "UPDATE P01PARAMETER SET P01LastNo=P01LastNo +" & 1 & " WHERE P01CODE='SP' and P01Com_Code='" & _ComCode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M09Supplier(M09Code,M09Status,M09Name,M09Address,M09Address1,M09TP,M09VAT,M09Fax,M09Contact_On,M09Type,M09Time,M09User,M09Active,M09Loc_Code)" & _
                                                                  " values('" & Trim(txtCode.Text) & "', '" & Trim(cboStatus.Text) & "','" & Trim(cboName.Text) & "','" & txtAddress.Text & "','" & txtAdd1.Text & "','" & txtTp.Text & "','" & txtVAT.Text & "','" & txtFax.Text & "','" & txtContact.Text & "','" & cboType.Text & "','" & Now & "','" & strDisname & "','A','" & _ComCode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M01Account_Master(M01Acc_Type,M01Acc_Code,M01Acc_Name,M01Address,M01Address2,M01TP,M01Acc_Limit,M01DOC,M01User,M01Status,M01year,M01Comm,M01Com_Code,M01ACC_OF,M01OB_Chq)" & _
                                                                  " values('SP', '" & Trim(txtCode.Text) & "','" & Trim(cboName.Text) & "','" & txtAddress.Text & "','" & txtAdd1.Text & "','" & txtTp.Text & "','0','" & Today & "','" & strDisname & "','A','" & Year(Today) & "','0','" & _ComCode & "' ,'MS','0')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information .......")
            transaction.Commit()

            Call Claer_Text()
            Call Load_Supp_Code()
            Call Load_Combo_Name()
            cboStatus.ToggleDropdown()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Type as [##],M09Code as [Supplier Code],M09Name as [supplier Name] from M09Supplier where M09Active='A' and M09Loc_Code='" & _ComCode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
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


    Function Load_Grid1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M09Type as [##],M09Code as [Supplier Code],M09Name as [supplier Name] from M09Supplier where M09Active='A' and M09Name like '" & txtFind.Text & "%' and M09Loc_Code='" & _ComCode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
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

    Private Sub txtFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyUp
        If e.KeyCode = Keys.Escape Then
            Call Load_Grid()
            OPR4.Visible = False
            cboStatus.ToggleDropdown()
        End If
    End Sub

    Private Sub txtFind_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFind.ValueChanged
        Call Load_Grid1()
    End Sub

    Function Search_Record(ByVal strCode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M09Supplier where M09Active='A' and M09Code='" & strCode & "' and M09Loc_Code='" & _ComCode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01
                    txtCode.Text = strCode
                    cboStatus.Text = .Tables(0).Rows(0)("M09Status")
                    cboName.Text = .Tables(0).Rows(0)("M09Name")
                    cboType.Text = .Tables(0).Rows(0)("M09Type")
                    txtAddress.Text = .Tables(0).Rows(0)("M09Address")
                    txtAdd1.Text = .Tables(0).Rows(0)("M09Address1")
                    txtContact.Text = .Tables(0).Rows(0)("M09Contact_On")
                    txtFax.Text = .Tables(0).Rows(0)("M09Fax")
                    txtTp.Text = .Tables(0).Rows(0)("M09TP")
                    txtVAT.Text = .Tables(0).Rows(0)("M09VAT")
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

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        ' On Error Resume Next
        Dim _RowIndex As Integer
        Dim _SupCode As String
        _RowIndex = UltraGrid1.ActiveRow.Index

        _SupCode = UltraGrid1.Rows(_RowIndex).Cells(1).Text
        Search_Record(_SupCode)
        OPR4.Visible = False
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
        Dim t01 As DataSet
        Dim A As String
        Try
            A = MsgBox("Are you sure you want to delete this Supplier", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete ......")
            If A = vbYes Then
                nvcFieldList1 = "UPDATE M09Supplier SET M09Active='I' WHERE M09Code='" & txtCode.Text & "' and M09Loc_Code='" & _ComCode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE M01Account_Master SET M01Status='I' WHERE M01Acc_Code='" & txtCode.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Record deleted successfully", MsgBoxStyle.Information, "Information ......")
            End If

            transaction.Commit()
            connection.Close()
            Call Claer_Text()
            Call Load_Combo_Name()
            Call Load_Supp_Code()
            cboStatus.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub
End Class