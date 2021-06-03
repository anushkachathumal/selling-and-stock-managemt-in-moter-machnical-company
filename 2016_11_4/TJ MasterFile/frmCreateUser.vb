Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports Microsoft.VisualBasic.FileIO
Public Class frmCreateUser
    Dim Clicked As String
    Dim n_Group As String
    Dim n_Category As String
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        txtEmp.Focus()
        cmdSave.Enabled = True
    End Sub

    Private Sub frmCreateUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M05name as [User Group] from M05UserGroup"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboEmpgroup
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 245
                '  .Rows.Band.Columns(1).Width = 265
            End With


            Sql = "select M04Name as [Description] from M04UserCategory"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                '  .Rows.Band.Columns(0).Width = 115
                .Rows.Band.Columns(0).Width = 140
            End With

         
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub txtEmp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEmp.KeyUp
        If e.KeyCode = 13 Then
            If txtEmp.Text <> "" Then

                Call Search_Records()
                txtUname.Focus()
            End If
        ElseIf e.KeyCode = 9 Then
            ' MsgBox("")
        End If
    End Sub

    Private Sub txtEmp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmp.LostFocus
        Call Search_Records()
    End Sub

    Private Sub txtEmp_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEmp.ValueChanged

    End Sub

    Private Sub txtUname_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUname.KeyUp
        If e.KeyCode = 13 Then
            txtLast.Focus()
        End If
    End Sub

    Private Sub txtUname_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUname.ValueChanged

    End Sub

    Private Sub txtLast_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLast.KeyUp
        If e.KeyCode = 13 Then
            txtEmail.Focus()
        End If
    End Sub

    Private Sub txtLast_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLast.ValueChanged

    End Sub

    Private Sub txtEmail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEmail.KeyUp
        If e.KeyCode = 13 Then
            cboEmpgroup.ToggleDropdown()
        End If
    End Sub

    Function Find_UserCategory() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Find_UserCategory = False
            Sql = "select * from M04UserCategory where M04Name='" & cboCategory.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then '
                Find_UserCategory = True
                n_Category = M01.Tables(0).Rows(0)("M04CatCode")
            Else
                Find_UserCategory = False
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Function Find_UserGroup() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Find_UserGroup = False
            Sql = "select * from M05UserGroup where M05Name='" & cboEmpgroup.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then '
                Find_UserGroup = True
                n_Group = M01.Tables(0).Rows(0)("M05Code")
            Else
                Find_UserGroup = False
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboEmpgroup_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEmpgroup.AfterCloseUp
        Call Find_UserGroup()
    End Sub

    Private Sub cboEmpgroup_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboEmpgroup.InitializeLayout

    End Sub

    Private Sub cboEmpgroup_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEmpgroup.KeyUp
        If e.KeyCode = 13 Then
            cboCategory.ToggleDropdown()
        End If
    End Sub

    Private Sub cboCategory_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCategory.AfterCloseUp
        Call Find_UserCategory()
    End Sub

    Private Sub cboCategory_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboCategory.InitializeLayout

    End Sub

    Private Sub cboCategory_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCategory.KeyUp
        If e.KeyCode = 13 Then
            txtPW.Focus()
        End If
    End Sub

    Private Sub txtConfirm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConfirm.KeyUp
        If e.KeyCode = 13 Then
            If cmdDelete.Enabled = True Then
                cmdDelete.Focus()
            Else
                cmdSave.Focus()
            End If
        End If
    End Sub

    Private Sub txtConfirm_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtConfirm.ValueChanged

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        cmdAdd.Focus()
    End Sub

    Private Sub txtPW_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPW.KeyUp
        If e.KeyCode = 13 Then
            txtConfirm.Focus()
        End If
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select FirstName,LastName,Email,M04Name,M05name from Users inner join M04UserCategory on M04CatCode=Department inner join M05UserGroup on UGroup=M05Code where EPFNo='" & Trim(txtEmp.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01
                    txtUname.Text = .Tables(0).Rows(0)("FirstName")
                    txtLast.Text = .Tables(0).Rows(0)("LastName")
                    txtEmail.Text = .Tables(0).Rows(0)("Email")
                    cboCategory.Text = .Tables(0).Rows(0)("M04name")
                    cboEmpgroup.Text = .Tables(0).Rows(0)("M05name")
                    cmdSave.Enabled = False
                    cmdDelete.Enabled = True
                    cmdEdit.Enabled = True


                End With

            Else
                cmdSave.Enabled = True
                cmdDelete.Enabled = False
                cmdEdit.Enabled = False
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

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
        Try

            If Find_UserCategory() = True Then
                If Find_UserGroup() = True Then
                    If Trim(txtEmp.Text) <> "" Then
                        If Trim(txtUname.Text) <> "" Then
                            If Trim(txtPW.Text) <> "" Then
                                If Trim(txtPW.Text) = Trim(txtConfirm.Text) Then
                                    nvcFieldList1 = "Insert Into Users(EPFNo,FirstName,LastName,Email,Department,UGroup,Username,Password)" & _
                                                               " values('" & Trim(txtEmp.Text) & "', '" & Trim(txtUname.Text) & "','" & Trim(txtLast.Text) & "','" & Trim(txtEmail.Text) & "','" & n_Group & "','" & n_Category & "','" & Trim(txtUname.Text) & "','" & Trim(txtPW.Text) & "')"
                                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                                    MsgBox("User Created successfully ", MsgBoxStyle.Information, "Information ......")
                                    transaction.Commit()
                                    DBEngin.CloseConnection(connection)
                                    connection.ConnectionString = ""
                                    common.ClearAll(OPR0)
                                    Clicked = ""
                                    cmdAdd.Enabled = True
                                    cmdSave.Enabled = False
                                    cmdEdit.Enabled = False
                                    cmdDelete.Enabled = False
                                    cmdAdd.Focus()
                                Else
                                    MsgBox("Password miss match", MsgBoxStyle.Exclamation, "Texturd Jersy ....")
                                    Exit Sub
                                End If
                            Else
                                MsgBox("Please enter the Password", MsgBoxStyle.Information, "Texturd Jersy .....")
                                txtPW.Focus()
                                Exit Sub
                            End If
                        Else
                            MsgBox("Please enter the User name", MsgBoxStyle.Information, "Texturd Jersy .......")
                            txtUname.Focus()
                            Exit Sub
                        End If
                    Else
                        MsgBox("Please enter the Emp No", MsgBoxStyle.Information, "Texturd Jersy .....")
                        txtEmp.Focus()
                        Exit Sub
                    End If
                Else
                    MsgBox("Please select the user Group", MsgBoxStyle.Information, "Texturd Jersy ....")
                End If
            Else
                MsgBox("Please select the user Category", MsgBoxStyle.Information, "Texturd Jersy ....")
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try

            If Find_UserCategory() = True Then
                If Find_UserGroup() = True Then
                    If Trim(txtEmp.Text) <> "" Then
                        If Trim(txtUname.Text) <> "" Then
                            If Trim(txtPW.Text) <> "" Then
                                If Trim(txtPW.Text) = Trim(txtConfirm.Text) Then
                                    nvcFieldList1 = "update Users set FirstName='" & Trim(txtUname.Text) & "',LastName='" & Trim(txtLast.Text) & "',Email='" & Trim(txtEmail.Text) & "',Department='" & n_Category & "',UGroup='" & n_Group & "',Username='" & Trim(txtUname.Text) & "',Password='" & Trim(txtPW.Text) & "' where EPFNo='" & Trim(txtEmp.Text) & "'"
                                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                                    MsgBox("User Update successfully ", MsgBoxStyle.Information, "Information ......")
                                    transaction.Commit()
                                    DBEngin.CloseConnection(connection)
                                    connection.ConnectionString = ""
                                    common.ClearAll(OPR0)
                                    Clicked = ""
                                    cmdAdd.Enabled = True
                                    cmdSave.Enabled = False
                                    cmdEdit.Enabled = False
                                    cmdDelete.Enabled = False
                                    cmdAdd.Focus()
                                Else
                                    MsgBox("Password miss match", MsgBoxStyle.Exclamation, "Texturd Jersy ....")
                                    Exit Sub
                                End If
                            Else
                                MsgBox("Please enter the Password", MsgBoxStyle.Information, "Texturd Jersy .....")
                                txtPW.Focus()
                                Exit Sub
                            End If
                        Else
                            MsgBox("Please enter the User name", MsgBoxStyle.Information, "Texturd Jersy .......")
                            txtUname.Focus()
                            Exit Sub
                        End If
                    Else
                        MsgBox("Please enter the Emp No", MsgBoxStyle.Information, "Texturd Jersy .....")
                        txtEmp.Focus()
                        Exit Sub
                    End If
                Else
                    MsgBox("Please select the user Group", MsgBoxStyle.Information, "Texturd Jersy ....")
                End If
            Else
                MsgBox("Please select the user Category", MsgBoxStyle.Information, "Texturd Jersy ....")
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
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

            A = MsgBox("Are you sure you want to cancel this user account", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Cancel User Account .......")
            If A = vbYes Then
                nvcFieldList1 = "delete from Users where EPFNo='" & Trim(txtEmp.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("User Cancel successfully ", MsgBoxStyle.Information, "Information ......")
                transaction.Commit()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""

                common.ClearAll(OPR0)
                Clicked = ""
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdAdd.Focus()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class