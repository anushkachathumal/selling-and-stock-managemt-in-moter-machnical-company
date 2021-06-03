Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration

Public Class frmLocation
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _Comcode As String

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = Keys.Enter Then
            Call Search_Records()
            If Trim(txtCode.Text) <> "" Then
                txtDescription.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_Records()
            txtDescription.Focus()

        End If
    End Sub

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M04Loc_Code as [Loc Code],M04Loc_Name as [Location Name] from M04Location where M04Com_Code='" & _Comcode & "'  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select * from M04Location where M04Loc_Code='" & Trim(txtCode.Text) & "' and M04Com_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                txtDescription.Text = dsUser.Tables(0).Rows(0)("M04Loc_Name")
                cmdAdd.Enabled = False
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = Keys.Enter Then
            If cmdAdd.Enabled = True Then
                cmdAdd.Focus()
            Else
                cmdEdit.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If cmdAdd.Enabled = True Then
                cmdAdd.Focus()
            Else
                cmdEdit.Focus()
            End If
        End If
    End Sub
    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        ' OPR2.Enabled = False
        'OPR1.Enabled = False
        OPR0.Enabled = True
        ' OPR3.Enabled = False
        cmdAdd.Enabled = True
        cmdAdd.Focus()
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        Call Load_Gride()

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
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

        Try
            If Trim(txtCode.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Location Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    Exit Sub
                End If
            End If

            '--------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Location Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    Exit Sub
                End If
            End If
            '--------------------------------------------------------------------
            nvcFieldList1 = "Insert Into M04Location(M04Loc_Code,M04Loc_Name,M04Com_Code)" & _
                                                          " values('" & (Trim(txtCode.Text)) & "', '" & (Trim(txtDescription.Text)) & "','" & _Comcode & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            OPR0.Enabled = True
            ' OPR3.Enabled = False
            cmdAdd.Enabled = True
            txtCode.Focus()
            ' cmdSave.Enabled = False
            cmdDelete.Enabled = False
            Call Load_Gride()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
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
        Dim MB51 As DataSet
        Dim i As Integer
        Dim result1 As String

        Try
            If Trim(txtCode.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Location Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    Exit Sub
                End If
            End If

            '--------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Location Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    Exit Sub
                End If
            End If
            '--------------------------------------------------------------------
            nvcFieldList1 = "update M04Location set M04Loc_Name='" & (Trim(txtDescription.Text)) & "' where M04Loc_Code='" & Trim(txtCode.Text) & "' and M04Com_Code='" & _Comcode & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            ' OPR0.Enabled = False
            ' OPR3.Enabled = False
            OPR0.Enabled = True
            cmdAdd.Enabled = True
            txtCode.Focus()
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Call Load_Gride()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim A As String
        Dim nvcFieldList As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Try
            A = MsgBox("Are you sure you want to Delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Technova .........")
            If A = vbYes Then
                nvcFieldList = "delete from M04Location where M04Loc_Code = '" & Trim(txtCode.Text) & "' and M04Com_Code='" & _Comcode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)


            End If
            transaction.Commit()
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            ' cmdSave.Enabled = False
            OPR0.Enabled = True
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            txtCode.Focus()
            Call Load_Gride()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Private Sub frmLocation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        _Comcode = ConfigurationManager.AppSettings("LOCCODE")
        Call Load_Gride()
    End Sub
End Class