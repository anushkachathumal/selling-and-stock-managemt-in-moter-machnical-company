Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmLIB
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = Keys.Enter Then
            Call Search_Records()
            txtDescription.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDescription.Focus()
        End If
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select * from M12LIB_Department where M12Code='" & Trim(txtCode.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                txtDescription.Text = M01.Tables(0).Rows(0)("M12Department")


                cmdEdit.Enabled = True
                cmdDelete.Enabled = True

            Else
                txtDescription.Text = ""
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                con.close()
            End If
        End Try
    End Function

    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = 13 Then
            cmdSave.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            cmdSave.Focus()
        End If
    End Sub

    Function LoadGride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            If Trim(txtDescription.Text) <> "" Then

                Sql = "select M12Code as [Ref Code],M12Department as [LIB Department] from M12LIB_Department where M12Department like '" & txtDescription.Text & "%'"
            Else
                Sql = "select M12Code as [Ref Code],M12Department as [LIB Department] from M12LIB_Department "
            End If
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmLIB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadGride()
    End Sub

    Private Sub txtDescription_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescription.ValueChanged
        Call LoadGride()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        ' OPR2.Enabled = False
        'OPR1.Enabled = False
        OPR0.Enabled = True
        ' OPR3.Enabled = False

        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        Call LoadGride()
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
        Dim MB51 As DataSet
        Dim i As Integer

        Try

            '-----------------------------------------------------------------------------------------
            If Trim(txtCode.Text) <> "" Then
            Else
                MsgBox("Please enter the code", MsgBoxStyle.Information, "Information .....")
                txtCode.Focus()
                Exit Sub
            End If
            '------------------------------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                MsgBox("Please enter the Department", MsgBoxStyle.Information, "Information ......")
                txtDescription.Focus()
                Exit Sub
            End If
            '------------------------------------------------------------------------------------------
            nvcFieldList1 = "Insert Into M12LIB_Department(M12Code,M12Department)" & _
                                                        " values('" & UCase(Trim(txtCode.Text)) & "', '" & (Trim(txtDescription.Text)) & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)



            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            connection.Close()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            OPR0.Enabled = True

            cmdSave.Enabled = True
            cmdDelete.Enabled = False
            Call LoadGride()

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
            A = MsgBox("Are you sure you want to Delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "BlueCloude.........")
            If A = vbYes Then
                nvcFieldList = "delete from M12LIB_Department where M12Code = '" & Trim(txtCode.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)


            End If
            transaction.Commit()
            common.ClearAll(OPR0)
            Clicked = ""
            OPR0.Enabled = True
            cmdSave.Enabled = True
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False

            Call LoadGride()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
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

        Try

            '-----------------------------------------------------------------------------------------
            If Trim(txtCode.Text) <> "" Then
            Else
                MsgBox("Please enter the code", MsgBoxStyle.Information, "Information .....")
                txtCode.Focus()
                Exit Sub
            End If
            '------------------------------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                MsgBox("Please enter the Department", MsgBoxStyle.Information, "Information ......")
                txtDescription.Focus()
                Exit Sub
            End If
            '------------------------------------------------------------------------------------------

            nvcFieldList1 = "update M12LIB_Department set M12Department='" & (Trim(txtDescription.Text)) & "' where M12Code='" & Trim(txtCode.Text) & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            OPR0.Enabled = True

            cmdSave.Enabled = True
            cmdDelete.Enabled = False
            cmdEdit.Enabled = False
            Call LoadGride()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub
End Class