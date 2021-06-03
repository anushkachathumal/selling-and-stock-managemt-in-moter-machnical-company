Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Public Class frmDowntime
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        txtCode.Focus()
    End Sub

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = 13 Then
            If txtCode.Text <> "" Then
                Call SearchRecords()
                txtDescription.Focus()
            End If

        End If
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        txtDescription.Text = ""
    End Sub

    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = 13 Then
            If txtDescription.Text <> "" Then
                If SearchRecords() = True Then
                    cmdDelete.Focus()
                Else
                    cmdSave.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub txtDescription_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.TextChanged
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        If txtDescription.Text <> "" Then
            Try
                Sql = "select * from M05DownTime where M05Code='" & Trim(txtCode.Text) & "'"
                dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If dsUser.Tables(0).Rows.Count > 0 Then
                    cmdSave.Enabled = False
                Else
                    cmdDelete.Enabled = False
                    cmdSave.Enabled = True
                End If

                DBEngin.CloseConnection(con)
            Catch returnMessage As Exception
                If returnMessage.Message <> Nothing Then
                    MessageBox.Show(returnMessage.Message)
                End If
            End Try
        End If
    End Sub

    Function SearchRecords() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        SearchRecords = False
        Try
            Sql = "select M05Name from M05Downtime where M05Code='" & Trim(txtCode.Text) & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If dsUser.Tables(0).Rows.Count > 0 Then
                SearchRecords = True

                txtDescription.Text = dsUser.Tables(0).Rows(0)("M05name")
                cmdSave.Enabled = False
                cmdDelete.Enabled = True
                'cmdEdit.Enabled = True
            Else
                SearchRecords = False
                cmdDelete.Enabled = False
                ' cmdEdit.Enabled = False
                cmdSave.Enabled = True
            End If

            DBEngin.CloseConnection(con)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function LoadGride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M05Code as [Down Time Type],M05Name as [Description] from M05Downtime "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370


            DBEngin.CloseConnection(con)
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Function

    Private Sub frmDowntime_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadGride()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
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

        Try
            If Trim(txtCode.Text) <> "" And Trim(txtDescription.Text) <> "" Then
                nvcFieldList1 = "Insert Into M05Downtime(M05Code,M05name)" & _
                                                         " values('" & Trim(txtCode.Text) & "', '" & Trim(txtDescription.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                MsgBox("Please enter the complete records", MsgBoxStyle.Information, "Textured Jersey ........")
            End If
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()
            Call LoadGride()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
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
            A = MsgBox("Are you sure you want to Delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Textured Jersey .........")
            If A = vbYes Then
                nvcFieldList = "delete from M05Downtime where  M05Code = '" & Trim(txtCode.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)


            End If
            transaction.Commit()

            DBEngin.CloseConnection(connection)
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()
            Call LoadGride()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub
End Class