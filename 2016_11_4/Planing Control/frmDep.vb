Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Public Class frmDep
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        Call Load_Code()
        txtCode.Focus()

    End Sub

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = 13 Then
            If txtCode.Text <> "" Then
                Call SearchRecords()
                cboDep.ToggleDropdown()
            End If

        End If
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        cboDep.Text = ""
    End Sub

    Function SearchRecords() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        SearchRecords = False
        Try
            Sql = "select * from M19Segrigrade where M19Code='" & Trim(txtCode.Text) & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If dsUser.Tables(0).Rows.Count > 0 Then
                SearchRecords = True

                cboDep.Text = dsUser.Tables(0).Rows(0)("M19Group")
                cboNext.Text = dsUser.Tables(0).Rows(0)("M19Dis")
                cmdSave.Enabled = False
                cmdDelete.Enabled = True
                'cmdEdit.Enabled = True
            Else
                SearchRecords = False
                cmdDelete.Enabled = False
                ' cmdEdit.Enabled = False
                cmdSave.Enabled = True
            End If

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
            Sql = "select M19Code as [Code],M19Dis as [Description] from M19Segrigrade "
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

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
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
                nvcFieldList = "delete from M19Segrigrade where M19Code = '" & Trim(txtCode.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)


            End If
            transaction.Commit()
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


    Function Load_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try
            Sql = "select * from P01Parameter where P01Code='DP'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                txtCode.Text = dsUser.Tables(0).Rows(0)("P01LastNo")
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmDep_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Code()
        Call LoadGride()
        Call Load_NextOp()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try
            Sql = "select M20Dis as [Department] from M20Department"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDep
                .DataSource = dsUser
                .Rows.Band.Columns(0).Width = 175
            End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Sub

    Function Load_NextOp()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Try
            Sql = "SELECT M18NextOparation as [Next Oparation] FROM M18WIP WHERE M18NextOparation NOT IN (SELECT M19dis FROM M19Segrigrade)"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboNext
                .DataSource = dsUser
                .Rows.Band.Columns(0).Width = 475
            End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Function

    Private Sub cboDep_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboDep.InitializeLayout

    End Sub

    Private Sub cboDep_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDep.KeyUp
        If e.KeyCode = 13 Then
            cboNext.ToggleDropdown()
        End If
    End Sub

    Private Sub cboNext_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboNext.InitializeLayout

    End Sub

    Private Sub cboNext_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboNext.KeyUp
        If e.KeyCode = 13 Then
            If cmdSave.Enabled = True Then
                cmdSave.Focus()
            Else
                cmdDelete.Focus()
            End If
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

        Try
            nvcFieldList1 = "select * from M20Department where M20Dis='" & cboDep.Text & "'"
            dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(dsUser) Then
            Else
                MsgBox("Please select the department", MsgBoxStyle.Information, "Textured Jersey ......")
                cboDep.ToggleDropdown()
                Exit Sub
            End If
            If Trim(txtCode.Text) <> "" And Trim(cboNext.Text) <> "" Then
                Call Load_Code()

                nvcFieldList1 = "Insert Into M19Segrigrade (M19Code,M19Dis,M19Group)" & _
                                                         " values('" & Trim(txtCode.Text) & "', '" & Trim(cboNext.Text) & "','" & cboDep.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+" & 1 & " where P01Code='DP'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                MsgBox("Please enter the complete records", MsgBoxStyle.Information, "Textured Jersey ........")
            End If
            transaction.Commit()
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