Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmCatogery_cnt
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
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M01Cat_Code as [Category Code],M01Description as [Category Name] from M01Category WHERE M01Status='A'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370

            con.ClearAllPools()
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Gride_1()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()

        Try
            Sql = "select 'DEACTIVE' AS [##],M01Cat_Code as [Category Code],M01Description as [Category Name] from M01Category WHERE M01Status='I' order by M01ID"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 80
            UltraGrid1.Rows.Band.Columns(1).Width = 110
            UltraGrid1.Rows.Band.Columns(2).Width = 320
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Load_Gride_2()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M01Cat_Code as [Category Code],M01Description as [Category Name] from M01Category WHERE M01Description like '%" & txtDescription.Text & "%' order by M01ID"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370
            ' UltraGrid1.Rows.Band.Columns(2).Width = 320
            con.ClearAllPools()
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()

        Try
            Sql = "select * from M01Category where M01Cat_Code='" & Trim(txtCode.Text) & "'  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                txtDescription.Text = dsUser.Tables(0).Rows(0)("M01Description")
                cmdAdd.Enabled = False
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
            End If

            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function
    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = Keys.Enter Then
            If txtDescription.Text <> "" Then
                If cmdAdd.Enabled = True Then
                    cmdAdd.Focus()
                Else
                    cmdEdit.Focus()
                End If
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If cmdAdd.Enabled = True Then
                cmdAdd.Focus()
            Else
                cmdEdit.Focus()
            End If
        End If
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
                result1 = MessageBox.Show("Please enter the Category Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    Exit Sub
                End If
            End If

            '--------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Category Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    Exit Sub
                End If
            End If
            '--------------------------------------------------------------------
            nvcFieldList1 = "select * from M01Category where M01Cat_Code='" & Trim(txtCode.Text) & "'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                MsgBox("This Category code alrady exsist", MsgBoxStyle.Information, "Information ......")
                connection.Close()
            Else

                nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='CT' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M01Category(M01Cat_Code,M01Description,M01Status)" & _
                                                              " values('" & (Trim(txtCode.Text)) & "', '" & (Trim(txtDescription.Text)) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                             " values('NEW_CATEGORY','SAVE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            End If
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            txtCode.Text = ""
            txtDescription.Text = ""
            txtDescription.Focus()
            Call Load_Entry()
            Call Load_Gride()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
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
                result1 = MessageBox.Show("Please enter the Category Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    Exit Sub
                End If
            End If

            '--------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Category Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    Exit Sub
                End If
            End If
            '--------------------------------------------------------------------
            nvcFieldList1 = "update M01Category set M01Description='" & (Trim(txtDescription.Text)) & "',M01Status='A' where M01Cat_Code='" & Trim(txtCode.Text) & "'  "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                             " values('NEW_CATEGORY','EDIT', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            Call Clear_Text()
            Call Load_Gride()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
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
                nvcFieldList = "UPDATE M01Category set M01Status='I' where M01Cat_Code = '" & Trim(txtCode.Text) & "'  "
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

                nvcFieldList = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                             " values('NEW_CATEGORY','DEACTIVE', '" & Now & "','" & strDisname & "','" & txtCode.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)
            End If
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
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
            connection.ClearAllPools()
            connection.Close()
        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Function Load_Entry()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='CT' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtCode.Text = "CT/00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtCode.Text = "CT/0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtCode.Text = "CT/" & M01.Tables(0).Rows(0)("P01No")
                End If
            End If


            con.ClearAllPools()
            con.close()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Private Sub frmCatogery_cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Gride()
        txtCode.ReadOnly = True
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Entry()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_Text()
    End Sub

    Function Clear_Text()
        Me.txtCode.Text = ""
        Me.txtDescription.Text = ""
        cmdAdd.Enabled = True
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
        Call Load_Entry()
        Call Load_Gride()
    End Function

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Clear_Text()
        Call Load_Gride()
    End Sub

    Private Sub DeactivateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeactivateToolStripMenuItem.Click
        Call Load_Gride_1()
    End Sub

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid1.ActiveRow.Index
        txtCode.Text = Trim(UltraGrid1.Rows(_Row).Cells(0).Text)
        Call Search_Records()
    End Sub

    Private Sub txtDescription_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescription.ValueChanged
        Call Load_Gride_2()
    End Sub
End Class