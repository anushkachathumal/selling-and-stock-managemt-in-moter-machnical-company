Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmKnittingMC
    Dim Clicked As String
    'Develop by Suranga R Wijesinghe
    'Developing Date - 2011/04/14
    'Time - 10.30 PM -
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function LoadGride()
        'Load Color data to gride
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Sql = "select M09Code as [Ref.Doc.No],M09Name as [Description] from M09Downtime"
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        UltraGrid1.DataSource = M01
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 270
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '  .DisplayLayout.Bands(0).Columns(2).Width = 90

        End With
    End Function

    Function LoadGride1()
        'Filter Color data to gride
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Sql = "select M09Code as [Ref.Doc.No],M09Name as [Description] from M09Downtime where  M09Name like '" & Trim(txtVoucher.Text) & "%'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        UltraGrid1.DataSource = M01
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 270
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            ' .DisplayLayout.Bands(0).Columns(2).Width = 90

        End With
    End Function

    Private Sub txtVoucher_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVoucher.KeyUp
        If e.KeyCode = 13 Then
            If txtVoucher.Text <> "" Then
                If cmdSave.Enabled = True Then
                    cmdSave.Focus()
                Else
                    cmdDelete.Focus()
                End If
            Else

            End If
        End If
    End Sub

    Private Sub txtVoucher_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVoucher.TextChanged
        Call LoadGride1()
    End Sub


    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        txtVoucher.Focus()
        cmdSave.Enabled = True

        txtDis.Focus()
    End Sub

    Function Serch_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select * from M09Downtime where M09Code='" & Trim(txtDis.Text) & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtVoucher.Text = M01.Tables(0).Rows(0)("M09Name")
                cmdDelete.Enabled = True
                cmdEdit.Enabled = True
                cmdSave.Enabled = False
            Else
                txtVoucher.Text = ""
                cmdDelete.Enabled = False
                cmdSave.Enabled = True
                cmdEdit.Enabled = False
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub txtDis_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDis.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtDis.Text) <> "" Then
                txtVoucher.Focus()
                Call Serch_Records()
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

        Dim P01Parameter As Integer
        Dim M01 As DataSet

        Try
            If Trim(txtVoucher.Text) <> "" Then
                'If Trim(txtDis.Text) <> "" Then
                'nvcFieldList1 = "select * from P01Parameter where P01Code='UD'"
                'M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(M01) Then
                '    P01Parameter = M01.Tables(0).Rows(0)("P01LastNo")
                'End If
                ''----------------------------------------------------------
                'P01Parameter = P01Parameter + 1

                nvcFieldList1 = "Insert Into M09Downtime(M09Code,M09Name)" & _
                                                           " values('" & Trim(txtDis.Text) & "', '" & Trim(txtVoucher.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '----------------------------------------------------------


                transaction.Commit()
                MsgBox("Record saved successfully ", MsgBoxStyle.Information, "Information ......")
                common.ClearAll(OPR0)
                Clicked = ""
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdAdd.Focus()
                Call LoadGride()

                'Else
                '    MsgBox("NX Number cannot be blank...! ", MsgBoxStyle.Information, "Information ....")
                '    txtDis.Focus()
                'End If
            Else
                MsgBox("Downtime Reason cannot be blank...! ", MsgBoxStyle.Information, "Information ....")
                txtVoucher.Focus()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub frmKnittingMC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadGride()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        cmdAdd.Focus()

        Call LoadGride()
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

        Dim P01Parameter As Integer
        Dim M01 As DataSet

        Try
            If Trim(txtVoucher.Text) <> "" Then
                ' If Trim(txtDis.Text) <> "" Then


                nvcFieldList1 = "update M09Downtime set M09Name='" & Trim(txtVoucher.Text) & "' where M09code='" & Trim(txtDis.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '----------------------------------------------------------

                transaction.Commit()
                MsgBox("Record Update successfully ", MsgBoxStyle.Information, "Information ......")
                common.ClearAll(OPR0)
                Clicked = ""
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdAdd.Focus()
                Call LoadGride()

                'MsgBox("Department Name cannot be blank...! ", MsgBoxStyle.Information, "Information ....")
                'txtVoucher.Focus()
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

        Dim P01Parameter As Integer
        Dim M01 As DataSet
        Dim A As String
        Try


            A = MsgBox("Are you sure you want to delete this Downtime", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete .......")
            If A = vbYes Then
                nvcFieldList1 = "delete from M09Downtime  where M09Code='" & Trim(txtDis.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '----------------------------------------------------------

                transaction.Commit()
                MsgBox("Record Deleted successfully ", MsgBoxStyle.Information, "Information ......")
                common.ClearAll(OPR0)
                Clicked = ""
                cmdAdd.Enabled = True
                cmdSave.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdAdd.Focus()
                Call LoadGride()
            End If
            'Else
            '    'MsgBox("NX Number cannot be blank...! ", MsgBoxStyle.Information, "Information ....")
            '    'txtDis.Focus()
            'End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub
End Class