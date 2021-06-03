
Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmTypeofDye
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        'OPR1.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        txtCode.Focus()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function FindRecords() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        FindRecords = False
        Try
            Sql = "select * from M03Dye_Type where M03Code='" & Trim(txtCode.Text) & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                FindRecords = True
                txtName.Text = dsUser.Tables(0).Rows(0)("M03Name")
                ' If Not IsDBNull(dsUser.Tables(0).Rows(0)("M03Address")) Then txtAddress.Text = dsUser.Tables(0).Rows(0)("M03Address")
                'If Not IsDBNull(dsUser.Tables(0).Rows(0)("M03contactNo")) Then txtContact.Text = dsUser.Tables(0).Rows(0)("M03contactNo")
                'If Not IsDBNull(dsUser.Tables(0).Rows(0)("M03Email")) Then txtEmail.Text = dsUser.Tables(0).Rows(0)("M03Email")
                'If Not IsDBNull(dsUser.Tables(0).Rows(0)("M03Bank")) Then txtBank.Text = dsUser.Tables(0).Rows(0)("M03Bank")
                'If Not IsDBNull(dsUser.Tables(0).Rows(0)("M03Branch")) Then txtBranch.Text = dsUser.Tables(0).Rows(0)("M03Branch")
                'If Not IsDBNull(dsUser.Tables(0).Rows(0)("M03AccountNo")) Then txtAccountNo.Text = dsUser.Tables(0).Rows(0)("M03AccountNo")
                cmdDelete.Enabled = True
                cmdEdit.Enabled = True

            Else
                FindRecords = False
                ' txtAddress.Text = ""
                txtName.Text = ""
                'txtContact.Text = ""
                'txtEmail.Text = ""
                'txtBank.Text = ""
                'txtBranch.Text = ""
                'txtAccountNo.Text = ""
                cmdSave.Enabled = False
                cmdDelete.Enabled = False
                cmdEdit.Enabled = False
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = 13 Then
            If txtCode.Text <> "" Then
                FindRecords()
                txtName.Focus()
            End If
        End If
    End Sub

    Private Sub txtCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.ValueChanged

    End Sub

    Private Sub txtName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyUp
        If e.KeyCode = 13 Then
            If cmdSave.Enabled = True Then
                cmdSave.Focus()
            Else
                cmdEdit.Focus()
            End If
        End If
    End Sub

    Private Sub txtName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.TextChanged
        If cmdDelete.Enabled = True Then
            cmdSave.Enabled = False
        Else
            If Trim(txtCode.Text) <> "" And Trim(txtName.Text) <> "" Then
                cmdSave.Enabled = True
            End If
        End If
    End Sub

    Private Sub txtName_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.ValueChanged

    End Sub





    Private Sub txtAccountNo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        Clicked = ""
        cmdAdd.Enabled = True
        cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
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


            If Trim(txtCode.Text) <> "" And Trim(txtName.Text) <> "" Then
                nvcFieldList1 = "Insert Into M03Dye_Type(M03Code,M03Name)" & _
                                                         " values('" & Trim(txtCode.Text) & "', '" & Trim(txtName.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                MsgBox("Please enter the complete records", MsgBoxStyle.Information, "Textured Jersey ........")
                Exit Sub
            End If
            MsgBox("Record Update Sucessfully", MsgBoxStyle.Information, "Textured Jersey .........")
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

                If FindRecords() = True Then
                    nvcFieldList = "delete from M03Dye_Type where  M03Code = '" & Trim(txtCode.Text) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList)
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, "Textured Jersey ........")
                    transaction.Commit()
                Else
                    MsgBox("Cant delete this records", MsgBoxStyle.Information, "Textured Jersey ..........")
                End If

            End If

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


            If Trim(txtCode.Text) <> "" And Trim(txtName.Text) <> "" Then
                nvcFieldList1 = "Update M03Dye_Type set M03Name='" & Trim(txtName.Text) & "' where M03Code='" & Trim(txtCode.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                MsgBox("Please enter the complete records", MsgBoxStyle.Information, "Textured Jersey ........")
            End If
            MsgBox("Record Update Sucessfully", MsgBoxStyle.Information, "Textured Jersey .........")
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

    Function LoadGride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M03Code as [Code],M03Name as [Description] from M03Dye_Type"
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

    Private Sub frmTypeofDye_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadGride()
    End Sub
End Class