Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Public Class frmMaterial_Group
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        cboFrom.ToggleDropdown()

    End Sub

    Function LoadGride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M13From as [From Material No],M13To as [To Material No] from M13Material_Category "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 150
            UltraGrid1.Rows.Band.Columns(1).Width = 150
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmMaterial_Group_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Call LoadGride()

        Try
            Sql = "select M11SAPCode as [Material No] from M11MRS group by M11SAPCode"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboFrom
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
                End With
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Sub

    Function Load_Comdo2(ByVal strMT As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet



        Try
            Sql = "select M11SAPCode as [Material No] from M11MRS where M11SAPCode<> '" & strMT & "' group by M11SAPCode"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboTo
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 175
                End With
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cboFrom_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFrom.AfterCloseUp
        Call Load_Comdo2(cboFrom.Text)
    End Sub

    

    Private Sub cboFrom_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboFrom.KeyUp
        If e.KeyCode = 13 Then
            Call Load_Comdo2(cboFrom.Text)
            cboTo.ToggleDropdown()
        End If
    End Sub

    Private Sub cboTo_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTo.AfterCloseUp
        Call Search_Data()
    End Sub

    Private Sub cboTo_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboTo.InitializeLayout

    End Sub

    Private Sub cboTo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboTo.KeyUp
        If e.KeyCode = 13 Then
            If cmdSave.Enabled = True Then
                cmdSave.Focus()
            Else
                cmdEdit.Focus()
            End If
        End If
    End Sub

    Private Sub cboTo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTo.TextChanged
        Call Search_Data()
    End Sub

    Function Search_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
            Sql = "select * from M13Material_Category where  M13From='" & cboFrom.Text & "' and M13To='" & cboTo.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cmdSave.Enabled = False
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
                cmdSave.Enabled = True
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
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


                nvcFieldList = "delete from M13Material_Category where  M13From = '" & Trim(cboFrom.Text) & "' and M13To='" & cboTo.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, "Textured Jersey ........")
                transaction.Commit()

            End If

            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False

            cmdAdd.Focus()
            Call LoadGride()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
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
        Dim M01 As DataSet

        Try
          
            nvcFieldList1 = "select * from M11MRS where M11SAPCode='" & Trim(cboFrom.Text) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
            Else
                MsgBox("Please select the correct Material No", MsgBoxStyle.Information, "Textured Jersey ........")
                cboFrom.ToggleDropdown()
                Exit Sub
            End If
            '--------------------------------------------------------------------------------------
            nvcFieldList1 = "select * from M11MRS where M11SAPCode='" & Trim(cboTo.Text) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
            Else
                MsgBox("Please select the correct Material No", MsgBoxStyle.Information, "Textured Jersey ........")
                cboTo.ToggleDropdown()
                Exit Sub
            End If
            '--------------------------------------------------------------------------------------
            nvcFieldList1 = "select * from M13Material_Category where M13To='" & Trim(cboTo.Text) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                MsgBox("Please select the correct Material No", MsgBoxStyle.Information, "Textured Jersey ........")
                cboFrom.ToggleDropdown()
                Exit Sub
            Else

            End If

            nvcFieldList1 = "Insert Into M13Material_Category(M13From,M13To)" & _
                                                     " values('" & Trim(cboFrom.Text) & "', '" & Trim(cboTo.Text) & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False

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