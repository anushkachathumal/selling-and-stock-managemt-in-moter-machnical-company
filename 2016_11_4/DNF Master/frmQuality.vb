Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin


Public Class frmQuality
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
                txtName.Focus()
            End If

        End If
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        txtName.Text = ""
    End Sub

  

    
    Function SearchRecords() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        SearchRecords = False
        Try
            Sql = "select * from M09Quality where M09Code='" & Trim(txtCode.Text) & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If dsUser.Tables(0).Rows.Count > 0 Then
                SearchRecords = True

                txtName.Text = dsUser.Tables(0).Rows(0)("M09Dis")
                cboPO.Text = dsUser.Tables(0).Rows(0)("M09Type")
                txtRTR.Text = dsUser.Tables(0).Rows(0)("M09RTR")
                txtCon.Text = dsUser.Tables(0).Rows(0)("M09MKG")

                cmdSave.Enabled = False
                cmdDelete.Enabled = True
                cmdEdit.Enabled = True
            Else
                SearchRecords = False
                cmdDelete.Enabled = False
                cmdEdit.Enabled = False
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
            Sql = "select M09Code as [Code],M09Dis as [Description],M09Type as [Type],M09RTR as [RTR],M09MKG as [Conversion factor] from M09Quality "
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

    Private Sub txtName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyUp
        If e.KeyCode = 13 Then
            cboPO.ToggleDropdown()
        End If
    End Sub

    Private Sub txtName_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.ValueChanged

    End Sub

    Private Sub cboPO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboPO.InitializeLayout

    End Sub

    Private Sub cboPO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPO.KeyUp
        If e.KeyCode = 13 Then
            txtRTR.Focus()
        End If
    End Sub

    Private Sub txtRTR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRTR.KeyUp
        If e.KeyCode = 13 Then
            txtCon.Focus()
        End If
    End Sub

    Private Sub txtRTR_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRTR.ValueChanged

    End Sub

    Private Sub txtCon_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCon.KeyUp
        If e.KeyCode = 13 Then
            If cmdSave.Enabled = True Then
                cmdSave.Focus()
            Else
                cmdEdit.Focus()
            End If
        End If
    End Sub

    Private Sub txtCon_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCon.ValueChanged

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
            If Trim(txtCode.Text) <> "" And Trim(txtName.Text) <> "" And Trim(cboPO.Text) <> "" And txtRTR.Text <> "" And txtCon.Text <> "" Then
                If IsNumeric(txtRTR.Text) Then
                Else
                    MsgBox("Please enter the correct RTR Value", MsgBoxStyle.Information, "Information.......")
                    Exit Sub
                End If

                If IsNumeric(txtCon.Text) Then
                Else
                    MsgBox("Please enter the correct Conversion factor", MsgBoxStyle.Information, "Information.......")
                    Exit Sub
                End If

                nvcFieldList1 = "Insert Into M09Quality(M09Code,M09Dis,M09Type,M09RTR,M09MKG)" & _
                                                         " values('" & Trim(txtCode.Text) & "', '" & Trim(txtName.Text) & "','" & cboPO.Text & "','" & txtRTR.Text & "','" & txtCon.Text & "')"
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

        Try
            If Trim(txtCode.Text) <> "" And Trim(txtName.Text) <> "" And Trim(cboPO.Text) <> "" And txtRTR.Text <> "" And txtCon.Text <> "" Then
                If IsNumeric(txtRTR.Text) Then
                Else
                    MsgBox("Please enter the correct RTR Value", MsgBoxStyle.Information, "Information.......")
                    Exit Sub
                End If

                If IsNumeric(txtCon.Text) Then
                Else
                    MsgBox("Please enter the correct Conversion factor", MsgBoxStyle.Information, "Information.......")
                    Exit Sub
                End If

                nvcFieldList1 = "Update M09Quality set M09Dis='" & txtName.Text & "',M09Type='" & cboPO.Text & "',M09RTR='" & txtRTR.Text & "',M09MKG='" & txtCon.Text & "' where M09Code='" & txtCode.Text & "'"
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

    Private Sub frmQuality_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Load_Combo()
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M09Type as [Type] from M09Quality group by M09Type"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With cboPO
                    .DataSource = M01
                    .Rows.Band.Columns(0).Width = 675
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

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
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


                nvcFieldList = "delete from M09Quality where  M09Code = '" & Trim(txtCode.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList)
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, "Textured Jersey ........")
                transaction.Commit()
            
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

    Private Sub txtCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.ValueChanged

    End Sub
End Class