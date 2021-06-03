Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmSetupCPI
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
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Sql = "select M11Quality as [Quality],M11CPIValue as [CPI ORIZIO],M11SAN as [CPI SUNTEC] from M11Quality_CPI"
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        UltraGrid1.DataSource = M01
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 140
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '  .DisplayLayout.Bands(0).Columns(2).Width = 90

        End With
        DBEngin.CloseConnection(con)
        con.ConnectionString = ""
    End Function

    Private Sub frmSetupCPI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadGride()
        Call Load_Quality()

    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Clicked = "ADD"
        OPR0.Enabled = True
        ' Call Clear_Text()
        cmdAdd.Enabled = False
        txtVoucher.Focus()
        cmdSave.Enabled = True

        cboCategory.ToggleDropdown()
    End Sub

    Function Load_Quality()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Quality as [Quality] from M03Knittingorder group by M03Quality"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 245
                '  .Rows.Band.Columns(1).Width = 265
            End With

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboCategory_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboCategory.InitializeLayout

    End Sub

    Private Sub cboCategory_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCategory.KeyUp
        If e.KeyCode = 13 Then
            If cboCategory.Text <> "" Then
                Call Search_Quality()
                txtVoucher.Focus()
            End If
        End If
    End Sub

    Private Sub txtVoucher_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVoucher.KeyUp
        If e.KeyCode = 13 Then
            txtSAN.Focus()
        End If



    End Sub

    Function Search_Quality() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from M11Quality_CPI where M11Quality='" & Trim(cboCategory.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Quality = True
                txtVoucher.Text = M01.Tables(0).Rows(0)("M11CPIValue")
                cmdSave.Enabled = False
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
            Else
                cmdDelete.Enabled = False
                cmdEdit.Enabled = False
                cmdSave.Enabled = True
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try

    End Function
    Private Sub txtVoucher_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVoucher.ValueChanged

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
            If Search_Quality() = True Then
                MsgBox("This Quality alrady exsist", MsgBoxStyle.Information, "Textued Jersey ........")
                Exit Sub
            End If

            If IsNumeric(txtVoucher.Text) Then
            Else
                MsgBox("Please enter the correct CPI Value", MsgBoxStyle.Information, "Textued Jersey ........")
                txtVoucher.Focus()
                Exit Sub
            End If

            If IsNumeric(txtSAN.Text) Then
            Else
                MsgBox("Please enter the correct CPI Value", MsgBoxStyle.Information, "Textued Jersey ........")
                txtSAN.Focus()
                Exit Sub
            End If


            nvcFieldList1 = "Insert Into M11Quality_CPI(M11Quality,M11CPIValue,M11SAN)" & _
                                                       " values('" & Trim(cboCategory.Text) & "', '" & Trim(txtVoucher.Text) & "'," & Val(txtSAN.Text) & ")"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            '----------------------------------------------------------


            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            MsgBox("Record saved successfully ", MsgBoxStyle.Information, "Information ......")
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()
            Call LoadGride()

          

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

        Dim P01Parameter As Integer
        Dim M01 As DataSet

        Try
            If IsNumeric(txtVoucher.Text) Then
            Else
                MsgBox("Please enter the correct CPI Value", MsgBoxStyle.Information, "Textued Jersey ........")
                txtVoucher.Focus()
                Exit Sub
            End If
            ' If Trim(txtDis.Text) <> "" Then


            nvcFieldList1 = "update M11Quality_CPI set M11CPIValue='" & Trim(txtVoucher.Text) & "',M11SAN='" & Val(txtSAN.Text) & "' where M11Quality='" & Trim(cboCategory.Text) & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            '----------------------------------------------------------

            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            MsgBox("Record Update successfully ", MsgBoxStyle.Information, "Information ......")
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Focus()
            Call LoadGride()

        

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


            A = MsgBox("Are you sure you want to delete this Knitter", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete .......")
            If A = vbYes Then
                nvcFieldList1 = "delete from M11Quality_CPI  where M11Quality='" & Trim(cboCategory.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '----------------------------------------------------------

                transaction.Commit()
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
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

    Private Sub txtSAN_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSAN.KeyUp
        If e.KeyCode = 13 Then
            If cmdSave.Enabled = True Then
                cmdSave.Focus()
            Else
                cmdEdit.Focus()

            End If
        End If
    End Sub

    Private Sub txtSAN_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSAN.ValueChanged

    End Sub
End Class