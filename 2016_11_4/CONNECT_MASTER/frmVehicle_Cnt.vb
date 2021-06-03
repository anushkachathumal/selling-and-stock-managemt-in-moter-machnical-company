Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmVehicle_Cnt
    Dim _PrintStatus As String

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Function Clear_text()
        Me.txtEng_Capacity.Text = ""
        Me.txtRemak.Text = ""
        Me.txtV_No.Text = ""
        Me.cboBrand.Text = ""
        Me.cboType.Text = ""
        txtV_No.Focus()
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_text()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Call Clear_text()
        OPR0.Visible = False
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Clear_text()
        OPR0.Visible = False
    End Sub

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        OPR0.Visible = True
        txtV_No.Focus()
    End Sub

    Private Sub txtV_No_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtV_No.KeyUp
        If e.KeyCode = 13 Then
            If Trim(txtV_No.Text) <> "" Then
                cboBrand.ToggleDropdown()
            End If
        End If
    End Sub

    Private Sub cboBrand_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBrand.KeyUp
        If e.KeyCode = 13 Then
            cboType.ToggleDropdown()
        End If
    End Sub

    Private Sub cboType_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboType.KeyUp
        If e.KeyCode = 13 Then
            txtEng_Capacity.Focus()
        End If
    End Sub

    Private Sub txtEng_Capacity_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEng_Capacity.KeyUp
        If e.KeyCode = 13 Then
            txtRemak.Focus()
        End If
    End Sub

    Private Sub txtRemak_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemak.KeyUp
        If e.KeyCode = Keys.F1 Then
            cmdAdd.Focus()
        End If
    End Sub

    Function Load_Brand()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M21Name as [##] from M21Vehicle_type WHERE M21Status='1' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboBrand
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 159
                '  .Rows.Band.Columns(1).Width = 160


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Type()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M21Name as [##] from M21Vehicle_type WHERE M21Status='2' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboType
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 148
                '  .Rows.Band.Columns(1).Width = 160


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Private Sub frmVehicle_Cnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Brand()
        Call Load_Type()
        Call Load_Grid()
    End Sub

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M06Vehi_No) ##,M06Vehi_No as [Vehicle No],M06Brand as [Brand Name],M06V_Type as [V.Type]  from M06Vehicle_Master where M06Status='A' order by M06ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 30
            UltraGrid1.Rows.Band.Columns(1).Width = 120
            UltraGrid1.Rows.Band.Columns(2).Width = 160
            UltraGrid1.Rows.Band.Columns(3).Width = 160
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Trim(txtV_No.Text) <> "" Then
        Else
            MsgBox("Please enter the Vehicle No", MsgBoxStyle.Information, "Information .......")
            txtV_No.Focus()
            Exit Sub
        End If

        If Trim(cboBrand.Text) <> "" Then
        Else
            MsgBox("Please select the Brand Name", MsgBoxStyle.Information, "Information ........")
            cboBrand.ToggleDropdown()
            Exit Sub
        End If

        If Trim(cboType.Text) <> "" Then
        Else
            MsgBox("Please select the Vehicle Type", MsgBoxStyle.Information, "Information ........")
            cboType.ToggleDropdown()
            Exit Sub
        End If

        If Trim(txtEng_Capacity.Text) <> "" Then
        Else
            txtEng_Capacity.Text = "-"
        End If

        If Trim(txtRemak.Text) <> "" Then
        Else
            txtRemak.Text = "-"
        End If
        Call Save_Data()
    End Sub

    Function Save_Data()
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
        Try

            nvcFieldList1 = "SELECT * FROM M06Vehicle_Master WHERE M06Vehi_No='" & Trim(txtV_No.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M06Vehicle_Master set M06Brand='" & Trim(cboBrand.Text) & "',M06V_Type='" & Trim(cboType.Text) & "',M06Eng_Capacity='" & Trim(txtEng_Capacity.Text) & "',M06Remark='" & Trim(txtRemak.Text) & "',M06Status='A' where M06Vehi_No='" & Trim(txtV_No.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_VEHICLE','EDIT', '" & Now & "','" & strDisname & "','" & txtV_No.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
               
                nvcFieldList1 = "Insert Into M06Vehicle_Master(M06Vehi_No,M06Brand,M06V_Type,M06Eng_Capacity,M06Remark,M06Status)" & _
                                                                  " values('" & UCase(txtV_No.Text) & "','" & Trim(cboBrand.Text) & "', '" & Trim(cboType.Text) & "','" & UCase(txtEng_Capacity.Text) & "','" & Trim(txtRemak.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                                " values('NEW_VEHICLE','SAVE', '" & Now & "','" & strDisname & "','" & txtV_No.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If
            MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.Close()
            Call Clear_text()
            Call Load_Grid()
            txtV_No.Focus()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid1.ActiveRow.Index
        OPR0.Visible = True
        txtV_No.Text = Trim(UltraGrid1.Rows(_Row).Cells(1).Text)
        Call Search_Records()
    End Sub

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  *  from M06Vehicle_Master where M06Status='A' and M06Vehi_No='" & Trim(txtV_No.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboBrand.Text = Trim(M01.Tables(0).Rows(0)("M06Brand"))
                cboType.Text = Trim(M01.Tables(0).Rows(0)("M06V_Type"))
                txtEng_Capacity.Text = Trim(M01.Tables(0).Rows(0)("M06Eng_Capacity"))
                txtRemak.Text = Trim(M01.Tables(0).Rows(0)("M06Remark"))

            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim A As String

        A = MsgBox("Are you sure you want to deactive Electrician", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Deactive Electrician.........")
        If A = vbYes Then
            Call Deactive_Data()
        End If
    End Sub

    Function Deactive_Data()
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
        Try

            nvcFieldList1 = "SELECT * FROM M06Vehicle_Master WHERE M06Vehi_No='" & Trim(txtV_No.Text) & "' "
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                nvcFieldList1 = "update M06Vehicle_Master set M06Status='I' where M06Vehi_No='" & Trim(txtV_No.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                               " values('NEW_VEHICLE','DELETE', '" & Now & "','" & strDisname & "','" & txtV_No.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If
            MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.Close()


            Call Load_Grid()
            Call Clear_Text()
            txtV_No.Focus()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

End Class