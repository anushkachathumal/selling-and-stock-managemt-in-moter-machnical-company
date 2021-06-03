Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmUML_Employee
    Dim c_dataCustomer1 As DataTable


    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Function Load_Status()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Status as [##] from M03Gender order by M03Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboStatus
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 60
                ' .Rows.Band.Columns(1).Width = 180


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_NAME()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Employee_Name as [##] from M01Employee_Master order by M01Employee_Name "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 340
                ' .Rows.Band.Columns(1).Width = 180


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Department()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M02Dis as [##] from M02Department  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboDepartment
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 170
                ' .Rows.Band.Columns(1).Width = 180


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub frmUML_Employee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Status()
        Call Load_Department()
        Call LoadGride()
        Call Load_NAME()
        txtBasic.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
    End Sub

    Function Clear_Text()
        Me.txtBasic.Text = ""
        Me.txtCode.Text = ""
        Me.cboDepartment.Text = ""
        Me.cboName.Text = ""
        Me.cboStatus.Text = ""
        txtCode.Focus()
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Clear_Text()
    End Sub

    Function LoadGride()
        'Load Color data to gride
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Sql = "select M01Emp_No as [Emp No],M01Employee_Name as [Employee Name] from M01Employee_Master where M01Status='A'"
        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
        UltraGrid1.DataSource = M01
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 270
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            '  .DisplayLayout.Bands(0).Columns(2).Width = 90

        End With
        DBEngin.CloseConnection(con)
        con.ConnectionString = ""
    End Function

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = 13 Then
            If txtCode.Text <> "" Then
                Call SEARCH_RECORDS("1")
                cboStatus.ToggleDropdown()
            End If
        End If
    End Sub

    Private Sub cboStatus_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboStatus.KeyUp
        If e.KeyCode = 13 Then
            cboName.ToggleDropdown()
        End If
    End Sub

    Private Sub cboName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboName.AfterCloseUp
        Call SEARCH_RECORDS("2")
    End Sub

    Private Sub cboName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboName.KeyUp
        If e.KeyCode = 13 Then
            cboDepartment.ToggleDropdown()
        End If
    End Sub

    Private Sub cboDepartment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboDepartment.KeyUp
        If e.KeyCode = 13 Then
            txtBasic.Focus()
        End If
    End Sub

    Private Sub txtBasic_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBasic.KeyUp
        Dim Value As Double

        If e.KeyCode = 13 Then
            If txtBasic.Text <> "" Then
                If IsNumeric(txtBasic.Text) Then
                    Value = txtBasic.Text
                    txtBasic.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)

                End If

            End If
            cmdSave.Focus()
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
            If txtCode.Text <> "" Then
            Else
                MsgBox("Please enter the Emp No", MsgBoxStyle.Information, "Information .....")
                connection.Close()
                Exit Sub
            End If

            If cboDepartment.Text <> "" Then
            Else
                MsgBox("Please select the Department", MsgBoxStyle.Information, "Information .....")
                connection.Close()
                cboDepartment.ToggleDropdown()
                Exit Sub
            End If


            If cboStatus.Text <> "" Then
            Else
                MsgBox("Please select the status", MsgBoxStyle.Information, "Information .....")
                connection.Close()
                cboStatus.ToggleDropdown()
                Exit Sub
            End If

            If cboName.Text <> "" Then
            Else
                MsgBox("Please enter the Employee Name", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                cboName.ToggleDropdown()
                Exit Sub
            End If

            If txtBasic.Text <> "" Then
                If IsNumeric(txtBasic.Text) Then
                Else
                    MsgBox("Please enter the correct Basic Salary", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    txtBasic.Focus()
                    Exit Sub
                End If
            Else
                txtBasic.Text = "0"
            End If

            nvcFieldList1 = "SELECT * FROM M01Employee_Master WHERE M01Emp_No='" & Trim(txtCode.Text) & "' AND M01Status='A'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                nvcFieldList1 = "UPDATE M01Employee_Master SET M01Gender='" & Trim(cboStatus.Text) & "',M01Employee_Name='" & Trim(cboName.Text) & "',M01Basic_Salary='" & txtBasic.Text & "',M01Depatment='" & Trim(cboDepartment.Text) & "' WHERE M01Emp_No='" & Trim(txtCode.Text) & "' AND M01Status='A'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "Insert Into M01Employee_Master(M01Emp_No,M01Gender,M01Employee_Name,M01Basic_Salary,M01Depatment,M01Status,M01User,M01Time)" & _
                                                           " values('" & Trim(txtCode.Text) & "', '" & Trim(cboStatus.Text) & "','" & Trim(cboName.Text) & "','" & txtBasic.Text & "','" & Trim(cboDepartment.Text) & "','A','" & strDisname & "','" & Now & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information .....")
            transaction.Commit()
            Call Clear_Text()
            Call LoadGride()
            Call Load_NAME()
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub


    Function SEARCH_RECORDS(ByVal STRCode As String)
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Try
            If STRCode = "1" Then
                Sql = "select * from M01Employee_Master WHERE M01Emp_No='" & Trim(txtCode.Text) & "' AND M01Status='A' "
            ElseIf STRCode = "2" Then
                Sql = "select * from M01Employee_Master WHERE M01Employee_Name='" & Trim(cboName.Text) & "' AND M01Status='A' "
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtCode.Text = Trim(M01.Tables(0).Rows(0)("M01Emp_No"))
                cboStatus.Text = Trim(M01.Tables(0).Rows(0)("M01Gender"))
                cboName.Text = Trim(M01.Tables(0).Rows(0)("M01Employee_Name"))
                cboDepartment.Text = Trim(M01.Tables(0).Rows(0)("M01Depatment"))
                Value = Trim(M01.Tables(0).Rows(0)("M01Basic_Salary"))
                txtBasic.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

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
            A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information ......")
            If A = vbYes Then

                nvcFieldList1 = "SELECT * FROM M01Employee_Master WHERE M01Emp_No='" & Trim(txtCode.Text) & "' AND M01Status='A'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "UPDATE M01Employee_Master SET M01Status='I' WHERE M01Emp_No='" & Trim(txtCode.Text) & "' AND M01Status='A'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                   
                End If

                MsgBox("Record update successfully", MsgBoxStyle.Information, "Information .....")
                transaction.Commit()
                Call Clear_Text()
                Call LoadGride()
                Call Load_NAME()
            End If
            connection.Close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub
End Class