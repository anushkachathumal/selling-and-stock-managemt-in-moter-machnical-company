Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmSetup_Planner
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _UnitCode As Integer


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Unit()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet

        Try
           

            Sql = "select M13Merchant as [Merchant] from M13Biz_Unit "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboUnit.DataSource = M01
                cboUnit.Rows.Band.Columns(0).Width = 175
                ' cboGroup.Rows.Band.Columns(1).Width = 270
                'cboGroup.Rows.Band.Columns(2).Width = 170
                'cboGroup.Rows.Band.Columns(3).Width = 130
            End If


            Sql = "select M29Name as [Planner] from M29Palnner "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboMerchant.DataSource = M01
                cboMerchant.Rows.Band.Columns(0).Width = 175
                ' cboGroup.Rows.Band.Columns(1).Width = 270
                'cboGroup.Rows.Band.Columns(2).Width = 170
                'cboGroup.Rows.Band.Columns(3).Width = 130
            End If


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub frmSetup_Planner_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Unit()
        Call LoadGride()
    End Sub

    Function LoadGride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M29Merchant as [Merchant],M29Name as [Planner's Name] from M29Palnner"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 270
            ' UltraGrid1.Rows.Band.Columns(2).Width = 110
            'UltraGrid1.Rows.Band.Columns(3).Width = 140
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function


    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        cboMerchant.Text = ""
        cboUnit.Text = ""
        Call Load_Unit()
    End Sub

    Function Search_Unit() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T03 As DataSet

        Try
            Sql = "SELECT * from M13Biz_Unit WHERE M13Merchant='" & Trim(cboUnit.Text) & "' "
            T03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(T03) Then
                Search_Unit = True
                ' _UnitCode = T03.Tables(0).Rows(0)("M14Code")
            Else
                Search_Unit = False
            End If

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
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
            If Search_Unit() = True Then
            Else
                MsgBox("Please select the Merchant", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If

            If cboMerchant.Text <> "" Then
            Else
                MsgBox("Please select the Planner", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If

            nvcFieldList1 = "select * from M29Palnner where M29Merchant='" & Trim(cboUnit.Text) & "'"
            dsUser = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(dsUser) Then
                nvcFieldList1 = "Update M29Palnner  set M29Merchant=" & cboUnit.Text & " where M29Merchant='" & Trim(cboUnit.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "Insert Into M29Palnner(M29Merchant,M29Name)" & _
                                                              " values('" & cboUnit.Text & "', '" & Trim(cboMerchant.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            transaction.Commit()
            Call LoadGride()
            Call Load_Unit()
            cboMerchant.Text = ""
            cboUnit.Text = ""
            cboUnit.ToggleDropdown()
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
        Dim A As String

        Try
            A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete Records .......")
            If A = vbYes Then
                nvcFieldList1 = "delete from M29Palnner  where M29Name='" & Trim(cboMerchant.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            transaction.Commit()
            Call LoadGride()
            Call Load_Unit()
            cboMerchant.Text = ""
            cboUnit.Text = ""
            cboUnit.ToggleDropdown()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub
End Class