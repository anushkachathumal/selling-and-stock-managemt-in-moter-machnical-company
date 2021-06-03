Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmDiscount
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _CountryCode As String
    Dim _Comcode As String

    Function Load_Gride_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name] from M03Item_Master  order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = dsUser
            UltraGrid2.Rows.Band.Columns(0).Width = 130
            UltraGrid2.Rows.Band.Columns(1).Width = 370
            ' UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = Keys.Enter Then
            Call Search_Records()
            If Trim(txtCode.Text) <> "" Then
                txtDescription.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_Records()
            txtDescription.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR5.Visible = True
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False

        End If
    End Sub

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M27Item_Code as [Item Code],M03Item_Name as [Item Name],CAST(M27Discount AS DECIMAL(16,2)) as [Discount %] from M27Discount_Items inner join M03Item_Master on M03Item_Code=M27Item_Code  "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370
            UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim Value As Double

        Try
            Sql = "select * from M03Item_Master where M03Item_Code='" & Trim(txtCode.Text) & "' and m03Status='A' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                txtDescription.Text = dsUser.Tables(0).Rows(0)("M03Item_Name")
            End If

            Sql = "select * from M27Discount_Items where M27Item_Code='" & Trim(txtCode.Text) & "'"
            T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(T01) Then
                Value = T01.Tables(0).Rows(0)("M27Discount")
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = Keys.Enter Then
            txtDiscount.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtDiscount.Focus()
        End If
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0)
        ' OPR2.Enabled = False
        'OPR1.Enabled = False
        OPR0.Enabled = True
        ' OPR3.Enabled = False
        cmdAdd.Enabled = True
        cmdAdd.Focus()
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        Call Load_Gride()

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
            nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & txtCode.Text & "' and  m03status='A'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
            Else
                MsgBox("Please enter the correct Item Code", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtCode.Focus()
                Exit Sub
            End If

            '--------------------------------------------------------------------
            If Trim(txtDiscount.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the correct Discount%", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDiscount.Focus()
                    connection.Close()
                    Exit Sub
                End If
            End If

            If IsNumeric(txtDiscount.Text) Then
            Else
                result1 = MessageBox.Show("Please enter the correct Discount%", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDiscount.Focus()
                    connection.Close()
                    Exit Sub
                End If
            End If
            '--------------------------------------------------------------------
            nvcFieldList1 = "Insert Into M27Discount_Items(M27Item_Code,M27Discount)" & _
                                                          " values('" & (Trim(txtCode.Text)) & "', '" & (Trim(txtDiscount.Text)) & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            OPR0.Enabled = True
            ' OPR3.Enabled = False
            cmdAdd.Enabled = True
            txtCode.Focus()
            ' cmdSave.Enabled = False
            cmdDelete.Enabled = False
            Call Load_Gride()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
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
            nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & txtCode.Text & "' and  m03status='A'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
            Else
                MsgBox("Please enter the correct Item Code", MsgBoxStyle.Information, "Information ........")
                connection.Close()
                txtCode.Focus()
                Exit Sub
            End If

            '--------------------------------------------------------------------
            If Trim(txtDiscount.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the correct Discount%", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDiscount.Focus()
                    connection.Close()
                    Exit Sub
                End If
            End If

            If IsNumeric(txtDiscount.Text) Then
            Else
                result1 = MessageBox.Show("Please enter the correct Discount%", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDiscount.Focus()
                    connection.Close()
                    Exit Sub
                End If
            End If
            '--------------------------------------------------------------------
            nvcFieldList1 = "update M27Discount_Items set M27Discount='" & (Trim(txtDiscount.Text)) & "' where M27Item_Code='" & Trim(txtCode.Text) & "' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0)
            ' OPR2.Enabled = False
            'OPR1.Enabled = False
            ' OPR0.Enabled = False
            ' OPR3.Enabled = False
            OPR0.Enabled = True
            cmdAdd.Enabled = True
            txtCode.Focus()
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Call Load_Gride()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
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
                nvcFieldList = "delete from M27Discount_Items where M27Item_Code = '" & Trim(txtCode.Text) & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList)


            End If
            transaction.Commit()
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

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Private Sub frmDiscount_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Gride()
        Call Load_Gride_Item()
        txtDescription.ReadOnly = True
    End Sub

    Function Load_Gride1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Item_Name  like '%" & txtFind.Text & "%' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = dsUser
            UltraGrid2.Rows.Band.Columns(0).Width = 130
            UltraGrid2.Rows.Band.Columns(1).Width = 370
            UltraGrid2.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtFind_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFind.ValueChanged
        Call Load_Gride1()
    End Sub

    Private Sub UltraGrid2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid2.DoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer

        _Rowindex = UltraGrid2.ActiveRow.Index
        txtCode.Text = UltraGrid2.Rows(_Rowindex).Cells(0).Text
        Call Search_Records()
        OPR5.Visible = False
        txtFind.Text = ""
    End Sub

  
    Private Sub txtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDiscount.KeyUp
        If e.KeyCode = 13 Then
            cmdAdd.Focus()
        End If
    End Sub
End Class