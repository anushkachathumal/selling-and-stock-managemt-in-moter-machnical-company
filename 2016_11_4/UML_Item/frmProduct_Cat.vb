Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmProduct_Cat

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        txtCode.Text = ""
        txtName.Text = ""
        Call Load_ItemCode()
    End Sub

    Function Load_ItemCode()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='CAT' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01NO") >= 1 And M01.Tables(0).Rows(0)("P01NO") < 10 Then
                    txtCode.Text = "CAT-0" & M01.Tables(0).Rows(0)("P01NO")
                Else
                    txtCode.Text = "CAT-" & M01.Tables(0).Rows(0)("P01NO")
                End If
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub frmProduct_Cat_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_ItemCode()
        txtCode.ReadOnly = True
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call LoadGride()
    End Sub

    Function LoadGride()
        'Load Color data to gride
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select M13Cat_Code as [##], M13Name as [Category Name] from M13Product_Category "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 220
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            

            End With
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try

    End Function

    Function LoadGride1()
        'Load Color data to gride
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select M13Cat_Code as [##], M13Name as [Category Name] from M13Product_Category where  M13Name like '" & txtName.Text & "%'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 220
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
              
                '  .DisplayLayout.Bands(0).Columns(2).Width = 90

            End With
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try

    End Function

    Private Sub txtName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.TextChanged
        Call LoadGride1()
    End Sub

    Private Sub txtName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyUp
        If e.KeyCode = 13 Then
            If txtName.Text <> "" Then
                cmdSave.Focus()
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
        Dim t01 As DataSet
        Try
          

            If txtName.Text <> "" Then
            Else
                MsgBox("Please select the Product Name", MsgBoxStyle.Information, "Information .....")
                connection.Close()

                Exit Sub
            End If

            nvcFieldList1 = "SELECT * FROM M13Product_Category WHERE M13Cat_Code='" & txtCode.Text & "' "
            t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(t01) Then
                nvcFieldList1 = "UPDATE M13Product_Category SET M13Name='" & Trim(txtName.Text) & "' WHERE M13Cat_Code='" & txtCode.Text & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "UPDATE P01PARAMETER SET P01NO=P01NO + " & 1 & " WHERE P01CODE='CAT'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M13Product_Category(M13Cat_Code,M13Name)" & _
                                                                 " values('" & Trim(txtCode.Text) & "', '" & Trim(txtName.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If

            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information .....")
            transaction.Commit()
            connection.Close()
            Me.txtName.Text = ""
            txtName.Focus()
            Call Load_ItemCode()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try

    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
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
            A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information .......")
            If A = vbYes Then
                nvcFieldList1 = "delete from  M13Product_Category WHERE M13Cat_Code='" & txtCode.Text & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                MsgBox("Record deleted successfully", MsgBoxStyle.Information, "Information .....")
                transaction.Commit()
            End If
            connection.Close()
            Me.txtName.Text = ""
            txtName.Focus()
            Call Load_ItemCode()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _rowindex As Integer

        _rowindex = UltraGrid1.ActiveRow.Index
        txtCode.Text = UltraGrid1.Rows(_rowindex).Cells(0).Text
        Call Search_Record()
    End Sub

    Function Search_Record()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select * from M13Product_Category where  M13Cat_Code = '" & txtCode.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                ' cboCategory.Text = Trim(M01.Tables(0).Rows(0)("M12Category"))
                txtName.Text = Trim(M01.Tables(0).Rows(0)("M13Name"))
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try
    End Function

 
End Class