Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmRow_Material
    Dim c_dataCustomer1 As DataTable
    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Function Load_ItemCode()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='IT' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01NO") >= 1 And M01.Tables(0).Rows(0)("P01NO") < 10 Then
                    txtCode.Text = "RI-0" & M01.Tables(0).Rows(0)("P01NO")
                Else
                    txtCode.Text = "RI-" & M01.Tables(0).Rows(0)("P01NO")
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

    Private Sub frmRow_Material_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_ItemCode()
        txtCode.ReadOnly = True
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Category()
        Call LoadGride()

    End Sub


    Function Load_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M08Name as [##] from M08Main_Category order by M08Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCategory
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 140
                ' .Rows.Band.Columns(1).Width = 180


            End With

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        txtName.Text = ""
        cboCategory.Text = ""
        Call Load_ItemCode()
    End Sub



    Private Sub cboCategory_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCategory.KeyUp
        If e.KeyCode = 13 Then
            If cboCategory.Text <> "" Then
                txtName.Focus()
            End If
        End If
    End Sub

    Private Sub txtName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyUp
        If e.KeyCode = 13 Then
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
        Dim t01 As DataSet
        Try
            If cboCategory.Text <> "" Then
            Else
                MsgBox("Please select the category", MsgBoxStyle.Information, "Information .....")
                connection.Close()
                cboCategory.ToggleDropdown()
                Exit Sub
            End If

            If txtName.Text <> "" Then
            Else
                MsgBox("Please select the Item Name", MsgBoxStyle.Information, "Information .....")
                connection.Close()
                txtName.Focus()
                Exit Sub
            End If

            nvcFieldList1 = "SELECT * FROM M12Row_Material WHERE M12Item_Code='" & txtCode.Text & "' "
            t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(t01) Then
                nvcFieldList1 = "UPDATE M12Row_Material SET M12Name='" & Trim(txtName.Text) & "',M12Category='" & cboCategory.Text & "' WHERE M12Item_Code='" & txtCode.Text & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "UPDATE P01PARAMETER SET P01NO=P01NO + " & 1 & " WHERE P01CODE='IT'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M12Row_Material(M12Item_Code,M12Name,M12Status,M12Category)" & _
                                                                 " values('" & Trim(txtCode.Text) & "', '" & Trim(txtName.Text) & "','A','" & cboCategory.Text & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            End If

            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information .....")
            transaction.Commit()
            connection.Close()
            Me.txtName.Text = ""
            Me.cboCategory.Text = ""
            cboCategory.ToggleDropdown()
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
                nvcFieldList1 = "UPDATE M12Row_Material SET M12Status='I' WHERE M12Item_Code='" & txtCode.Text & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                MsgBox("Record deleted successfully", MsgBoxStyle.Information, "Information .....")
                transaction.Commit()
            End If
            connection.Close()
            Me.txtName.Text = ""
            Me.cboCategory.Text = ""
            cboCategory.ToggleDropdown()
            Call Load_ItemCode()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub


    Function LoadGride()
        'Load Color data to gride
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select M12Category as [Category], M12Item_Code as [Item Code],M12Name as [Item Name] from M12Row_Material where M12Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 180
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
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

    Function LoadGride1()
        'Load Color data to gride
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select M12Category as [Category], M12Item_Code as [Item Code],M12Name as [Item Name] from M12Row_Material where M12Status='A' and M12Name like '" & txtName.Text & "%'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 180
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
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

    Function Search_Record()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select * from M12Row_Material where M12Status='A' and M12Item_Code = '" & txtCode.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboCategory.Text = Trim(M01.Tables(0).Rows(0)("M12Category"))
                txtName.Text = Trim(M01.Tables(0).Rows(0)("M12Name"))
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try
    End Function

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _rowindex As Integer

        _rowindex = UltraGrid1.ActiveRow.Index
        txtCode.Text = UltraGrid1.Rows(_rowindex).Cells(1).Text
        Call Search_Record()
    End Sub

   
   

  
End Class