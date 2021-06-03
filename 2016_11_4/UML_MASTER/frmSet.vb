Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmSet
    Dim c_dataCustomer1 As DataTable
    Dim _Product_Category As String
    Dim _ItemCode As String

    Private Sub frmSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Grid()
        Call Load_Combo()
        Call Load_Combo_Category()
        txtAvg_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRetail.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Function Load_Grid()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_SetCreation
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M14Item_Name as [Item Name] from M14Product_Item where M14Status='A' order by M14Item_Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
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

    Function Load_Combo_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M13Name as [##] from M13Product_Category  "
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

    Private Sub cboName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboName.AfterCloseUp
        Call Search_Category()
        ' Call Search_Records()
    End Sub

    Private Sub cboName_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboName.InitializeLayout

    End Sub

    Private Sub cboName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboName.KeyUp
        If e.KeyCode = 13 Then
            txtAvg_Cost.Focus()
        End If
    End Sub

    Private Sub txtAvg_Cost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAvg_Cost.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtAvg_Cost.Text) Then
                Value = txtAvg_Cost.Text
                txtAvg_Cost.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
            End If
            txtRetail.Focus()
        End If
    End Sub

    Private Sub txtRetail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRetail.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtRetail.Text) Then
                Value = txtRetail.Text
                txtRetail.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
            End If
            cboItem.ToggleDropdown()
        End If
    End Sub

    Private Sub cboItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItem.KeyUp
        If e.KeyCode = 13 Then
            txtQty.Focus()
        End If
    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Try
            If e.KeyCode = 13 Then
                If txtQty.Text <> "" Then
                    If Search_Item() = True Then
                        Dim newRow As DataRow = c_dataCustomer1.NewRow
                        newRow("Product Code") = _ItemCode
                        newRow("Product Name") = cboItem.Text
                        newRow("QTY") = txtQty.Text
                        newRow("Qty") = txtQty.Text
                        c_dataCustomer1.Rows.Add(newRow)
                        cboItem.Text = ""
                        txtQty.Text = ""
                        cboItem.ToggleDropdown()

                    End If
                End If
            End If
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Function Search_Item() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            'GSA
            Sql = "select * from M14Product_Item where M14Item_Name='" & cboItem.Text & "' and M14Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Item = True
                _ItemCode = Trim(M01.Tables(0).Rows(0)("M14Item_Code"))
            End If

            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.CLOSE()
            End If
        End Try
    End Function


    Function Search_Category() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            'GSA
            Sql = "select * from M13Product_Category where M13Name='" & cboName.Text & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Category = True
                _Product_Category = Trim(M01.Tables(0).Rows(0)("M13Cat_Code"))
            End If

            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.CLOSE()
            End If
        End Try
    End Function

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
        Dim i As Integer

        Try
            If txtAvg_Cost.Text <> "" Then
                If IsNumeric(txtAvg_Cost.Text) Then
                Else
                    MsgBox("Please enter the correct Cost", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    txtAvg_Cost.Focus()
                    Exit Sub
                End If
            Else
                txtAvg_Cost.Text = "0"
            End If

            If txtRetail.Text <> "" Then
                If IsNumeric(txtRetail.Text) Then
                Else
                    MsgBox("Please enter the correct Retail Price", MsgBoxStyle.Information, "Information ......")
                    connection.Close()
                    txtRetail.Focus()
                    Exit Sub
                End If
            Else
                txtRetail.Text = "0"
            End If

            If Search_Category() = True Then
            Else
                MsgBox("Please select the product set", MsgBoxStyle.Information, "Information .....")
                cboName.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If


            If UltraGrid1.Rows.Count > 0 Then
            Else
                MsgBox("Please enter the product items", MsgBoxStyle.Information, "Information .......")
                cboItem.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If

            nvcFieldList1 = "select * from M15Product_Set where M15Cat_Code='" & cboName.Text & "'"
            t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(t01) Then
                nvcFieldList1 = "DELETE FROM M16Item_for_Set WHERE M16Product_Code='" & _Product_Category & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    nvcFieldList1 = "Insert Into M16Item_for_Set(M16Product_Code,M16Item_Code,M16Qty,M16Status,M16User)" & _
                                                                  " values('" & _Product_Category & "', '" & UltraGrid1.Rows(i).Cells(0).Value & "','" & UltraGrid1.Rows(i).Cells(2).Value & "','A','" & strDisname & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    i = i + 1
                Next

                nvcFieldList1 = "UPDATE M15Product_Set SET M15Cost_Price='" & txtAvg_Cost.Text & "',M15Retail='" & txtRetail.Text & "',M15Status='A' WHERE M15Cat_Code='" & _Product_Category & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            Else
                nvcFieldList1 = "Insert Into M15Product_Set(M15Cat_Code,M15Cost_Price,M15Retail,M15Status,M15User)" & _
                                                                   " values('" & _Product_Category & "', '" & Trim(txtAvg_Cost.Text) & "','" & txtRetail.Text & "','A','" & strDisname & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    nvcFieldList1 = "Insert Into M16Item_for_Set(M16Product_Code,M16Item_Code,M16Qty,M16Status,M16User)" & _
                                                                  " values('" & _Product_Category & "', '" & UltraGrid1.Rows(i).Cells(0).Value & "','" & UltraGrid1.Rows(i).Cells(2).Value & "','A','" & strDisname & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    i = i + 1
                Next

                nvcFieldList1 = "Insert Into S02Set_Stock(S02Tr_Type,S02Date,S02Pr_Code,S02Qty,S02Location,S02Status,S02User,S02Product_Status,S02Remark)" & _
                                                           " values('OB', '" & Today & "','" & _Product_Category & "','0','MS','A','" & strDisname & "','GOOD','-')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            connection.Close()
            Call Clear_Text()
            cboName.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Function Clear_Text()
        Call Load_Grid()
        Me.cboItem.Text = ""
        Me.txtQty.Text = ""
        Me.cboName.Text = ""
        Me.txtAvg_Cost.Text = ""
        Me.txtRetail.Text = ""
    End Function


    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Dim i As Integer

        Try
            'GSA
            Sql = "select * from M15Product_Set inner join M13Product_Category on m15cat_code=M13Cat_Code where M13Name='" & cboName.Text & "'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                _Product_Category = Trim(M01.Tables(0).Rows(0)("M13Cat_Code"))
                Value = Trim(M01.Tables(0).Rows(0)("M15cost_price"))
                txtAvg_Cost.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                Value = Trim(M01.Tables(0).Rows(0)("M15Retail"))
                txtRetail.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)

            End If

            Call Load_Grid()
            i = 0
            Sql = "select * from M16Item_for_Set inner join M14Product_Item on M16Item_Code=M14Item_Code where M16Product_Code='" & _Product_Category & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow2 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Product Code") = Trim(M01.Tables(0).Rows(i)("M14Item_Code"))
                newRow("Product Name") = Trim(M01.Tables(0).Rows(i)("M14Item_Name"))
                newRow("QTY") = Trim(M01.Tables(0).Rows(i)("M16Qty"))

                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            con.CLOSE()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.CLOSE()
            End If
        End Try
    End Function

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Call Search_Records()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean
        Dim T01 As DataSet

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim A As String
        Try
            A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information .......")
            If A = vbYes Then

                nvcFieldList1 = "SELECT * FROM M13Product_Category WHERE M13Name='" & cboName.Text & "'"
                T01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(T01) Then
                    _Product_Category = Trim(T01.Tables(0).Rows(0)("M13Cat_Code"))
                End If
                nvcFieldList1 = "UPDATE M15Product_Set SET M15Status='I' WHERE M15Cat_Code='" & _Product_Category & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE M16Item_for_Set SET M16Status='I' WHERE M16Product_Code='" & _Product_Category & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Record deleted successfully", MsgBoxStyle.Information, "Information .....")
                transaction.Commit()
            End If
            connection.Close()
            Call Clear_Text()
            Call Load_Grid()
            cboName.ToggleDropdown()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Clear_Text()
        Call Load_Grid()
    End Sub
End Class