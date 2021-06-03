Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmPr_Item

    Dim _BoxNo As String

    Private Sub frmPr_Item_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtArt.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCost1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCost2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRetail.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Grid()
        Call Load_Category()
        Call Load_Box()
    End Sub

    Function Load_Category()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M05Dis as [##] from M05Main_Production  "
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

    Function Load_Box()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M07Name as [##] from M07Product_Box where M07Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboBox
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

    Function Load_Grid()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select M06Category as [##],M06Art_No as [Art No], M06Name as [Production Name] from M06Production_Item where M06Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            With UltraGrid1
                .DisplayLayout.Bands(0).Columns(0).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(2).Width = 230
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False

            End With
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try

    End Function

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Private Sub txtArt_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtArt.KeyUp
        If e.KeyCode = 13 Then
            If txtArt.Text <> "" Then
                Call Search_Records()
                cboCategory.ToggleDropdown()
            End If
        End If
    End Sub


    Private Sub cboCategory_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCategory.KeyUp
        If e.KeyCode = 13 Then
            cboBox.ToggleDropdown()
        End If
    End Sub

    Private Sub cboName_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboName.InitializeLayout

    End Sub

    Private Sub cboName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboName.KeyUp
        If e.KeyCode = 13 Then
            txtCost1.Focus()
        End If
    End Sub

    Private Sub txtCost1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCost1.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtCost1.Text) Then
                Value = txtCost1.Text
                txtCost1.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
            End If

            txtCost2.Focus()
        End If
    End Sub

    Private Sub txtCost2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCost2.KeyUp
        Dim Value As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtCost2.Text) Then
                Value = txtCost2.Text
                txtCost2.Text = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
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

            cmdSave.Focus()
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
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
        Dim T02 As DataSet
        Dim T03 As DataSet
        Dim I As Integer
        Dim x As Integer


        Try
            If txtCost1.Text <> "" Then
            Else
                txtCost1.Text = "0"
            End If

            If IsNumeric(txtCost1.Text) Then
            Else
                MsgBox("Please enter the correct Cost Price", MsgBoxStyle.Information, "Information .....")
                connection.Close()
                txtCost1.Focus()
                Exit Sub
            End If

            If IsNumeric(txtCost2.Text) Then
            Else
                MsgBox("Please enter the correct Sub Manufacture Cost Price", MsgBoxStyle.Information, "Information .....")
                txtCost2.Focus()
                connection.Close()
                Exit Sub
            End If


            If IsNumeric(txtRetail.Text) Then
            Else
                MsgBox("Please enter the correct Retail Price", MsgBoxStyle.Information, "Information .....")
                txtRetail.Focus()
                connection.Close()
                Exit Sub
            End If

            If cboCategory.Text <> "" Then
            Else
                MsgBox("Please select the Category", MsgBoxStyle.Information, "Information .....")
                cboCategory.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If

            If cboName.Text <> "" Then
            Else
                MsgBox("Please enter the Product Name", MsgBoxStyle.Information, "Information .....")
                cboName.Focus()
                connection.Close()
                Exit Sub
            End If

            If txtArt.Text <> "" Then
            Else
                MsgBox("Please enter the Art No", MsgBoxStyle.Information, "Information .....")
                txtArt.Focus()
                connection.Close()
                Exit Sub
            End If

            If Search_Box() = True Then
            Else
                MsgBox("Please select the Box Name", MsgBoxStyle.Information, "Information .....")
                cboBox.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If



            nvcFieldList1 = "select * from M06Production_Item where M06Art_No='" & Trim(txtArt.Text) & "'"
            t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(t01) Then
                nvcFieldList1 = "UPDATE M06Production_Item SET M06Name='" & Trim(cboName.Text) & "',M06Category='" & cboCategory.Text & "',M06Reorder='0',M06Cost='" & txtCost1.Text & "',M06Sub_Price='" & txtCost2.Text & "',M06Retail='" & txtRetail.Text & "',M06Status='A',M06Box='" & _BoxNo & "' WHERE M06Art_No='" & Trim(txtArt.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            Else

                nvcFieldList1 = "Insert Into M06Production_Item(M06Art_No,M06Name,M06Category,M06Reorder,M06Cost,M06Retail,M06Sub_Price,M06User,M06Status,M06Box)" & _
                                                                    " values('" & Trim(txtArt.Text) & "', '" & Trim(cboName.Text) & "','" & cboCategory.Text & "','0','" & txtCost1.Text & "','" & txtRetail.Text & "','" & txtCost2.Text & "','" & strDisname & "','A','" & _BoxNo & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "SELECT * FROM M04Color "
                T02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                I = 0
                For Each DTRow1 As DataRow In T02.Tables(0).Rows
                    If Trim(cboCategory.Text) = "BOYES" Then
                        x = 0
                        nvcFieldList1 = "select * from M13Product_Size where M13category='" & Trim(cboCategory.Text) & "'"
                        T03 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        For Each DTRow2 As DataRow In T03.Tables(0).Rows
                            nvcFieldList1 = "Insert Into S02Production_Stock(S02Tr_Type,S02Date,S02Pr_Code,S02Colour,S02Size,S02Qty,S02Location,S02Remark,S02Status,S02User,S02Product_Status)" & _
                                                                   " values('OB', '" & Today & "','" & Trim(txtArt.Text) & "','" & Trim(T02.Tables(0).Rows(I)("M04Dis")) & "','" & Trim(T03.Tables(0).Rows(x)("M13Size")) & "','0','MS','-','A','" & strDisname & "','-')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                            x = x + 1
                        Next
                    ElseIf Trim(cboCategory.Text) = "GENTS" Then
                        x = 0
                        nvcFieldList1 = "select * from M13Product_Size where M13category='" & Trim(cboCategory.Text) & "'"
                        T03 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                        For Each DTRow2 As DataRow In T03.Tables(0).Rows
                            nvcFieldList1 = "Insert Into S02Production_Stock(S02Tr_Type,S02Date,S02Pr_Code,S02Colour,S02Size,S02Qty,S02Location,S02Remark,S02Status,S02User,S02Product_Status)" & _
                                                                   " values('OB', '" & Today & "','" & Trim(txtArt.Text) & "','" & Trim(T02.Tables(0).Rows(I)("M04Dis")) & "','" & Trim(T03.Tables(0).Rows(x)("M13Size")) & "','0','MS','-','A','" & strDisname & "','-')"
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                            x = x + 1
                        Next
                    End If
                    I = I + 1
                Next
            End If

            MsgBox("Record Update successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            connection.Close()
            Call Clear_Text()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Function Clear_Text()
        Me.cboCategory.Text = ""
        Me.txtArt.Text = ""
        Me.txtCost1.Text = ""
        Me.txtCost2.Text = ""
        Me.txtRetail.Text = ""
        Me.cboName.Text = ""
        Me.cboBox.Text = ""
        Call Load_Grid()
        txtArt.Focus()
    End Function

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Clear_Text()
    End Sub

    Private Sub cboBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboBox.KeyUp
        If e.KeyCode = 13 Then
            cboName.ToggleDropdown()
        End If
    End Sub


    Function Search_Box() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Try
            Sql = "select * from M07Product_Box where M07Status='A' and M07Name='" & Trim(cboBox.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Box = True
                _BoxNo = Trim(M01.Tables(0).Rows(0)("M07Ref"))
            End If
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try
    End Function

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim value As Double
        Try
            Sql = "select * from M06Production_Item inner join M07Product_Box on m06box=M07Ref where M06Art_No='" & Trim(txtArt.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                With M01
                    cboCategory.Text = Trim(.Tables(0).Rows(0)("M06Category"))
                    cboName.Text = Trim(.Tables(0).Rows(0)("M06Name"))
                    cboBox.Text = Trim(.Tables(0).Rows(0)("M07Name"))
                    value = Trim(.Tables(0).Rows(0)("M06Cost"))
                    txtCost1.Text = value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                    value = Trim(.Tables(0).Rows(0)("M06Sub_Price"))
                    txtCost2.Text = value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                    value = Trim(.Tables(0).Rows(0)("M06Retail"))
                    txtRetail.Text = value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                End With
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
        Dim _Rowindex As Integer

        _Rowindex = UltraGrid1.ActiveRow.Index
        txtArt.Text = UltraGrid1.Rows(_Rowindex).Cells(1).Text
        Call Search_Records()
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
            A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete Records ........")
            If A = vbYes Then
                nvcFieldList1 = "UPDATE M06Production_Item SET M06Status='I' WHERE M06Art_No='" & txtArt.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S02Production_Stock SET S02Status='I' WHERE S02Pr_Code='" & txtArt.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, "Information ......")
                transaction.Commit()
            End If
            connection.Close()
            Call Clear_Text()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub
End Class