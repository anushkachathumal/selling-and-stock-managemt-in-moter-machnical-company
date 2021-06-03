Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Public Class frmProduct_Item
    Dim c_dataCustomer1 As DataTable

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Private Sub frmProduct_Item_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'txtAvg_Cost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtRetail.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCode.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCode.ReadOnly = True
        Call Load_Parameter()
        Call Load_Combo()
        Call Load_Gride()
    End Sub

    Function Clear_Text()
        Me.txtCode.Text = ""
        Me.cboName.Text = ""
        Call Load_Parameter()
        Call Load_Combo()
        cboName.ToggleDropdown()
        Call Load_Gride()
    End Function

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        Dim dsUSER As DataSet

        con = DBEngin.GetConnection()

        Try
            Sql = "select M07Ref as [Ref Code],M07Name as [Box Name] from M07Product_Box WHERE M07Status='A' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 240
            'UltraGrid1.Rows.Band.Columns(2).Width = 90
            'UltraGrid1.Rows.Band.Columns(3).Width = 90
            'UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride1()
        Dim Sql As String
        Dim con = New SqlConnection()
        Dim dsUSER As DataSet

        con = DBEngin.GetConnection()

        Try
            Sql = "select M07Ref as [Ref Code],M07Name as [Box Name] from M07Product_Box WHERE M07Status='A' and M07Name like '%" & cboName.Text & "%'"
            dsUSER = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUSER
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 240
            'UltraGrid1.Rows.Band.Columns(2).Width = 90
            'UltraGrid1.Rows.Band.Columns(3).Width = 90
            'UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function
    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Call Clear_Text()
    End Sub

    Function Load_Parameter()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='PR'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01NO") <= 10 Then
                    txtCode.Text = "PR00" & M01.Tables(0).Rows(0)("P01NO")
                ElseIf M01.Tables(0).Rows(0)("P01NO") > 10 And M01.Tables(0).Rows(0)("P01NO") <= 100 Then
                    txtCode.Text = "PR0" & M01.Tables(0).Rows(0)("P01NO")
                Else
                    txtCode.Text = "PR" & M01.Tables(0).Rows(0)("P01NO")
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

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M07Name as [Item Name] from M07Product_Box where M07Status='A' order by M07Name "
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

    Function SEARCH_RECORD()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim VALUE As Double

        Try
            Sql = "select * from M07Product_Box where M07name='" & cboName.Text & "' AND M07Status='A'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                txtCode.Text = M01.Tables(0).Rows(0)("M07Ref")
                'VALUE = M01.Tables(0).Rows(0)("M14Cost")
                'txtAvg_Cost.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtAvg_Cost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                'VALUE = M01.Tables(0).Rows(0)("M14Retail")
                'txtRetail.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
            End If




            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                con.close()
            End If
        End Try
    End Function

    Function SEARCH_RECORD1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim VALUE As Double

        Try
            Sql = "select * from M07Product_Box where M07ref='" & txtCode.Text & "' AND M07Status='A'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                cboName.Text = M01.Tables(0).Rows(0)("M07name")
                'VALUE = M01.Tables(0).Rows(0)("M14Cost")
                'txtAvg_Cost.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtAvg_Cost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))

                'VALUE = M01.Tables(0).Rows(0)("M14Retail")
                'txtRetail.Text = (VALUE.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", VALUE))
            End If




            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                con.close()
            End If
        End Try
    End Function

    Private Sub cboName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboName.AfterCloseUp
        Call SEARCH_RECORD()
    End Sub
    Private Sub cboName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboName.KeyUp
        If e.KeyCode = 13 Then
            If cboName.Text <> "" Then
                cmdSave.Focus()
            End If
        End If
    End Sub

    Private Sub cboName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboName.TextChanged
        Call Load_Gride1()
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
            If cboName.Text <> "" Then
            Else
                MsgBox("Please enter the correct Product Name", MsgBoxStyle.Information, "Information ......")
                cboName.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If

           

            nvcFieldList1 = "SELECT * FROM M07Product_Box WHERE M07ref='" & txtCode.Text & "' AND M07Status='A'"
            t01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(t01) Then
                nvcFieldList1 = "UPDATE M07Product_Box SET M07name='" & Trim(cboName.Text) & "' WHERE M07ref='" & txtCode.Text & "' AND M07Status='A' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            Else
                nvcFieldList1 = "UPDATE P01PARAMETER SET P01NO=P01NO + " & 1 & " WHERE P01CODE='PR'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into M07Product_Box(M07ref,M07name,M07Status)" & _
                                                                 " values('" & Trim(txtCode.Text) & "', '" & Trim(cboName.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S03Box_Stock(S03Tr_Type,S03Date,S03Box_No,S03Qty,S03Loc_Type,S03Status,S03User)" & _
                                                            " values('OB', '" & Today & "','" & txtCode.Text & "','0','MS','A','" & strDisname & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            MsgBox("Record update successfully", MsgBoxStyle.Information, "Information .....")
            transaction.Commit()
            Call Clear_Text()
            connection.Close()
            cboName.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub



    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _RowCount As Integer

        _RowCount = UltraGrid1.ActiveRow.Index
        txtCode.Text = UltraGrid1.Rows(_RowCount).Cells(0).Text
        Call SEARCH_RECORD1()
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
        Dim A As String
        Try
            A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information .......")
            If A = vbYes Then
                nvcFieldList1 = "UPDATE M07Product_Box SET M07Status='I' WHERE M07ref='" & txtCode.Text & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                MsgBox("Record deleted successfully", MsgBoxStyle.Information, "Information .....")
                transaction.Commit()
            End If
            connection.Close()
            Call Clear_Text()
            Call Load_Gride()
            cboName.ToggleDropdown()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub
End Class