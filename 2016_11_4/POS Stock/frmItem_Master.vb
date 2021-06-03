Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmItem_Master
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Dim _Supcode As String
    Dim _Loc As String

    Function COmmision_rate()
        On Error Resume Next
        Dim Vale As Double
        If IsNumeric(txtRetail.Text) And IsNumeric(txtCost.Text) Then
            Vale = CDbl(txtCost.Text) / CDbl(txtRetail.Text)
            Vale = Vale * 100
            Vale = 100 - Vale
            txtComm.Text = (Vale.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtComm.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Vale))
        End If
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

        End If
    End Sub

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M11Part_No,M07Sup_Name,M09Cat_Name,M11Part_Name,CAST(M11Retail_price AS DECIMAL(16,2)) as Rate from M11Product_Item inner join M07Supplier on M07Sup_Code=M11Supp_Code inner join view_category on M09Cat_Code=M11Cat_Code where M11Status='A' order by M11id"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 110
            UltraGrid1.Rows.Band.Columns(1).Width = 120
            UltraGrid1.Rows.Band.Columns(2).Width = 120
            UltraGrid1.Rows.Band.Columns(3).Width = 240
            UltraGrid1.Rows.Band.Columns(4).Width = 90
            '  UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CAST(M03Cost_Price AS DECIMAL(16,2)) as [Cost Price],CAST(M03Retail_Price AS DECIMAL(16,2)) as [Retail Price] from M03Item_Master where m03Status='A' and M03Item_Name like '%" & txtDescription.Text & "%' and M03Com_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 130
            UltraGrid1.Rows.Band.Columns(1).Width = 370
            UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Category() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M02Category where M02Cat_Name='" & Trim(cboMain.Text) & "' and M02Com_Code='" & _Comcode & "' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                _Category = dsUser.Tables(0).Rows(0)("M02Cat_Code")
                Search_Category = True
            Else
                Search_Category = False
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M04Location where M04Loc_Name='" & Trim(cboLocation.Text) & "' and M04Com_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                _Loc = dsUser.Tables(0).Rows(0)("M04Loc_Code")
                Search_Location = True
            Else
                Search_Location = False
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function


    Function Search_Supplier() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M01Account_Master where M01Acc_Name='" & Trim(cboSupplier.Text) & "'  and M01Acc_Type='SP' and M01Com_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                _Supcode = dsUser.Tables(0).Rows(0)("M01Acc_Code")
                Search_Supplier = True
            Else
                Search_Supplier = False
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double
        Dim _FromDate As Date
        Dim M01 As DataSet
        Dim i As Integer
        Dim M02 As DataSet
        Dim M03 As DataSet
        Dim M04 As DataSet

        Try
            Sql = "select * from M03Item_Master inner join M02Category on M03Cat_Code=M02Cat_Code inner join M01Account_Master on M01Acc_Code=M03Supplier inner join M04Location on M03Location=M04Loc_Code where M03Item_Code='" & Trim(txtCode.Text) & "'  and M01Acc_Type='SP' and M03Com_Code='" & _Comcode & "' and M03Status='A'"
            M04 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M04) Then
                txtDescription.Text = M04.Tables(0).Rows(0)("M03Item_Name")
                cboSupplier.Text = M04.Tables(0).Rows(0)("M01Acc_Name")
                cboEx_Date.Text = M04.Tables(0).Rows(0)("M03ExPair")
                txtReorder.Text = M04.Tables(0).Rows(0)("M03Reorder")
                cboMain.Text = M04.Tables(0).Rows(0)("M02Cat_Name")
                cboLocation.Text = M04.Tables(0).Rows(0)("M04Loc_Name")

                Value = M04.Tables(0).Rows(0)("M03Cost_Price")
                txtCost.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M04.Tables(0).Rows(0)("M03Retail_Price")
                txtRetail.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M04.Tables(0).Rows(0)("M03MRP")
                txtMrp.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMrp.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Call COmmision_rate()
                cmdAdd.Enabled = False
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True

                Call Load_Gride2()
                Sql = "select * from M04Location "
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                i = 0
                For Each DTRow2 As DataRow In M01.Tables(0).Rows
                    Sql = "select * from S01Stock_Balance where S01Item_Code='" & Trim(txtCode.Text) & "' and S01Trans_Type='OB' and S01Loc_Code='" & Trim(M01.Tables(0).Rows(i)("M04Loc_Code")) & "' "
                    M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(M02) Then
                        Sql = "select sum(S01Qty) as Qty from S01Stock_Balance where S01Loc_Code='" & Trim(M01.Tables(0).Rows(i)("M04Loc_Code")) & "' and S01Item_Code='" & Trim(txtCode.Text) & "'   group by S01Item_Code"
                        M03 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                        If isValidDataset(M03) Then
                            Dim newRow As DataRow = c_dataCustomer1.NewRow
                            newRow("Loc Code") = Trim(M01.Tables(0).Rows(i)("M04Loc_Code"))
                            newRow("Current Qty") = M03.Tables(0).Rows(0)("Qty")


                            c_dataCustomer1.Rows.Add(newRow)
                        Else
                            Dim newRow As DataRow = c_dataCustomer1.NewRow
                            newRow("Loc Code") = Trim(M01.Tables(0).Rows(i)("M04Loc_Code"))
                            newRow("Current Qty") = "0"


                            c_dataCustomer1.Rows.Add(newRow)
                        End If
                    End If
                    i = i + 1
                Next
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub txtDescription_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDescription.KeyUp
        If e.KeyCode = Keys.Enter Then
            txtReorder.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtReorder.Focus()
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
        Call Load_Gride2()

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
        Dim M01 As DataSet

        Try
            If Search_Location() = True Then
            Else
                result1 = MessageBox.Show("Please enter the correct Location", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    connection.Close()
                    Exit Sub
                End If
            End If
            If Search_Category() = True Then
            Else

                result1 = MessageBox.Show("Please enter the correct Category", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboMain.ToggleDropdown()
                    connection.Close()
                    Exit Sub
                End If
            End If

            If txtMrp.Text <> "" Then
            Else
                MsgBox("Please enter the MRP Value", MsgBoxStyle.Information, "Information ......")
                connection.Close()
                Exit Sub
            End If

            If IsNumeric(txtMrp.Text) Then
            Else
                MsgBox("Please enter the correct MRP value", MsgBoxStyle.Information, "Information  .......")
                connection.Close()
                Exit Sub
            End If
            If Search_Supplier() = True Then
            Else

                result1 = MessageBox.Show("Please enter the correct Supplier", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboSupplier.ToggleDropdown()
                    connection.Close()
                    Exit Sub
                End If
            End If


            If Trim(txtCode.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Item Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    connection.Close()
                    Exit Sub
                End If
            End If

            '--------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Item Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    connection.Close()
                    Exit Sub
                End If
            End If

            If txtReorder.Text <> "" Then
                If IsNumeric(txtReorder.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Reorder Level", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtReorder.Focus()
                        connection.Close()
                        Exit Sub
                    End If
                End If
            Else
                txtReorder.Text = "0"
            End If

            If txtComm.Text <> "" Then
            Else
                MsgBox("Please enter the Commission %", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                connection.Close()
                Exit Sub
            End If

            If IsNumeric(txtComm.Text) Then
            Else
                MsgBox("Please enter the correct Commission %", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub
            End If
            If txtCost.Text <> "" Then
                If IsNumeric(txtCost.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Cost Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtReorder.Focus()
                        connection.Close()
                        Exit Sub
                    End If
                End If
            Else
                txtCost.Text = "0"
            End If

            If IsNumeric(txtRetail.Text) Then
                If Val(txtRetail.Text) < 0 Then
                    result1 = MessageBox.Show("Retail price must be >0", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtReorder.Focus()
                        connection.Close()
                        Exit Sub
                    End If
                End If
            Else
                result1 = MessageBox.Show("Please enter the Correct Retail Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtRetail.Focus()
                    connection.Close()
                    Exit Sub
                End If
            End If

            If cboEx_Date.Text <> "" Then
            Else
                cboEx_Date.Text = "NO"
            End If
            '--------------------------------------------------------------------
            nvcFieldList1 = "Insert Into M03Item_Master(M03Item_Code,M03Item_Name,M03Cat_Code,M03Cost_Price,M03Reorder,M03Retail_Price,M03Com_Code,M03Supplier,M03Status,M03ExPair,M03Location,M03Comm,M03MRP)" & _
                                                          " values('" & (Trim(txtCode.Text)) & "', '" & (Trim(txtDescription.Text)) & "','" & Trim(_Category) & "','" & txtCost.Text & "','" & txtReorder.Text & "','" & txtRetail.Text & "','" & _Comcode & "','" & _Supcode & "','A','" & cboEx_Date.Text & "','" & Trim(_Loc) & "','" & txtComm.Text & "','" & txtMrp.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'i = 0
            'nvcFieldList1 = "select * from M04Location"
            'M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            'For Each DTRow2 As DataRow In M01.Tables(0).Rows
            '    '--------------------------------------------------------------------
            '    nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Free_Issue,S01Com_Code,S01Status)" & _
            '                                                  " values('" & Trim(M01.Tables(0).Rows(i)("M04Loc_Code")) & "', '" & (Trim(txtCode.Text)) & "','" & Today & "','OB','0','0','" & _Comcode & "','A')"
            '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            '    i = i + 1
            'Next

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            Me.txtComm.Text = ""
            Me.txtDescription.Text = ""
            Me.cboEx_Date.Text = ""
            Me.txtCode.Text = ""
            Me.txtDescription.Text = ""
            Me.txtRetail.Text = ""
            Me.txtCost.Text = ""
            Me.txtReorder.Text = ""
            Me.txtMrp.Text = ""
            txtCode.Focus()
            'common.ClearAll(OPR0)
            '' OPR2.Enabled = False
            ''OPR1.Enabled = False
            'OPR0.Enabled = True
            '' OPR3.Enabled = False
            cmdAdd.Enabled = True
            cboMain.ToggleDropdown()
            ' cmdSave.Enabled = False
            cmdDelete.Enabled = False
            Call Load_Gride()
            Call Load_Gride2()
            cboLocation.ToggleDropdown()
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
            If Search_Location() = True Then
            Else
                result1 = MessageBox.Show("Please enter the correct Location", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Search_Category() = True Then
            Else

                result1 = MessageBox.Show("Please enter the correct Category", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboMain.ToggleDropdown()
                    Exit Sub
                End If
            End If


            If Search_Supplier() = True Then
            Else

                result1 = MessageBox.Show("Please enter the correct Supplier", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboSupplier.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Trim(txtCode.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Item Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    Exit Sub
                End If
            End If

            If txtComm.Text <> "" Then
            Else
                MsgBox("Please enter the Commission %", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub
            End If

            If IsNumeric(txtComm.Text) Then
            Else
                MsgBox("Please enter the correct Commission %", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub
            End If
            If txtMrp.Text <> "" Then
            Else
                MsgBox("Please enter the MRP Value", MsgBoxStyle.Information, "Information ......")
                Exit Sub
            End If

            If IsNumeric(txtMrp.Text) Then
            Else
                MsgBox("Please enter the correct MRP value", MsgBoxStyle.Information, "Information  .......")
                Exit Sub
            End If
            '--------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Item Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    Exit Sub
                End If
            End If

            If txtReorder.Text <> "" Then
                If IsNumeric(txtReorder.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Reorder Level", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtReorder.Focus()
                        Exit Sub
                    End If
                End If
            Else
                txtReorder.Text = "0"
            End If

            If txtCost.Text <> "" Then
                If IsNumeric(txtCost.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Cost Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtReorder.Focus()
                        Exit Sub
                    End If
                End If
            Else
                txtCost.Text = "0"
            End If

            If IsNumeric(txtRetail.Text) Then
                If Val(txtRetail.Text) < 0 Then
                    result1 = MessageBox.Show("Retail price must be >0", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtReorder.Focus()
                        Exit Sub
                    End If
                End If
            Else
                result1 = MessageBox.Show("Please enter the Correct Retail Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtRetail.Focus()
                    Exit Sub
                End If
            End If

            If cboEx_Date.Text <> "" Then
            Else
                cboEx_Date.Text = "NO"
            End If
            nvcFieldList1 = "update M03Item_Master set M03Item_Name='" & (Trim(txtDescription.Text)) & "',M03Cat_Code='" & _Category & "',M03Cost_Price='" & txtCost.Text & "',M03Reorder='" & txtReorder.Text & "',M03Retail_Price='" & txtRetail.Text & "',M03Supplier='" & _Supcode & "',M03Status='A',M03ExPair='" & Trim(cboEx_Date.Text) & "',M03Location='" & Trim(_Loc) & "',M03Comm='" & txtComm.Text & "',M03MRP='" & txtMrp.Text & "' where M03Item_Code='" & Trim(txtCode.Text) & "' and M03Com_Code='" & _Comcode & "'"
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
            Call Load_Gride2()
            OPR0.Enabled = True
            cmdAdd.Enabled = True
            cboMain.ToggleDropdown()
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            Call Load_Gride()
            cboLocation.ToggleDropdown()
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
                nvcFieldList = "Update M03Item_Master set M03Status='I' where M03Item_Code = '" & Trim(txtCode.Text) & "' and M03Com_Code='" & _Comcode & "' "
                ExecuteNonQueryText(connection, transaction, nvcFieldList)

                MsgBox("Records Deleted Successfully", MsgBoxStyle.Information, "Information .....")
            End If
            transaction.Commit()
            common.ClearAll(OPR0)
            Clicked = ""
            cmdAdd.Enabled = True
            ' cmdSave.Enabled = False
            OPR0.Enabled = True
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cboMain.ToggleDropdown()
            Call Load_Gride2()
            Call Load_Gride()
            cboLocation.ToggleDropdown()
        Catch ex As Exception
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)

        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub

    Private Sub frmItem_Master_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride()
        Call Load_Gride2()
        Call Load_Combo()
        Call Load_Supplier()
        txtReorder.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRetail.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtComm.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMrp.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_ComboEX()
        Call Load_Location()
    End Sub
    Function Load_Location()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [##] from M04Location where M04Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 130
                '  .Rows.Band.Columns(1).Width = 160


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

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M02Cat_Name as [Catogory] from M02Category where M02Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboMain
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 280
                '  .Rows.Band.Columns(1).Width = 160


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

    Function Load_ComboEX()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M25Dis as [##] from M25Ex_Status"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboEx_Date
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 110
                '  .Rows.Band.Columns(1).Width = 160


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

    Function Load_Supplier()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Name as [Supplier Name] from M01Account_Master where  M01Acc_Type='SP' and M01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSupplier
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 280
                '  .Rows.Band.Columns(1).Width = 160


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

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableItem
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function


    Private Sub cboMain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMain.KeyUp
        If e.KeyCode = 13 Then
            cboSupplier.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboSupplier.ToggleDropdown()
        End If
    End Sub

    Private Sub txtDescription_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.TextChanged
        Call Load_Gride1()
    End Sub

    Private Sub txtReorder_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtReorder.KeyUp
        If e.KeyCode = Keys.Enter Then
            cboEx_Date.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboEx_Date.ToggleDropdown()
        End If
    End Sub


    Private Sub txtCost_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCost.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then

            Value = txtCost.Text
            txtCost.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtCost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            txtRetail.Focus()
            Call COmmision_rate()
        ElseIf e.KeyCode = Keys.Tab Then

            Value = txtCost.Text
            txtCost.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtCost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            txtRetail.Focus()

        End If
    End Sub


    Private Sub txtRetail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRetail.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then

            Value = txtRetail.Text
            txtRetail.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            'If cmdEdit.Enabled = True Then
            '    cmdEdit.Focus()
            'Else
            '    cmdAdd.Focus()
            'End If
            txtComm.Focus()
            Call COmmision_rate()
        ElseIf e.KeyCode = Keys.Tab Then
            Value = txtRetail.Text
            txtRetail.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            If cmdEdit.Enabled = True Then
                cmdEdit.Focus()
            Else
                cmdAdd.Focus()
            End If
        End If
    End Sub

 

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim B As New ReportDocument
        Dim A As String
        Try
            'A = ConfigurationManager.AppSettings("ReportPath") + "\ItemList.rpt"
            'B.Load(A.ToString)
            'B.SetDatabaseLogon("sa", "sainfinity")
            'frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            'frmReport.CrystalReportViewer1.DisplayToolbar = True
            '' frmReport.CrystalReportViewer1.SelectionFormula = "{R01Report.R01Com}='" & _Comcode & "' and {R01Report.R01Remark}='SB' "
            'frmReport.Refresh()
            '' frmReport.CrystalReportViewer1.PrintReport()
            '' B.PrintToPrinter(1, True, 0, 0)
            'frmReport.MdiParent = MDIMain
            'frmReport.Show()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(connection)
                'connection.ConnectionString = ""
            End If
        End Try
    End Sub

  

    Private Sub cboSupplier_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSupplier.KeyUp
        If e.KeyCode = 13 Then
            txtCode.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtCode.Focus()

        End If
    End Sub

    Private Sub cboEx_Date_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEx_Date.KeyUp
        If e.KeyCode = 13 Then
            txtCost.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtCost.Focus()
        End If
    End Sub

    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            cboMain.ToggleDropdown()
        End If
    End Sub

   

    Private Sub txtComm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtComm.KeyUp
        If e.KeyCode = 13 Then
            txtMrp.Focus()
        End If
    End Sub

  

    Private Sub txtMrp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMrp.KeyUp
        Dim Value As Double

        If e.KeyCode = 13 Then
            If IsNumeric(txtMrp.Text) Then
                Value = txtMrp.Text
                txtMrp.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMrp.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
            If cmdEdit.Enabled = True Then
                cmdEdit.Focus()
            Else
                cmdAdd.Focus()
            End If
        End If
    End Sub

   
End Class