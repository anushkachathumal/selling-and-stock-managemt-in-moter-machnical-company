Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports System.IO
Public Class frmBom_Creation
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Supplier As String
    Dim _Comcode As String
    Dim _cat_Code As String

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub frmBom_Creation_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmView_Items_cnt.Close()
    End Sub

    Private Sub frmBom_Creation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Item_Code()
        Call Load_Item_Name()
        Call Load_Row_Code()
        Call Load_Row_Name()
        Call Load_Gride2()
        Call Load_Grid_Data()
    End Sub

    Function Clear_text()
        Call Load_Gride2()
        Me.cboRow_Name.Text = ""
        Me.cboRow_Code.Text = ""
        Me.cboItemName.Text = ""
        Me.cboCode.Text = ""
        Me.txtQty.Text = ""
        cboCode.ToggleDropdown()
    End Function

    Function Load_Item_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_No as [##] from M11Product_Item where M11Status='A' and M11Type='PRODUCT ITEM'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCode
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 124
                '  .Rows.Band.Columns(1).Width = 160


            End With


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Item_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_Name as [##] from M11Product_Item where M11Status='A' and M11Type='PRODUCT ITEM'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItemName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 384
                '  .Rows.Band.Columns(1).Width = 160


            End With


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Row_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_No as [##] from M11Product_Item where M11Status='A' and M11Type='ROW ITEMS'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRow_Code
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 119
                '  .Rows.Band.Columns(1).Width = 160


            End With


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Row_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_Name as [##] from M11Product_Item where M11Status='A' and M11Type='ROW ITEMS'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboRow_Name
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 360
                '  .Rows.Band.Columns(1).Width = 160


            End With


            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboCode_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.AfterCloseUp
        Call Search_Records()
        Call Search_Itemcode()

    End Sub

    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCode.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Records()
            Call Search_Itemcode()
            cboItemName.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F1 Then
            strWindowName = Me.Name
            strWinStatus = "PRODUCT"
            frmView_Items_cnt.Close()
            frmView_Items_cnt.Show()
            Call frmView_Items_cnt.Load_Grid_PRODUCT()
        End If
    End Sub

    Private Sub cboItemName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItemName.AfterCloseUp
        Call Search_ItemName()
    End Sub

    Private Sub cboItemName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItemName.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemName()
            cboRow_Code.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F1 Then
            strWindowName = Me.Name
            strWinStatus = "PRODUCT"
            frmView_Items_cnt.Close()
            frmView_Items_cnt.Show()
            Call frmView_Items_cnt.Load_Grid_PRODUCT()
        End If
    End Sub

    Private Sub cboRow_Code_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRow_Code.AfterCloseUp
        Call Search_Rowcode()
    End Sub

    Private Sub cboRow_Code_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboRow_Code.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Rowcode()
            cboRow_Name.ToggleDropdown()
        End If
    End Sub

    Private Sub cboRow_Name_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRow_Name.AfterCloseUp
        Call Search_RowName()
    End Sub

    Private Sub cboRow_Name_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboRow_Name.KeyUp
        If e.KeyCode = 13 Then
            Call Search_RowName()
            txtQty.Focus()
        End If
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_BOM
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 80
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 170
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70


            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center



            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit



        End With
    End Function

    Private Sub UsingDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsingDateToolStripMenuItem.Click
        Call Clear_text()
        cboCode.ToggleDropdown()
        OPR0.Visible = True
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Clear_text()
    End Sub

    Private Sub UltraButton30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton30.Click
        OPR0.Visible = False
        Call Clear_text()
    End Sub

    Function Search_Rowcode() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_Name  from M11Product_Item where M11Status='A' and M11Type='ROW ITEMS' and M11Part_No='" & Trim(cboRow_Code.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Rowcode = True
                cboRow_Name.Text = Trim(M01.Tables(0).Rows(0)("M11Part_Name"))
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_Itemcode() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_Name  from M11Product_Item where M11Status='A' and M11Type='PRODUCT ITEM' and M11Part_No='" & Trim(cboCode.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Itemcode = True
                cboItemName.Text = Trim(M01.Tables(0).Rows(0)("M11Part_Name"))
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_ItemName()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_No  from M11Product_Item where M11Status='A' and M11Type='PRODUCT ITEM' and M11Part_Name='" & Trim(cboItemName.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                cboCode.Text = Trim(M01.Tables(0).Rows(0)("M11Part_No"))
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Search_RowName()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Part_No  from M11Product_Item where M11Status='A' and M11Type='ROW ITEMS' and M11Part_Name='" & Trim(cboItemName.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                cboRow_Name.Text = Trim(M01.Tables(0).Rows(0)("M11Part_No"))
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_Grid_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select  ROW_NUMBER() OVER (ORDER BY M13Part_No) ##,(M13Part_No) as [Item Code],max(M11Part_Name) as [Item Name] from M11Product_Item inner join M13BOM_Creation on M13Part_No=M11Part_No where M13Status='A' group by M13Part_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 40
            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 240
           
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Try
            If e.KeyCode = 13 Then
                If Trim(txtQty.Text) <> "" Then
                    If Search_Rowcode() = True Then
                    Else
                        MsgBox("Please select the correct Row Code", MsgBoxStyle.Information, "Information .......")
                        cboRow_Code.ToggleDropdown()
                        Exit Sub
                    End If

                    If IsNumeric(txtQty.Text) Then
                    Else
                        MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information ..........")
                        txtQty.Focus()
                        Exit Sub
                    End If

                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Item Code") = Trim(cboRow_Code.Text)
                    newRow("Item Name") = cboRow_Name.Text
                    newRow("Qty") = txtQty.Text
                    c_dataCustomer1.Rows.Add(newRow)
                    Me.cboRow_Code.Text = ""
                    Me.cboRow_Name.Text = ""
                    Me.txtQty.Text = ""
                    cboRow_Code.ToggleDropdown()
                End If
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
        End Try
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Search_Itemcode() = True Then
        Else
            MsgBox("Please enter the correct Item Code", MsgBoxStyle.Information, "Information ........")
            cboCode.ToggleDropdown()
            Exit Sub

            If UltraGrid2.Rows.Count > 0 Then
            Else
                MsgBox("Please enter the Row material details", MsgBoxStyle.Information, "Information .......")
                Exit Sub
            End If
        End If

        Call SAVE_DATA()
    End Sub

    Function SAVE_DATA()
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
            nvcFieldList1 = "delete from M13BOM_Creation where M13Part_No='" & Trim(cboCode.Text) & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                       " values('NEW_BOM','SAVE', '" & Now & "','" & strDisname & "','" & cboCode.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                nvcFieldList1 = "Insert Into M13BOM_Creation(M13Part_No,M13Row_Code,M13Row_Name,M13Qty,M13Status)" & _
                                                                    " values('" & (Trim(cboCode.Text)) & "', '" & (Trim(UltraGrid2.Rows(i).Cells(0).Text)) & "','" & (Trim(UltraGrid2.Rows(i).Cells(1).Text)) & "','" & (Trim(UltraGrid2.Rows(i).Cells(2).Text)) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                i = i + 1
            Next
            MsgBox("BOM creation successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            connection.Close()
            Call Load_Gride2()
            Call Load_Grid_Data()
            Me.cboCode.Text = ""
            Me.cboItemName.Text = ""
            Me.txtQty.Text = ""
            Me.cboRow_Code.Text = ""
            Me.cboRow_Name.Text = ""
            cboCode.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Function DELETE_DATA()
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
            nvcFieldList1 = "delete from M13BOM_Creation where M13Part_No='" & Trim(cboCode.Text) & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                       " values('NEW_BOM','DELETE', '" & Now & "','" & strDisname & "','" & cboCode.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            MsgBox("BOM DELETED successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            connection.Close()
            Call Load_Gride2()
            Call Load_Grid_Data()
            Me.cboCode.Text = ""
            Me.cboItemName.Text = ""
            Me.txtQty.Text = ""
            Me.cboRow_Code.Text = ""
            Me.cboRow_Name.Text = ""
            cboCode.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim I As Integer
        Try
            Sql = "select *  from M13BOM_Creation where M13Status='A' and M13Part_No='" & Trim(cboCode.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                Call Search_Itemcode()
                I = 0
                Call Load_Grid_Data()
                For Each DTRow2 As DataRow In M01.Tables(0).Rows
                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Item Code") = Trim(M01.Tables(0).Rows(I)("M13Row_Code"))
                    newRow("Item Name") = Trim(M01.Tables(0).Rows(I)("M13Row_Name"))
                    newRow("Qty") = CInt(M01.Tables(0).Rows(I)("M13Qty"))
                    c_dataCustomer1.Rows.Add(newRow)
                    I = I + 1
                Next
            End If

            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _row As Integer

        _row = UltraGrid1.ActiveRow.Index
        OPR0.Visible = True
        cboCode.Text = Trim(UltraGrid1.Rows(_row).Cells(1).Text)
        Call Search_Records()
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim A As String

        A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete ........")
        If A = vbYes Then
            Call DELETE_DATA()
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call Clear_text()
    End Sub


   
End Class


