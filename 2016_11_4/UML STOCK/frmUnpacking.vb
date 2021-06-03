Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmUnpacking
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _ProItemcode As String
    Dim _Itemcode As String

    Function Load_Gride_Item()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Unpacking
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 320
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 90
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(3).Width = 90
            .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(4).Width = 220
            .DisplayLayout.Bands(0).Columns(4).AutoEdit = False

            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            '.DisplayLayout.Bands(0).Columns(1).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            '.DisplayLayout.Bands(0).Columns(0).CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            ' .DisplayLayout.Bands(0).Columns(1).
            ' .DisplayLayout.Bands(0).Header.Height = 60

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
            Sql = "select M14Item_Name as [##] from View_Production_Items where M14Status='A' and Category='PS' order by M14Item_Code "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 220
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

    Function Load_Parameter()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01PARAMETER where P01CODE='UPK'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtBatch.Text = M01.Tables(0).Rows(0)("P01NO")
            End If
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

    Private Sub frmUnpacking_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Load_Combo()
        Call Load_Gride_Item()
        Call Load_Parameter()
        txtItem_Qty.ReadOnly = True
        txtItem_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCurrent.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRecycle.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtBatch.ReadOnly = True
        txtBatch.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate.Text = Today
        Call Load_Gride()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Me.Close()
    End Sub

    Function Search_Praduct_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try

            Sql = "select * from View_Production_Items where m14Item_name='" & cboItem.Text & "' and category='PS' and M14Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _ProItemcode = Trim(M01.Tables(0).Rows(0)("M14Item_code"))
            End If

            Sql = "select m14item_name as [##] from M16Item_for_Set inner join M14Product_Item on M16item_code=m14item_code where M16Product_Code='" & _ProItemcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPr_Item
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 220
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

    Private Sub cboItem_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItem.AfterCloseUp
        Call Search_Praduct_Code()
        Call Load_Gride_Item()
        Call Load_Date()
    End Sub

    Private Sub cboItem_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboItem.InitializeLayout

    End Sub

    Private Sub cboItem_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItem.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Praduct_Code()
            Call Load_Gride_Item()
            Call Load_Date()
            txtCurrent.Focus()

        ElseIf e.KeyCode = Keys.F1 Then
            UltraGrid2.Visible = True
            OPRView.Visible = True
        ElseIf e.KeyCode = Keys.Escape Then
            UltraGrid2.Visible = False
            OPRView.Visible = False
            cmdSave.Enabled = True
        End If
    End Sub



    Function Load_Date()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _QTY As Integer
        Try
            'Product Item 
            '---------------------------------------------->> OB (OPERNING BALANCE)
            '---------------------------------------------->> DR (SALES TRANSACTION)
            '---------------------------------------------->> RN (RETURN) REMARK USERBALE
            '---------------------------------------------->> PK (PACKING)
            '---------------------------------------------->> SI (STOCK IN)
            _QTY = 0
            i = 0

            Sql = "select * from M16Item_for_Set inner join M14Product_Item on M16item_code=m14item_code where M16Product_Code='" & _ProItemcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows


                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = M01.Tables(0).Rows(i)("M14Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                newRow("Qty") = "0"
                newRow("Recycle Qty") = "0"
                newRow("Remark") = "-"

                c_dataCustomer1.Rows.Add(newRow)




                i = i + 1
            Next

            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub txtCurrent_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrent.KeyUp
        If e.KeyCode = 13 Then
            Call Change_Gide_Data()
            cboPr_Item.ToggleDropdown()
        End If
    End Sub

    Function Change_Gide_Data()
        Dim Sql As String

        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _QTY As Integer
        Dim x As Integer

        Try
           
            _QTY = 0
            i = 0

            Sql = "select * from M16Item_for_Set inner join M14Product_Item on M16item_code=m14item_code where M16Product_Code='" & _ProItemcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                x = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    If Trim(M01.Tables(0).Rows(i)("M14Item_Code")) = Trim(UltraGrid1.Rows(x).Cells(0).Text) Then
                        UltraGrid1.Rows(x).Cells(2).Value = M01.Tables(0).Rows(i)("M16Qty") * txtCurrent.Text
                        Exit For
                    End If
                    x = x + 1
                Next


                i = i + 1
            Next

            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub cboPr_Item_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPr_Item.AfterCloseUp
        Call Search_ItemCode()
    End Sub

    Private Sub cboPr_Item_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPr_Item.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemCode()
            txtRecycle.Focus()
        End If
    End Sub

    Function Search_ItemCode()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet


        Try

        
            Sql = "select * from M16Item_for_Set inner join M14Product_Item on M16item_code=m14item_code where M16Product_Code='" & _ProItemcode & "' and m14Item_Name='" & cboPr_Item.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Itemcode = Trim(M01.Tables(0).Rows(0)("M14Item_Code"))
                If IsNumeric(txtCurrent.Text) Then
                    txtItem_Qty.Text = M01.Tables(0).Rows(0)("M16Qty") * txtCurrent.Text
                End If
            End If

            con.close()

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Private Sub txtRecycle_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRecycle.KeyUp
        If e.KeyCode = 13 Then
            If txtRecycle.Text <> "" Then
                txtRemark.Focus()
            End If
        End If
    End Sub

    Private Sub txtRemark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyUp
        Dim i As Integer
        Dim _Qty As Integer
        Try
            If e.KeyCode = 13 Then
                If IsNumeric(txtCurrent.Text) Then
                    Call Change_Gide_Data()
                Else
                    MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information .....")
                    txtCurrent.Focus()
                    Exit Sub
                End If
                If IsNumeric(txtRecycle.Text) Then
                    If CDbl(txtRecycle.Text) > CDbl(txtItem_Qty.Text) Then
                        MsgBox("Please enter the correct Recycle Qty", MsgBoxStyle.Information, "Information ....")
                        txtRecycle.Focus()
                        Exit Sub
                    End If
                    _Qty = 0
                    i = 0
                    For Each uRow As UltraGridRow In UltraGrid1.Rows
                        If _Itemcode = Trim(UltraGrid1.Rows(i).Cells(0).Text) Then
                            UltraGrid1.Rows(i).Cells(3).Value = txtRecycle.Text
                            UltraGrid1.Rows(i).Cells(4).Value = txtRemark.Text
                            Exit For
                        End If
                        i = i + 1
                    Next
                    txtItem_Qty.Text = ""
                    txtRecycle.Text = ""
                    txtRemark.Text = ""
                    cboPr_Item.Text = ""
                    cboPr_Item.ToggleDropdown()
                Else
                    MsgBox("Please enter the correct Recycle Qty", MsgBoxStyle.Information, "Information ......")
                    txtRecycle.Focus()
                End If
            End If

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                ' con.close()
            End If
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
        Dim MB51 As DataSet
        Dim _count As Integer
        Dim i As Integer
        Dim A As String
        Dim X As Integer
        Dim T01 As DataSet

        Try
            nvcFieldList1 = "UPDATE P01PARAMETER SET P01NO=P01NO + " & 1 & " WHERE P01CODE='UPK'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "SELECT * FROM View_Production_Items WHERE M14Item_name='" & cboItem.Text & "' AND m14status='A' and Category='PS'"
            MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(MB51) Then
                _ProItemcode = Trim(MB51.Tables(0).Rows(0)("m14Item_code"))
            Else
                MsgBox("Please select the correct Production Item", MsgBoxStyle.Information, "Information .....")
                cboItem.ToggleDropdown()
                connection.Close()
                Exit Sub
            End If

            nvcFieldList1 = "Insert Into S02Set_Stock(S02Tr_Type,S02Date,S02Pr_Code,S02Qty,S02Location,S02Status,S02User,S02Product_Status,S02Ref_No)" & _
                                                           " values('UPK', '" & txtDate.Text & "','" & _ProItemcode & "','" & -(txtCurrent.Text) & "','MS','A','" & strDisname & "','-','" & txtBatch.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            nvcFieldList1 = "Insert Into T05Unpacking_Header(T05Ref_No,T05Date,T05Pro_Code,T05Qty,T05Status,T05User,T05Time)" & _
                                                         " values('" & txtBatch.Text & "', '" & txtDate.Text & "','" & _ProItemcode & "','" & txtCurrent.Text & "','A','" & strDisname & "','" & Now & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If (UltraGrid1.Rows(i).Cells(3).Text) > (UltraGrid1.Rows(i).Cells(2).Text) Then
                    MsgBox("Recycle Qty grater than qty please check again -" & (UltraGrid1.Rows(i).Cells(1).Text), MsgBoxStyle.Information, "Information ....")
                    connection.Close()
                    Exit Sub
                End If
                i = i + 1
            Next
            _count = 0
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                If IsNumeric(UltraGrid1.Rows(i).Cells(2).Text) Then
                    _count = _count + 1

                    nvcFieldList1 = "SELECT * FROM View_Production_Items WHERE M14Item_Code='" & UltraGrid1.Rows(i).Cells(0).Value & "' AND m14status='A' and Category='PI'"
                    MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(MB51) Then
                    
                        nvcFieldList1 = "Insert Into S01Product_Stock(S01Tr_Type,S01Date,S01Item_Code,S01Qty,S01Location,S01Status,S01User,S01Product_Status,S01Ref_No)" & _
                                                       " values('UPK', '" & txtDate.Text & "','" & MB51.Tables(0).Rows(0)("M14Item_Code") & "','" & (UltraGrid1.Rows(i).Cells(2).Value) - (UltraGrid1.Rows(i).Cells(3).Value) & "','MS','A','" & strDisname & "','GOOD','" & txtBatch.Text & "')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        nvcFieldList1 = "Insert Into T06Unpacking_Fluter(T06Ref_No,T06Item_Code,T06Qty,T06Recycle,T06Remark,T06Status)" & _
                                                     " values('" & txtBatch.Text & "','" & UltraGrid1.Rows(i).Cells(0).Value & "','" & UltraGrid1.Rows(i).Cells(2).Value & "','" & UltraGrid1.Rows(i).Cells(3).Value & "','" & UltraGrid1.Rows(i).Cells(4).Value & "','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                    End If
                End If
                i = i + 1
            Next

            MsgBox(_count & " Items successfully updated", MsgBoxStyle.Information, "Information ..........")
            transaction.Commit()
            connection.Close()
            Call Clear_Text()
            cboItem.ToggleDropdown()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub

    Function Clear_Text()
        Me.txtCurrent.Text = ""
        Me.txtBatch.Text = ""
        Me.cboPr_Item.Text = ""
        Me.cboItem.Text = ""
        Me.txtRecycle.Text = ""
        Me.txtRemark.Text = ""
        Me.txtItem_Qty.Text = ""
        Call Load_Parameter()
        Call Load_Gride_Item()
        Call Load_Gride()
        OPRView.Visible = False
        cmdSave.Enabled = True
    End Function

    Function Load_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select T05Ref_No as [Un Packing No],T05Pro_Code as [Product Code],M14Item_Name as [Product Name],T05Qty as [Qty] from T05Unpacking_Header inner join View_Production_Items on M14Item_code=T05Pro_Code where T05Status='A' and category='PS' and T05Date='" & Today & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = M01
            With UltraGrid2
                .DisplayLayout.Bands(0).Columns(0).Width = 70
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(1).Width = 90
                .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

                .DisplayLayout.Bands(0).Columns(2).Width = 180
                .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(3).Width = 90
                .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
                .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            End With
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try
    End Function

    Private Sub UltraGrid2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid2.DoubleClick
        On Error Resume Next
        Dim _rowIndex As Integer
        _rowIndex = UltraGrid2.ActiveRow.Index
        txtBatch.Text = UltraGrid2.Rows(_rowIndex).Cells(0).Text
        Call Search_Records()
        OPRView.Visible = False
        cmdSave.Enabled = False
    End Sub
    Function Search_Records()
        Dim Sql As String
        Dim M01 As DataSet
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim i As Integer
        Try
            Sql = "select * from T05Unpacking_Header inner join View_Production_Items on M14Item_Code=T05Pro_Code where T05Ref_No='" & txtBatch.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboItem.Text = Trim(M01.Tables(0).Rows(0)("M14item_name"))
                txtDate.Text = M01.Tables(0).Rows(0)("T05Date")
                txtCurrent.Text = M01.Tables(0).Rows(0)("T05Qty")
            End If

            Sql = "select * from T06Unpacking_Fluter inner join M14Product_Item on M14Item_Code=T06Item_Code where T06Ref_No='" & txtBatch.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            Call Load_Gride_Item()
            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = M01.Tables(0).Rows(i)("M14Item_Code")
                newRow("Item Name") = M01.Tables(0).Rows(i)("M14Item_Name")
                newRow("Qty") = M01.Tables(0).Rows(i)("T06Qty")
                newRow("Recycle Qty") = M01.Tables(0).Rows(i)("T06Recycle")
                newRow("Remark") = M01.Tables(0).Rows(i)("T06Remark")

                c_dataCustomer1.Rows.Add(newRow)

                i = i + 1
            Next
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

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
            A = MsgBox("Are you sure you want to delete this Records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete Records .......")
            If A = vbYes Then
                nvcFieldList1 = "UPDATE T06Unpacking_Fluter SET T06Status='I' WHERE T06Ref_No='" & txtBatch.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T05Unpacking_Header SET T05Status='I' WHERE T05Ref_No='" & txtBatch.Text & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S02Set_Stock SET S02Status='I' WHERE S02Ref_No='" & txtBatch.Text & "' AND S02Tr_Typ='UPK'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S01Product_Stock SET S01Status='I' WHERE S01Ref_No='" & txtBatch.Text & "' AND S01Tr_Typ='UPK'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records deleted successfully", MsgBoxStyle.Information, "Information ........")
            End If

            transaction.Commit()
            connection.Close()
            Call Clear_Text()
            cboItem.ToggleDropdown()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Sub
End Class