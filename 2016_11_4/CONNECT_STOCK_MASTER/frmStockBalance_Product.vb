Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports Infragistics.Win
Imports Infragistics.Win.UltraWinToolTip
Imports Infragistics.Win.FormattedLinkLabel
Imports Infragistics.Win.Misc
'Imports Infragistics.Win.UltraWinSchedule

Public Class frmStockBalance_Product
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _LocCode As String
    Dim _ItemCode As String 
    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Name from M11Common where M11Status='LC' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Location = True
                ' _LocCode = Trim(M01.Tables(0).Rows(0)("M11ID"))
            End If


            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function
    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_Stock_Adjestment_1
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 250
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 60
            .DisplayLayout.Bands(0).Columns(5).Width = 60
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            ' .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit



        End With
    End Function

    Function Load_Location()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M11Name as [##] from M11Common where M11Status='LC'  "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 366
                '  .Rows.Band.Columns(1).Width = 242


            End With

            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function

    Private Sub frmStockBalance_Product_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmView_Items_cnt.Close()
    End Sub

    Private Sub frmStockBalance_Product_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCurrent.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCurrent.ReadOnly = True
        txtNew.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLast_Date.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtLast_Date.ReadOnly = True
        txtLast_OB.ReadOnly = True
        txtLast_OB.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtLoading.ReadOnly = True
        'txtLoading.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtunloading.ReadOnly = True
        'txtunloading.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtGrn.ReadOnly = True
        txtGrn.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtMk_Return.ReadOnly = True
        txtMk_Return.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtWastage.ReadOnly = True
        txtWastage.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtSales.ReadOnly = True
        txtSales.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTransfer.ReadOnly = True
        txtTransfer.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Gride2()
        txtItem_Name.ReadOnly = True
        Call Load_Location()
        'Call Load_Item_Code()
        '  Call Load_Item_Name()

        cboLocation.ToggleDropdown()
        txtRack.Appearance.TextHAlign = HAlign.Center
        txtCell.Appearance.TextHAlign = HAlign.Center
        Call Load_Grid_ROW()
    End Sub

    Function Search_Itemcode() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim m02 As DataSet

        Try
            Sql = "select M05Description  from M05Item_Master where M05Status='A'  and M05Ref_No='" & _ItemCode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Itemcode = True
                txtItem_Name.Text = Trim(M01.Tables(0).Rows(0)("M05Description"))
            End If
            '========================================================================
            Sql = "select M12Rack,M12Cell  from M12Store_Location where M12Item_Code='" & _ItemCode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtRack.Text = Trim(M01.Tables(0).Rows(0)("M12Rack"))
                txtCell.Text = Trim(M01.Tables(0).Rows(0)("M12Cell"))
            End If

            '=======================================================================
            Sql = "SELECT * FROM M11Common WHERE  M11Status='LC'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                'current balance
                Sql = "SELECT sum(S01Qty) as S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & _ItemCode & "' AND S01Status='A'   group by S01Item_Code"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then
                    txtCurrent.Text = Trim(m02.Tables(0).Rows(0)("S01Qty"))
                Else
                    txtCurrent.Text = "0"
                End If

                'LAST O/B
                Sql = "SELECT S01Date,S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & _ItemCode & "' AND S01Status='A' AND S01Tr_Type='OB'"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then
                    txtLast_Date.Text = Microsoft.VisualBasic.Day(m02.Tables(0).Rows(0)("S01Date")) & "/" & Month(m02.Tables(0).Rows(0)("S01Date")) & "/" & Year(m02.Tables(0).Rows(0)("S01Date"))
                    txtLast_OB.Text = CInt(m02.Tables(0).Rows(0)("S01Qty"))
                   
                Else
                    txtLast_Date.Text = "-"
                    txtLast_OB.Text = "0"
                End If
                '=======================================================================
                'GRN
                Sql = "SELECT SUM(S01Qty) as S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & _ItemCode & "' AND S01Status='A' AND S01Tr_Type='GRN'  GROUP BY S01Item_Code,S01Tr_Type"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtGrn.Text = CInt(m02.Tables(0).Rows(0)("S01Qty"))
                Else

                    txtGrn.Text = "0"
                End If
                
                '=================================================================================================================================================
                'WASTAGE
                Sql = "SELECT SUM(S01Qty) as S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & _ItemCode & "' AND S01Status='A' AND S01Tr_Type='WST'  GROUP BY S01Item_Code,S01Tr_Type"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtWastage.Text = CInt(m02.Tables(0).Rows(0)("S01Qty"))
                Else

                    txtWastage.Text = "0"
                End If
             
                '=====================================================================================================================
                'Market Return
                Sql = "SELECT SUM(S01Qty)as S01Qty FROM S01Stock_Balance WHERE S01Item_Code='" & _ItemCode & "' AND S01Status='A' AND S01Tr_Type='SP_RETURN' GROUP BY S01Item_Code,S01Tr_Type"
                m02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(m02) Then

                    txtMk_Return.Text = CInt(m02.Tables(0).Rows(0)("S01Qty"))
                Else

                    txtMk_Return.Text = "0"
                End If
            End If

            con.ClearAllPools()
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.CLOSE()
            End If
        End Try
    End Function


    Function clear_text()
        Me.txtItem_Name.Text = ""
        Me.txtCode.Text = ""
        Me.cboLocation.Text = ""
        Me.txtCurrent.Text = ""
        Me.txtNew.Text = ""
        ' Me.txtStock_IN.Text = ""
        Me.txtLast_Date.Text = ""
        Me.txtLast_OB.Text = ""
        '  Me.txtLoading.Text = ""
        Me.txtWastage.Text = ""
        Me.txtTransfer.Text = ""
        Me.txtMk_Return.Text = ""
        Me.txtGrn.Text = ""
        Me.txtCell.Text = ""
        Me.txtRack.Text = ""
        UltraGrid1.Visible = False
        _ItemCode = ""
        '  Me.txtunloading.Text = ""
        cboLocation.ToggleDropdown()
    End Function

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call clear_text()
    End Sub

    Function Load_Grid_ROW()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select m05ref_no as  ##, max(m05item_code) as [Part No],max(M05Brand_Name) as [Brand Name],MAX(tmpDescription) as [Description],max(CAST(Retail AS DECIMAL(16,2))) as [Retail Price],sum(qty) as [Current Stock],max(rack) as [Rack No],max(cell) as [Cell No] from View_Product_Stock  group by m05ref_no having max(m05item_code) like '" & txtCode.Text & "%'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = M01
            UltraGrid1.Rows.Band.Columns(0).Width = 40

            UltraGrid1.Rows.Band.Columns(1).Width = 90
            UltraGrid1.Rows.Band.Columns(2).Width = 110
            UltraGrid1.Rows.Band.Columns(3).Width = 210
            UltraGrid1.Rows.Band.Columns(4).Width = 80
            UltraGrid1.Rows.Band.Columns(5).Width = 80
            UltraGrid1.Rows.Band.Columns(6).Width = 80
            UltraGrid1.Rows.Band.Columns(7).Width = 80
            '  UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            UltraGrid1.Rows.Band.Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            UltraGrid1.Rows.Band.Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            con.ClearAllPools()
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.ClearAllPools()
                con.close()

            End If
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Dim A As String
        A = MsgBox("Are you sure you want to exit this", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Exit .......")
        If A = vbYes Then
            Me.Close()
        End If
    End Sub

    Private Sub cboCode_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Search_Itemcode()
    End Sub

    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            If Trim(txtCode.Text) <> "" Then
                txtRack.Focus()
            Else
                ' cboItemName.ToggleDropdown()
            End If

        ElseIf e.KeyCode = Keys.F1 Then
            strWindowName = Me.Name
            strWinStatus = "PRODUCT"
            frmView_Items_cnt.Close()
            frmView_Items_cnt.Show()
            Call frmView_Items_cnt.Load_Grid_PRODUCT()
        ElseIf e.KeyCode = Keys.Escape Then
            strWindowName = Me.Name
            strWinStatus = "PRODUCT"
            frmView_Items_cnt.Close()
            frmView_Items_cnt.Show()
            Call frmView_Items_cnt.Load_Grid_PRODUCT()
        End If
    End Sub



    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            txtCode.Focus()
        End If
    End Sub

  

    Private Sub UltraButton1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraButton1.MouseHover
        Dim image As String = FormattedLinkEditor.EncodeImage(Me.ImageList1.Images(0))
        Dim TipInfo As New UltraToolTipInfo()
        TipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
        Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraButton1, TipInfo)
        Me.UltraToolTipManager1.DisplayStyle = ToolTipDisplayStyle.BalloonTip

        TipInfo.ToolTipTextFormatted = "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "font-weight:bold; " + _
     "text-smoothing-mode:AntiAlias;'> " + _
    "Techno Help</p> " + _
     "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "text-Click the button to add data to gride;'> " + _
    "<img data='" + image + "' " + _
     "align='left' " + _
     "HSpace='5'/> " + _
    "Click the button to add data to gride.</p>"
    End Sub

    'Private Sub cboCode_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim image As String = FormattedLinkEditor.EncodeImage(Me.ImageList1.Images(0))
    '    Dim TipInfo As New UltraToolTipInfo()
    '    TipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
    '    Me.UltraToolTipManager1.SetUltraToolTip(Me.txtCode.Text, TipInfo)
    '    Me.UltraToolTipManager1.DisplayStyle = ToolTipDisplayStyle.BalloonTip

    '    TipInfo.ToolTipTextFormatted = "<p style='color:Black; " + _
    ' "font-family:verdana; " + _
    ' "font-weight:bold; " + _
    ' "text-smoothing-mode:AntiAlias;'> " + _
    '"Techno Help</p> " + _
    ' "<p style='color:Black; " + _
    ' "font-family:verdana; " + _
    ' "text-Click the button to add data to gride;'> " + _
    '"<img data='" + image + "' " + _
    ' "align='left' " + _
    ' "HSpace='5'/> " + _
    '"Press F1 or Esc button for Item search</p>"
    'End Sub

  
  
    Private Sub cboItemName_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim image As String = FormattedLinkEditor.EncodeImage(Me.ImageList1.Images(0))
        Dim TipInfo As New UltraToolTipInfo()
        TipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
        Me.UltraToolTipManager1.SetUltraToolTip(Me.txtItem_Name, TipInfo)
        Me.UltraToolTipManager1.DisplayStyle = ToolTipDisplayStyle.BalloonTip

        TipInfo.ToolTipTextFormatted = "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "font-weight:bold; " + _
     "text-smoothing-mode:AntiAlias;'> " + _
    "Techno Help</p> " + _
     "<p style='color:Black; " + _
     "font-family:verdana; " + _
     "text-Click the button to add data to gride;'> " + _
    "<img data='" + image + "' " + _
     "align='left' " + _
     "HSpace='5'/> " + _
    "Press F1 or Esc button for Item search</p>"
    End Sub

    

    Private Sub txtNew_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNew.KeyUp
        Dim i As Integer

        If e.KeyCode = 13 Then
            UltraButton1.Focus()
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If Search_Location() = True Then
        Else
            MsgBox("Please select the Location", MsgBoxStyle.Information, "Information .......")
            cboLocation.ToggleDropdown()
            Exit Sub
        End If

        If UltraGrid2.Rows.Count > 0 Then
        Else
            MsgBox("Please enter the Item Details", MsgBoxStyle.Information, "Information ......")
            txtCode.Focus()
            Exit Sub
        End If

        Call Save_Data()
    End Sub

    Function Save_Data()
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


          

            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows

                'nvcFieldList1 = "UPDATE  M05Item_Master SET M05Status='A' where M05Item_Code='" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "'"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE  S01Stock_Balance SET S01Status='CLOSE' where S01Item_Code='" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Item_Code,S01Ref_No,S01Date,S01Time,S01Tr_Type,S01Qty,S01Status)" & _
                                                                    " values('" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "', '-','" & Today & "','" & Now & "','OB','" & Trim(UltraGrid2.Rows(i).Cells(3).Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "select * from M12Store_Location where M12Item_Code='" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "' "
                MB51 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(MB51) Then
                    nvcFieldList1 = "update M12Store_Location set M12Rack='" & Trim(UltraGrid2.Rows(i).Cells(4).Text) & "',M12Cell='" & Trim(UltraGrid2.Rows(i).Cells(5).Text) & "' where M12Item_Code='" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "Insert Into M12Store_Location(M12Item_Code,M12Rack,M12Cell)" & _
                                                                        " values('" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "','" & Trim(UltraGrid2.Rows(i).Cells(4).Text) & "', '" & Trim(UltraGrid2.Rows(i).Cells(5).Text) & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                nvcFieldList1 = "Insert Into tmpMaster_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                     " values('NEW_O/BALANCE','SAVE', '" & Now & "','" & strDisname & "','" & Trim(UltraGrid2.Rows(i).Cells(0).Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = i + 1
            Next
            MsgBox("Stock Adjestment creation successfully", MsgBoxStyle.Information, "Information ........")
            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            Call Load_Gride2()
            Call clear_text()
            txtCode.Focus()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Private Sub cboCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Search_Itemcode()
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        On Error Resume Next
        Dim i As Integer

        If txtNew.Text <> "" Then
            If Search_Itemcode() = True Then
            Else
                MsgBox("Please enter the correct Item code", MsgBoxStyle.Information, "Information ........")
                txtCode.Focus()
                Exit Sub
            End If

            If IsNumeric(txtNew.Text) Then
            Else
                MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information .......")
                txtNew.Focus()
                Exit Sub
            End If

            If Trim(txtRack.Text) <> "" Then
            Else
                MsgBox("Please enter the Rack No", MsgBoxStyle.Information, "Information ........")
                txtRack.Focus()
                Exit Sub
            End If
            If Trim(txtCell.Text) <> "" Then
            Else
                MsgBox("Please enter the Cell No", MsgBoxStyle.Information, "Information ........")
                txtCell.Focus()
                Exit Sub
            End If

            i = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                If (Trim(UltraGrid2.Rows(i).Cells(0).Text)) = Trim(txtCode.Text) Then
                    UltraGrid2.Rows(i).Cells(2).Value = CInt(UltraGrid2.Rows(i).Cells(2).Text) + CInt(txtNew.Text)
                    Me.txtNew.Text = ""
                    Me.txtCode.Text = ""
                    Me.txtItem_Name.Text = ""
                    txtCode.Focus()
                    Exit Sub
                End If
                i = i + 1
            Next
            Dim newRow As DataRow = c_dataCustomer1.NewRow
            newRow("Ref.Code") = _ItemCode
            newRow("Part No") = Trim(txtCode.Text)
            newRow("Item Name") = txtItem_Name.Text
            newRow("Qty") = txtNew.Text
            newRow("#Rack No") = txtRack.Text
            newRow("#Cell No") = txtCell.Text
            c_dataCustomer1.Rows.Add(newRow)
            Me.txtNew.Text = ""
            Me.txtCode.Text = ""
            Me.txtItem_Name.Text = ""
            Me.txtRack.Text = ""
            Me.txtCell.Text = ""
            _ItemCode = ""
            txtCode.Focus()
        End If
    End Sub

    Private Sub txtRack_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRack.KeyUp
        If e.KeyCode = 13 Then
            If txtRack.Text <> "" Then
                txtCell.Focus()
            End If
        End If
    End Sub

    Private Sub txtCell_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCell.KeyUp
        If e.KeyCode = 13 Then
            If txtCell.Text <> "" Then
                txtNew.Focus()
            End If
        End If
    End Sub

    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = Keys.Escape Then
            UltraGrid1.Visible = False
        ElseIf e.KeyCode = 13 Then
            If Trim(txtCode.Text) <> "" Then
                If UltraGrid1.Visible = True Then
                    UltraGrid1.Focus()
                Else
                    txtRack.Focus()
                End If
            End If
            End If
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        If Trim(txtCode.Text) <> "" Then
            If UltraGrid1.Visible = True Then
                Call Load_Grid_ROW()
            Else
                UltraGrid1.Visible = True
                Call Load_Grid_ROW()
            End If
        End If
    End Sub

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _Row As Integer

        _Row = UltraGrid1.ActiveRow.Index
        _ItemCode = Trim(UltraGrid1.Rows(_Row).Cells(0).Text)
        Call Search_Itemcode()
        UltraGrid1.Visible = False
        txtCode.Focus()
    End Sub

   
    Private Sub UltraGrid1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid1.KeyUp
        On Error Resume Next
        Dim _Row As Integer
        If e.KeyCode = 13 Then
            _Row = UltraGrid1.ActiveRow.Index
            _ItemCode = Trim(UltraGrid1.Rows(_Row).Cells(0).Text)
            Call Search_Itemcode()
            UltraGrid1.Visible = False
            txtCode.Focus()
        End If
    End Sub

   
End Class