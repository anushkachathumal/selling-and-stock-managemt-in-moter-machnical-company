Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmDirect_Sales
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _Emp As String
    Dim _Location As Integer
    Dim _LogStaus As Boolean
    Dim _UserLevel As String
    Dim _CusNo As String
    Dim _CounterPC As String
    Dim _CusType As String
    ' Dim _ItemCode As String

    Function Load_VNO()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M07V_No as [##],M06Name as [Customer Name] from M07Vehicle_Master inner join M06Customer_Master on M06Code=M07Cus_Code where M07Status='A'  order by M07ID"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboV_no
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 159
                .Rows.Band.Columns(1).Width = 210

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

    Function Search_Records()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim Value As Double
        Dim i As Integer
        Dim _St As String
        Try
            Sql = "select * from T08Sales_Header where T08Invo_No='" & Trim(txtEnter1.Text) & "' and T08Tr_Type='DIRECT_SALE'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtDate1.Text = M01.Tables(0).Rows(0)("T08Date")
                cboV_no.Text = M01.Tables(0).Rows(0)("T08V_No")
                txtMtr.Text = M01.Tables(0).Rows(0)("T08St_Mtr")

                cmdDelete1.Enabled = True
                Sql = "select * from M06Customer_Master where M06Code='" & Trim(M01.Tables(0).Rows(0)("T08Cus_NO")) & "' "
                M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(M02) Then
                    cboCustomer.Text = Trim(M02.Tables(0).Rows(0)("M06Name"))
                End If
            End If
            '-----------------------------------------------------------------------
            Call Load_Gride2()
            Sql = "select * from T09Sales_Flutter inner join View_Product_Item on M05Ref_No=T09Item_Ref where T09Inv_No='" & Trim(txtEnter1.Text) & "'"
            M02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            For Each DTRow2 As DataRow In M02.Tables(0).Rows
                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("#Ref.No") = Trim(M02.Tables(0).Rows(i)("T09Item_Ref"))
                newRow("#Part No") = Trim(M02.Tables(0).Rows(i)("T09Item_Code"))
                newRow("Item Name") = Trim(M02.Tables(0).Rows(i)("T09Item_Name"))
                Value = Trim(M02.Tables(0).Rows(i)("T09Retail"))
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Retail Price") = _St
                newRow("Qty") = Trim(M02.Tables(0).Rows(i)("T09Qty"))
                newRow("Free Issue") = Trim(M02.Tables(0).Rows(i)("T09Free"))
                newRow("Discount%") = Trim(M02.Tables(0).Rows(i)("T09Discount"))
                Value = CDbl(M02.Tables(0).Rows(i)("T09Retail")) * CDbl(M02.Tables(0).Rows(i)("T09Qty"))
                Value = Value - (Value * CDbl(M02.Tables(0).Rows(i)("T09Discount")) / 100)
                _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Total") = _St
                c_dataCustomer1.Rows.Add(newRow)
                i = i + 1
            Next
            txtCount1.Text = UltraGrid2.Rows.Count
            Call Calculation_Net()
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
    Function Search_Discount()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select * from M05Item_Master where  M05Item_Code='" & Trim(cboPartno.Text) & "' and M05Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then

                Value = Trim(M01.Tables(0).Rows(0)("M05Discount"))
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            Else
                txtDiscount.Text = "0"
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

    Function Search_Item() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Search_Item = False
            Sql = "select * from View_Product_Stock where  M05Ref_No='" & strItem_Code & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Item = True
                cboPartNo.Text = Trim(M01.Tables(0).Rows(0)("M05Item_Code"))
                txtItem_Name.Text = Trim(M01.Tables(0).Rows(0)("tmpDescription"))
                Value = Trim(M01.Tables(0).Rows(0)("Retail"))
                txtRate1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtRate1.ReadOnly = True
            Else
                txtRate1.Text = ""
                txtRate1.ReadOnly = False
            End If
            Call Search_Discount()
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

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub
    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='" & _CounterPC + "-D" & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                    txtEnter1.Text = _CounterPC & "-D-00" & M01.Tables(0).Rows(0)("P01No")
                ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                    txtEnter1.Text = _CounterPC & "-D-0" & M01.Tables(0).Rows(0)("P01No")
                Else
                    txtEnter1.Text = _CounterPC & "-D-" & M01.Tables(0).Rows(0)("P01No")
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

    Function SEARCH_CUSTOMER() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select * from M06Customer_Master WHERE M06Status='A' AND M06Name='" & Trim(cboCustomer.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                SEARCH_CUSTOMER = True
                _CusNo = Trim(M01.Tables(0).Rows(0)("M06Code"))
                txtAddress1.Text = Trim(M01.Tables(0).Rows(0)("M06Address"))
                txtTP1.Text = Trim(M01.Tables(0).Rows(0)("M06Mobile_No"))
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
    Function Load_cUSTOMER()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M06Name as [##] from M06Customer_Master WHERE M06Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCustomer
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 384


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

    Private Sub frmDirect_Sales_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmPayMain_DS.Close()
    End Sub

  
    Private Sub frmDirect_Sales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _CounterPC = ConfigurationManager.AppSettings("MC")
        txtEnter1.ReadOnly = True
        txtEnter1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDate1.Text = Today
        Call Load_Gride2()
        txtNet1.ReadOnly = True
        txtNet1.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtCount1.ReadOnly = True
        txtCount1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRate1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRate1.ReadOnly = True
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDiscount.ReadOnly = True
        txtTotal.ReadOnly = True
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFree.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtFree.ReadOnly = True
        txtMtr.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtAddress1.ReadOnly = True
        '  txtTP1.ReadOnly = True
        Call Load_EntryNo()
        Call Load_cUSTOMER()
        '  Call Load_Item()
        Call Load_VNO()
    End Sub
    Function Load_Grid_ROW()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select m05ref_no as  ##, max(m05item_code) as [Part No],max(M05Brand_Name) as [Brand Name],MAX(tmpDescription) as [Description],max(CAST(Retail AS DECIMAL(16,2))) as [Retail Price],sum(qty) as [Current Stock],max(rack) as [Rack No],max(cell) as [Cell No] from View_Product_Stock  group by m05ref_no having  max(m05item_code) like '" & Trim(cboPartNo.Text) & "%'"
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

    Function Search_Vehicle_No() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Search_Vehicle_No = False
            Sql = "select * from M07Vehicle_Master inner join M06Customer_Master on M06Code=M07Cus_Code where M07Status='A'  and M07V_No='" & Trim(cboV_no.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Vehicle_No = True
                _CusNo = Trim(M01.Tables(0).Rows(0)("M06Code"))
                'cboBrand.Text = Trim(M01.Tables(0).Rows(0)("M07Brand_Name"))
                'cbov_Type.Text = Trim(M01.Tables(0).Rows(0)("M07Type"))
                txtTP1.Text = Trim(M01.Tables(0).Rows(0)("M06Mobile_No"))
                cboCustomer.Text = Trim(M01.Tables(0).Rows(0)("M06Name"))
                txtAddress1.Text = Trim(M01.Tables(0).Rows(0)("M06Address"))
                _CusType = Trim(M01.Tables(0).Rows(0)("M06Cus_Type"))
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

    Function Load_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        SqlConnection.ClearAllPools()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Try
            Sql = "select M05Item_Code as [##],MAX(tmpDescription) as [Item Name],SUM(qTY) as [Quantity],MAX(Rack) as [#Rack],MAX(Cell) as[#Cell] from View_Product_Stock GROUP BY M05Item_Code"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPartno
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 180
                .Rows.Band.Columns(1).Width = 280
                .Rows.Band.Columns(2).Width = 80
                .Rows.Band.Columns(3).Width = 80
                .Rows.Band.Columns(4).Width = 80
                .Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

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


  
    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTable_itemIssue_Uniq
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(0).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(2).Width = 290
            .DisplayLayout.Bands(0).Columns(2).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(6).Width = 70
            .DisplayLayout.Bands(0).Columns(7).Width = 110

            .DisplayLayout.Bands(0).Columns(0).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit
        End With
    End Function

    Private Sub cboCustomer_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustomer.AfterCloseUp
        Call SEARCH_CUSTOMER()
    End Sub

    Private Sub cboCustomer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCustomer.KeyUp
        If e.KeyCode = 13 Then
            If SEARCH_CUSTOMER() = True Then
                cboPartNo.ToggleDropdown()
            Else
                txtAddress1.Focus()
            End If
        End If
    End Sub
    Function Calculation_Net()
        On Error Resume Next
        Dim Value As Double
        Dim i As Integer
        i = 0
        Value = 0
        For Each uRow As UltraGridRow In UltraGrid2.Rows
            If IsNumeric(UltraGrid2.Rows(i).Cells(7).Text) Then
                Value = Value + CDbl(UltraGrid2.Rows(i).Cells(7).Text)
            End If

            i = i + 1
        Next

        txtNet1.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
        txtNet1.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
    End Function

    Function Calculation_total()
        On Error Resume Next
        Dim Value As Double
        If IsNumeric(txtRate1.Text) And IsNumeric(txtQty.Text) And IsNumeric(txtDiscount.Text) Then
            Value = CDbl(txtRate1.Text) * CDbl(txtQty.Text)
            Value = Value - (Value * CDbl(txtDiscount.Text) / 100)
            txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
        End If
    End Function

    Private Sub cboPartNo_AfterDropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPartNo.AfterDropDown
        'Call Search_Item()
    End Sub

    Private Sub cboPartNo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboPartNo.KeyUp
        If e.KeyCode = 13 Then
            If Trim(cboPartNo.Text) <> "" Then
                If UltraGrid1.Visible = True Then
                    UltraGrid1.Focus()
                Else
                    txtQty.Focus()
                End If
            End If
        ElseIf e.KeyCode = Keys.Escape Then
            UltraGrid1.Visible = False
        End If
    End Sub

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Try
            If e.KeyCode = 13 Then

                If Search_Item() = True Then
                Else
                    MsgBox("Please enter the Part No", MsgBoxStyle.Information, "Information ........")
                    ' cboPartno.ToggleDropdown()
                    Exit Sub
                End If

             

                If txtQty.Text <> "" Then
                Else
                    txtQty.Text = "0"
                End If

                If IsNumeric(txtQty.Text) Then
                Else
                    MsgBox("Please enter the correct Qty", MsgBoxStyle.Information, "Information .........")
                    Exit Sub
                End If

                If txtFree.Text <> "" Then
                Else
                    txtFree.Text = "0"
                End If

                If IsNumeric(txtFree.Text) Then
                Else
                    MsgBox("Please enter the correct Free Issue", MsgBoxStyle.Information, "Information .........")
                    Exit Sub
                End If

                If txtDiscount.Text <> "" Then
                Else
                    txtDiscount.Text = "0"
                End If
                Call Calculation_total()

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("#Ref.No") = strItem_Code
                newRow("#Part No") = Trim(cboPartNo.Text)
                newRow("Item Name") = Trim(txtItem_Name.Text)
                newRow("Retail Price") = txtRate1.Text
                newRow("Qty") = txtQty.Text
                newRow("Free Issue") = txtFree.Text
                newRow("Discount%") = txtDiscount.Text
                newRow("Total") = txtTotal.Text
                c_dataCustomer1.Rows.Add(newRow)
                Me.cboPartNo.Text = ""
                Me.txtItem_Name.Text = ""
                Me.txtQty.Text = "0"
                Me.txtRate1.Text = "00.00"
                Me.txtFree.Text = "0"
                Me.txtTotal.Text = "00.00"
                Me.txtDiscount.Text = "0"
                strItem_Code = ""
                cboPartNo.ToggleDropdown()
                txtCount1.Text = UltraGrid2.Rows.Count
            End If

            Call Calculation_Net()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub txtQty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQty.TextChanged
        Call Calculation_total()
    End Sub

    Private Sub UltraGrid2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid2.AfterRowsDeleted
        Call Calculation_Net()
        txtCount1.Text = UltraGrid2.Rows.Count
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click


        If UltraGrid2.Rows.Count > 0 Then
        Else
            MsgBox("Please enter the item Details", MsgBoxStyle.Information, "Information ........")
            Exit Sub
        End If

        If txtRemark.Text <> "" Then
        Else
            txtRemark.Text = "-"
        End If

        If Trim(cboV_no.Text) <> "" Then
        Else
            cboV_no.Text = "-"
        End If

        If Trim(txtMtr.Text) <> "" Then
        Else
            txtMtr.Text = "-"
        End If

        frmPayMain_DS.Close()
        frmPayMain_DS.txtBill_Amount.Text = txtNet1.Text

        frmPayMain_DS.Show()
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
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim A As String
        Dim B As New ReportDocument

        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim I As Integer
        Dim _CREDIT As Double

        Dim _SERVICE_AMOUNT As Double
        Dim _Remark As String

        Try
         

            _GetDate = Month(txtDate1.Text) & "/" & Microsoft.VisualBasic.Day(txtDate1.Text) & "/" & Year(txtDate1.Text)

            _Get_Time = Month(Now) & "/" & Microsoft.VisualBasic.Day(Now) & "/" & Year(Now) & " " & Hour(Now) & ":" & Minute(Now)

            nvcFieldList1 = "select * from T08Sales_Header  where T08Invo_No='" & Trim(txtEnter1.Text) & "'  "
            M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M02) Then
                MsgBox("This invoice no alrady exsist.", MsgBoxStyle.Information, "Information ......")
                connection.ClearAllPools()
                connection.Close()
                Exit Function
                'If _LogStaus = True Then
                '    nvcFieldList1 = "UPDATE T08Sales_Header SET T08Date='" & _GetDate & "',T08Cus_NO='" & _CusNo & "',T08Net_Amount='" & CDbl(txtNet1.Text) & "',T08Status='A' WHERE T08Invo_No='" & txtEnter1.Text & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '    nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                '                                  " values('DIRECT_SALE','EDIT', '" & Now & "','" & strDisname & "','" & Trim(txtEnter1.Text) & "')"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '    '=========================================================================================
                '    _CREDIT = 0
                '    If SEARCH_CUSTOMER() = True Then
                '        _CREDIT = CDbl(frmPayMain_DS.txtBalance.Text)

                '    Else
                '        _CREDIT = 0
                '    End If
                '    nvcFieldList1 = "UPDATE T11Income_Summery SET T11Net_Amount='" & CDbl(txtNet1.Text) & "',T11Cash='" & CDbl(frmPayMain_DS.txtCash.Text) & "',T11Credit='" & _CREDIT & "',T11Chq='" & CDbl(frmPayMain_DS.txtChq_Total.Text) & "',T11Card='" & frmPayMain_DS.txtTotal.Text & "',T11Status='A' WHERE T11Invo_No='" & Trim(txtEnter1.Text) & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '    '==========================================================================================

                '    nvcFieldList1 = "SELECT * FROM T15Outstanding_Collection WHERE T15Inv_No='" & Trim(txtEnter1.Text) & "' AND T15Status='PAID'"
                '    M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                '    If isValidDataset(M02) Then
                '        MsgBox("This invoice alrady paid.can't Delete", MsgBoxStyle.Information, "Information .........")
                '        connection.ClearAllPools()
                '        connection.Close()
                '        Exit Function
                '    End If

                '    nvcFieldList1 = "DELETE FROM FROM T12OutStanding WHERE T12Inv_No='" & Trim(txtEnter1.Text) & "' AND T12Status='A'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '    If CDbl(frmPayMain_uniq.txtBalance.Text) > 0 Then
                '        If SEARCH_CUSTOMER() = True Then
                '            nvcFieldList1 = "select * from T12OutStanding where T12Inv_No='" & Trim(txtEnter1.Text) & "'"
                '            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                '            If isValidDataset(M01) Then
                '            Else
                '                nvcFieldList1 = "Insert Into T12OutStanding(T12Inv_No,T12Cus_No,T12Date,T12Pay_No,T12Chq_No,T12Cr,T12Dr,T12Status)" & _
                '                                               " values('" & Trim(txtEnter1.Text) & "','" & _CusNo & "','" & _GetDate & "','-','-','" & CDbl(frmPayMain_DS.txtBalance.Text) & "','0','A')"
                '                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '            End If
                '        Else
                '            MsgBox("Please select the customer name", MsgBoxStyle.Information, "Information .......")
                '            connection.ClearAllPools()
                '            connection.Close()
                '            Exit Function
                '        End If
                '    End If
                '    '=================================================================================================================

                '    nvcFieldList1 = "DELETE FROM T09Sales_Flutter WHERE T09Inv_No='" & txtEnter1.Text & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '    nvcFieldList1 = "DELETE FROM S01Stock_Balance WHERE S01Ref_No='" & txtEnter1.Text & "' AND S01Tr_Type='DIRECT_SALE'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '    I = 0
                '    For Each uRow As UltraGridRow In UltraGrid2.Rows
                '        nvcFieldList1 = "SELECT * FROM M05Item_Master WHERE M05Item_Code ='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "'"
                '        M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                '        If isValidDataset(M02) Then
                '            nvcFieldList1 = "Insert Into S01Stock_Balance(S01Item_Code,S01Ref_No,S01Date,S01Time,S01Tr_Type,S01Qty,S01Status)" & _
                '                                    " values('" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "','" & txtEnter1.Text & "', '" & _GetDate & "','" & _Get_Time & "','DIRECT_SALES','" & -(CDbl(UltraGrid2.Rows(I).Cells(4).Text) + CDbl(UltraGrid2.Rows(I).Cells(5).Text)) & "','A')"
                '            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '            nvcFieldList1 = "Insert Into T09Sales_Flutter(T09Inv_No,T09Department,T09Item_Code,T09Item_Name,T09Cost,T09Retail,T09Qty,T09Discount,T09Free,T09Status)" & _
                '                                  " values('" & txtEnter1.Text & "','-','" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(1).Text) & "','" & M02.Tables(0).Rows(0)("M05Cost") & "', '" & Trim(UltraGrid2.Rows(I).Cells(2).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(3).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(4).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(5).Text) & "','A')"
                '            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '        End If
                '        I = I + 1
                '    Next

                '    nvcFieldList1 = "DELETE FROM T13Chq_Transaction WHERE T13Ref_No='" & Trim(txtEnter1.Text) & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '    nvcFieldList1 = "DELETE FROM T14Credit_Card_TR WHERE T14Inv_No='" & Trim(txtEnter1.Text) & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '    'CHQ TRANSACTION
                '    If Trim(frmPayMain_DS.cboTearms.Text) = "CHQUE" Then
                '        If SEARCH_CUSTOMER() = True Then
                '        Else
                '            MsgBox("Please select the customer name", MsgBoxStyle.Information, "Information ......")
                '            connection.ClearAllPools()
                '            connection.Close()
                '            Exit Function
                '        End If
                '    End If
                '    I = 0
                '    With frmPayMain_DS
                '        For Each uRow As UltraGridRow In .UltraGrid2.Rows
                '            nvcFieldList1 = "Insert Into T13Chq_Transaction(T13Ref_No,T13Cus_Code,T13Date,T13Bank,T13Chq_No,T13DOR,T13Amount,T13Return_Status,T13Status,T13Tr_Type)" & _
                '                                          " values('" & Trim(txtEnter1.Text) & "','" & _CusNo & "','" & _GetDate & "','" & .UltraGrid2.Rows(I).Cells(0).Text & "','" & .UltraGrid2.Rows(I).Cells(1).Text & "','" & .UltraGrid2.Rows(I).Cells(2).Text & "','" & .UltraGrid2.Rows(I).Cells(3).Text & "','-','A','JOB_INVOICE')"
                '            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '            I = I + 1
                '        Next
                '    End With
                '    '===================================================================================================
                '    'CREDIT CARD TRANSACTION
                '    I = 0
                '    With frmPayMain_DS
                '        For Each uRow As UltraGridRow In .UltraGrid1.Rows
                '            nvcFieldList1 = "Insert Into T14Credit_Card_TR(T14Inv_No,T14Date,T14Type,T14Card_No,T14Amount,T14Status)" & _
                '                                          " values('" & Trim(txtEnter1.Text) & "','" & _GetDate & "','" & .UltraGrid1.Rows(I).Cells(0).Text & "','" & .UltraGrid1.Rows(I).Cells(1).Text & "','" & .UltraGrid1.Rows(I).Cells(2).Text & "','A')"
                '            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                '            I = I + 1
                '        Next
                '    End With
                '    '==================================================================================================

                '    nvcFieldList1 = "UPDATE T04Profit_Loss SET T04Date='" & _GetDate & "',T04Cr='" & CDbl(txtNet1.Text) & "',T04Status='A' WHERE T04Ref_No='" & Trim(txtEnter1.Text) & "'"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                '    ' MsgBox("Invoice Change successfully", MsgBoxStyle.Information, "Information ........")
                'Else
                '    OPRUser.Visible = True
                '    txtPassword.Text = ""
                '    txtUserName.Text = ""
                '    txtUserName.Focus()
                '    _LogStaus = False
                '    connection.Close()
                '    Exit Function
                'End If
            Else

            Call Load_EntryNo()
            If Search_Vehicle_No() = True Then

            Else

                If SEARCH_CUSTOMER() = True Then
                    nvcFieldList1 = "Insert Into M07Vehicle_Master(M07V_No,M07Cus_Code,M07Type,M07Brand_Name,M07Status)" & _
                                                         " values('" & UCase(Trim(cboV_no.Text)) & "','" & _CusNo & "', '-','-','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                Else
                    nvcFieldList1 = "select * from P01Parameter where P01Code='CU'"
                    M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                    If isValidDataset(M01) Then
                        If M01.Tables(0).Rows(0)("P01No") >= 1 And M01.Tables(0).Rows(0)("P01No") < 10 Then
                            _CusNo = "CU-00" & M01.Tables(0).Rows(0)("P01No")
                        ElseIf M01.Tables(0).Rows(0)("P01No") >= 10 And M01.Tables(0).Rows(0)("P01No") < 100 Then
                            _CusNo = "CU-0" & M01.Tables(0).Rows(0)("P01No")
                        Else
                            _CusNo = "CU-" & M01.Tables(0).Rows(0)("P01No")
                        End If
                        nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='CU' "
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        nvcFieldList1 = "Insert Into M06Customer_Master(M06Code,M06Name,M06Address,M06Contact_No,M06Mobile_No,M06Email,M06Cus_Type,M06Credit_Limit,M06Status)" & _
                                              " values('" & _CusNo & "','" & Trim(cboCustomer.Text) & "', '" & Trim(txtAddress1.Text) & "','" & Trim(txtTP1.Text) & "','-','-','Private','0','A')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    'SAVE VEHICLE DETAILES
                    nvcFieldList1 = "Insert Into M07Vehicle_Master(M07V_No,M07Cus_Code,M07Type,M07Brand_Name,M07Status)" & _
                                                             " values('" & UCase(Trim(cboV_no.Text)) & "','" & _CusNo & "', '-','-','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                End If
            End If
            nvcFieldList1 = "update P01Parameter set P01No=P01No+ " & 1 & " where P01Code='" & _CounterPC & "-D" & "' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'SAVE INVOICE 
            nvcFieldList1 = "Insert Into T08Sales_Header(T08Invo_No,T08Job_No,T08Tr_Type,T08Service_on,T08St_Mtr,T08End_mtr,T08Date,T08V_No,T08Cus_NO,T08Status,T08Net_Amount,T08Time)" & _
                                                     " values('" & Trim(txtEnter1.Text) & "','-', 'DIRECT_SALE','-','" & txtMtr.Text & "','0','" & _GetDate & "','" & Trim(cboV_no.Text) & "','" & _CusNo & "','A','" & CDbl(txtNet1.Text) & "','" & _Get_Time & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                                " values('DIRECT_SALES','SAVE', '" & _Get_Time & "','" & strDisname & "','" & Trim(txtEnter1.Text) & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            '===============================================================================================
            I = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                    nvcFieldList1 = "select * from M05Item_Master where M05Ref_No='" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "'"
                M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M02) Then
                    nvcFieldList1 = "Insert Into S01Stock_Balance(S01Item_Code,S01Ref_No,S01Date,S01Time,S01Tr_Type,S01Qty,S01Status)" & _
                                             " values('" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "','" & txtEnter1.Text & "', '" & _GetDate & "','" & _Get_Time & "','DIRECT_SALES','" & -(CDbl(UltraGrid2.Rows(I).Cells(4).Text) + CDbl(UltraGrid2.Rows(I).Cells(5).Text)) & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                        nvcFieldList1 = "Insert Into T09Sales_Flutter(T09Inv_No,T09Department,T09Item_Ref,T09Item_Code,T09Item_Name,T09Cost,T09Retail,T09Qty,T09Discount,T09Free,T09Status)" & _
                                                          " values('" & Trim(txtEnter1.Text) & "','-','" & Trim(UltraGrid2.Rows(I).Cells(0).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(1).Text) & "','" & Trim(UltraGrid2.Rows(I).Cells(2).Text) & "','" & Trim(M02.Tables(0).Rows(0)("M05Cost")) & "','" & Trim(Trim(UltraGrid2.Rows(I).Cells(3).Text)) & "','" & Trim(Trim(UltraGrid2.Rows(I).Cells(5).Text)) & "','" & Trim(Trim(UltraGrid2.Rows(I).Cells(4).Text)) & "','" & Trim(Trim(UltraGrid2.Rows(I).Cells(6).Text)) & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
                I = I + 1
            Next



            '=================================================================================================
            If CDbl(frmPayMain_DS.txtBalance.Text) < 0 Then
                frmPayMain_DS.txtBalance.Text = "0"
            End If
            'PAY HEADER
            nvcFieldList1 = "select * from T11Income_Summery where T11Invo_No='" & Trim(txtEnter1.Text) & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
            Else
                nvcFieldList1 = "Insert Into T11Income_Summery(T11Invo_No,T11Job_No,T11Tr_Type,T11Date,T11Net_Amount,T11Cash,T11Chq,T11Card,T11Credit,T11Status)" & _
                                               " values('" & Trim(txtEnter1.Text) & "','-','DIRECT_SALE','" & _GetDate & "','" & CDbl(txtNet1.Text) & "','" & CDbl(frmPayMain_DS.txtCash.Text) & "','" & CDbl(frmPayMain_DS.txtChq_Total.Text) & "','" & CDbl(frmPayMain_DS.txtTotal.Text) & "','" & CDbl(frmPayMain_DS.txtBalance.Text) & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
            '===================================================================================================
            'OUTSTANDING 
            If CDbl(frmPayMain_DS.txtBalance.Text) > 0 Then
                If SEARCH_CUSTOMER() = True Then
                Else
                    MsgBox("Please select the customer name", MsgBoxStyle.Information, "Information .......")
                    connection.ClearAllPools()
                    connection.Close()
                    Exit Function
                End If
                nvcFieldList1 = "select * from T12OutStanding where T12Inv_No='" & Trim(txtEnter1.Text) & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                Else
                    nvcFieldList1 = "Insert Into T12OutStanding(T12Inv_No,T12Cus_No,T12Date,T12Pay_No,T12Chq_No,T12Cr,T12Dr,T12Status)" & _
                                                   " values('" & Trim(txtEnter1.Text) & "','" & _CusNo & "','" & _GetDate & "','-','-','" & CDbl(frmPayMain_DS.txtBalance.Text) & "','0','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If
            End If
            '===================================================================================================
            'CHQ TRANSACTION
            I = 0
            With frmPayMain_DS
                For Each uRow As UltraGridRow In .UltraGrid2.Rows
                    nvcFieldList1 = "Insert Into T13Chq_Transaction(T13Ref_No,T13Cus_Code,T13Date,T13Bank,T13Chq_No,T13DOR,T13Amount,T13Return_Status,T13Status,T13Tr_Type)" & _
                                                  " values('" & Trim(txtEnter1.Text) & "','" & _CusNo & "','" & _GetDate & "','" & .UltraGrid2.Rows(I).Cells(0).Text & "','" & .UltraGrid2.Rows(I).Cells(1).Text & "','" & .UltraGrid2.Rows(I).Cells(2).Text & "','" & .UltraGrid2.Rows(I).Cells(3).Text & "','-','A','DIRECT_SALES')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    I = I + 1
                Next
            End With
            '===================================================================================================
            'CREDIT CARD TRANSACTION
            I = 0
            With frmPayMain_DS
                For Each uRow As UltraGridRow In .UltraGrid1.Rows
                    nvcFieldList1 = "Insert Into T14Credit_Card_TR(T14Inv_No,T14Date,T14Type,T14Card_No,T14Amount,T14Status)" & _
                                                  " values('" & Trim(txtEnter1.Text) & "','" & _GetDate & "','" & .UltraGrid1.Rows(I).Cells(0).Text & "','" & .UltraGrid1.Rows(I).Cells(1).Text & "','" & .UltraGrid1.Rows(I).Cells(2).Text & "','A')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    I = I + 1
                Next
            End With
            '==================================================================================================
            'DAILY INCOME
            _SERVICE_AMOUNT = CDbl(frmPayMain_DS.txtCash.Text) + CDbl(frmPayMain_DS.txtTotal.Text)

            If _SERVICE_AMOUNT > 0 Then
                _Remark = "Direct Sales Income -" & Trim(txtEnter1.Text)

                nvcFieldList1 = "Insert Into T04Profit_Loss(T04Date,T04Tr_Type,T04Ref_No,T04Description,T04Cr,T04Dr,T04Status)" & _
                                                 " values('" & _GetDate & "','DIRECT_SALE','" & txtEnter1.Text & "','" & _Remark & "','" & _SERVICE_AMOUNT & "','0','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If
                End If
            transaction.Commit()
            A = MsgBox("Are you sure you want to print this invoice", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information ..........")
            If A = vbYes Then

            End If
                connection.ClearAllPools()
                connection.Close()
                Call CLEAR_TEXT()
            cboV_no.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                SqlClient.SqlConnection.ClearAllPools()
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Function CLEAR_TEXT()
        Me.txtFree.Text = ""
        Me.txtTP1.Text = ""
        Me.txtAddress1.Text = ""
        Me.cboCustomer.Text = ""
        Me.txtRemark.Text = ""
        Me.cboPartNo.Text = ""
        Me.txtItem_Name.Text = ""
        Me.txtRate1.Text = ""
        Me.txtQty.Text = ""
        Me.txtDiscount.Text = ""
        Me.txtNet1.Text = ""
        Me.txtCount1.Text = ""
        Me.cboV_no.Text = ""
        Me.txtMtr.Text = ""
        strItem_Code = ""
        Me.cmdDelete1.Enabled = True
        Call Load_cUSTOMER()
        Call Load_VNO()
        _LogStaus = False
        Call Load_EntryNo()
        Call Load_Gride2()
        UltraGrid1.Visible = False
    End Function

    Private Sub ItemLookupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemLookupToolStripMenuItem.Click
        strWindowName = Me.Name
        frmView_Item_Uniq.Close()
        frmView_Item_Uniq.Show()
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call CLEAR_TEXT()
    End Sub

    Private Sub DeactivateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeactivateToolStripMenuItem.Click
        frmView_Direct_Sales.Close()
        frmView_Direct_Sales.Show()
    End Sub

    Private Sub cmdDelete1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete1.Click
        If _LogStaus = True Then
            Call Delete_Records()
        Else
            OPRUser.Visible = True
            txtPassword.Text = ""
            txtUserName.Text = ""
            txtUserName.Focus()
        End If
    End Sub

    Function Delete_Records()
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
        Dim _GetDate As DateTime
        Dim _Get_Time As DateTime
        Dim A As String
        Dim B As New ReportDocument

        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim I As Integer
        Dim _CREDIT As Double

        Dim _SERVICE_AMOUNT As Double
        Dim _Remark As String

        Try


            _GetDate = Month(txtDate1.Text) & "/" & Microsoft.VisualBasic.Day(txtDate1.Text) & "/" & Year(txtDate1.Text)

            _Get_Time = Month(Now) & "/" & Microsoft.VisualBasic.Day(Now) & "/" & Year(Now) & " " & Hour(Now) & ":" & Minute(Now)
            A = MsgBox("Are you sure you want to delete this records", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Delete Invoice ...........")
            If A = vbYes Then
                nvcFieldList1 = "SELECT * FROM T15Outstanding_Collection WHERE T15Inv_No='" & Trim(txtEnter1.Text) & "' AND T15Status='PAID'"
                M02 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M02) Then
                    MsgBox("This invoice alrady paid.can't Delete", MsgBoxStyle.Information, "Information .........")
                    connection.ClearAllPools()
                    connection.Close()
                    Exit Function
                End If
                '========================================================================
                nvcFieldList1 = "UPDATE T08Sales_Header SET T08status='CANCEL' WHERE T08Invo_No='" & txtEnter1.Text & "' AND T08Tr_Type='DIRECT_SALE'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmpStatus,tmpProcess,tmpTime,tmpUser,tmpCode)" & _
                                            " values('DIRECT_SALE','CANCEL', '" & Now & "','" & strDisname & "','" & Trim(txtEnter1.Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T11Income_Summery SET T11Status='CANCEL' WHERE T11Invo_No='" & Trim(txtEnter1.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T12OutStanding SET T12Status='CANCEL' WHERE T12Inv_No='" & Trim(txtEnter1.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T09Sales_Flutter SET T09Status='CANCEL' WHERE T09Inv_No='" & Trim(txtEnter1.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S01Stock_Balance SET S01Status='CANCEL' WHERE S01Ref_No='" & Trim(txtEnter1.Text) & "' AND S01Tr_Type='DIRECT_SALES'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T13Chq_Transaction SET T13Status='CANCEL' WHERE T13Ref_No='" & Trim(txtEnter1.Text) & "' AND T13Tr_Type='DIRECT_SALES'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T14Credit_Card_TR SET T14Status='CANCEL' WHERE T14Inv_No='" & Trim(txtEnter1.Text) & "' AND T14Type='DIRECT_SALES'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE T04Profit_Loss SET T04Status='CANCEL' WHERE T04Ref_No='" & Trim(txtEnter1.Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records deleted successfully", MsgBoxStyle.Information, "Information .......")
            End If

           

            transaction.Commit()
            connection.ClearAllPools()
            connection.Close()
            Call CLEAR_TEXT()
            cboCustomer.ToggleDropdown()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)

                SqlClient.SqlConnection.ClearAllPools()
                connection.ClearAllPools()
                connection.Close()
            End If
        End Try
    End Function

    Private Sub txtUserName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.KeyUp
        If e.KeyCode = 13 Then
            txtPassword.Focus()
        End If
    End Sub

  

    Private Sub txtPassword_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyUp
        If e.KeyCode = 13 Then
            OK.Focus()
        End If
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim SQL As String
        Dim con = New SqlConnection()
        Dim M01 As DataSet

        Try
            SqlConnection.ClearAllPools()
            con = DBEngin.GetConnection()
            SQL = "SELECT * FROM users WHERE (NAME ='" & txtUserName.Text & "')and Password='" & txtPassword.Text & "' and UType in ('ADMIN','Manger','Accountant','MD') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(M01) Then
                _LogStaus = True
                ' _permissionLevel = Trim(M01.Tables(0).Rows(0)("UType"))
                ' _AthzUser = Trim(txtUserName.Text)
                OPRUser.Visible = False
                UltraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True
            Else
                MsgBox("User name and pasword combination not found", "Information ......")
                txtUserName.Focus()
                con.ClearAllPools()
                con.CLOSE()
                Exit Sub
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
    End Sub

    Private Sub cboPartNo_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboPartNo.InitializeLayout

    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call CLEAR_TEXT()
    End Sub

    Private Sub cboV_no_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboV_no.AfterCloseUp
        Call Search_Vehicle_No()
    End Sub

    Private Sub cboV_no_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboV_no.KeyUp
        If e.KeyCode = 13 Then
            Call Search_Vehicle_No()
            txtMtr.Focus()
        End If
    End Sub

    Private Sub txtMtr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMtr.KeyUp
        If e.KeyCode = 13 Then
            If Trim(cboCustomer.Text) <> "" Then
                cboPartNo.ToggleDropdown()
            Else
                cboCustomer.ToggleDropdown()
            End If
        End If
    End Sub

    Private Sub txtTP1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTP1.KeyUp
        If e.KeyCode = 13 Then
            cboPartNo.ToggleDropdown()
        End If
    End Sub

    Private Sub cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel.Click
        txtPassword.Text = ""
        txtUserName.Text = ""
        OPRUser.Visible = False
        _LogStaus = False
    End Sub

    Private Sub cboPartNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPartNo.TextChanged
        If UltraGrid1.Visible = True Then
            Call Load_Grid_ROW()
        Else
            UltraGrid1.Visible = True
            Call Load_Grid_ROW()
        End If
    End Sub

    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        Dim _Row As Integer

        _Row = UltraGrid1.ActiveRow.Index
        strItem_Code = Trim(UltraGrid1.Rows(_Row).Cells(0).Text)
        Call Search_Item()
        UltraGrid1.Visible = False
        txtQty.Focus()
    End Sub

    Private Sub UltraGrid1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid1.KeyUp
        Dim _Row As Integer
        If e.KeyCode = 13 Then
            _Row = UltraGrid1.ActiveRow.Index
            strItem_Code = Trim(UltraGrid1.Rows(_Row).Cells(0).Text)
            Call Search_Item()
            UltraGrid1.Visible = False
            txtQty.Focus()
        End If
    End Sub
End Class