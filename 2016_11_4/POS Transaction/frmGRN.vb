Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine


Public Class frmGRN
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _Category As String
    Dim _Comcode As String
    Dim _Loccode As String
    Dim _SupCode As String
    Dim _SupCode1 As String
    Dim _ItemLoc As String
    Dim _RefNo As Integer
    Dim _LogStaus As Boolean
    Dim _UserLevel As String
    Dim _AthzUser As String

    Function Search_Supplier() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M01Account_Master where M01Acc_Name='" & Trim(cboLocation.Text) & "'  and M01Acc_Type='SP' and M01Com_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                _Supcode = dsUser.Tables(0).Rows(0)("M01Acc_Code")
                Search_Supplier = True
            Else
                Search_Supplier = False
            End If

            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Acc_Name as [From Location] from M01Account_Master where  M01Acc_Type='SP' and M01Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboLocation
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 270
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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Load_Category()
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

    Function Load_Supplier_1()
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

    Private Sub frmGRN_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmViewGRN.Close()
    End Sub

    Private Sub frmGRN_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtMRP.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        _Comcode = ConfigurationManager.AppSettings("LOCCode")
        Call Load_Gride_Item()
        ' txtEx_Date.Text = Today
        txtDate.Text = Today
        txtRate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDis_Rate.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        ' txtPO.ReadOnly = True
        Call Load_Location()
        txtReorder.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCost.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtRetail.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtSales.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtFree.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtFree_Amount.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtGross.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtMarket.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNett.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        'txtRe_Qty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtEntry.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtVAT.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtNBT.Appearance.TextHAlign = Infragistics.Win.HAlign.Right

        txtEntry.ReadOnly = True
        txtCount.ReadOnly = True
        txtTotal.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtTotal.ReadOnly = True

        Call Load_Combo()
        Call Load_EntryNo()
        Call Load_ToLocation()
        Call Load_Item_Code()
        Call Load_Item_Name()

        Call Load_Gride2()
        Call Load_Data()


        Call Load_Category()
        Call Load_ComboEX()
        Call Load_Supplier_1()

    End Sub

    Private Sub cboMain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboMain.KeyUp
        If e.KeyCode = 13 Then
            cboSupplier.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboSupplier.ToggleDropdown()
        End If
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
        ElseIf e.KeyCode = Keys.Tab Then

            Value = txtCost.Text
            txtCost.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtCost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            txtRetail.Focus()

        End If
    End Sub
    Function Load_Location()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [##] from M04Location WHERE M04Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItem_Loc
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


    Function Search_ItemLocation() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M04Location where M04Loc_Name='" & Trim(cboItem_Loc.Text) & "' AND M04Com_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                _ItemLoc = dsUser.Tables(0).Rows(0)("M04Loc_Code")
                Search_ItemLocation = True
            Else
                Search_ItemLocation = False
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

    Private Sub txtRetail_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRetail.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then

            Value = txtRetail.Text
            txtRetail.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            txtQty.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Value = txtRetail.Text
            txtRetail.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            txtQty.Focus()
        End If
    End Sub

    Private Sub cboSupplier_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSupplier.KeyUp
        If e.KeyCode = 13 Then
            txtCode.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtCode.Focus()

        End If
    End Sub

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
            Sql = "select * from M03Item_Master inner join M02Category on M03Cat_Code=M02Cat_Code inner join M01Account_Master on M01Acc_Code=M03Supplier where M03Item_Code='" & Trim(txtCode.Text) & "'  and M01Acc_Type='SP' and M03Com_Code='" & _Comcode & "'"
            M04 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M04) Then
                txtDescription.Text = M04.Tables(0).Rows(0)("M03Item_Name")
                cboSupplier.Text = M04.Tables(0).Rows(0)("M01Acc_Name")
                cboEx_Date.Text = M04.Tables(0).Rows(0)("M03ExPair")
                txtReorder.Text = M04.Tables(0).Rows(0)("M03Reorder")
                cboMain.Text = M04.Tables(0).Rows(0)("M02Cat_Name")

                Value = M04.Tables(0).Rows(0)("M03Cost_Price")
                txtCost.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtCost.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M04.Tables(0).Rows(0)("M03Retail_Price")
                txtRetail.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRetail.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

               
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

    Private Sub cboEx_Date_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboEx_Date.KeyUp
        If e.KeyCode = 13 Then
            txtCost.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtCost.Focus()
        End If
    End Sub
    Function Load_EntryNo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='GR'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                If M01.Tables(0).Rows(0)("P01LastNo") >= 1 And M01.Tables(0).Rows(0)("P01LastNo") < 10 Then
                    txtEntry.Text = "GRN-00" & M01.Tables(0).Rows(0)("P01LastNo")
                ElseIf M01.Tables(0).Rows(0)("P01LastNo") >= 10 And M01.Tables(0).Rows(0)("P01LastNo") < 100 Then
                    txtEntry.Text = "GRN-0" & M01.Tables(0).Rows(0)("P01LastNo")
                Else
                    txtEntry.Text = "GRN-" & M01.Tables(0).Rows(0)("P01LastNo")
                End If
            End If

            'Sql = "select * from M04Location"
            'M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            'If isValidDataset(M01) Then
            '    cboTo.Text = M01.Tables(0).Rows(0)("M04Loc_Name")
            'End If
            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function

    Function Load_EntryNo1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select * from P01Parameter where  P01Code='IN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtEntry.Text = M01.Tables(0).Rows(0)("P01LastNo")
            End If

            Sql = "select * from M04Location"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboTo.Text = M01.Tables(0).Rows(0)("M04Loc_Name")
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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Load_ToLocation()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M04Loc_Name as [To Location] from M04Location WHERE M04Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboTo
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 270
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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Load_Item_Code()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Code as [Item Code] from M03Item_Master where M03Status='A' AND M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCode
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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Load_Item_Name()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M03Item_Name as [Item Name] from M03Item_Master where M03status='A' AND M03Com_Code='" & _Comcode & "' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboItemName
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 370
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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_ItemCode()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double

        Try
            Sql = "select * from M03Item_Master where M03Item_Name='" & Trim(cboItemName.Text) & "' and M03status='A' AND M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboCode.Text = M01.Tables(0).Rows(0)("M03Item_Code")
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M01.Tables(0).Rows(0)("M03Retail_Price")
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M01.Tables(0).Rows(0)("M03MRP")
                txtMRP.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMRP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_ItemName() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Try
            Sql = "select * from M03Item_Master where  M03Item_Code='" & Trim(cboCode.Text) & "' and M03status='A' AND M03Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_ItemName = True
                cboItemName.Text = M01.Tables(0).Rows(0)("M03Item_Name")
                Value = M01.Tables(0).Rows(0)("M03Cost_Price")
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M01.Tables(0).Rows(0)("M03Retail_Price")
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = M01.Tables(0).Rows(0)("M03MRP")
                txtMRP.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMRP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub cboItemName_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItemName.AfterCloseUp
        Call Search_ItemCode()
    End Sub

    Private Sub cboCode_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCode.AfterCloseUp
        Call Search_ItemName()
    End Sub


    Private Sub cboCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCode.KeyUp
        If e.KeyCode = Keys.Enter Then
            If cboCode.Text <> "" Then
                Call Search_ItemName()
                txtRate.Focus()
            Else
                txtVAT.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_ItemName()
            txtRate.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            Panel1.Visible = False
            GRP1.Visible = False
        ElseIf e.KeyCode = Keys.F2 Then
            OPR5.Visible = True
            txtFind.Text = ""
            Call Load_Gride_Item()
            txtFind.Focus()
            Panel1.Visible = False
            GRP1.Visible = False
        ElseIf e.KeyCode = Keys.Escape Then
            OPR4.Visible = False
            OPR5.Visible = False
            Panel1.Visible = False
            GRP1.Visible = False
        ElseIf e.KeyCode = Keys.F5 Then
            'Call Clear_Item()
            'Panel1.Visible = True
            'GRP1.Visible = True
            'cboMain.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Function Clear_Item()
        Me.cboMain.Text = ""
        Me.cboSupplier.Text = ""
        Me.txtReorder.Text = ""
        Me.txtCost.Text = ""
        Me.txtRetail.Text = ""
        Me.cboEx_Date.Text = ""
        Me.txtCode.Text = ""
        Me.txtDescription.Text = ""

    End Function

    Private Sub cboTo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboTo.KeyUp
        If e.KeyCode = Keys.Enter Then
            txtRemark.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtRemark.Focus()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Private Sub txtRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRate.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtRate.Text) Then
                Value = txtRate.Text
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtSales.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If IsNumeric(txtRate.Text) Then
                Value = txtRate.Text
                txtRate.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtRate.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtSales.Focus()
            End If
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Function Calculation()
        Dim Value As Double

        If IsNumeric(txtRate.Text) Then
            If IsNumeric(txtQty.Text) Then
                Value = CDbl(txtRate.Text) * CDbl(txtQty.Text)
                txtTotal.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtTotal.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
        Else
            txtTotal.Text = ""
        End If
    End Function
  
    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Dim result1 As String

        If e.KeyCode = Keys.Enter Then
            Call Calculation()
            If IsNumeric(txtQty.Text) Then
                'txtRe_Qty.Text = txtQty.Text
            Else
                'result1 = MessageBox.Show("Please enter the Correct Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'If result1 = Windows.Forms.DialogResult.OK Then
                '    txtQty.Focus()
                '    Exit Sub
                'End If
            End If
            txtFree.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Calculation()
            If IsNumeric(txtQty.Text) Then
                'txtRe_Qty.Text = txtQty.Text
            Else
                'result1 = MessageBox.Show("Please enter the Correct Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'If result1 = Windows.Forms.DialogResult.OK Then
                '    txtQty.Focus()
                '    Exit Sub
                'End If
            End If
            txtFree.Focus()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableGRN
        UltraGrid1.DataSource = c_dataCustomer1
        With UltraGrid1
            .DisplayLayout.Bands(0).Columns(0).Width = 90
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 70
            .DisplayLayout.Bands(0).Columns(5).Width = 70
            .DisplayLayout.Bands(0).Columns(6).Width = 80
            .DisplayLayout.Bands(0).Columns(7).Width = 90
            ' .DisplayLayout.Bands(0).Columns(8).Width = 90

            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(7).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(6).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(1).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(2).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(3).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(4).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(5).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(6).CellActivation = Activation.NoEdit
            .DisplayLayout.Bands(0).Columns(7).CellActivation = Activation.NoEdit


            .DisplayLayout.Bands(0).Columns(0).CellActivation = Activation.NoEdit
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Private Sub txtFree_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFree.KeyUp
        If e.KeyCode = 13 Then
            txtMRP.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtMRP.Focus()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If

    End Sub


    Private Sub txtRe_Qty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            txtFree.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtFree.Focus()
        End If
    End Sub

    Private Sub txtTotal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTotal.KeyUp
        Dim result1 As String
        Dim Value As Double
        Dim SQL As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim T01 As DataSet
        Dim _CostStatus As Boolean
        Try
            If e.KeyCode = 13 Then
                If IsNumeric(txtMRP.Text) Then
                Else
                    MsgBox("Please enter the MRP", MsgBoxStyle.Information, "Information .....")
                    con.close()
                    Exit Sub
                End If
                Call Calculation()
                'If Trim(txtRe_Qty.Text) <> "" Then
                'Else
                '    If IsNumeric(txtRe_Qty.Text) Then

                '        result1 = MessageBox.Show("Please enter the Correct Rec.Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                '        If result1 = Windows.Forms.DialogResult.OK Then
                '            txtRe_Qty.Focus()
                '            Exit Sub
                '        End If
                '    End If
                'End If

                If txtFree.Text <> "" Then
                Else
                    txtFree.Text = "0"
                End If

                If Trim(txtQty.Text) <> "" Then

                    If IsNumeric(txtQty.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the Correct Qty", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtQty.Focus()
                            Exit Sub
                        End If
                    End If
                End If

                If txtRetail.Text <> "" Then
                Else
                    txtRetail.Text = "0"
                End If

                If Trim(txtQty.Text) <> "" Then

                    If IsNumeric(txtQty.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the Correct Retail Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtRetail.Focus()
                            Exit Sub
                        End If
                    End If
                End If

                If Trim(txtFree.Text) <> "" Then

                    If IsNumeric(txtFree.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the Correct Free Issue", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtFree.Focus()
                            Exit Sub
                        End If
                    End If
                End If

                If txtCost.Text <> "" Then
                Else
                    txtCost.Text = "0"
                End If
                If Trim(txtRate.Text) <> "" Then

                    If IsNumeric(txtRate.Text) Then
                    Else
                        result1 = MessageBox.Show("Please enter the Correct cost price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If result1 = Windows.Forms.DialogResult.OK Then
                            txtRate.Focus()
                            Exit Sub
                        End If
                    End If
                Else
                    result1 = MessageBox.Show("Please enter the Correct cost price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtRate.Focus()

                        Exit Sub
                    End If
                End If

                Dim newRow As DataRow = c_dataCustomer1.NewRow
                newRow("Item Code") = Trim(cboCode.Text)
                newRow("Item Name") = cboItemName.Text
                newRow("Cost Price") = txtRate.Text
                newRow("Retail Price") = txtSales.Text
                newRow("Qty") = txtQty.Text
                ' newRow("Rec.Qty") = txtRe_Qty.Text
                newRow("Free Issue") = txtFree.Text
                newRow("Total") = txtTotal.Text

                Value = txtMRP.Text
                ' Value = M01.Tables(0).Rows(0)("M03MRP")
                txtMRP.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMRP.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("MRP") = txtMRP.Text

                SQL = "select * from M03Item_Master where M03Item_Code='" & cboCode.Text & "' and m03Status='A' and LEFT(M03ExPair,1)='Y' AND M03Com_Code='" & _Comcode & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then
                  

                    If CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) <> CDbl(txtRate.Text) Then
                        result1 = MsgBox("Previous cost price is Rs." & CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) & ".do you want to change new one", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Information ........")
                        If result1 = vbYes Then
                            newRow("##") = True
                            _CostStatus = True
                        Else
                            newRow("##") = False
                        End If
                    End If
                End If

                SQL = "select * from M03Item_Master where M03Item_Code='" & cboCode.Text & "' and m03Status='A' AND M03Com_Code='" & _Comcode & "' "
                T01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
                If isValidDataset(T01) Then

                    If CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) <> CDbl(txtRate.Text) Then
                        result1 = MsgBox("Previous cost price is Rs." & CDbl(T01.Tables(0).Rows(0)("M03Cost_Price")) & ".do you want to change new one", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Information ........")
                        If result1 = vbYes Then
                            newRow("##") = True
                        Else
                            newRow("##") = False
                        End If

                    Else
                        newRow("##") = False
                    End If
                Else
                    MsgBox("Please enter the correct item code", MsgBoxStyle.Information, "Information .......")
                    cboCode.Focus()
                    Exit Sub
                End If
                c_dataCustomer1.Rows.Add(newRow)

                If _CostStatus = True Then
                    Dim _lastRow As Integer

                    _lastRow = UltraGrid1.Rows.Count - 1
                    UltraGrid1.Rows(_lastRow).Cells(8).Value = True
                End If
                txtCount.Text = Val(txtCount.Text) + 1

                If txtNett.Text <> "" Then
                Else
                    txtNett.Text = "0"
                End If
                'Value = Double.TryParse(txtNett.Text, Value) + Double.TryParse(txtTotal.Text, Value)
                Value = CDbl(txtNett.Text) + CDbl(txtTotal.Text)

                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                'If txtFree_Amount.Text <> "" Then
                'Else
                '    txtFree_Amount.Text = "0"
                'End If

                If txtMarket.Text <> "" Then
                Else
                    txtMarket.Text = "0"
                End If

                If txtVAT.Text <> "" Then
                Else
                    txtVAT.Text = "0"
                End If

                If txtDiscount.Text <> "" Then
                Else
                    txtDiscount.Text = "0"
                End If
                Value = Value - Val(txtMarket.Text) - Val(txtVAT.Text)
                Value = Value - Val(txtDiscount.Text)

                txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                cboItemName.Text = ""
                cboCode.Text = ""
                txtRate.Text = ""
                txtFree.Text = ""
                txtQty.Text = ""
                txtSales.Text = ""
                txtTotal.Text = ""
                Me.txtMRP.Text = ""
                Me.txtRetail.Text = ""
                'Me.txtEx_Date.Appearance.BackColor = Color.White
                ' txtEx_Date.Text = ""
                cboCode.Focus()
            ElseIf e.KeyCode = Keys.F3 Then
                frmViewGRN.Show()
            End If

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub

   

    Private Sub UltraGrid1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.AfterRowsDeleted
        Dim i As Integer
        Dim value As Double
        txtCount.Text = UltraGrid1.Rows.Count
        Try
            i = 0
            txtNett.Text = "0.00"
            txtGross.Text = "0.00"
            value = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Double.TryParse(txtNett.Text, value))
                value = value + CDbl((UltraGrid1.Rows(i).Cells(7).Value))
                i = i + 1
            Next

            txtNett.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))


            value = txtNett.Text
            value = value - (CDbl(txtVAT.Text) + CDbl(txtMarket.Text) + CDbl(txtDiscount.Text))
            txtGross.Text = (value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", value))
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub
    Function Calculate_Gross()

        Dim Value As Double
        Try
            If txtNett.Text <> "" Then
            Else
                txtNett.Text = "0"
            End If
            If txtMarket.Text <> "" Then
            Else
                txtMarket.Text = "0"
            End If

            If txtDiscount.Text <> "" Then
            Else
                txtDiscount.Text = "0"
            End If
            Value = txtNett.Text
            If IsNumeric(txtVAT.Text) Then
                Value = Value + CDbl(txtVAT.Text)
            End If

            If IsNumeric(txtNBT.Text) Then
                Value = Value + CDbl(txtNBT.Text)
            End If
            Value = Value - (CDbl(txtMarket.Text) + CDbl(txtDiscount.Text))
            txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub txtVAT_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVAT.KeyUp
        Dim Value As Double
        Dim result1 As String
        If e.KeyCode = 13 Then
            If Trim(txtVAT.Text) <> "" Then
                If IsNumeric(txtVAT.Text) Then
                    Call Calculate_Gross()
                    Value = txtVAT.Text
                    txtVAT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtVAT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    txtMarket.Focus()
                Else
                    result1 = MessageBox.Show("Please enter the VAT Amount", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtVAT.Focus()
                        Exit Sub
                    End If
                End If
            End If

        ElseIf e.KeyCode = Keys.Tab Then
            If Trim(txtVAT.Text) <> "" Then
                If IsNumeric(txtVAT.Text) Then
                    Call Calculate_Gross()
                    Value = txtVAT.Text
                    txtVAT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtVAT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    txtMarket.Focus()
                Else
                    result1 = MessageBox.Show("Please enter the VAT Amount", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtVAT.Focus()
                        Exit Sub
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub txtMarket_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMarket.KeyUp
        Dim Value As Double
        Dim result1 As String
        If e.KeyCode = 13 Then
            If Trim(txtMarket.Text) <> "" Then
                If IsNumeric(txtMarket.Text) Then
                    Call Calculate_Gross()
                    Value = txtMarket.Text
                    txtMarket.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtMarket.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    txtDis_Rate.Focus()
                Else
                    result1 = MessageBox.Show("Please enter the Market Return", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtMarket.Focus()
                        Exit Sub
                    End If
                End If
            End If

        ElseIf e.KeyCode = Keys.Tab Then

            If Trim(txtMarket.Text) <> "" Then
                If IsNumeric(txtMarket.Text) Then
                    Call Calculate_Gross()
                    Value = txtMarket.Text
                    txtMarket.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtMarket.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    txtDis_Rate.Focus()
                Else
                    result1 = MessageBox.Show("Please enter the Market Return ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtMarket.Focus()
                        Exit Sub
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub txtDis_Rate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDis_Rate.KeyUp
        Dim Value As Double
        Dim result1 As String

        If e.KeyCode = 13 Then
            If Trim(txtDis_Rate.Text) <> "" Then
                If IsNumeric(txtDis_Rate.Text) Then
                    Value = (CDbl(txtNett.Text) * CDbl(txtDis_Rate.Text)) / 100
                    txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    Call Calculate_Gross()
                Else
                    result1 = MessageBox.Show("Please enter the Discount Rate ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtDis_Rate.Focus()
                        Exit Sub
                    End If
                End If
            End If
            'txtFree_Amount.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            If Trim(txtDis_Rate.Text) <> "" Then
                If IsNumeric(txtDis_Rate.Text) Then
                    Value = (CDbl(txtNett.Text) * CDbl(txtDis_Rate.Text)) / 100
                    txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                Else
                    result1 = MessageBox.Show("Please enter the Discount Rate ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtDis_Rate.Focus()
                        Exit Sub
                    End If
                End If
            End If
            'txtFree_Amount.Focus()
        End If
    End Sub

    Private Sub txtDis_Rate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDis_Rate.ValueChanged

    End Sub

  


    Function Search_Location() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet

        Dim i As Integer
        Try
            Sql = "select * from M04Location where  M04Loc_Name='" & Trim(cboTo.Text) & "' AND M04Com_Code='" & _Comcode & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                _Loccode = Trim(M01.Tables(0).Rows(0)("M04Loc_Code"))
                Search_Location = True
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
        Dim Value As Double
        Dim A As String
        Dim A1 As String
        Dim B As New ReportDocument

        Try
            If txtPO.Text <> "" Then
            Else
                txtPO.Text = " "
            End If

            If txtNBT.Text <> "" Then
            Else
                txtNBT.Text = "0"
            End If

            If IsNumeric(txtNBT.Text) Then
            Else
                MsgBox("Please enter the Correct NBT Amount", MsgBoxStyle.Information, "Information .......")
                connection.Close()
                Exit Sub
            End If
            If Trim(txtDis_Rate.Text) <> "" Then
                If IsNumeric(txtDis_Rate.Text) Then
                    Value = (CDbl(txtNett.Text) * CDbl(txtDis_Rate.Text)) / 100
                    txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    Call Calculate_Gross()
                Else
                    result1 = MessageBox.Show("Please enter the Discount Rate", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtDis_Rate.Focus()
                        Exit Sub
                    End If
                End If
            End If

            If Trim(txtCom_Invoice.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the company invoice ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCom_Invoice.Focus()
                    Exit Sub
                End If
            End If

            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = " "
            End If

            If Search_Location() = True Then
            Else
                result1 = MessageBox.Show("Please Select the Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboTo.ToggleDropdown()
                    Exit Sub
                End If
            End If

            If Search_Supplier() = True Then
            Else
                result1 = MessageBox.Show("Please Select the From Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboLocation.ToggleDropdown()
                    Exit Sub
                End If
            End If

            '----------------------------------------------------------------------------------
            Call Load_EntryNo()
            If UltraGrid1.Rows.Count > 0 Then
            Else
                result1 = MessageBox.Show("Please enter the Items ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboCode.ToggleDropdown()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------------
            nvcFieldList1 = "select * from P01Parameter where P01Code='IN'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = M01.Tables(0).Rows(0)("P01LastNo")
            End If
            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='IN' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='GR' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'UPDATE T01 TRANSACTION
            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                Dim _Status As String
                Dim _Exdate As String

                _Exdate = " "
                'If (UltraGrid1.Rows(i).Cells(3).Value) = (UltraGrid1.Rows(i).Cells(4).Value) Then
                '    _Status = "OK"
                'Else
                _Status = "A"
                'End If

                'If IsDate((UltraGrid1.Rows(i).Cells(6).Value)) Then
                '    _Exdate = (UltraGrid1.Rows(i).Cells(6).Value)
                'End If
                nvcFieldList1 = "Insert Into T02Transaction_Flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Commition,T02Com_Discount,T02Qty,T02Rec_Qty,T02Free_Issue,T02Status,T02Item_Received,T02Com_Code,T02Total,T02MRP,T02Count)" & _
                                                              " values('" & _RefNo & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','0','0','" & (UltraGrid1.Rows(i).Cells(4).Value) & "','0','" & (UltraGrid1.Rows(i).Cells(5).Value) & "','A','" & _Status & "','" & _Comcode & "','" & CDbl(UltraGrid1.Rows(i).Cells(7).Value) & "','" & (UltraGrid1.Rows(i).Cells(6).Value) & "','" & i + 1 & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Ref_No,S01Com_Code,S01STATUS)" & _
                                                              " values('" & _Loccode & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','GRN','" & CDbl(UltraGrid1.Rows(i).Cells(4).Value) + CDbl(UltraGrid1.Rows(i).Cells(5).Value) & "','" & _RefNo & "','" & _Comcode & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S04Stock_Price(S04Location,S04Tr_Type,S04Date,S04Item_Code,S04Qty,S04Ref_No,S04Status,S04Rate,S04Com_Code,S04Cost)" & _
                                                               " values('" & _Loccode & "','GRN','" & txtDate.Text & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & CDbl(UltraGrid1.Rows(i).Cells(4).Text) + CDbl(UltraGrid1.Rows(i).Cells(5).Text) & "','" & _RefNo & "','A','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & _Comcode & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'EXPAIRE STOCK UPDATE
                'nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "' and m03Status='A' and M03ExPair='YES'"
                'M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                'If isValidDataset(M01) Then
                '    nvcFieldList1 = "Insert Into S03Ex_Stock(S03Loc_Code,S03Tr_Type,S03Item_Code,S03Qty,S03Ex_Date,S03Status,S03Ref_No)" & _
                '                                                " values('" & _Loccode & "','GRN', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & CDbl(UltraGrid1.Rows(i).Cells(4).Text) + CDbl(UltraGrid1.Rows(i).Cells(5).Text) & "','" & UltraGrid1.Rows(i).Cells(6).Text & "','A','" & _RefNo & "')"
                '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'End If

                If (UltraGrid1.Rows(i).Cells(8).Value) = True Then
                    nvcFieldList1 = "UPDATE M03Item_Master SET M03Cost_Price='" & (UltraGrid1.Rows(i).Cells(2).Value) & "',M03MRP='" & (UltraGrid1.Rows(i).Cells(6).Value) & "' WHERE M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "' AND M03Com_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If

                i = i + 1
            Next

            nvcFieldList1 = "Insert Into T01Transaction_Header(T01Trans_Type,T01Ref_No,T01Date,T01Invoice_No,T01FromLoc_Code,T01To_Loc_Code,T01Net_Amount,T01Com_Discount,T01DisRate,T01Vat,T01FreeIssue,T01Market_Return,T01Profit,T01Com_Code,T01Discount,T01User,T01Remark,T01PO_NO,T01Grn_No,T01Status,T01NBT)" & _
                                                              " values('GRN', '" & _RefNo & "','" & txtDate.Text & "','" & txtCom_Invoice.Text & "','" & _SupCode & "','" & _Loccode & "','" & CDbl(txtNett.Text) & "','" & txtDiscount.Text & "','" & txtDis_Rate.Text & "','" & txtVAT.Text & "','0','" & txtMarket.Text & "','0','" & _Comcode & "','0','" & strDisname & "','" & txtRemark.Text & "','" & txtPO.Text & "','" & txtEntry.Text & "','A','" & txtNBT.Text & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            '-------------------------------------------------------------------
            ''Account
            'Dim _StrRemark As String

            '_StrRemark = "Good Received invo.No -" & txtCom_Invoice.Text
            'nvcFieldList1 = "Insert Into T05Acc_Trans(T05Acc_Type,T05Ref_No,T05Date,T05Acc_No,T05Remark,T05Credit,T05Debit,T05Com_Code,T05User,T05Status)" & _
            '                                                 " values('SP', '" & txtEntry.Text & "','" & txtDate.Text & "','" & _SupCode & "','" & _StrRemark & "','0','" & CDbl(txtGross.Text) & "','" & _Comcode & "','" & strDisname & "','GR')"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'If CDbl(txtGross.Text) > 0 Then
            '    nvcFieldList1 = "Insert Into T07Supplier_Payment(T07RefNo,T07Date,T07InvoiceAmount,T07Paid_Amount,T07Status,T07Com_Code,T07Paid_Voucher)" & _
            '                                                 " values('" & txtEntry.Text & "','" & txtDate.Text & "','" & CDbl(txtGross.Text) & "','0','N','" & _SupCode & "','0')"
            '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            'End If
            'MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            A = MsgBox("Are you sure you want to print dispatch note", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Print Dispatch Note .....")
            If A = vbYes Then
                A1 = ConfigurationManager.AppSettings("ReportPath") + "\GRNDispatch.rpt"
                B.Load(A1.ToString)
                B.SetDatabaseLogon("sa", "tommya")
                'B.SetParameterValue("To", _To)
                'B.SetParameterValue("From", _From)
                '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
                frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
                frmReport.CrystalReportViewer1.DisplayToolbar = True
                frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Ref_No}=" & _RefNo & " "
                frmReport.Refresh()
                ' frmReport.CrystalReportViewer1.PrintReport()
                ' B.PrintToPrinter(1, True, 0, 0)
                frmReport.MdiParent = MDIMain
                frmReport.Show()
            End If

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0, OPR1, OPR2, OPR3)
            OPR2.Enabled = True
            OPR1.Enabled = True
            OPR0.Enabled = True
            OPR3.Enabled = True
            cmdAdd.Enabled = True
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cboLocation.ToggleDropdown()
            ' cmdSave.Enabled = False
            cmdDelete.Enabled = False
            '  Call Load_Gride()
            Call Load_Gride2()
            Call Load_EntryNo()
            Call Load_Combo()
            Call Load_Data()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        common.ClearAll(OPR0, OPR1, OPR2, OPR3)
        OPR2.Enabled = True
        OPR1.Enabled = True
        OPR0.Enabled = True
        OPR3.Enabled = True
        cmdAdd.Enabled = True
        cboLocation.ToggleDropdown()
        ' cmdSave.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        '  Call Load_Gride()
        Call Load_Gride2()
        Call Load_EntryNo()
        Call Load_Combo()
        OPRUser.Visible = False
        'Panel2.Visible = False
        _UserLevel = ""
        _LogStaus = False
    End Sub

    Function Search_RecordsUsing_Entry()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet
        Dim i As Integer
        Dim Value As Double


        Try
            Sql = "select * from T01Transaction_Header inner join T02Transaction_Flutter  on T01Ref_No=T02Ref_No  inner join M03Item_Master on T02Item_Code=M03Item_Code  inner join M04Location on M04Loc_Code=T01To_Loc_Code  inner join M09Supplier on M09Code=T01FromLoc_Code where T01Grn_No='" & Trim(txtEntry.Text) & "' and T01Trans_Type='GRN' and T01To_Loc_Code='" & _Comcode & "'"
            Sql = " select * from T01Transaction_Header inner join T02Transaction_Flutter  on T01Ref_No=T02Ref_No and T02Com_Code=T01Com_Code inner join M03Item_Master on T02Item_Code=M03Item_Code and M03Com_Code=T02Com_Code inner join M04Location on M04Loc_Code=T01To_Loc_Code inner join M09Supplier on M09Code=T01FromLoc_Code where T01Com_Code='" & _Comcode & "' and T01Grn_No='" & Trim(txtEntry.Text) & "' and T01Trans_Type='GRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboLocation.Text = Trim(M01.Tables(0).Rows(0)("M09Name"))
                cboTo.Text = Trim(M01.Tables(0).Rows(0)("M04Loc_Name"))
                txtCom_Invoice.Text = Trim(M01.Tables(0).Rows(0)("T01Invoice_No"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("T01Date"))
                txtEntry.Text = Trim(M01.Tables(0).Rows(0)("T01Grn_No"))
                txtRemark.Text = Trim(M01.Tables(0).Rows(0)("T01Remark"))
                _RefNo = Trim(M01.Tables(0).Rows(0)("T01Ref_No"))

                Value = Trim(M01.Tables(0).Rows(0)("T01Net_Amount"))
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Vat"))
                txtVAT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtVAT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01NBT"))
                txtNBT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNBT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Market_Return"))
                txtMarket.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMarket.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Com_Discount"))
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtDis_Rate.Text = Trim(M01.Tables(0).Rows(0)("T01DisRate"))

                'Value = Trim(M01.Tables(0).Rows(0)("T01FreeIssue"))
                'txtFree_Amount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtFree_Amount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = (CDbl(txtNett.Text) + CDbl(txtVAT.Text) + CDbl(txtNBT.Text)) - (CDbl(txtDiscount.Text) + Val(txtMarket.Text))
                txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtCount.Text = M01.Tables(0).Rows.Count

                Dim _St As String
                Call Load_Gride2()

                i = 0
                For Each DTRow2 As DataRow In M01.Tables(0).Rows

                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Item Code") = Trim(M01.Tables(0).Rows(i)("M03Item_Code"))
                    newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("M03Item_Name"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Cost Price") = _St
                    newRow("Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    'If IsDate(Trim(M01.Tables(0).Rows(i)("T02Ex_Date"))) Then
                    '    newRow("Ex Date") = Trim(M01.Tables(0).Rows(i)("T02Ex_Date"))
                    'End If
                    Value = Trim(M01.Tables(0).Rows(i)("T02MRP"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("MRP") = _St

                    Value = Trim(M01.Tables(0).Rows(i)("T02Retail_Price"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Retail Price") = _St
                    ' newRow("Rec.Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    newRow("Free Issue") = Trim(M01.Tables(0).Rows(i)("T02Free_Issue"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Qty")) * Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Total") = _St
                    newRow("##") = False
                    c_dataCustomer1.Rows.Add(newRow)


                    i = i + 1
                Next
                cmdAdd.Enabled = False
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True

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
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_RecordsUsing_ComInvoice()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _From As Date
        Dim M03 As DataSet
        Dim i As Integer
        Dim Value As Double


        Try
            Sql = "select * from T01Transaction_Header inner join T02Transaction_Flutter  on T01Ref_No=T02Ref_No and T02Com_Code=T01Com_Code inner join M03Item_Master on T02Item_Code=M03Item_Code and M03Com_Code=T02Com_Code inner join M04Location on M04Loc_Code=T01To_Loc_Code and T01Com_Code=M04Com_Code where T01Com_Code='" & _Comcode & "' and T01Invoice_No='" & Trim(txtCom_Invoice.Text) & "' and T01Trans_Type='GRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                cboTo.Text = Trim(M01.Tables(0).Rows(0)("M04Loc_Name"))
                txtCom_Invoice.Text = Trim(M01.Tables(0).Rows(0)("T01Invoice_No"))
                txtDate.Text = Trim(M01.Tables(0).Rows(0)("T01Date"))
                txtEntry.Text = Trim(M01.Tables(0).Rows(0)("T01Grn_No"))


                Value = Trim(M01.Tables(0).Rows(0)("T01Net_Amount"))
                txtNett.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtNett.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Vat"))
                txtVAT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtVAT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Market_Return"))
                txtMarket.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtMarket.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = Trim(M01.Tables(0).Rows(0)("T01Com_Discount"))
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtDis_Rate.Text = Trim(M01.Tables(0).Rows(0)("T01DisRate"))

                'Value = Trim(M01.Tables(0).Rows(0)("T01FreeIssue"))
                'txtFree_Amount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtFree_Amount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                Value = CDbl(txtNett.Text) - (CDbl(txtDiscount.Text) + CDbl(txtVAT.Text) + Val(txtMarket.Text))
                txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                txtCount.Text = M01.Tables(0).Rows.Count

                Dim _St As String

                i = 0
                For Each DTRow2 As DataRow In M01.Tables(0).Rows

                    Dim newRow As DataRow = c_dataCustomer1.NewRow
                    newRow("Item Code") = Trim(M01.Tables(0).Rows(i)("M03Item_Code"))
                    newRow("Item Name") = Trim(M01.Tables(0).Rows(i)("M03Item_Name"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Rate") = _St
                    newRow("Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    newRow("Rec.Qty") = Trim(M01.Tables(0).Rows(i)("T02Qty"))
                    newRow("Free Issue") = Trim(M01.Tables(0).Rows(i)("T02Free_Issue"))
                    Value = Trim(M01.Tables(0).Rows(i)("T02Qty")) * Trim(M01.Tables(0).Rows(i)("T02Cost"))
                    _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    newRow("Total") = _St

                    c_dataCustomer1.Rows.Add(newRow)


                    i = i + 1
                Next
                cmdAdd.Enabled = False
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True

            End If

            con.close()
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Private Sub txtSearch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Call Search_RecordsUsing_Entry()
            cboTo.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            Call Search_RecordsUsing_Entry()
            cboTo.ToggleDropdown()
        End If
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
        Dim M01 As DataSet
        Dim Value As Double

        Try

            txtUserName.Text = ""
            txtPassword.Text = ""
            If _LogStaus = False Then
                'Panel2.Visible = True
                OPRUser.Visible = True
                connection.Close()
                txtUserName.Focus()
                Exit Sub
            End If

            If txtPO.Text <> "" Then
            Else
                txtPO.Text = " "
            End If

            If Trim(txtDis_Rate.Text) <> "" Then
                If IsNumeric(txtDis_Rate.Text) Then
                    Value = (CDbl(txtNett.Text) * CDbl(txtDis_Rate.Text)) / 100
                    txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                    txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                    Call Calculate_Gross()
                Else
                    result1 = MessageBox.Show("Please enter the Discount Rate", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtDis_Rate.Focus()
                        Exit Sub
                    End If
                End If
            End If

            If Trim(txtCom_Invoice.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the company invoice ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCom_Invoice.Focus()
                    Exit Sub
                End If
            End If

            If Search_Location() = True Then
            Else
                result1 = MessageBox.Show("Please Select the Location ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboTo.ToggleDropdown()
                    Exit Sub
                End If
            End If

            '----------------------------------------------------------------------------------
            '  Call Load_EntryNo()
            If UltraGrid1.Rows.Count > 0 Then
            Else
                result1 = MessageBox.Show("Please enter the Items ", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboCode.ToggleDropdown()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------------
            If txtRemark.Text <> "" Then
            Else
                txtRemark.Text = " "
            End If

            'UPDATE T01 TRANSACTION
            i = 0

            nvcFieldList1 = "delete from T02Transaction_Flutter where T02Ref_No='" & _RefNo & "' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from S01Stock_Balance where S01Ref_No='" & _RefNo & "'  and S01Trans_Type='GRN'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from S03Ex_Stock where S03Ref_No='" & _RefNo & "'  and S03Tr_Type='GRN'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            nvcFieldList1 = "delete from S04Stock_Price where S04Ref_No='" & _RefNo & "'  and S04Tr_Type='GRN'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            i = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                Dim _Status As String
                Dim _Exdate As String

                _Exdate = " "
                'If (UltraGrid1.Rows(i).Cells(3).Value) = (UltraGrid1.Rows(i).Cells(4).Value) Then
                '    _Status = "OK"
                'Else
                _Status = "A"
                'End If

                If txtNBT.Text <> "" Then
                Else
                    txtNBT.Text = "0"
                End If

                If IsNumeric(txtNBT.Text) Then
                Else
                    MsgBox("Please enter the correct NBT Amount", MsgBoxStyle.Information, "Information .......")
                    connection.Close()
                    Exit Sub
                End If
                If IsDate((UltraGrid1.Rows(i).Cells(6).Value)) Then
                    _Exdate = (UltraGrid1.Rows(i).Cells(6).Value)
                End If
                nvcFieldList1 = "Insert Into T02Transaction_Flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Commition,T02Com_Discount,T02Qty,T02Rec_Qty,T02Free_Issue,T02Status,T02Item_Received,T02Com_Code,T02Total,T02Ex_Date,T02Count)" & _
                                                              " values('" & _RefNo & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & (UltraGrid1.Rows(i).Cells(2).Value) & "','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','0','0','" & (UltraGrid1.Rows(i).Cells(4).Value) & "','0','" & (UltraGrid1.Rows(i).Cells(5).Value) & "','A','" & _Status & "','" & _Comcode & "','" & CDbl(UltraGrid1.Rows(i).Cells(7).Value) & "','" & _Exdate & "','" & i + 1 & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Ref_No,S01Com_Code,s01status)" & _
                                                              " values('" & _Loccode & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & txtDate.Text & "','GRN','" & CDbl(UltraGrid1.Rows(i).Cells(4).Value) + CDbl(UltraGrid1.Rows(i).Cells(5).Value) & "','" & _RefNo & "','" & _Comcode & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into S04Stock_Price(S04Location,S04Tr_Type,S04Date,S04Item_Code,S04Qty,S04Ref_No,S04Status,S04Rate,S04Com_Code)" & _
                                                               " values('" & _Loccode & "','GRN','" & txtDate.Text & "', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & CDbl(UltraGrid1.Rows(i).Cells(4).Text) + CDbl(UltraGrid1.Rows(i).Cells(5).Text) & "','" & _RefNo & "','A','" & (UltraGrid1.Rows(i).Cells(3).Value) & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'EXPAIRE STOCK UPDATE
                nvcFieldList1 = "select * from M03Item_Master where M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "' and m03Status='A' and M03ExPair='YES' AND M03Com_Code='" & _Comcode & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    nvcFieldList1 = "Insert Into S03Ex_Stock(S03Loc_Code,S03Tr_Type,S03Item_Code,S03Qty,S03Ex_Date,S03Status,S03Ref_No,S03Com_Code)" & _
                                                                " values('" & _Loccode & "','GRN', '" & (UltraGrid1.Rows(i).Cells(0).Value) & "','" & CDbl(UltraGrid1.Rows(i).Cells(4).Text) + CDbl(UltraGrid1.Rows(i).Cells(5).Text) & "','" & UltraGrid1.Rows(i).Cells(6).Text & "','A','" & _RefNo & "','" & _Comcode & "')"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                End If

                If (UltraGrid1.Rows(i).Cells(8).Value) = True Then
                    nvcFieldList1 = "UPDATE M03Item_Master SET M03Cost_Price='" & (UltraGrid1.Rows(i).Cells(2).Value) & "' WHERE M03Item_Code='" & (UltraGrid1.Rows(i).Cells(0).Value) & "' and M03Com_Code='" & _Comcode & "'"
                    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                End If

                i = i + 1
            Next

            nvcFieldList1 = "update T01Transaction_Header set T01To_Loc_Code='" & _Loccode & "',T01Net_Amount='" & txtNett.Text & "',T01Com_Discount='" & txtDiscount.Text & "',T01DisRate='" & txtDis_Rate.Text & "',T01Vat='" & txtVAT.Text & "',T01FreeIssue='0',T01Market_Return='" & txtMarket.Text & "',T01Remark='" & txtRemark.Text & "',T01PO_NO='" & txtPO.Text & "',T01Invoice_No='" & txtCom_Invoice.Text & "',T01NBT='" & txtNBT.Text & "' where T01Ref_No='" & _RefNo & "' and T01Trans_Type='GRN' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '-------------------------------------------------------------------

            nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpAthz,tmpUser,tmpLog)" & _
                                                           " values('GRN', 'EDIT','" & txtEntry.Text & "','" & Now & "','" & _AthzUser & "','" & strDisname & "','" & _Comcode & "')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            'Account
            Dim _StrRemark As String

            '_StrRemark = "Good Received invo.No -" & txtCom_Invoice.Text

            'nvcFieldList1 = "update T05Acc_Trans set T05Credit='" & txtGross.Text & "' where T05Ref_No='" & txtEntry.Text & "' and T05Acc_Type='SP' and T05Acc_No='" & _Comcode & "'"
            'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            MsgBox("Records update successfully", MsgBoxStyle.Information, "Information ...........")
            transaction.Commit()
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            common.ClearAll(OPR0, OPR1, OPR2, OPR3)
            OPR2.Enabled = True
            OPR1.Enabled = True
            OPR0.Enabled = True
            OPR3.Enabled = True

            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Enabled = True
            cboLocation.ToggleDropdown()
            ' cmdSave.Enabled = False

            '  Call Load_Gride()
            Call Load_Gride2()
            Call Load_EntryNo()
            Call Load_Combo()
            Call Load_Data()
            _LogStaus = False

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
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
        Dim MB51 As DataSet
        Dim strDelete As String
        Try
            txtUserName.Text = ""
            txtPassword.Text = ""

            If _LogStaus = False Then
                'Panel2.Visible = True
                OPRUser.Visible = True
                connection.Close()
                txtUserName.Focus()
                Exit Sub
            End If
            strDelete = MessageBox.Show("Are you sure you want to delete this records ", "Information ....", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If strDelete = Windows.Forms.DialogResult.Yes Then
                nvcFieldList1 = "update T01Transaction_Header set T01status='I' where T01Ref_No='" & _RefNo & "'and T01Trans_Type='GRN'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "update T02Transaction_Flutter set T02status='I' where T02Ref_No='" & _RefNo & "'  "
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'nvcFieldList1 = "Update T05Acc_Trans where T05Ref_No='" & txtEntry.Text & "'  and T05Acc_Type='SP'"
                'ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                Call Search_Location()
                nvcFieldList1 = "UPDATE S01Stock_Balance SET S01STATUS='CLOSE' where S01Loc_Code='" & _Loccode & "'  and S01Ref_No='" & _RefNo & "' and S01Trans_Type='GRN'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S03Ex_Stock SET S03STATUS='CLOSE' where S03Loc_Code='" & _Loccode & "'  and S03Ref_No='" & _RefNo & "' and S03Tr_Type='GRN'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "UPDATE S04Stock_Price SET S04STATUS='CLOSE' where  S04Ref_No='" & _RefNo & "' and S04Tr_Type='GRN' AND S04Location='" & _Loccode & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                nvcFieldList1 = "Insert Into tmpTransaction_Log(tmp_TR,tmpProcess,tmpRef_No,tmpTime,tmpAthz,tmpUser,tmpLog)" & _
                                                          " values('GRN', 'DELETE','" & txtEntry.Text & "','" & Now & "','" & _AthzUser & "','" & strDisname & "','" & _Comcode & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                MsgBox("Records delete successfully", MsgBoxStyle.Information, "Information ...........")
                transaction.Commit()
            End If

            connection.Close()
            common.ClearAll(OPR0, OPR1, OPR2, OPR3)
            OPR2.Enabled = True
            OPR1.Enabled = True
            OPR0.Enabled = True
            OPR3.Enabled = True

            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdAdd.Enabled = True
            cboLocation.ToggleDropdown()
            ' cmdSave.Enabled = False

            '  Call Load_Gride()
            Call Load_Gride2()
            Call Load_EntryNo()
            Call Load_Combo()
            Call Load_Data()
            _LogStaus = False
            Call Load_Data()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Sub

    Private Sub txtCom_Invoice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCom_Invoice.KeyUp
        If e.KeyCode = 13 Then
            Call Search_RecordsUsing_ComInvoice()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Dim B As New ReportDocument
        Dim A As String

        Try
            A = ConfigurationManager.AppSettings("ReportPath") + "\GRNDispatch.rpt"
            B.Load(A.ToString)
            B.SetDatabaseLogon("sa", "tommya")
            ' B.SetParameterValue("To", txtTo.Value)
            'B.SetParameterValue("From", txtDate.Value)
            '  frmReport.CrystalReportViewer1.SelectionFormula = "{T01Transaction_Header.T01RefNo} =" & P01 & ""
            frmReport.CrystalReportViewer1.ReportSource = B 'intanance System\CrystalReport1" 'B ' "f:\salesinvo1.rpt" 'A.ToString '"F:\Stock Maintanance System\Report\salesinvo1.rpt"
            frmReport.CrystalReportViewer1.DisplayToolbar = True
            frmReport.CrystalReportViewer1.SelectionFormula = "{View_GRN_Header.T01Ref_No}=" & _RefNo & " "
            frmReport.Refresh()
            ' frmReport.CrystalReportViewer1.PrintReport()
            ' B.PrintToPrinter(1, True, 0, 0)
            frmReport.MdiParent = MDIMain
            frmReport.Show()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(connection)
                'connection.ConnectionString = ""
            End If
        End Try
    End Sub


    Private Sub txtRemark_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyUp
        If e.KeyCode = Keys.Enter Then
            cboCode.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboCode.ToggleDropdown()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub


   
    Private Sub cboItemName_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboItemName.KeyUp
        If e.KeyCode = 13 Then
            Call Search_ItemCode()
            txtRate.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
        ElseIf e.KeyCode = Keys.F2 Then
            OPR5.Visible = True
            txtFind.Text = ""
            Call Load_Gride_Item()
            txtFind.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR4.Visible = False
            OPR5.Visible = False
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Function Load_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select T01Grn_No as [GRN No],M09Name as [Supplier Name],T01Invoice_No as [Invoice No],CONVERT(varchar,CAST(T01Net_Amount AS money), 1) as [Net Amount] from T01Transaction_Header inner join M09Supplier on M09Code=T01FromLoc_Code where t01trans_type='grn' and t01date='" & Today & "' AND T01STATUS='A' and  T01To_Loc_Code='" & _Comcode & "' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid2.DataSource = dsUser
            UltraGrid2.Rows.Band.Columns(0).Width = 90
            UltraGrid2.Rows.Band.Columns(1).Width = 220
            UltraGrid2.Rows.Band.Columns(2).Width = 90
            UltraGrid2.Rows.Band.Columns(3).Width = 90
            'ltraGrid1.Rows.Band.Columns(4).Width = 110
            UltraGrid2.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid1.Rows.Band.Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            'UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboLocation_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboLocation.KeyUp
        If e.KeyCode = 13 Then
            cboTo.Focus()
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
        ElseIf e.KeyCode = Keys.Escape Then
            OPR4.Visible = False
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Function Load_Gride_Item3()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Item_Name  like '%" & txtFind.Text & "%' and M03Status='A' and M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 130
            UltraGrid3.Rows.Band.Columns(1).Width = 370
            UltraGrid3.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Gride_Item()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim dsUser As DataSet
        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name] from M03Item_Master where M03Status='A' and M03Com_Code='" & _Comcode & "' order by M03Item_Code"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid3.DataSource = dsUser
            UltraGrid3.Rows.Band.Columns(0).Width = 130
            UltraGrid3.Rows.Band.Columns(1).Width = 370
            ' UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()
        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub txtFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyUp
        If e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            cboCode.Focus()
        End If
    End Sub
    Private Sub txtFind_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFind.TextChanged
        Call Load_Gride_Item3()
    End Sub

   
    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        Panel1.Visible = False
        GRP1.Visible = False
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        Call Clear_Item()
    End Sub
    Function Search_Category() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M02Category where M02Cat_Name='" & Trim(cboMain.Text) & "' and M02Com_Code='" & _Comcode & "'"
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

    Function SAVE_ITEMMASTER()
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

            If Search_ItemLocation() = True Then
            Else

                result1 = MessageBox.Show("Please enter the Location", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboItem_Loc.ToggleDropdown()
                    Exit Function
                End If
            End If


            If Search_Category() = True Then
            Else

                result1 = MessageBox.Show("Please enter the correct Category", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboMain.ToggleDropdown()
                    Exit Function
                End If
            End If

            If Search_Supplier1() = True Then
            Else

                result1 = MessageBox.Show("Please enter the correct Supplier", "Information .....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    cboSupplier.ToggleDropdown()
                    Exit Function
                End If
            End If


            If Trim(txtCode.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Item Code", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtCode.Focus()
                    Exit Function
                End If
            End If

            '--------------------------------------------------------------------
            If Trim(txtDescription.Text) <> "" Then
            Else
                result1 = MessageBox.Show("Please enter the Item Name", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtDescription.Focus()
                    Exit Function
                End If
            End If

            If txtReorder.Text <> "" Then
                If IsNumeric(txtReorder.Text) Then
                Else
                    result1 = MessageBox.Show("Please enter the Correct Reorder Level", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If result1 = Windows.Forms.DialogResult.OK Then
                        txtReorder.Focus()
                        Exit Function
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
                        Exit Function
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
                        Exit Function
                    End If
                End If
            Else
                result1 = MessageBox.Show("Please enter the Correct Retail Price", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtRetail.Focus()
                    Exit Function
                End If
            End If

            If cboEx_Date.Text <> "" Then
            Else
                cboEx_Date.Text = "NO"
            End If
            '--------------------------------------------------------------------
            nvcFieldList1 = "Insert Into M03Item_Master(M03Item_Code,M03Item_Name,M03Cat_Code,M03Cost_Price,M03Reorder,M03Retail_Price,M03Com_Code,M03Supplier,M03Status,M03ExPair,M03Location)" & _
                                                          " values('" & (Trim(txtCode.Text)) & "', '" & (Trim(txtDescription.Text)) & "','" & Trim(_Category) & "','" & txtCost.Text & "','" & txtReorder.Text & "','" & txtRetail.Text & "','" & _Comcode & "','" & _SupCode1 & "','A','" & cboEx_Date.Text & "','" & _ItemLoc & "')"
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
            connection.Close()
            Call Clear_Item()
            Panel1.Visible = False
            GRP1.Visible = False
            cboCode.ToggleDropdown()


        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_Supplier1() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim Value As Double

        Try
            Sql = "select * from M01Account_Master where M01Acc_Name='" & Trim(cboSupplier.Text) & "'  and M01Acc_Type='SP' and M01Com_Code='" & _Comcode & "'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(dsUser) Then
                _SupCode1 = dsUser.Tables(0).Rows(0)("M01Acc_Code")
                Search_Supplier1 = True
            Else
                Search_Supplier1 = False
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
        If e.KeyCode = 13 Then
            txtReorder.Focus()
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call SAVE_ITEMMASTER()
    End Sub

    Private Sub txtRetail_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRetail.ValueChanged

    End Sub

    Private Sub txtSales_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSales.KeyUp
        Dim Value As Double
        If e.KeyCode = Keys.Enter Then
            If IsNumeric(txtSales.Text) Then
                Value = txtSales.Text
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtQty.Focus()
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If IsNumeric(txtSales.Text) Then
                Value = txtSales.Text
                txtSales.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtSales.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                txtQty.Focus()
            End If
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Private Sub UltraGrid3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraGrid3.KeyUp
        On Error Resume Next
        Dim _Rowindex As Integer
        If e.KeyCode = 13 Then

            _Rowindex = UltraGrid3.ActiveRow.Index


            cboCode.Text = UltraGrid3.Rows(_Rowindex).Cells(0).Text
            Search_ItemName()
            OPR5.Visible = False
            txtRate.Focus()
        End If
    End Sub

    Private Sub UltraGrid3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid3.MouseDoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer
        _Rowindex = UltraGrid3.ActiveRow.Index
        cboCode.Text = UltraGrid3.Rows(_Rowindex).Cells(0).Text
        Search_ItemName()
        OPR5.Visible = False
        txtRate.Focus()
    End Sub

    Private Sub UltraGrid2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles UltraGrid2.MouseDoubleClick
        On Error Resume Next
        Dim _Rowindex As Integer

        _Rowindex = UltraGrid2.ActiveRow.Index
        txtEntry.Text = UltraGrid2.Rows(_Rowindex).Cells(0).Text
        Call Search_RecordsUsing_Entry()
        OPR4.Visible = False
    End Sub



   
    Private Sub txtEx_Date_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            txtTotal.Focus()
        ElseIf e.KeyCode = Keys.Tab Then
            txtTotal.Focus()
        ElseIf e.KeyCode = Keys.F3 Then
            frmViewGRN.Show()
        End If
    End Sub

    Private Sub txtDiscount_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDiscount.ValueChanged
        On Error Resume Next
        Dim Value As Double

        If IsNumeric(txtDiscount.Text) Then
            If CDbl(txtNett.Text) > 0 Then
                Value = CDbl(txtNett.Text) - CDbl(txtDiscount.Text)
                txtGross.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtGross.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            End If
        End If
    End Sub

 
    Private Sub cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel.Click
        'Panel2.Visible = False
        OPRUser.Visible = False
    End Sub

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
            con = DBEngin.GetConnection()
            SQL = "SELECT * FROM users WHERE (NAME ='" & txtUserName.Text & "')and Password='" & txtPassword.Text & "' and UType in ('ADMIN','SH/MANEGER','ACCOUNT') "
            M01 = DBEngin.ExecuteDataset(con, Nothing, SQL)
            If isValidDataset(M01) Then
                _LogStaus = True
                _UserLevel = Trim(M01.Tables(0).Rows(0)("UType"))
                _AthzUser = Trim(txtUserName.Text)
                OPRUser.Visible = False
            Else
                MsgBox("User name and pasword combination not found", "Information ......")
                txtUserName.Focus()
                con.close()
                Exit Sub
            End If
            con.CLOSE()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.CLOSE()
            End If

        End Try

    End Sub

    Private Sub txtVAT_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVAT.TextChanged
        On Error Resume Next
        Dim Value As Double
        Dim result1 As String

        If Trim(txtVAT.Text) <> "" Then
            If IsNumeric(txtVAT.Text) Then
                Call Calculate_Gross()
                'Value = txtVAT.Text
                'txtVAT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtVAT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                ' txtMarket.Focus()
            Else
                result1 = MessageBox.Show("Please enter the VAT Amount", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtVAT.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub cboLocation_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboLocation.InitializeLayout

    End Sub

    Private Sub txtUserName_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUserName.ValueChanged

    End Sub

    Private Sub txtVAT_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVAT.ValueChanged

    End Sub

    Private Sub txtNBT_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNBT.ValueChanged
        On Error Resume Next
        Dim Value As Double
        Dim result1 As String

        If Trim(txtNBT.Text) <> "" Then
            If IsNumeric(txtNBT.Text) Then
                Call Calculate_Gross()
                'Value = txtVAT.Text
                'txtVAT.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                'txtVAT.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                ' txtMarket.Focus()
            Else
                result1 = MessageBox.Show("Please enter the NBT Amount", "Information ....", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If result1 = Windows.Forms.DialogResult.OK Then
                    txtNBT.Focus()
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub txtTotal_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotal.ValueChanged

    End Sub

    Private Sub txtFree_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFree.ValueChanged

    End Sub

    Private Sub txtMRP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMRP.KeyUp
        If e.KeyCode = 13 Then
            txtTotal.Focus()
        End If
    End Sub
End Class