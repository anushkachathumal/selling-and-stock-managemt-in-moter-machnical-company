Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Public Class frmSales_Trans
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim c_dataCustomer2 As DataTable
    Dim _itemRate As Double
    Private Sub frmSales_Trans_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblDate.Text = Today
        lblUser.Text = strDisname
        txtQty.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        lblDate.TextAlign = ContentAlignment.BottomCenter
        txtDiscount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        Call Load_Gride2()
        Call Load_Data()
        Call Load_Combo()
        Call Load_Gride3()
        txtTotal_Card.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
        txtTotal_Card.ReadOnly = True
        txtCard_Amount.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        Call Load_Amount()
    End Sub
    Function Load_Combo()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M26Dis as [##] from M26Ez_Payment "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboCard
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 170
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

    Function Search_Item() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim Value As Double
        Try
            Sql = "select * from M03Item_Master where  M03Item_Code='" & Trim(txtCode.Text) & "' and M03status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            lblItem.Text = M01.Tables(0).Rows(0)("M03Item_Name")
            If isValidDataset(M01) Then
                Search_Item = True
                _itemRate = M01.Tables(0).Rows(0)("M03Retail_Price")
            End If

            Sql = "select * from M27Discount_Items where M27Item_Code='" & txtCode.Text & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("M27Discount")
                txtDiscount.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                txtDiscount.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
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
    Function Load_Gride2()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer1 = CustomerDataClass.MakeDataTableSales
        UltraGrid2.DataSource = c_dataCustomer1
        With UltraGrid2
            .DisplayLayout.Bands(0).Columns(0).Width = 70
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 210
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 70
            .DisplayLayout.Bands(0).Columns(3).Width = 70
            .DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            .DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function

    Function Load_Gride3()
        Dim CustomerDataClass As New DAL_InterLocation()
        c_dataCustomer2 = CustomerDataClass.MakeDataTableCard
        UltraGrid3.DataSource = c_dataCustomer2
        With UltraGrid3
            .DisplayLayout.Bands(0).Columns(0).Width = 110
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False
            .DisplayLayout.Bands(0).Columns(1).Width = 100
            .DisplayLayout.Bands(0).Columns(1).AutoEdit = False

            .DisplayLayout.Bands(0).Columns(2).Width = 90
            '.DisplayLayout.Bands(0).Columns(3).Width = 70
            '.DisplayLayout.Bands(0).Columns(4).Width = 90
            .DisplayLayout.Bands(0).Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '.DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '.DisplayLayout.Bands(0).Columns(4).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            '  .DisplayLayout.Bands(0).Columns(5).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '  .DisplayLayout.Bands(0).Columns(6).Width = 90
            ' .DisplayLayout.Bands(0).Columns(7).Width = 90

            ' .DisplayLayout.Bands(0).Columns(3).Width = 300
            '.DisplayLayout.Bands(0).Columns(4).Width = 300
        End With
    End Function
    Private Sub txtCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCode.KeyUp
        If e.KeyCode = Keys.F15 Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F5 Then
            Call Load_Gride2()
            lblBill.Text = "000.00"
            txtCode.Text = ""
            txtQty.Text = ""
            lblItem.Text = ""
            OPR5.Visible = False
            Call Load_Gride3()
            OPR4.Visible = False
        ElseIf e.KeyCode = 13 Then
            If OPR4.Visible = True Then
                UltraGrid1.Focus()
            Else
                If txtCode.Text <> "" Then
                    If Search_Item() = True Then
                        txtQty.Focus()
                    Else
                        MsgBox("Please enter the Correct Item Code", MsgBoxStyle.Information, "Information .....")
                    End If
                End If
            End If
            ElseIf e.KeyCode = Keys.F1 Then
                OPR4.Visible = True
                Call Load_Data()
            ElseIf e.KeyCode = Keys.Escape Then
                If OPR4.Visible = True Then
                    OPR4.Visible = False
            End If
        ElseIf e.KeyCode = Keys.F2 Then
            Call Save_DATA()
        ElseIf e.KeyCode = Keys.F3 Then
            OPR5.Visible = True
            cboCard.ToggleDropdown()
            Call Load_Gride3()
        End If
    End Sub

  

    Function Load_Amount()
        Dim Value As Double
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim _St As String
        Try
            Sql = "select sum(T01Net_Amount) as Qty from T01Transaction_Header where T01Date='" & Today & "' and T01Trans_Type='DR' and T01Terminal='POS01' group by T01Terminal"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Value = M01.Tables(0).Rows(0)("qty")
                _St = Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                MDIMain.UltraToolbarsManager1.Tools(21).SharedProps.Caption = "POS01 " & _St
            End If



            con.close()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.Close()
            End If
        End Try

    End Function

    Private Sub txtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQty.KeyUp
        Dim Value As Double
        Dim _St As String
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F5 Then
            Call Load_Gride2()
            lblBill.Text = "000.00"
            txtCode.Text = ""
            txtQty.Text = ""
            lblItem.Text = ""
        ElseIf e.KeyCode = 13 Then
            If txtQty.Text <> "" Then
                If IsNumeric(txtQty.Text) Then
                    txtDiscount.Focus()
                Else
                    MsgBox("Please enter the correct Qty", MsgBoxStyle.Critical, "Error ....")
                End If
            End If
        End If
    End Sub

    Private Sub txtQty_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQty.ValueChanged

    End Sub


    Function Load_Data()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Status='A' "
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 220
            UltraGrid1.Rows.Band.Columns(2).Width = 90
            '  UltraGrid2.Rows.Band.Columns(3).Width = 90
            'ltraGrid1.Rows.Band.Columns(4).Width = 110
            UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Function Load_Data1()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()

        Try
            Sql = "select M03Item_Code as [Item Code],M03Item_Name as [Item Name],CONVERT(varchar,CAST(M03Retail_Price AS money), 1) as [Retail Price] from M03Item_Master where M03Status='A' and M03Item_Name like '%" & txtCode.Text & "%'"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            UltraGrid1.DataSource = dsUser
            UltraGrid1.Rows.Band.Columns(0).Width = 90
            UltraGrid1.Rows.Band.Columns(1).Width = 220
            UltraGrid1.Rows.Band.Columns(2).Width = 90
            '  UltraGrid2.Rows.Band.Columns(3).Width = 90
            'ltraGrid1.Rows.Band.Columns(4).Width = 110
            UltraGrid1.Rows.Band.Columns(2).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
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

    Private Sub txtCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        On Error Resume Next
        If OPR4.Visible = True Then
            Call Load_Data1()
        End If
    End Sub

    Private Sub UltraGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.DoubleClick
        On Error Resume Next
        Dim _rowind As Integer
        _rowind = UltraGrid1.ActiveRow.Index
        txtCode.Text = UltraGrid1.Rows(_rowind).Cells(0).Value
        Call Search_Item()
        txtQty.Focus()
        OPR4.Visible = False
    End Sub

    Function Save_DATA()
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
        Dim _PROFIT As Double
        Dim _PROFITCOM As Double
        Dim T01 As DataSet
        Dim _StrRemark As String
        Dim _Balance As Double
        Dim _RefNo As Integer
        Dim _INVONO As String
        Dim _CASH As Double
        Dim A As String
        Try
            If txtTotal_Card.Text <> "" Then
            Else
                txtTotal_Card.Text = "0"
            End If
           
            _CASH = CDbl(lblBill.Text) - CDbl(txtTotal_Card.Text)

            nvcFieldList1 = "SELECT * FROM P01Parameter WHERE P01Code='IN'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _RefNo = M01.Tables(0).Rows(0)("P01LastNo")

            End If

            nvcFieldList1 = "update P01Parameter set P01LastNo=P01LastNo+ " & 1 & " where P01Code='IN' "
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '-------------------------------------------------------------------------------------------------------------------------------

            nvcFieldList1 = "SELECT * FROM T01Transaction_Header WHERE T01Trans_Type='DR' AND T01Date='" & Today & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then
                _INVONO = Year(Today) & Month(Today) & Microsoft.VisualBasic.Day(Today) & M01.Tables(0).Rows.Count + 1
            Else
                _INVONO = Year(Today) & Month(Today) & Microsoft.VisualBasic.Day(Today) & 1
            End If

            i = 0
            _PROFIT = 0
            For Each uRow As UltraGridRow In UltraGrid2.Rows
                Dim _COSTPRICE As Double
                Dim _Item_Profit As Double
                Dim _Discount As Double

                _Item_Profit = 0
                nvcFieldList1 = "SELECT * FROM M03Item_Master WHERE M03Item_Code='" & Trim(UltraGrid2.Rows(i).Cells(0).Value) & "'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
                If isValidDataset(M01) Then
                    _COSTPRICE = M01.Tables(0).Rows(0)("M03Cost_Price")
                End If
                'MsgBox(UltraGrid1.Rows(i).Cells(6).Value)
                _Item_Profit = (CDbl(UltraGrid2.Rows(i).Cells(2).Value) - _COSTPRICE) * (UltraGrid2.Rows(i).Cells(3).Value)
                ' _Item_Profit = _Item_Profit - ((_Item_Profit * Val(UltraGrid1.Rows(i).Cells(6).Value)) / 100)
                If IsNumeric((UltraGrid2.Rows(i).Cells(4).Value)) Then
                    _Discount = (UltraGrid2.Rows(i).Cells(4).Value)
                Else
                    _Discount = 0
                End If
                _PROFIT = _PROFIT + _Item_Profit
                nvcFieldList1 = "Insert Into T02Transaction_Flutter(T02Ref_No,T02Item_Code,T02Cost,T02Retail_Price,T02Com_Discount,T02Qty,T02Rec_Qty,T02Free_Issue,T02Status,T02Item_Received,T02Com_Code,T02Total,T02Count)" & _
                                                              " values('" & _RefNo & "', '" & (UltraGrid2.Rows(i).Cells(0).Value) & "','" & _COSTPRICE & "','" & (UltraGrid2.Rows(i).Cells(2).Value) & "','" & _Discount & "','" & (UltraGrid2.Rows(i).Cells(3).Value) & "','" & (UltraGrid2.Rows(i).Cells(3).Value) & "','0','A','A','-','" & (UltraGrid2.Rows(i).Cells(5).Value) & "','" & i + 1 & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)


                nvcFieldList1 = "Insert Into S01Stock_Balance(S01Loc_Code,S01Item_Code,S01Date,S01Trans_Type,S01Qty,S01Free_Issue,S01Ref_No,S01Com_Code)" & _
                                                              " values('MS', '" & (UltraGrid2.Rows(i).Cells(0).Value) & "','" & Today & "','DR','" & -Val(UltraGrid2.Rows(i).Cells(3).Value) & "','0','" & _RefNo & "','-')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                i = i + 1
            Next


            nvcFieldList1 = "Insert Into T01Transaction_Header(T01Trans_Type,T01Ref_No,T01Date,T01FromLoc_Code,T01To_Loc_Code,T01Net_Amount,T01Com_Discount,T01DisRate,T01Vat,T01FreeIssue,T01Market_Return,T01Profit,T01Com_Code,T01Discount,T01User,T01Customer,T01Card_Type,T01Card_No,T01Tendered,T01Invoice_No,T01Terminal)" & _
                                                                " values('DR', '" & _RefNo & "','" & Today & "','MS','-','" & lblBill.Text & "','0','0','0','0','0','" & _PROFIT & "','-','0','" & strDisname & "','-','-','-','" & _CASH & "','" & _INVONO & "','POS01')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            'Outstanding

            nvcFieldList1 = "Insert Into T03Pay_Main(T03Ref_No,T03Trans_Type,T03Net_Amt,T03Cash,T03Credit,T03Status,T03Com_Code,T03CHQ,T03Pay_Method)" & _
                                                         " values('" & _RefNo & "', 'DR','" & lblBill.Text & "','0','0','A','-','0','-')"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)


            i = 0
            For Each uRow As UltraGridRow In UltraGrid3.Rows
                nvcFieldList1 = "Insert Into T11Credit_Card(T11Ref_No,T11Card_No,T11Card_Name,T11Amount,T11Status)" & _
                                                        " values('" & _RefNo & "', '" & UltraGrid3.Rows(i).Cells(0).Text & "','" & UltraGrid3.Rows(i).Cells(1).Text & "','" & UltraGrid3.Rows(i).Cells(2).Text & "','A')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                i = i + 1
            Next
            transaction.Commit()
            txtCode.Text = ""
            txtQty.Text = ""
            lblBill.Text = "000.00"
            Call Load_Gride2()
            OPR4.Visible = False
            OPR5.Visible = False
            Call Load_Amount()

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
            End If
        End Try
    End Function
    Private Sub UltraGrid1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles UltraGrid1.DoubleClickRow
        On Error Resume Next
        Dim _rowind As Integer
        _rowind = UltraGrid1.ActiveRow.Index
        txtCode.Text = UltraGrid1.Rows(_rowind).Cells(0).Value
        Call Search_Item()
        txtQty.Focus()
        OPR4.Visible = False

    End Sub

  
    Private Sub OPR4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OPR4.Click

    End Sub

    Private Sub txtCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.ValueChanged

    End Sub

    Private Sub cboCard_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboCard.InitializeLayout

    End Sub

    Private Sub cboCard_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCard.KeyUp
        If e.KeyCode = 13 Then
            txtCrad_No.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            Call Load_Gride3()
            txtTotal_Card.Text = ""
        ElseIf e.KeyCode = Keys.F2 Then
            Call Save_DATA()
        End If
    End Sub

    Private Sub txtCrad_No_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCrad_No.KeyUp
        If e.KeyCode = 13 Then
            txtCard_Amount.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            OPR5.Visible = False
            Call Load_Gride3()
            txtTotal_Card.Text = ""
        ElseIf e.KeyCode = Keys.F2 Then
            Call Save_DATA()
        End If
    End Sub

    Private Sub txtCard_Amount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCard_Amount.KeyUp
        On Error Resume Next
        Dim Value As Double
        Dim _St As String
        If e.KeyCode = 13 Then
            If cboCard.Text <> "" Then
            Else
                MsgBox("Please enter the card name", MsgBoxStyle.Information, "Information .....")
                cboCard.ToggleDropdown()
                Exit Sub
            End If

            If txtCrad_No.Text <> "" Then
            Else
                MsgBox("Please enter the Card No", MsgBoxStyle.Information, "Information .....")
                txtCrad_No.Focus()
                Exit Sub
            End If

            If txtCard_Amount.Text <> "" Then
            Else
                MsgBox("Please enter the Card Amount", MsgBoxStyle.Information, "Information .....")
                txtCard_Amount.Focus()
                Exit Sub
            End If

            If IsNumeric(txtCard_Amount.Text) Then
            Else
                MsgBox("Please enter the Card Amount", MsgBoxStyle.Information, "Information .....")
                txtCard_Amount.Focus()
                Exit Sub
            End If

            If txtTotal_Card.Text <> "" Then
            Else
                txtTotal_Card.Text = "0"
            End If

            Dim newRow As DataRow = c_dataCustomer2.NewRow
            newRow("Card No") = Trim(txtCrad_No.Text)
            newRow("Card Name") = cboCard.Text
            Value = txtCard_Amount.Text
            ' newRow("Rec.Qty") = txtRe_Qty.Text
            ' newRow("Free Issue") = txtFree.Text
            ' Value = _itemRate * txtQty.Text
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow("Amount") = _St

            c_dataCustomer2.Rows.Add(newRow)

            Value = CDbl(txtTotal_Card.Text) + txtCard_Amount.Text
            txtTotal_Card.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal_Card.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
           
        End If
    End Sub

    Private Sub UltraGrid3_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid3.AfterRowsDeleted
        Try
            Dim I As Integer
            Dim Value As Double
            I = 0
            Value = 0
            For Each uRow As UltraGridRow In UltraGrid3.Rows

                Value = UltraGrid3.Rows(I).Cells(0).Value + Value
            Next
            txtTotal_Card.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            txtTotal_Card.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

        Catch returnMessage As Exception
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                'DBEngin.CloseConnection(con)
                'con.ConnectionString = ""
            End If
        End Try
    End Sub


    Private Sub txtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDiscount.KeyUp
        Dim Value As Double
        Dim _St As String
        Dim _discount As Double
        If e.KeyCode = 13 Then
            If IsNumeric(txtDiscount.Text) Then
            Else
                MsgBox("Please enter the correct Discount%", MsgBoxStyle.Information, "Information .......")
                txtDiscount.Focus()
                Exit Sub
            End If
            If Search_Item() = True Then

            Else
                MsgBox("Please enter the correct Item Code", MsgBoxStyle.Information, "Information .......")
                txtCode.Focus()
                Exit Sub
            End If

            Value = _itemRate
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            Dim newRow As DataRow = c_dataCustomer1.NewRow
            newRow("Item Code") = Trim(txtCode.Text)
            newRow("Item Name") = lblItem.Text
            newRow("Rate") = _St
            newRow("Qty") = txtQty.Text
            newRow("Dis%") = txtDiscount.Text
            ' newRow("Free Issue") = txtFree.Text
            If IsNumeric(txtDiscount.Text) Then
                _discount = (_itemRate * txtQty.Text)
                _discount = _discount * (txtDiscount.Text / 100)
                Value = (_itemRate * txtQty.Text) - _discount
            Else
                Value = _itemRate * txtQty.Text
            End If
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            newRow("Total") = _St

            c_dataCustomer1.Rows.Add(newRow)
            Value = CDbl(lblBill.Text) + Value
            _St = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            _St = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
            lblBill.Text = _St
            txtCode.Text = ""
            txtQty.Text = ""
            lblItem.Text = ""
            txtDiscount.Text = ""
            txtCode.Focus()
        ElseIf e.KeyCode = Keys.F15 Then
            Me.Close()
        ElseIf e.KeyCode = Keys.F5 Then
            Call Load_Gride2()
            lblBill.Text = "000.00"
            txtCode.Text = ""
            txtQty.Text = ""
            lblItem.Text = ""
            OPR5.Visible = False
            Call Load_Gride3()
            OPR4.Visible = False
        ElseIf e.KeyCode = Keys.F1 Then
            OPR4.Visible = True
            Call Load_Data()
        ElseIf e.KeyCode = Keys.Escape Then
            If OPR4.Visible = True Then
                OPR4.Visible = False
            End If
        ElseIf e.KeyCode = Keys.F2 Then
            Call Save_DATA()
        ElseIf e.KeyCode = Keys.F3 Then
            OPR5.Visible = True
            cboCard.ToggleDropdown()
            Call Load_Gride3()
        End If
    End Sub
End Class