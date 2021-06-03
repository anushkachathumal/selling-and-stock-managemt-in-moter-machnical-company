'// Sales ordering module for the Merchant
'// Development Date - 07.24.2014
'// Developed by - Suranga Wijesinghe
'// Audit by     - Amila Priyankara - TJL
'// Referance Table - M01Sales_Order_SAP (Master Table)
'//                 - P01PARAMETER (For add referance No)
'//                 - T01Delivary_Request
'//                 - USERS     
'//---------------------------------------------------------->>>
'Automate the Email  send by merchant to  Panner & data merchant's order status Excell migration

Imports System.Data.SqlClient
Imports Infragistics.Win.UltraWinGrid
Imports DBLotVbnet.common
Imports System.Net
Imports DBLotVbnet.DBEngin
Imports DBLotVbnet.DAL_InterLocation
Imports DBLotVbnet.DAL_Distributors
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
'Imports Office = Microsoft.Office.Core
Imports Microsoft.Office.Interop.Outlook
Imports System.Drawing
Imports Spire.XlS
Imports System.Xml
Imports Infragistics.Win.UltraWinGrid.RowLayoutStyle.GroupLayout
Imports Infragistics.Win.UltraWinToolTip
Imports Infragistics.Win.FormattedLinkLabel
Imports Infragistics.Win.Misc
Imports System.Diagnostics

Public Class frmDelivary_Request
    Inherits System.Windows.Forms.Form
    Dim dsUser As DataSet
    Dim Clicked As String
    Dim c_dataCustomer1 As DataTable
    Dim _EPF As String
    Dim _Email As String
    Dim _LeadTime As String

    Dim c_dataCustomer As DataTable
    Dim c_dataCustomer2 As DataTable
    'Dim xlApp As New Excel.Application
    'Dim xlWBook As Excel.Workbook


    'Function Load_Gride_SalesOrder()
    '    Dim CustomerDataClass As New DAL_InterLocation()
    '    c_dataCustomer = CustomerDataClass.MakeDataTable_Sales_Order
    '    UltraGrid1.DataSource = c_dataCustomer
    '    With UltraGrid1
    '        .DisplayLayout.Bands(0).Columns(1).Width = 50
    '        .DisplayLayout.Bands(0).Columns(2).AutoEdit = False
    '        .DisplayLayout.Bands(0).Columns(3).Width = 190
    '        .DisplayLayout.Bands(0).Columns(3).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
    '        .DisplayLayout.Bands(0).Columns(3).AutoEdit = False
    '        .DisplayLayout.Bands(0).Columns(3).Width = 60
    '        .DisplayLayout.Bands(0).Columns(4).Width = 60
    '        .DisplayLayout.Bands(0).Columns(6).Width = 60
    '        .DisplayLayout.Bands(0).Columns(9).Width = 60
    '        .DisplayLayout.Bands(0).Columns(8).Width = 70
    '        .DisplayLayout.Bands(0).Columns(11).Width = 60
    '        .DisplayLayout.Bands(0).Columns(18).Width = 60
    '        .DisplayLayout.Bands(0).Columns(19).Width = 130
    '        .DisplayLayout.Bands(0).Columns(20).Width = 90
    '    End With
    'End Function

    Function Load_Sales_Order()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load sales order to cboSO combobox

        Try
            Sql = "select CONVERT(INT,M01Sales_Order) as [Sales Order],max(M01Cuatomer_Name) as [Customer] from M01Sales_Order_SAP WHERE M01SO_Qty<>M01Delivary_Qty and M01Status='A' group by M01Sales_Order "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboSO
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                .Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Search_Parameter()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Search Referance No via the P01PARAMETER Table
        Try
            Sql = "select * from P01PARAMETER where P01code='SO'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                txtRefNo.Text = Trim(M01.Tables(0).Rows(0)("P01no"))
            End If
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""


        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_PLANNER()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        'Load Planer to cboPlanner combobox

        Try
            Sql = "select Name as [Planner's Name] from users where Designation='PLANNER' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboPlaner
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 160
                '.Rows.Band.Columns(1).Width = 260


            End With
            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            'With txtNSL
            '    .DataSource = M01
            '    .Rows.Band.Columns(0).Width = 225
            'End With

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Function Load_Quality_Allocation()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            Sql = "select M01Quality_No as [Quality No] from M01Sales_Order_SAP where convert(int,M01Sales_Order)='" & Trim(cboSO.Text) & "' AND M01Status='A' group by M01Quality_No"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            With cboAllocate_Qulity
                .DataSource = M01
                .Rows.Band.Columns(0).Width = 245
            End With
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try
    End Function
    Function Load_Data_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        Dim con1 = New SqlConnection()

        con = DBEngin.GetConnection(True)
        ' con1 = DBEngin1.GetConnection(True)
        Dim M01 As DataSet
        Dim T02 As DataSet

        Dim i As Integer
        Dim Value As Double
        Dim _Qty As String
        Dim VCWhere As String
        Dim X As Integer

        Try
            'Search Data to M01Sales_Order_SAP table
            ' Call Load_Gride_SalesOrder()
            Sql = "select * from M01Sales_Order_SAP where convert(int,M01Sales_Order)='" & Trim(cboSO.Text) & "' AND M01Status='A'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            cmdSave.Enabled = True

            For Each DTRow1 As DataRow In M01.Tables(0).Rows
                txtPO.Text = M01.Tables(0).Rows(0)("M01PO")
                txtCustomer.Text = M01.Tables(0).Rows(0)("M01Cuatomer_Name")

                Dim newRow As DataRow = c_dataCustomer.NewRow
                _Qty = 0
                newRow("##") = False
                newRow("Line Item") = M01.Tables(0).Rows(i)("M01Line_Item")
                newRow("Quality") = M01.Tables(0).Rows(i)("M01Quality")
                Value = M01.Tables(0).Rows(i)("M01SO_Qty")
                _Qty = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                _Qty = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                newRow("Quantity") = _Qty
                newRow("Retailer") = M01.Tables(0).Rows(i)("M01Department")

                VCWhere = "M01Line_Item=" & Trim(M01.Tables(0).Rows(i)("M01Line_Item")) & " and M01Sales_Order=" & cboSO.Text & ""
                T02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "QRCD"), New SqlParameter("@vcWhereClause1", VCWhere))
                If isValidDataset(T02) Then
                    VCWhere = "M42Rcode='" & T02.Tables(0).Rows(0)("m16R_Code") & "'"
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "STCD"), New SqlParameter("@vcWhereClause1", VCWhere))
                    If isValidDataset(dsUser) Then
                        newRow("1st Bulk/Repeat") = "REPEAT"
                        newRow("PP") = "APPROVED"
                        newRow("NPL") = "APPROVED"
                        newRow("Lab Dye") = "APPROVED"
                    Else
                        newRow("1st Bulk/Repeat") = "1ST BULK"
                        VCWhere = "M14Order='" & T02.Tables(0).Rows(0)("m16R_Code") & "'  AND M14Status<>''"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCDE1"), New SqlParameter("@vcWhereClause1", VCWhere))
                        If isValidDataset(dsUser) Then
                            newRow("Lab Dye") = "APPROVED"
                        Else
                            newRow("Lab Dye") = "NOT APP"
                        End If
                    End If

                Else
                    newRow("1st Bulk/Repeat") = "1ST BULK"
                End If

                newRow("Lead Time") = "Normal"
                'If IsNumeric(Microsoft.VisualBasic.Right(Trim(M01.Tables(0).Rows(i)("M01Quality")), 5)) Then
                '    Sql = "select * from CL01_Customer_Submisions where CL01RCode='" & Microsoft.VisualBasic.Right(Trim(M01.Tables(0).Rows(i)("M01Quality")), 5) & "'"
                '    T02 = DBEngin1.ExecuteDataset(con1, Nothing, Sql)

                '    If isValidDataset(T02) Then
                '        If Trim(T02.Tables(0).Rows(0)("CL01Status")) = "A" Then
                '            newRow("Lab Dye") = "Approved"
                '        ElseIf Trim(T02.Tables(0).Rows(0)("CL01Status")) = "R" Then
                '            newRow("Lab Dye") = "Reject"
                '        End If
                '    End If
                '    c_dataCustomer.Rows.Add(newRow)
                'End If
                ' newRow("PP") = False
                c_dataCustomer.Rows.Add(newRow)

                i = i + 1
            Next

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

            DBEngin1.CloseConnection(con1)
            con1.ConnectionString = ""

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub cboSO_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSO.AfterCloseUp
        strSales_Order = cboSO.Text
        Call Load_Data_Gride()
        Me.BindUltraDropDown2()
        Call Delete_Projection(strSales_Order)

    End Sub

    Private Sub cboSO_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSO.LostFocus
        Call Search_Salrs_Order()
        Call Load_Data_Gride()
        Me.BindUltraDropDown2()

    End Sub

    Private Sub BindUltraDropDown1()

        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("Dis", GetType(String))

        dt.Rows.Add(New Object() {"APPROVED"})
        dt.Rows.Add(New Object() {"NOT APPROVED"})
        ' dt.Rows.Add(New Object() {"SESANAL"})
        dt.AcceptChanges()

        Me.UltraDropDown2.SetDataBinding(dt, Nothing)
        '  Me.UltraDropDown1.ValueMember = "ID"
        Me.UltraDropDown2.DisplayMember = "Dis"
    End Sub

    Private Sub BindUltraDropDown2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim I As Integer
        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        Try
            If Trim(cboSO.Text) <> "" Then
                dt.Columns.Add("Dis", GetType(String))
                Sql = "select * from M01Sales_Order_SAP where convert(int,M01Sales_Order)='" & Trim(cboSO.Text) & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                I = 0
                'cmdSave.Enabled = True

                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    dt.Rows.Add(New Object() {M01.Tables(0).Rows(I)("M01Line_Item")})
                    ' dt.Rows.Add(New Object() {"NOT APPROVED"})
                    ' dt.Rows.Add(New Object() {"SESANAL"})
                    dt.AcceptChanges()
                    I = I + 1
                Next
                Me.UltraDropDown3.SetDataBinding(dt, Nothing)
                '  Me.UltraDropDown1.ValueMember = "ID"
                Me.UltraDropDown3.DisplayMember = "Dis"
            End If

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Sub

    Private Sub BindUltraDropDown()

        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("Dis", GetType(String))

        dt.Rows.Add(New Object() {"1st BULK"})
        dt.Rows.Add(New Object() {"REPEAT"})
        dt.Rows.Add(New Object() {"SESANAL"})
        dt.AcceptChanges()

        Me.UltraDropDown1.SetDataBinding(dt, Nothing)
        '  Me.UltraDropDown1.ValueMember = "ID"
        Me.UltraDropDown1.DisplayMember = "Dis"
    End Sub

    Private Sub BindUltraDropDown_LeadTime()

        Dim dt As DataTable = New DataTable()
        ' dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("##", GetType(String))

        dt.Rows.Add(New Object() {"Normal"})
        dt.Rows.Add(New Object() {"Short Lead Time"})
        dt.Rows.Add(New Object() {"Speed Order"})
        dt.AcceptChanges()

        Me.UltraDropDown4.SetDataBinding(dt, Nothing)
        '  Me.UltraDropDown1.ValueMember = "ID"
        Me.UltraDropDown4.DisplayMember = "##"
    End Sub

    Private Sub frmDelivary_Request_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Call Delete_Projection(strSales_Order)
    End Sub

    Private Sub frmDelivary_Request_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Call Load_Gride_SalesOrder()
        Me.BindUltraDropDown()
        Me.BindUltraDropDown1()
        Me.BindUltraDropDown2()
        Me.BindUltraDropDown_LeadTime()
        Panel1.Visible = False

        Call Load_Sales_Order()
        Call Search_Parameter()
        Call Load_PLANNER()
        ' Call Load_Combo_Lead_Time()

        chkNPL2.Checked = True

        txtRefNo.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtQuality.ReadOnly = True
        txtQuality.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtCLR.ReadOnly = True
        txtCLR.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYC1.ReadOnly = True
        txtYC1.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtYC2.ReadOnly = True '
        txtYC2.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        txtQuality_Dis.ReadOnly = True
        txtLine_Item.ReadOnly = True
        txtLine_Item.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        txtPO.ReadOnly = True
        Dim TipInfo As New UltraToolTipInfo()
        Dim TipInfo1 As New UltraToolTipInfo()
        Dim TipInfo2 As New UltraToolTipInfo()

        TipInfo.ToolTipText = "Close"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraButton3, TipInfo)
        Me.UltraToolTipManager1.DisplayStyle = Infragistics.Win.ToolTipDisplayStyle.BalloonTip

        TipInfo1.ToolTipText = "Minimize"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraButton4, TipInfo1)
        Me.UltraToolTipManager1.DisplayStyle = Infragistics.Win.ToolTipDisplayStyle.BalloonTip

        TipInfo2.ToolTipText = "Maximize"
        Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraButton5, TipInfo2)
        Me.UltraToolTipManager1.DisplayStyle = Infragistics.Win.ToolTipDisplayStyle.BalloonTip

    End Sub

    Function Projection_Allocation()
        Dim _RowIndex As Integer
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _30Class As String
        Dim Value As Double
        Dim vcWhere As String
        Dim characterToRemove As String
        Dim X As Integer
        Dim agroup1 As UltraGridGroup
        Dim agroup2 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim agroup4 As UltraGridGroup
        Dim agroup5 As UltraGridGroup
        Dim _Retailer As String

        Try


            If UltraGroupBox1.Visible = False Then
                If UltraGrid1.Rows.Count > 0 Then
                    _RowIndex = (UltraGrid1.ActiveRow.Index)
                    If Trim(UltraGrid1.Rows(_RowIndex).Cells(0).Value) = True Then
                        txtLine_Item.Text = Trim(UltraGrid1.Rows(_RowIndex).Cells(1).Text)
                        txtQuality_Dis.Text = Trim(UltraGrid1.Rows(_RowIndex).Cells(2).Text)

                        vcWhere = "M01Sales_Order=" & cboSO.Text & " and M01Line_Item=" & txtLine_Item.Text & ""
                        M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M01) Then
                            txtQuality.Text = M01.Tables(0).Rows(0)("M01Quality_No")
                            _30Class = Trim(M01.Tables(0).Rows(0)("M01Material_No"))
                            characterToRemove = "-"

                            'MsgBox(Trim(fields(9)))
                            _30Class = (Replace(_30Class, characterToRemove, ""))

                            vcWhere = "M16Material='" & _30Class & "'"
                            M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCODE"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(M02) Then
                                txtCLR.Text = M02.Tables(0).Rows(0)("M16Shade_Type")
                            End If


                        End If
                        'Search Block Quality
                        'Developd by suranga wijesinghe
                        'Date : 01/03/2016
                        'Check Available quality used for other merchant for projection allocation
                        If Search_Block_Quality() = True Then
                            Exit Function
                        End If

                        _Retailer = ""
                        vcWhere = "M43Quality='" & txtQuality.Text & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            _Retailer = dsUser.Tables(0).Rows(0)("M43Retailler")
                        End If
                        '---------------------------------------------------------END Description
                        vcWhere = "M22Quality='" & txtQuality.Text & "'"
                        dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "VWT"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(dsUser) Then
                            'MsgBox(Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(0)("M22Yarn"), 7))
                            If _Retailer <> "" Then
                                If Microsoft.VisualBasic.Left(txtQuality.Text, 1) = "Y" Then
                                    vcWhere = "M43Quality<>'" & txtQuality.Text & "' and gauge='" & dsUser.Tables(0).Rows(0)("gauge") & "' and Ptype='" & dsUser.Tables(0).Rows(0)("Ptype") & "' and M22Product_Type='" & dsUser.Tables(0).Rows(0)("M22Product_Type") & "' and M22Yarn_Cons='" & dsUser.Tables(0).Rows(0)("M22Yarn_Cons") & "' and left(M22Yarn,7)='" & Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(0)("M22Yarn"), 7) & "' and M43Retailler='" & _Retailer & "' and M43Year>=" & Year(Today) & " and  M43Product_Month>=" & Month(Today) & ""
                                Else
                                    vcWhere = "M43Quality<>'" & txtQuality.Text & "' and gauge='" & dsUser.Tables(0).Rows(0)("gauge") & "' and Ptype='" & dsUser.Tables(0).Rows(0)("Ptype") & "' and M22Product_Type='" & dsUser.Tables(0).Rows(0)("M22Product_Type") & "' and M22Yarn_Cons='" & dsUser.Tables(0).Rows(0)("M22Yarn_Cons") & "' and left(M22Yarn,7)='" & Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(0)("M22Yarn"), 7) & "' and M43Retailler='" & _Retailer & "' and M43Year>=" & Year(Today) & " and  M43Product_Month>=" & Month(Today) & ""
                                End If
                            Else
                                vcWhere = "M43Quality<>'" & txtQuality.Text & "' and gauge='" & dsUser.Tables(0).Rows(0)("gauge") & "' and Ptype='" & dsUser.Tables(0).Rows(0)("Ptype") & "' and M22Product_Type='" & dsUser.Tables(0).Rows(0)("M22Product_Type") & "' and M22Yarn_Cons='" & dsUser.Tables(0).Rows(0)("M22Yarn_Cons") & "' and left(M22Yarn,7)='" & Microsoft.VisualBasic.Left(dsUser.Tables(0).Rows(0)("M22Yarn"), 7) & "' and M43Year>=" & Year(Today) & " and  M43Product_Month>=" & Month(Today) & " "
                            End If

                            M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRO1"), New SqlParameter("@vcWhereClause1", vcWhere))
                            With cboStealing
                                .DataSource = M02
                                .Rows.Band.Columns(0).Width = 160
                            End With
                        End If
                        If chkAllocate.Checked = False Then
                            Value = Trim(UltraGrid1.Rows(_RowIndex).Cells(3).Text)
                            lblQty.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            lblQty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            lblBalance.Text = lblQty.Text
                            Call Load_Gride2()
                            UltraGroupBox1.Visible = True
                        Else

                            X = 0
                            Value = 0
                            For Each uRow As UltraGridRow In UltraGrid1.Rows
                                vcWhere = "M01Sales_Order=" & cboSO.Text & " and M01Line_Item='" & Trim(UltraGrid1.Rows(X).Cells(1).Text) & "' and M01Quality_No='" & txtQuality.Text & "'"
                                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                                If isValidDataset(M01) Then
                                    If Trim(UltraGrid1.Rows(X).Cells(0).Value) = True Then
                                        Value = Value + Trim(UltraGrid1.Rows(X).Cells(3).Text)
                                    End If
                                End If
                                X = X + 1
                            Next

                            lblQty.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            lblQty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            lblBalance.Text = lblQty.Text
                            Call Load_Gride2()
                            UltraGroupBox1.Visible = True
                        End If
                    End If
                    con.close()
                End If
            Else
                _RowIndex = (UltraGrid1.ActiveRow.Index)
                Dim TipInfo As New UltraToolTipInfo()
                TipInfo.ToolTipText = "Projection Allocation Screen alrady exist"
                Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraGrid1, TipInfo)
                Me.UltraToolTipManager1.DisplayStyle = Infragistics.Win.ToolTipDisplayStyle.BalloonTip
            End If
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function

    Function Projection_AllocationCapacity_stealing()
        Dim _RowIndex As Integer
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _30Class As String
        Dim Value As Double
        Dim vcWhere As String
        Dim characterToRemove As String
        Dim X As Integer
        Dim agroup1 As UltraGridGroup
        Dim agroup2 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim agroup4 As UltraGridGroup
        Dim agroup5 As UltraGridGroup
        Try


            'If UltraGroupBox1.Visible = False Then
            If UltraGrid1.Rows.Count > 0 Then
                _RowIndex = (UltraGrid1.ActiveRow.Index)
                If Trim(UltraGrid1.Rows(_RowIndex).Cells(0).Value) = True Then
                    txtLine_Item.Text = Trim(UltraGrid1.Rows(_RowIndex).Cells(1).Text)
                    txtQuality_Dis.Text = Trim(UltraGrid1.Rows(_RowIndex).Cells(2).Text)

                    vcWhere = "M01Sales_Order=" & cboSO.Text & " and M01Line_Item=" & txtLine_Item.Text & ""
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        txtQuality.Text = M01.Tables(0).Rows(0)("M01Quality_No")
                        _30Class = Trim(M01.Tables(0).Rows(0)("M01Material_No"))
                        characterToRemove = "-"

                        'MsgBox(Trim(fields(9)))
                        _30Class = (Replace(_30Class, characterToRemove, ""))

                        vcWhere = "M16Material='" & _30Class & "'"
                        M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCODE"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            txtCLR.Text = M02.Tables(0).Rows(0)("M16Shade_Type")
                        End If


                    End If
                    'Search Block Quality
                    'Developd by suranga wijesinghe
                    'Date : 01/03/2016
                    'Check Available quality used for other merchant for projection allocation
                    If Search_Block_Quality() = True Then
                        Exit Function
                    End If

                    '---------------------------------------------------------END Description
                    Dim _Str As String
                    _Str = ""
                    If cboStealing.Text <> "" Then
                        _Str = cboStealing.Text
                    End If
                    'vcWhere = "M43Quality<>'" & txtQuality.Text & "'"
                    'M02 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRO"), New SqlParameter("@vcWhereClause1", vcWhere))
                    'With cboStealing
                    '    .DataSource = M02
                    '    .Rows.Band.Columns(0).Width = 160
                    'End With
                    If _Str <> "" Then
                        cboStealing.Text = _Str
                    End If
                    If chkAllocate.Checked = False Then
                        Value = Trim(UltraGrid1.Rows(_RowIndex).Cells(3).Text)
                        lblQty.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        lblQty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        lblBalance.Text = lblQty.Text
                        Call Load_Gride2()
                        UltraGroupBox1.Visible = True
                    Else

                        X = 0
                        Value = 0
                        For Each uRow As UltraGridRow In UltraGrid1.Rows
                            vcWhere = "M01Sales_Order=" & cboSO.Text & " and M01Line_Item='" & Trim(UltraGrid1.Rows(X).Cells(1).Text) & "' and M01Quality_No='" & txtQuality.Text & "'"
                            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(M01) Then
                                If Trim(UltraGrid1.Rows(X).Cells(0).Value) = True Then
                                    Value = Value + Trim(UltraGrid1.Rows(X).Cells(3).Text)
                                End If
                            End If
                            X = X + 1
                        Next

                        lblQty.Text = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        lblQty.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                        lblBalance.Text = lblQty.Text
                        Call Load_Gride2()
                        UltraGroupBox1.Visible = True
                    End If
                End If
                con.close()
            End If
            'Else
            '    _RowIndex = (UltraGrid1.ActiveRow.Index)
            '    Dim TipInfo As New UltraToolTipInfo()
            '    TipInfo.ToolTipText = "Projection Allocation Screen alrady exist"
            '    Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraGrid1, TipInfo)
            '    Me.UltraToolTipManager1.DisplayStyle = Infragistics.Win.ToolTipDisplayStyle.BalloonTip
            'End If
            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                con.close()
            End If
        End Try

    End Function
    Function Search_Block_Quality() As Boolean
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True

        Dim M01 As DataSet
        Dim vcWhere As String
        Dim ncQryType As String

        Try
            vcWhere = "tmpQuality='" & txtQuality.Text & "'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "BPAL"), New SqlParameter("@vcWhereClause1", vcWhere))
            If isValidDataset(M01) Then
                If strDisname = Trim(M01.Tables(0).Rows(0)("tmpUser")) Then

                Else
                    Search_Block_Quality = True
                    MsgBox("The Quality used by " & Trim(M01.Tables(0).Rows(0)("tmpUser")), MsgBoxStyle.Exclamation, "Block Projection Allocation")
                End If
            Else
                Search_Block_Quality = False
                ncQryType = "ADD"
                vcWhere = ""
                nvcFieldList1 = "(tmpQuality," & "tmpUser," & "tmpDate) " & "values('" & Trim(txtQuality.Text) & "','" & strDisname & "','" & Now & "')"
                up_GetSetBlock_Projection(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                transaction.Commit()
            End If
            connection.Close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                connection.Close()
            End If
        End Try
    End Function

    Function Clear_Gride()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim vcWhere As String
        Dim _Code As Integer

        Try
            vcWhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1

            ' UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)
            If cboStealing.Text <> "" Then
                If Microsoft.VisualBasic.Day(Today) > 10 Then
                    vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "' and M43Product_Month<>" & Month(Today) & "  and M43Count_No=" & _Code & ""
                Else
                    vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "'  and M43Count_No=" & _Code & ""
                End If
            Else
                If Microsoft.VisualBasic.Day(Today) > 10 Then
                    vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "' and M43Product_Month<>" & Month(Today) & "  and M43Count_No=" & _Code & ""
                Else
                    vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "'  and M43Count_No=" & _Code & ""
                End If
            End If
            ' MsgBox(UltraGrid3.DisplayLayout.Bands(0).Groups.Clear)

            UltraGrid3.DisplayLayout.Bands(0).Groups.Clear()
            UltraGrid3.DisplayLayout.Bands(0).Columns.Dispose()
            UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)
            UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(1)
            UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(2)
            UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(3)

            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LOMN"), New SqlParameter("@vcWhereClause1", vcWhere))
            I = 0
            For Each DTRow3 As DataRow In M01.Tables(0).Rows

                If I = 0 Then

                    UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(1)
                ElseIf I = 1 Then

                    UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(2)
                ElseIf I = 2 Then

                    UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(3)
                ElseIf I = 3 Then
                    UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(4)
                End If
                I = I + 1
            Next

            con.close()
        Catch returnMessage As ExecutionEngineException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
            End If
        End Try
    End Function

    Function Search_Data_Gride() As Boolean
        On Error Resume Next
        Dim I As Integer
        Dim p, Grid, Table, RowId, ColumnName, Value
        'p = Sys.Process("SamplesExplorer")
        'Grid = p.WinFormsObject("frmDelivary_Request").WinFormsObject("UltraGrid1")
        Grid = UltraGrid1
        Table = Grid.DataSource
        ' Table = Grid.DataSource.Tables.Item(0)
        ColumnName = "NPL"
        Value = "NOT APPROVED"
        RowId = FindRow(Table, ColumnName, Value)
        If RowId >= 0 Then
            'MsgBox(Trim(UltraGrid1.Rows(RowId).Cells(10).Value))
            'If Trim(UltraGrid1.Rows(RowId).Cells(10).Value) <> "" Then
            '    With UltraGrid1
            '        For I = 0 To 15
            '            .Rows(RowId).Cells(I).Appearance.BackColor = Color.White
            '        Next
            '    End With
            'Else
            If Trim(UltraGrid1.Rows(RowId).Cells(10).Text) <> "" Then
                With UltraGrid1
                    For I = 0 To 15
                        .Rows(RowId).Cells(I).Appearance.BackColor = Color.White
                    Next
                End With
            Else
                MsgBox("Please enter the NPL Approved Date", MsgBoxStyle.Information, "Information ......")
                With UltraGrid1
                    For I = 0 To 15
                        .Rows(RowId).Cells(I).Appearance.BackColor = Color.Blue
                    Next
                End With
                Search_Data_Gride = True
                Exit Function
            End If
        Else
            With UltraGrid1
                For I = 0 To 15
                    .Rows(RowId).Cells(I).Appearance.BackColor = Color.White
                Next
            End With
        End If
    End Function
    Private Sub UltraGrid1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles UltraGrid1.AfterRowUpdate
        Call Search_Data_Gride()

    End Sub

    Function FindRow(ByVal Table, ByVal ColumnName, ByVal Value)
        Dim View
        View = Table.DefaultView
        View.Sort = ColumnName
        FindRow = View.Find(Value)
    End Function

    Private Sub UltraGrid1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraGrid1.Click
        On Error Resume Next
        Dim _RowIndex As Integer
        _RowIndex = (UltraGrid1.ActiveRow.Index)
        ' MsgBox(Trim(UltraGrid1.Rows(_RowIndex).Cells(0).Value))
        If Trim(UltraGrid1.Rows(_RowIndex).Cells(0).Value) = True Then
            Call Projection_Allocation()
        End If
    End Sub

    Private Sub UltraGrid1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout
        e.Layout.Bands(0).Columns("1st Bulk/Repeat").ValueList = Me.UltraDropDown1
        e.Layout.Bands(0).Columns("NPL").ValueList = Me.UltraDropDown2
        e.Layout.Bands(0).Columns("PP").ValueList = Me.UltraDropDown2
        e.Layout.Bands(0).Columns("Matching With").ValueList = Me.UltraDropDown3
        e.Layout.Bands(0).Columns("Lead Time").ValueList = Me.UltraDropDown4
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        Call Delete_Projection(cboSO.Text)
        common.ClearAll(OPR0, OPR12)
        Clicked = ""
        OPR0.Enabled = True
        OPR12.Enabled = True
        Call Load_Data_Gride()
        '  Call Load_Gride_SalesOrder()
        Call Search_Parameter()
        chkAllocate.Checked = False
        Panel1.Visible = False
        OPR11.Visible = False
        Me.cboAllocate_Qulity.Text = ""
        cboStealing.Text = ""

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
        Dim M01 As DataSet
        Dim I As Integer
        Dim A As String
        Dim vcWhere As String
        Dim characterToRemove As String
        Dim _COMMENT As String
        Dim _PackBiz As String
        Dim _WeekNo As Integer

        Try



            'I = 0
            'vcWhere = "select * from tmpDye_Plan_Boad order by tmpRef_No "
            'M01 = DBEngin.ExecuteDataset(connection, transaction, vcWhere)
            'For Each DTRow3 As DataRow In M01.Tables(0).Rows
            '    Dim remain As Integer
            '    Dim noOfWeek As Integer
            '    Dim userDate As Date
            '    Dim _LastDate As Date
            '    Dim _TimeSpan As TimeSpan
            '    Dim _STDate As Date
            '    _STDate = Month(M01.Tables(0).Rows(I)("tmpStart_Date")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(I)("tmpStart_Date")) & "/" & Year(M01.Tables(0).Rows(I)("tmpStart_Date"))

            '    userDate = _STDate
            '    ' MsgBox(WeekdayName(Weekday(userDate)))
            '    If WeekdayName(Weekday(userDate)) = "Sunday" Then
            '        userDate = userDate.AddDays(-3)
            '    ElseIf WeekdayName(Weekday(userDate)) = "Monday" Then
            '        userDate = userDate.AddDays(-4)
            '    ElseIf WeekdayName(Weekday(userDate)) = "Tuesday" Then
            '        userDate = userDate.AddDays(-5)
            '    ElseIf WeekdayName(Weekday(userDate)) = "Thusday" Then
            '        ' userDate = userDate.AddDays(-1)
            '    ElseIf WeekdayName(Weekday(userDate)) = "Friday" Then
            '        userDate = userDate.AddDays(-1)
            '    ElseIf WeekdayName(Weekday(userDate)) = "Saturday" Then
            '        userDate = userDate.AddDays(-2)

            '    End If
            '    _LastDate = "1/1/" & Year(_STDate)
            '    If WeekdayName(Weekday(_LastDate)) = "Sunday" Then
            '        _LastDate = _LastDate.AddDays(-3)
            '    ElseIf WeekdayName(Weekday(_LastDate)) = "Monday" Then
            '        _LastDate = _LastDate.AddDays(-4)
            '    ElseIf WeekdayName(Weekday(_LastDate)) = "Tuesday" Then
            '        _LastDate = _LastDate.AddDays(-5)
            '    ElseIf WeekdayName(Weekday(_LastDate)) = "Thusday" Then
            '        ' userDate = userDate.AddDays(-1)
            '    ElseIf WeekdayName(Weekday(_LastDate)) = "Friday" Then
            '        _LastDate = _LastDate.AddDays(-1)
            '    ElseIf WeekdayName(Weekday(_LastDate)) = "Saturday" Then
            '        _LastDate = _LastDate.AddDays(-2)

            '    End If

            '    _TimeSpan = userDate.Subtract(_LastDate)
            '    _WeekNo = _TimeSpan.Days / 7

            '    nvcFieldList1 = "update tmpDye_Plan_Boad set tmpWeek=" & _WeekNo & ",tmpYear=" & Year(userDate) & " where tmpRef_No=" & M01.Tables(0).Rows(I)("tmpRef_No") & ""
            '    ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '    I = I + 1
            'Next
            'transaction.Commit()
            nvcFieldList1 = "select * from M01Sales_Order_SAP where convert(int,M01Sales_Order)='" & Trim(cboSO.Text) & "' and M01Status='A'"
            M01 = DBEngin.ExecuteDataset(connection, transaction, nvcFieldList1)
            If isValidDataset(M01) Then

            Else
                MsgBox("Please select the Sales Order No", MsgBoxStyle.Information, "Information ........")
                Exit Sub
            End If

            Call Search_Parameter()

            If Search_Planner(cboPlaner.Text) = True Then
            Else
                MsgBox("Please select the Planner's Name", MsgBoxStyle.Information, "Information .....")
                cboPlaner.ToggleDropdown()
                Exit Sub
            End If

            If chkNPL1.Checked = True Then
                _PackBiz = "YES"
            ElseIf chkNPL2.Checked = False Then
                _PackBiz = "NO"
            End If


            '            If Search_Lead_Time() = True Then
            '            Else
            '                Dim result1 As String
            '                result1 = MessageBox.Show("Please select the Lead Time ", "Information .....", _
            'MessageBoxButtons.OK, MessageBoxIcon.Information)
            '                If result1 = Windows.Forms.DialogResult.OK Then
            '                    cboLeadTime.ToggleDropdown()
            '                    Exit Sub
            '                End If
            '            End If

            'INCRESE PARAMETER FOR SALES ORDER
            'SO ------------->>> SALESE ORDER
            nvcFieldList1 = "update P01PARAMETER set P01NO=P01NO +" & 1 & " where P01CODE='SO'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            '-------------------------------------------------------------------------------------


            'CHECK NPL AND PP
            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Trim(UltraGrid1.Rows(I).Cells(8).Text))
                If Trim(UltraGrid1.Rows(I).Cells(9).Text) <> "" Then
                Else
                    MsgBox("Please select the NPL on Line Item -" & Trim(UltraGrid1.Rows(I).Cells(1).Value), MsgBoxStyle.Exclamation, "Exclamation ....")
                    Exit Sub
                End If

                I = I + 1
            Next

            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Trim(UltraGrid1.Rows(I).Cells(8).Text))
                If Trim(UltraGrid1.Rows(I).Cells(18).Text) <> "" Then
                Else
                    MsgBox("Please select the Lead time -" & Trim(UltraGrid1.Rows(I).Cells(1).Value), MsgBoxStyle.Exclamation, "Exclamation ....")
                    Exit Sub
                End If

                I = I + 1
            Next

            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Trim(UltraGrid1.Rows(I).Cells(9).Text))
                If Trim(UltraGrid1.Rows(I).Cells(9).Text) = "NOT APPROVED" And Trim(UltraGrid1.Rows(I).Cells(10).Text) = "" Then

                    MsgBox("Please select the NPL Approved Date on Line Item -" & Trim(UltraGrid1.Rows(I).Cells(1).Value), MsgBoxStyle.Exclamation, "Exclamation ....")
                    Exit Sub
                End If

                I = I + 1
            Next

            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Trim(UltraGrid1.Rows(I).Cells(8).Text))
                If Trim(UltraGrid1.Rows(I).Cells(11).Text) <> "" Then
                Else
                    MsgBox("Please select the PP on Line Item -" & Trim(UltraGrid1.Rows(I).Cells(1).Value), MsgBoxStyle.Exclamation, "Exclamation ....")
                    Exit Sub
                End If
                I = I + 1
            Next

            '--------------------------------------------------------------------
            'CHECK 1ST BULK

            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Trim(UltraGrid1.Rows(I).Cells(8).Text))
                If Trim(UltraGrid1.Rows(I).Cells(6).Text) <> "" Then
                Else
                    MsgBox("Please select the 1st Bulk on Line Item -" & Trim(UltraGrid1.Rows(I).Cells(1).Value), MsgBoxStyle.Exclamation, "Exclamation ....")
                    Exit Sub
                End If
                I = I + 1
            Next
            '------------------------------------------------------------------
            'CHECK REQUEST DEL DATE

            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Trim(UltraGrid1.Rows(I).Cells(8).Text))
                If Trim(UltraGrid1.Rows(I).Cells(20).Text) <> "" Then
                Else
                    MsgBox("Please select the Req.Del.Date on Line Item -" & Trim(UltraGrid1.Rows(I).Cells(1).Value), MsgBoxStyle.Exclamation, "Exclamation ....")
                    Exit Sub
                End If
                I = I + 1
            Next
            '---------------------------------------------------------------------------------------------------
            'POSSIBLE APPROVED DATE - LAB DYE

            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                ' MsgBox(Trim(UltraGrid1.Rows(I).Cells(8).Text))
                If Trim(UltraGrid1.Rows(I).Cells(6).Text) = "1st BULK" Then

                    If Trim(UltraGrid1.Rows(I).Cells(7).Text) = "Approved" Then
                    Else

                        If Trim(UltraGrid1.Rows(I).Cells(8).Text) <> "" Then
                        Else
                            MsgBox("Please select the Possible Approved Date on Line Item -" & Trim(UltraGrid1.Rows(I).Cells(1).Value), MsgBoxStyle.Exclamation, "Exclamation ....")
                            Exit Sub

                        End If
                    End If
                End If
                I = I + 1
            Next

            'INSERT DATA TO T01Delivary_Request TABLE
            I = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows

                nvcFieldList1 = "Insert Into T01Delivary_Request(T01RefNo,T01Sales_Order,T01Line_Item,T01Qty,T01Date,T01Maching,T01Lab_Dye,T01Bulk,T01POD,T01NPL,T01NPL_AppDate,T01PP,T01RQD,T01Balance,T01User,T01Planner,T01Status,T01Cab_TB_Tkn,T01GO_Rolling,T01GO_Non_Rolling,T01STRO_Repelenish,T01STRO_Reduce,T01Lead_Time,T01Pack_Biz,T01Comment,T01PP_AppDate)" & _
                                                         " values(" & Trim(txtRefNo.Text) & ", '" & Trim(cboSO.Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(1).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(3).Text) & "','" & Today & "','" & Trim(UltraGrid1.Rows(I).Cells(4).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(7).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(6).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(8).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(9).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(10).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(11).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(20).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(3).Text) & "','" & strDisname & "','" & cboPlaner.Text & "','A','" & Trim(UltraGrid1.Rows(I).Cells(13).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(14).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(15).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(16).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(17).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(18).Text) & "','" & _PackBiz & "','" & Trim(UltraGrid1.Rows(I).Cells(19).Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(20).Text) & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                'UPDATE STATUS FOR M01Sales_Order_SAP
                '------------------------------------------------------------------------------------------
                nvcFieldList1 = "update M01Sales_Order_SAP set M01Status='I' WHERE convert(int,M01Sales_Order)='" & Trim(cboSO.Text) & "' AND M01Line_Item='" & Trim(UltraGrid1.Rows(I).Cells(1).Text) & "'"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)

                I = I + 1
            Next

            '------------------------------------------------------------------
            'ADD SPECIAL COMMENT FOR PLANNER
            If txtComment.Text <> "" Then
                _COMMENT = Trim(txtComment.Text)
                characterToRemove = "'"

                _COMMENT = (Replace(_COMMENT, characterToRemove, ""))

                characterToRemove = ","

                _COMMENT = (Replace(_COMMENT, characterToRemove, ""))

                nvcFieldList1 = "Insert Into T01_1Special_Comment_Planner(T01_1Ref_No,T01_1Comment)" & _
                                                        " values(" & Trim(txtRefNo.Text) & ", '" & _COMMENT & "')"
                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            End If

            I = 0
            Dim X As Integer
            X = 0
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                X = 0
                vcWhere = "tmpSales_Order='" & cboSO.Text & "' and tmpLine_Item=" & Trim(UltraGrid1.Rows(I).Cells(1).Text) & ""
                M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP3"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                Else

                    MsgBox("Please allocate the Projection for the Line Item - " & Trim(UltraGrid1.Rows(I).Cells(1).Value), MsgBoxStyle.Information, "Information ......")
                    Exit Sub
                End If
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    vcWhere = "T15Sales_Order='" & cboSO.Text & "' and T15Line_Item=" & Trim(UltraGrid1.Rows(I).Cells(1).Text) & " and T15Code='" & Trim(M01.Tables(0).Rows(X)("tmpCode")) & "' and T15Shade='" & Trim(M01.Tables(0).Rows(X)("tmpShade")) & "' and T15Month=" & Trim(M01.Tables(0).Rows(X)("tmpMonth")) & " and T15Year=" & Trim(M01.Tables(0).Rows(X)("tmpYear")) & ""
                    dsUser = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP4"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                    Else
                        nvcFieldList1 = "Insert Into T15Projection_Allocation(T15RefNo,T15Sales_Order,T15Line_Item,T15Month,T15Year,T15Quality,T15Code,T15Shade,T15Qty,T15Planner,T15User,T15Status)" & _
                                                       " values(" & Trim(txtRefNo.Text) & ", '" & Trim(cboSO.Text) & "','" & Trim(UltraGrid1.Rows(I).Cells(1).Text) & "'," & Trim(M01.Tables(0).Rows(X)("tmpMonth")) & "," & Trim(M01.Tables(0).Rows(X)("tmpYear")) & ",'" & Trim(M01.Tables(0).Rows(X)("tmpQuality")) & "','" & Trim(M01.Tables(0).Rows(X)("tmpCode")) & "','" & Trim(M01.Tables(0).Rows(X)("tmpShade")) & "','" & Trim(M01.Tables(0).Rows(X)("tmpQty")) & "','" & cboPlaner.Text & "','" & strDisname & "','N')"
                        ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                    End If
                    X = X + 1
                Next
                I = I + 1
            Next

            nvcFieldList1 = "delete from tmpProjection_Allocation where tmpSales_Order='" & cboSO.Text & "' and tmpUser='" & strDisname & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()

            A = MsgBox("Are you sure you want to send e-mail to Planner", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Information .........")
            If A = vbYes Then
                Call Send_Email() '------------------SENDING EMAIL


            End If
            common.ClearAll(OPR0, OPR12)
            Clicked = ""
            OPR0.Enabled = True
            OPR12.Enabled = True
            ' cmdSave.Enabled = False
            Call Load_Sales_Order()
            Call Search_Parameter()
            Call Load_PLANNER()
            '  Call Load_Combo_Lead_Time()
            ' Call Load_Gride_SalesOrder()

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch ex As EvaluateException
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)
            connection.Close()
        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Sub
    Function Delete_Projection(ByVal strSO As String)
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Try
            nvcFieldList1 = "delete from tmpProjection_Allocation where tmpSales_Order='" & strSO & "' and tmpUser='" & strDisname & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
            transaction.Commit()

            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
        Catch ex As EvaluateException
            If transactionCreated = False Then transaction.Rollback()
            MessageBox.Show(Me, ex.ToString)
            connection.Close()
        Finally
            If connectionCreated Then DBEngin.CloseConnection(connection)
        End Try
    End Function

    Function Send_Email()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim _RefNo As String
        Dim T01 As DataSet
        Dim T02 As DataSet

      
        Dim exc As New Microsoft.Office.Interop.Excel.Application
        Dim workbooks As Microsoft.Office.Interop.Excel.Workbooks = exc.Workbooks
        Dim workbook As Microsoft.Office.Interop.Excel._Workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)
        Dim sheets As Microsoft.Office.Interop.Excel.Sheets = workbook.Worksheets
        Dim worksheet1 As Microsoft.Office.Interop.Excel._Worksheet = CType(sheets.Item(1), Microsoft.Office.Interop.Excel._Worksheet)
        Dim range1 As Microsoft.Office.Interop.Excel.Range

        Dim objApp As Object
        Dim objEmail As Object
        If Microsoft.VisualBasic.Len(txtRefNo.Text) = 1 Then
            _RefNo = "000" & Trim(txtRefNo.Text)
        ElseIf Microsoft.VisualBasic.Len(txtRefNo.Text) = 2 Then
            _RefNo = "00" & Trim(txtRefNo.Text)
        ElseIf Microsoft.VisualBasic.Len(txtRefNo.Text) = 3 Then
            _RefNo = "0" & Trim(txtRefNo.Text)
        Else
            _RefNo = Trim(txtRefNo.Text)
        End If
        objApp = CreateObject("Outlook.Application")
        objEmail = objApp.CreateItem(0)

        With objEmail
            .To = _Email
            .Subject = Trim(cboSO.Text) & "-" & _RefNo
            '.body = "Dear " & Trim(cboPlaner.Text) & "," & vbCr & vbCr _
            '         & "Please Quote best possible delivery for below" & vbCr & vbCr _
            '& "---------------------------------------------------------------------------------------------------------------------------" & vbCr _
            '& "| Line Item | "
          
            If exc.Visible = True Then
                exc.Visible = False
                exc.Visible = True
            Else
                ' exc.Visible = False
                exc.Visible = True
            End If
            worksheet1.Rows(5).Font.size = 10
            worksheet1.Rows(5).Font.bold = True
            'worksheet1.Rows(5).width = 23
            Dim A As Integer
            Dim i As Integer

            worksheet1.Cells(5, 15) = "Use Greige Order"
            worksheet1.Cells(5, 15).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            ' worksheet1.Columns("A").ColumnWidth = 12
            worksheet1.Range("o5:p5").MergeCells = True
            worksheet1.Range("o5:p5").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

            worksheet1.Cells(5, 17) = "Use Strategic Order"
            worksheet1.Cells(5, 17).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            ' worksheet1.Columns("A").ColumnWidth = 12
            worksheet1.Range("q5:r5").MergeCells = True
            worksheet1.Range("q5:r5").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

            A = 97

            For i = 1 To 18
                worksheet1.Range(Chr(A) & "5:" & Chr(A) & "5").Interior.Color = RGB(0, 112, 192)
                A = A + 1
            Next

            A = 97
            i = 0
            For i = 1 To 18
                worksheet1.Range(Chr(A) & "5", Chr(A) & "5").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "5", Chr(A) & "5").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "5", Chr(A) & "5").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "5", Chr(A) & "5").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next

            A = 97

            worksheet1.Rows(6).Font.size = 10
            worksheet1.Rows(6).Font.bold = True

            worksheet1.Cells(6, 1) = "S/O"
            worksheet1.Cells(6, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(6).Font.size = 10
            worksheet1.Columns("A").ColumnWidth = 10

            worksheet1.Cells(6, 2) = "Line Item"
            worksheet1.Cells(6, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("B").ColumnWidth = 12

            worksheet1.Cells(6, 3) = "Material"
            worksheet1.Cells(6, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("C").ColumnWidth = 20

            worksheet1.Cells(6, 4) = "Quality"
            worksheet1.Cells(6, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("D").ColumnWidth = 30

            worksheet1.Cells(6, 5) = "Quantity"
            worksheet1.Cells(6, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("E").ColumnWidth = 8

            worksheet1.Cells(6, 6) = "Matching"
            worksheet1.Cells(6, 6).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("F").ColumnWidth = 10

            worksheet1.Cells(6, 7) = "Retailer"
            worksheet1.Cells(6, 7).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("G").ColumnWidth = 15

            worksheet1.Cells(6, 8) = "1st Bulk"
            worksheet1.Cells(6, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("H").ColumnWidth = 15

            worksheet1.Cells(6, 9) = "Lab dye"
            worksheet1.Cells(6, 9).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("I").ColumnWidth = 10

            worksheet1.Cells(6, 10) = "P/App Date"
            worksheet1.Cells(6, 10).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("J").ColumnWidth = 10

            worksheet1.Cells(6, 11) = "NPL"
            worksheet1.Cells(6, 11).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("K").ColumnWidth = 8

            worksheet1.Cells(6, 12) = "NPL App Date"
            worksheet1.Cells(6, 12).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("L").ColumnWidth = 8


            worksheet1.Cells(6, 13) = "PP"
            worksheet1.Cells(6, 13).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("M").ColumnWidth = 10

            worksheet1.Cells(6, 14) = "Reg.Del.Date"
            worksheet1.Cells(6, 14).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("N").ColumnWidth = 10

            worksheet1.Cells(6, 15) = "Cap_TB_Tkn"
            worksheet1.Cells(6, 15).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("O").ColumnWidth = 10

            worksheet1.Cells(6, 16) = "GO_Rolling"
            worksheet1.Cells(6, 16).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("P").ColumnWidth = 10

            worksheet1.Cells(6, 17) = "GO_Non_Rolling"
            worksheet1.Cells(6, 17).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("Q").ColumnWidth = 10


            worksheet1.Cells(6, 18) = "STRO_Repelenish"
            worksheet1.Cells(6, 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("R").ColumnWidth = 10

            worksheet1.Cells(6, 19) = "STRO_Reduce"
            worksheet1.Cells(6, 19).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Columns("S").ColumnWidth = 10

            'Dim A As Integer
            'Dim i As Integer

            A = 97

            For i = 1 To 19
                worksheet1.Range(Chr(A) & "6:" & Chr(A) & "6").Interior.Color = RGB(0, 112, 192)
                A = A + 1
            Next

            A = 97
            i = 0
            For i = 1 To 19
                worksheet1.Range(Chr(A) & "6", Chr(A) & "6").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "6", Chr(A) & "6").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "6", Chr(A) & "6").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & "6", Chr(A) & "6").Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next
            '------------------------------------------------------------------------------------------------------------------------
            Sql = "select * from T01Delivary_Request  where T01RefNo=" & txtRefNo.Text & ""
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            Dim X As Integer
            Dim Y As Integer

            X = 7
            Y = 1
            For Each DTRow As DataRow In M01.Tables(0).Rows
                Y = 1
                worksheet1.Rows(X).Font.size = 8
                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Sales_Order")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Line_Item")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                Sql = "SELECT * FROM M01Sales_Order_SAP WHERE convert(int,M01Sales_Order)='" & Trim(M01.Tables(0).Rows(i)("T01Sales_Order")) & "' AND M01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T01Line_Item")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    worksheet1.Cells(X, Y) = T01.Tables(0).Rows(0)("M01Material_No")
                    ' worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1


                    worksheet1.Cells(X, Y + 3) = T01.Tables(0).Rows(0)("M01Department")
                    worksheet1.Cells(X, Y + 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter


                    worksheet1.Cells(X, Y) = T01.Tables(0).Rows(0)("M01Quality")
                    ' worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1
                End If
                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Qty")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                If Trim(M01.Tables(0).Rows(i)("T01Maching")) <> "" Then
                    worksheet1.Cells(X, Y) = "Yes"
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                End If
                Y = Y + 1

                

                worksheet1.Cells(X, Y + 1) = M01.Tables(0).Rows(i)("T01Bulk")
                worksheet1.Cells(X, Y + 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 2

                If Trim(M01.Tables(0).Rows(i)("T01Lab_Dye")) <> "" Then

                    worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Lab_Dye")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Else

                    worksheet1.Cells(X, Y) = "NOT APPROVED"
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                End If
                Y = Y + 1

                If Trim(M01.Tables(0).Rows(i)("T01Bulk")) = "1st BULK" Then
                    worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01POD")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                End If
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01NPL")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1
                If Trim(M01.Tables(0).Rows(i)("T01NPL")) = "NOT APPROVED" Then
                    worksheet1.Cells(X, Y) = Month(M01.Tables(0).Rows(i)("T01NPL_AppDate")) & "/" & Microsoft.VisualBasic.Day(M01.Tables(0).Rows(i)("T01NPL_AppDate")) & "/" & Microsoft.VisualBasic.Year(M01.Tables(0).Rows(i)("T01NPL_AppDate"))
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                End If
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01PP")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01RQD")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Cab_TB_Tkn")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01GO_Rolling")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01GO_Non_Rolling")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01STRO_Repelenish")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01STRO_Reduce")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1

                ' Dim Z As Integer
                A = 97
                ' i = 0
                Dim Z As Integer
                For Z = 1 To 19
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    A = A + 1
                Next

                X = X + 1
                i = i + 1
            Next

            X = X + 2
            Y = 5
            worksheet1.Cells(X, Y) = "Body Order Info"
            worksheet1.Range(worksheet1.Cells(X, Y), worksheet1.Cells(X, Y + 2)).Merge()
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.BOLD = True

            worksheet1.Cells(X, Y + 4) = "Trim Order Info"
            worksheet1.Range(worksheet1.Cells(X, Y + 4), worksheet1.Cells(X, Y + 6)).Merge()
            worksheet1.Cells(X, Y + 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
           
            worksheet1.Cells(X, Y + 8) = "T/B Ratio"
            A = 101
            i = 0
            Dim Z1 As Integer

            For Z1 = 5 To 13
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next

            A = 101

            For Z1 = 5 To 13
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)
                A = A + 1
            Next

            X = X + 1
            Y = 5
            worksheet1.Cells(X, Y) = "L/Item"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.BOLD = True
            Y = Y + 1
            worksheet1.Cells(X, Y) = "Quality"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            Y = Y + 1
            worksheet1.Cells(X, Y) = "Qty"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            Y = Y + 2
            worksheet1.Cells(X, Y) = "L/Item"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(X).Font.size = 10
            worksheet1.Rows(X).Font.BOLD = True
            Y = Y + 1
            worksheet1.Cells(X, Y) = "Quality"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            Y = Y + 1
            worksheet1.Cells(X, Y) = "Qty"
            worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            Y = Y + 2
            A = 101
            For Z1 = 5 To 13
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next
            A = 104
            worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)
            A = 108
            worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)

            Sql = "select * from T01Delivary_Request where T01RefNo=" & txtRefNo.Text & "  and T01Maching<>''"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0

            For Each DTRow As DataRow In M01.Tables(0).Rows
                X = X + 1
                Y = 5
                worksheet1.Rows(X).Font.size = 8

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Line_Item")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 1
                Sql = "select * from M01Sales_Order_SAP where convert(int,M01Sales_Order)='" & Trim(M01.Tables(0).Rows(i)("T01Sales_Order")) & "' and M01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T01Line_Item")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    worksheet1.Cells(X, Y) = T01.Tables(0).Rows(0)("M01Quality")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1
                End If

                worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Qty")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                Y = Y + 2

                Sql = "select * from T01Delivary_Request where T01Sales_Order='" & Trim(M01.Tables(0).Rows(i)("T01Sales_Order")) & "' and T01Line_Item='" & Trim(M01.Tables(0).Rows(i)("T01Maching")) & "'"
                T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                If isValidDataset(T01) Then
                    worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("T01Maching")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 1

                    Sql = "select * from M01Sales_Order_SAP where convert(int,M01Sales_Order)='" & Trim(T01.Tables(0).Rows(0)("T01Sales_Order")) & "' and M01Line_Item='" & Trim(T01.Tables(0).Rows(0)("T01Line_Item")) & "'"
                    T02 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                    If isValidDataset(T02) Then
                        worksheet1.Cells(X, Y) = T02.Tables(0).Rows(0)("M01Quality")
                        worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        Y = Y + 1
                    End If
                    worksheet1.Cells(X, Y) = T01.Tables(0).Rows(0)("T01Qty")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    Y = Y + 2

                    worksheet1.Cells(X, Y) = "=K" & X & "/G" & X
                    range1 = worksheet1.Cells(X, Y)
                    range1.NumberFormat = "0.00"
                End If

                A = 101
                For Z1 = 5 To 13
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    A = A + 1
                Next

                A = 104
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)
                A = 108
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)

                i = i + 1
            Next

            X = X + 2
            worksheet1.Cells(X, 4) = "Projection Allocation"
            worksheet1.Range("D" & X & ":D" & X).MergeCells = True
            worksheet1.Range("D" & X & ":D" & X).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            worksheet1.Rows(X).font.size = 10
            worksheet1.Rows(X).font.bold = True

            X = X + 2
            worksheet1.Cells(X, 4) = "Code"
            worksheet1.Range("D" & X & ":D" & X).MergeCells = True
            worksheet1.Range("D" & X & ":D" & X).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Rows(X).font.size = 10
            worksheet1.Rows(X).font.bold = True

            worksheet1.Range("D" & X & ":D" & X).Interior.Color = RGB(0, 112, 192)

            worksheet1.Cells(X, 5) = "Line Item"
            worksheet1.Range("E" & X & ":E" & X).MergeCells = True
            worksheet1.Range("E" & X & ":E" & X).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            worksheet1.Cells(X, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            worksheet1.Range("E" & X & ":E" & X).Interior.Color = RGB(0, 112, 192)

            Sql = "select T15Month from T15Projection_Allocation where T15RefNo=" & txtRefNo.Text & " group by T15Year,T15Month order by T15Year,T15Month"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            i = 0
            Y = 6
            A = 70
            For Each DTRow As DataRow In M01.Tables(0).Rows
                worksheet1.Cells(X, Y) = MonthName(M01.Tables(0).Rows(i)("T15Month"))
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).MergeCells = True
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                worksheet1.Rows(X).font.size = 10
                worksheet1.Rows(X).font.bold = True
                worksheet1.Range(Chr(A) & X & ":" & Chr(A) & X).Interior.Color = RGB(0, 112, 192)
                Y = Y + 1
                A = A + 1
                i = i + 1
            Next

            A = 68
            For Z1 = 6 To Y + 1
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                A = A + 1
            Next
            X = X + 1
            Z1 = 0
            Y = 4
            Sql = "select T15Code,T15Line_Item from T15Projection_Allocation where T15RefNo=" & txtRefNo.Text & " group by T15Code,T15Line_Item"
            dsUser = DBEngin.ExecuteDataset(con, Nothing, Sql)
            For Each DTRow As DataRow In dsUser.Tables(0).Rows
                A = 68
                Y = 4
                worksheet1.Rows(X).Font.size = 8
                worksheet1.Cells(X, Y) = dsUser.Tables(0).Rows(Z1)("T15Code")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

                Y = Y + 1
                A = A + 1
                worksheet1.Cells(X, Y) = dsUser.Tables(0).Rows(Z1)("T15Line_Item")
                worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

                Sql = "select SUM(T15Qty) AS QTY from T15Projection_Allocation where T15RefNo=" & txtRefNo.Text & "  AND T15Code ='" & dsUser.Tables(0).Rows(Z1)("T15Code") & "' and T15Line_Item =" & dsUser.Tables(0).Rows(Z1)("T15Line_Item") & " group by T15Year,T15Month order by T15Year,T15Month"
                M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
                i = 0
                Y = 6
                A = A + 1
                For Each DTRow1 As DataRow In M01.Tables(0).Rows
                    worksheet1.Cells(X, Y) = M01.Tables(0).Rows(i)("QTY")
                    worksheet1.Cells(X, Y).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    range1 = worksheet1.Cells(X, Y)
                    range1.NumberFormat = "0.00"

                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                    worksheet1.Range(Chr(A) & X, Chr(A) & X).Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

                    Y = Y + 1
                    A = A + 1
                    i = i + 1
                Next
                X = X + 1
                Z1 = Z1 + 1
            Next
            '------------------------------------------------------------------------------------------------------------------------
            Dim xlRn As Excel.Range
            Dim Connect As String
            Dim strbody As String

            'strBody = "This is a test " & vbCrLf & vbCrLf & "Thanks Michael"
            '  RangetoHTML(xlRn)

            Connect = worksheet1.Range("A5:S" & X - 1).Copy()
            xlRn = worksheet1.Range("A5:S" & X + 1)
            'xlRn.Copy()

            '.HTMLBody = " Dear" & Trim(cboPlaner.Text) & "," & vbNewLine & _
            '              "Please Quote best possible delivery for below" & Chr(10) _
            '                                   & RangetoHTML(xlRn)
            If Trim(_LeadTime) = "01" Then
                .HTMLBody = "Dear " & Trim(cboPlaner.Text) & ",<br>Please Quote best possible delivery for below " & RangetoHTML(xlRn)
            Else
                .HTMLBody = "Dear " & Trim(cboPlaner.Text) & ",<br>Please Quote best possible delivery for below Speed Order" & RangetoHTML(xlRn)
            End If
            .display()
        End With
        objEmail = Nothing
        objApp = Nothing

        workbooks.Close()

    End Function

    Function RangetoHTML(ByVal rng As Microsoft.Office.Interop.Excel.Range)
        ' Changed by Ron de Bruin 28-Oct-2006
        ' Working in Office 2000-2010
        Dim fso As Object
        Dim ts As Object
        Dim TempFile As String
        ' Dim TempWB As Microsoft.Office.Interop.Excel.Workbook

        Dim exc As New Microsoft.Office.Interop.Excel.Application
        Dim TempWB1 As Microsoft.Office.Interop.Excel.Workbooks = exc.Workbooks
        Dim TempWB As Microsoft.Office.Interop.Excel._Workbook = TempWB1.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)

        TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

        'Copy the range and create a new workbook to past the data in
        rng.Copy()
        'TempWB = Microsoft.Office.Interop.Excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet)




        With TempWB.Sheets(1)
            .Cells(1).PasteSpecial(Paste:=8)
            ' Microsoft.Office.Interop.Excel.XlPastef
            '.Cells(1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, , False, False)
            '.Cells(1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, , False, False)
            '.Cells(1).Select()
            'Application.CutCopyMode = False
            On Error Resume Next
            .DrawingObjects.Visible = True
            .DrawingObjects.Delete()
            On Error GoTo 0
        End With


        'Publish the sheet to a htm file
        With TempWB.PublishObjects.Add( _
             SourceType:=Microsoft.Office.Interop.Excel.XlSourceType.xlSourceRange, _
             Filename:=TempFile, _
             Sheet:=TempWB.Sheets(1).Name, _
             Source:=TempWB.Sheets(1).UsedRange.Address, _
             HtmlType:=Microsoft.Office.Interop.Excel.XlHtmlType.xlHtmlStatic)
            .Publish(True)
        End With

        'Read all data from the htm file into RangetoHTML
        fso = CreateObject("Scripting.FileSystemObject")
        ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
        RangetoHTML = ts.ReadAll
        ts.Close()
        RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                              "align=left x:publishsource=")

        'Close TempWB
        TempWB.Close(savechanges:=False)

        'Delete the htm file we used in this function
        Kill(TempFile)

        ts = Nothing
        fso = Nothing
        TempWB = Nothing
    End Function

    Function Search_Planner(ByVal strName As String) As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet

        Try
            'SERCH PLANNER'S EPF NO
            Sql = "select * from users where Name='" & Trim(cboPlaner.Text) & "'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                Search_Planner = True
                _EPF = M01.Tables(0).Rows(0)("EPF")
                _Email = Trim(M01.Tables(0).Rows(0)("email"))

            End If


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

    Private Sub UltraLabel4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    'Function Load_Combo_Lead_Time()
    '    Dim Sql As String
    '    Dim con = New SqlConnection()
    '    con = DBEngin.GetConnection(True)
    '    Dim M01 As DataSet
    '    'Load sales order to cboSO combobox

    '    Try
    '        Sql = "select M02Dis as [Lead Time] from M02Lead_Time_Master where M02Code in ('01','03')"
    '        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
    '        With cboLeadTime
    '            .DataSource = M01
    '            .Rows.Band.Columns(0).Width = 180
    '            '   .Rows.Band.Columns(1).Width = 260


    '        End With
    '        DBEngin.CloseConnection(con)
    '        con.ConnectionString = ""
    '        'With txtNSL
    '        '    .DataSource = M01
    '        '    .Rows.Band.Columns(0).Width = 225
    '        'End With

    '    Catch returnMessage As EvaluateException
    '        If returnMessage.Message <> Nothing Then
    '            MessageBox.Show(returnMessage.Message)
    '        End If
    '    End Try
    ' End Function

    Function Load_Gride2()
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection()
        Dim M01 As DataSet
        Dim I As Integer
        Dim vcWhere As String
        Dim _Code As Integer
        Dim T01 As DataSet
        Dim agroup1 As UltraGridGroup
        Dim agroup2 As UltraGridGroup
        Dim agroup3 As UltraGridGroup
        Dim agroup4 As UltraGridGroup
        Dim agroup5 As UltraGridGroup
        Dim _Date As Date

        Try
            'Dim agroup1 As UltraGridGroup
            'Dim agroup2 As UltraGridGroup
            'Dim agroup3 As UltraGridGroup
            'Dim agroup4 As UltraGridGroup
            'Dim agroup5 As UltraGridGroup
            '  Dim agroup6 As UltraGridGroup

            'If UltraGrid3.DisplayLayout.Bands(0).GroupHeadersVisible = True Then
            'Else
            '  agroup1.Key = ""
            '  agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Remove("GroupH")
            agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Add("GroupH")

            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("Line", "Line Item")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("Line").Group = agroup1
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("Line").Width = 50

            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("##", "##")
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("##").Group = agroup1
            'Me.UltraGrid3.DisplayLayout.Bands(0).Columns("##").Width = 120
            ''  End If
            ' agroup1 = UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)
          
            vcWhere = "select * from P01PARAMETER where P01CODE='PRN'"
            M01 = DBEngin.ExecuteDataset(con, Nothing, vcWhere)
            If isValidDataset(M01) Then
                _Code = M01.Tables(0).Rows(0)("P01NO")
            End If

            _Code = _Code - 1
            agroup1.Header.Caption = "Code"
            If cboStealing.Text <> "" Then
                If Microsoft.VisualBasic.Day(Today) > 10 Then
                    _Date = Today.AddDays(+30)
                    _Date = Month(_Date) & "/1/" & Year(_Date)
                    vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  And M43Count_No = " & _Code & """"
                Else
                    _Date = Today
                    _Date = Month(_Date) & "/1/" & Year(_Date)
                    vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "' "
                End If
            Else
                If Microsoft.VisualBasic.Day(Today) > 10 Then
                    _Date = Today.AddDays(+30)
                    _Date = Month(_Date) & "/1/" & Year(_Date)
                    vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                Else
                    _Date = Today
                    _Date = Month(_Date) & "/1/" & Year(_Date)
                    vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "' "
                End If
            End If
            agroup1.Width = 110
            Dim dt As DataTable = New DataTable()
            ' dt.Columns.Add("ID", GetType(Integer))
            Dim colWork As New DataColumn("##", GetType(String))
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True
            colWork = New DataColumn("Shade", GetType(String))
            colWork.MaxLength = 250
            dt.Columns.Add(colWork)
            colWork.ReadOnly = True

            'dt.Columns.Add("##", GetType(String))
            ' dt.Columns.Add("Shade", GetType(String))
            I = 0
            If cboStealing.Text <> "" Then
                If Microsoft.VisualBasic.Day(Today) > 10 Then
                    vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "' and M43Count_No=" & _Code & ""
                Else
                    vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "' "
                End If
            Else
                If Microsoft.VisualBasic.Day(Today) > 10 Then
                    vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                Else
                    vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "'"
                End If
            End If
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRSU"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                dt.Rows.Add(M01.Tables(0).Rows(I)("Code"), UCase(M01.Tables(0).Rows(I)("M43Shade")))
                I = I + 1
            Next

            Me.UltraGrid3.SetDataBinding(dt, Nothing)
            Me.UltraGrid3.DisplayLayout.Bands(0).Columns(0).Group = agroup1
            Me.UltraGrid3.DisplayLayout.Bands(0).Columns(1).Group = agroup1
            Me.UltraGrid3.DisplayLayout.Bands(0).Columns(1).CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            Me.UltraGrid3.DisplayLayout.Bands(0).Columns(0).Width = 180
            Me.UltraGrid3.DisplayLayout.Bands(0).Columns(1).Width = 50
            Dim _Group As String
            'agroup2.Key = ""
            'agroup3.Key = ""
            'agroup4.Key = ""
            '' agroup5.Key = ""

            I = 0
            'T01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LOMN"), New SqlParameter("@vcWhereClause1", vcWhere))
            For Each DTRow3 As DataRow In M01.Tables(0).Rows
                _Group = "Group" & I + 1
                If I = 0 Then
                    'agroup2.Key = ""
                    agroup2 = UltraGrid3.DisplayLayout.Bands(0).Groups.Add("Group1")

                    agroup2.Header.Caption = Trim(M01.Tables(0).Rows(I)("M13Name"))
                    agroup2.Width = 220
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("ProjectColumn", "Projection")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn").Group = agroup2
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("UseColumn", "Used")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn").Group = agroup2
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("AlcColumn", "Allocate")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn").Group = agroup2
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ElseIf I = 1 Then
                    agroup3 = UltraGrid3.DisplayLayout.Bands(0).Groups.Add("Group2")
                    agroup3.Header.Caption = Trim(M01.Tables(0).Rows(I)("M13Name"))
                    agroup3.Width = 220
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("ProjectColumn1", "Projection")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn1").Group = agroup3
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn1").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("UseColumn1", "Used")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn1").Group = agroup3
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn1").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("AlcColumn1", "Allocate")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn1").Group = agroup3
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn1").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn1").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ElseIf I = 2 Then
                    agroup4 = UltraGrid3.DisplayLayout.Bands(0).Groups.Add("Group3")
                    agroup4.Header.Caption = Trim(M01.Tables(0).Rows(I)("M13Name"))
                    agroup4.Width = 220
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("ProjectColumn2", "Projection")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn2").Group = agroup4
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn2").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn2").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("UseColumn2", "Used")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn2").Group = agroup4
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn2").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn2").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("AlcColumn2", "Allocate")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn2").Group = agroup4
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn2").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn2").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ElseIf I = 3 Then
                    agroup5 = UltraGrid3.DisplayLayout.Bands(0).Groups.Add("Group4")
                    agroup5.Header.Caption = Trim(M01.Tables(0).Rows(I)("M13Name"))
                    agroup5.Width = 220
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("ProjectColumn3", "Projection")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn3").Group = agroup5
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn3").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("ProjectColumn3").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("UseColumn3", "Used")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn3").Group = agroup5
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn3").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("UseColumn3").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns.Add("AlcColumn3", "Allocate")
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn3").Group = agroup5
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn3").Width = 70
                    Me.UltraGrid3.DisplayLayout.Bands(0).Columns("AlcColumn3").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                End If
                I = I + 1
            Next

            Dim _Coloum_Count As Integer
            Dim x As Integer
            x = 0
            For Each uRow As UltraGridRow In UltraGrid3.Rows
                _Coloum_Count = 2
                If cboStealing.Text <> "" Then
                    If Microsoft.VisualBasic.Day(Today) > 10 Then
                        vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                    Else
                        vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "'"
                    End If
                Else
                    If Microsoft.VisualBasic.Day(Today) > 10 Then
                        vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                    Else
                        vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "'"
                    End If
                End If
                I = 0
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LOM1"), New SqlParameter("@vcWhereClause1", vcWhere))
                For Each DTRow3 As DataRow In M01.Tables(0).Rows
                    Dim Value As Double
                    Dim _String As String
                    Dim Prjcell As UltraGridCell
                    'Dim _Str As String
                    '_Str = "ProjectColumn" & I + 1

                    If cboStealing.Text <> "" Then
                        vcWhere = "code='" & Trim(UltraGrid3.Rows(x).Cells(0).Text) & "' and M43Count_No=" & _Code & " and M43Shade='" & Trim(UltraGrid3.Rows(x).Cells(1).Text) & "' and M43Quality='" & cboStealing.Text & "' and M43Product_Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                    Else
                        vcWhere = "code='" & Trim(UltraGrid3.Rows(x).Cells(0).Text) & "' and M43Count_No=" & _Code & " and M43Shade='" & Trim(UltraGrid3.Rows(x).Cells(1).Text) & "' and M43Quality='" & txtQuality.Text & "' and M43Product_Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                    End If
                    dsUser = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS1"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(dsUser) Then
                        Value = dsUser.Tables(0).Rows(0)("Qty")
                        _String = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                        _String = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))

                        UltraGrid3.Rows(x).Cells(_Coloum_Count).Value = _String

                        'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION TMPPROJECTION TABLE
                        'DEVELOPED BY SURANGA WIJESINGHE ON 01/04/2016
                        Value = 0
                        If cboStealing.Text <> "" Then
                            vcWhere = "tmpCode='" & Trim(UltraGrid3.Rows(x).Cells(0).Text) & "' and tmpShade='" & Trim(UltraGrid3.Rows(x).Cells(1).Text) & "' and tmpQuality='" & cboStealing.Text & "' and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & " and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                        Else
                            vcWhere = "tmpCode='" & Trim(UltraGrid3.Rows(x).Cells(0).Text) & "' and tmpShade='" & Trim(UltraGrid3.Rows(x).Cells(1).Text) & "' and tmpQuality='" & txtQuality.Text & "' and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & " and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                        End If
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP1"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T01) Then
                            Value = T01.Tables(0).Rows(0)("tmpQty")
                        End If
                        '===========================================================END STATMENT
                        'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION T15Projection
                        'DEVELOPED BY SURANGA WIJESINGHE ON 01/04/2016
                        If cboStealing.Text <> "" Then
                            vcWhere = "T15Code='" & Trim(UltraGrid3.Rows(x).Cells(0).Text) & "' and T15Shade='" & Trim(UltraGrid3.Rows(x).Cells(1).Text) & "' and T15Quality='" & cboStealing.Text & "' and T15Year=" & M01.Tables(0).Rows(I)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                        Else
                            vcWhere = "T15Code='" & Trim(UltraGrid3.Rows(x).Cells(0).Text) & "' and T15Shade='" & Trim(UltraGrid3.Rows(x).Cells(1).Text) & "' and T15Quality='" & txtQuality.Text & "' and T15Year=" & M01.Tables(0).Rows(I)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                        End If
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP2"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T01) Then
                            Value = Value + T01.Tables(0).Rows(0)("T15Qty")
                        End If
                        '===========================================================END STATMENT

                        If Value > 0 Then
                            _String = (Value.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            _String = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value))
                            UltraGrid3.Rows(x).Cells(_Coloum_Count + 1).Value = _String
                        End If
                        Prjcell = UltraGrid3.Rows(x).Cells(_Coloum_Count)
                        Prjcell.Activation = Activation.NoEdit

                        Prjcell = UltraGrid3.Rows(x).Cells(_Coloum_Count + 1)
                        Prjcell.Activation = Activation.NoEdit

                        _Coloum_Count = _Coloum_Count + 3

                    Else
                        Prjcell = UltraGrid3.Rows(x).Cells(_Coloum_Count)
                        Prjcell.Activation = Activation.NoEdit

                        Prjcell = UltraGrid3.Rows(x).Cells(_Coloum_Count + 1)
                        Prjcell.Activation = Activation.NoEdit

                        _Coloum_Count = _Coloum_Count + 3
                        'Prjcell = UltraGrid3.Rows(I).Cells(_Str)
                        'Prjcell.Activation = Activation.NoEdit
                    End If
                    I = I + 1
                Next
                x = x + 1
            Next
            Dim Value1 As Double

            I = 0
            'vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "'  and M43Count_No=" & _Code & "  and M43Product_Month>='" & Month(_Date) & "' and M43Year>= '" & Year(_Date) & "'"
            'T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "DRK"), New SqlParameter("@vcWhereClause1", vcWhere))

            For Each uRow As UltraGridRow In UltraGrid1.Rows
                'For Each DTRow3 As DataRow In T01.Tables(0).Rows
                If UltraGrid1.Rows(I).Cells(0).Value = True Then
                    Value1 = 0
                    vcWhere = "M01Sales_Order=" & cboSO.Text & " and M01Line_Item=" & Trim(UltraGrid1.Rows(I).Cells(1).Value) & ""
                    M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                    If isValidDataset(M01) Then
                        vcWhere = "M16Material='" & M01.Tables(0).Rows(0)("M0130Class") & "'"
                        T01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "RCODE"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(T01) Then
                            ' txtCLR.Text = T01.Tables(0).Rows(0)("M16Shade_Type")
                            ' MsgBox(Trim(T01.Tables(0).Rows(0)("M16Shade_Type")))
                            If Trim(T01.Tables(0).Rows(0)("M16Shade_Type")) = "Dark" Then
                                Value1 = M01.Tables(0).Rows(0)("M01SO_Qty") + CDbl(lblDark.Text)
                                lblDark.Text = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                lblDark.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value1))
                            ElseIf Trim(T01.Tables(0).Rows(0)("M16Shade_Type")) = "Light" Then
                                Value1 = M01.Tables(0).Rows(0)("M01SO_Qty") + CDbl(lblLight.Text)
                                lblLight.Text = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                lblLight.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value1))

                            ElseIf Trim(T01.Tables(0).Rows(0)("M16Shade_Type")) = "Medium" Then
                                Value1 = M01.Tables(0).Rows(0)("M01SO_Qty") + CDbl(lblLight.Text)
                                lblLight.Text = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                lblLight.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value1))

                            ElseIf Trim(T01.Tables(0).Rows(0)("M16Shade_Type")) = "White" Then
                                Value1 = M01.Tables(0).Rows(0)("M01SO_Qty") + CDbl(lblWhite.Text)
                                lblWhite.Text = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                lblWhite.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value1))
                            ElseIf Trim(T01.Tables(0).Rows(0)("M16Shade_Type")) = "Black" Then
                                Value1 = M01.Tables(0).Rows(0)("M01SO_Qty") + CDbl(lblBlack.Text)
                                lblBlack.Text = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                lblBlack.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value1))
                            ElseIf Trim(T01.Tables(0).Rows(0)("M16Shade_Type")) = "Marl" Then
                                Value1 = M01.Tables(0).Rows(0)("M01SO_Qty") + CDbl(lblMarl.Text)
                                lblMarl.Text = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                                lblMarl.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value1))
                            End If
                        End If
                    End If
                End If
                I = I + 1
            Next

            con.close()
        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()

            End If
        End Try
    End Function

    Private Sub cboSO_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboSO.KeyUp
        If e.KeyCode = 13 Then
            strSales_Order = cboSO.Text
            Call Search_Salrs_Order()
            cboPlaner.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            strSales_Order = cboSO.Text
            Call Search_Salrs_Order()
            cboPlaner.ToggleDropdown()
        End If
    End Sub

    Function Search_Salrs_Order() As Boolean
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim i As Integer

        Try
            'SERCH SALES ORDER
            Sql = "select T01Sales_Order  from T01Delivary_Request  where T01RefNo='" & Trim(txtRefNo.Text) & "' and T01Planner='" & strDisname & "' and T01Status='A' "
            M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
            If isValidDataset(M01) Then
                strSales_Order = M01.Tables(0).Rows(0)("T01Sales_Order")
                cboSO.Text = M01.Tables(0).Rows(0)("T01Sales_Order")
                Search_Salrs_Order = True

                cmdSave.Enabled = True
            Else
                Search_Salrs_Order = False
            End If
            '----------------------------------------------------------------------------------


            DBEngin.CloseConnection(con)
            con.ConnectionString = ""

        Catch returnMessage As EvaluateException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
            End If
        End Try
    End Function

  

    'Function Search_Lead_Time() As Boolean

    '    Dim Sql As String
    '    Dim con = New SqlConnection()
    '    con = DBEngin.GetConnection(True)
    '    Dim M01 As DataSet
    '    'Search Referance No via the P01PARAMETER Table
    '    Try
    '        Sql = "select * from M02Lead_Time_Master where M02Dis='" & Trim(cboLeadTime.Text) & "'"
    '        M01 = DBEngin.ExecuteDataset(con, Nothing, Sql)
    '        If isValidDataset(M01) Then
    '            _LeadTime = Trim(M01.Tables(0).Rows(0)("M02Code"))
    '            Search_Lead_Time = True
    '        Else
    '            Search_Lead_Time = False
    '        End If
    '        DBEngin.CloseConnection(con)
    '        con.ConnectionString = ""


    '    Catch returnMessage As EvaluateException
    '        If returnMessage.Message <> Nothing Then
    '            MessageBox.Show(returnMessage.Message)
    '        End If
    '    End Try
    'End Function


    Private Sub cboLeadTime_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = 13 Then
            cboPlaner.ToggleDropdown()
        ElseIf e.KeyCode = Keys.Tab Then
            cboPlaner.ToggleDropdown()
        End If
    End Sub

   
    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        cboStealing.Text = ""

        'Call Clear_Gride()
        '' agroup1.Key.Remove(0)
        'UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)
        cboStealing.Text = ""
        UltraGrid3.DisplayLayout.Bands(0).Groups.Clear()
        UltraGrid3.DisplayLayout.Bands(0).Columns.Dispose()
        UltraGroupBox1.Visible = False
    End Sub

    Private Sub UltraButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton3.Click
        'UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)
        UltraGrid3.DisplayLayout.Bands(0).Groups.Clear()
        UltraGrid3.DisplayLayout.Bands(0).Columns.Dispose()
        UltraGroupBox1.Visible = False

        Dim TipInfo As New UltraToolTipInfo()
        TipInfo.ToolTipText = ""
        Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraGrid1, TipInfo)
        Me.UltraToolTipManager1.DisplayStyle = Infragistics.Win.ToolTipDisplayStyle.BalloonTip
    End Sub

    Private Sub UltraButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton5.Click
        With UltraGroupBox1
            .Width = 919
            .Height = 325
        End With
    End Sub

    Private Sub UltraButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton4.Click
        With UltraGroupBox1
            .Width = 130
            .Height = 40
        End With
    End Sub

    Private Sub UltraButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton7.Click

        With UltraGroupBox1
            .Location = New Point(103, 173)

            .Width = 919
            .Height = 325
        End With

        With UltraButton1
            .Location = New Point(843, 282)
        End With
        With UltraButton2
            .Location = New Point(760, 282)
        End With

        With UltraButton8
            .Location = New Point(630, 282)
        End With

        With UltraLabel11
            .Location = New Point(6, 282)
        End With

        With UltraLabel12
            .Location = New Point(166, 282)
        End With

        With lblBalance
            .Location = New Point(236, 282)
        End With


        With lblQty
            .Location = New Point(82, 283)
        End With
        With UltraGrid3
            .Height = 151
        End With
    End Sub

    Private Sub UltraButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton6.Click
        With UltraGroupBox1
            .Location = New Point(90, 10)

            .Width = 919
            .Height = 476
        End With
        With UltraButton1
            .Location = New Point(843, 420)
        End With
        With UltraButton2
            .Location = New Point(760, 420)
        End With

        With UltraButton8
            .Location = New Point(630, 420)
        End With

        With UltraLabel11
            .Location = New Point(6, 420)
        End With

        With UltraLabel12
            .Location = New Point(166, 420)
        End With

        With lblBalance
            .Location = New Point(236, 420)
        End With

        With lblQty
            .Location = New Point(82, 420)
        End With

        With UltraGrid3
            .Height = 280
        End With
    End Sub

    Private Sub UltraButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Projection_Allocation()
    End Sub

    Private Sub UltraButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton10.Click
        Panel1.Visible = False
        OPR11.Visible = False
        Me.cboAllocate_Qulity.Text = ""
    End Sub

    Private Sub chkAllocate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllocate.CheckedChanged
        If chkAllocate.Checked = True Then
            Call Load_Quality_Allocation()
            Panel1.Visible = True
            OPR11.Visible = True
        Else
            Panel1.Visible = False
            OPR11.Visible = False
        End If
    End Sub

    Private Sub UltraButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton9.Click
        Dim x As Integer
        Dim Sql As String
        Dim con = New SqlConnection()
        con = DBEngin.GetConnection(True)
        Dim M01 As DataSet
        Dim vcWhere As String
        x = 0
        Try
            For Each uRow As UltraGridRow In UltraGrid1.Rows
                vcWhere = "M01Sales_Order=" & cboSO.Text & " and M01Line_Item='" & Trim(UltraGrid1.Rows(x).Cells(1).Text) & "' and M01Quality_No='" & cboAllocate_Qulity.Text & "'"
                M01 = DBEngin.ExecuteDataset(con, Nothing, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "LSTD"), New SqlParameter("@vcWhereClause1", vcWhere))
                If isValidDataset(M01) Then
                    UltraGrid1.Rows(x).Cells(0).Value = True
                Else
                    UltraGrid1.Rows(x).Cells(0).Value = False
                End If
                x = x + 1
            Next

            Panel1.Visible = False
            OPR11.Visible = False
            cboAllocate_Qulity.Text = ""

            DBEngin.CloseConnection(con)
            con.ConnectionString = ""
            con.close()

        Catch returnMessage As ExecutionEngineException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(con)
                con.ConnectionString = ""
                con.close()
            End If
        End Try
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click
        Dim X As Integer
        Dim I As Integer
        Dim vcWhere As String
        Dim nvcFieldList1 As String

        Dim connection As SqlClient.SqlConnection
        Dim transaction As SqlClient.SqlTransaction
        Dim transactionCreated As Boolean
        Dim connectionCreated As Boolean

        connection = DBEngin.GetConnection(True)
        connectionCreated = True
        transaction = connection.BeginTransaction()
        transactionCreated = True
        Dim M01 As DataSet
        Dim M02 As DataSet
        Dim _ColumCount As Integer
        Dim _RowColunt As Integer
        Dim _Used As Double
        Dim _Code As Integer
        Dim ncQryType As String
        Dim Y As Integer
        Dim _Allocation As Double
        Dim T01 As DataSet
        Dim _Balance As Double
        Dim _LineItemQty As Double
        Dim _Loop1 As Boolean
        Dim _Loop2 As Boolean
        Dim _Loop3 As Boolean
        Dim _Date As Date

        Try
            If CDbl(lblQty.Text) >= CDbl(lblBalance.Text) Then

            Else
                MsgBox("Please chech the Allocation Qty again", MsgBoxStyle.Information, "Information ....")
            End If

            If chkAllocate.Checked = True Then
                _RowColunt = 0
                vcWhere = "select * from P01PARAMETER where P01CODE='PRN'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, vcWhere)
                If isValidDataset(M01) Then
                    _Code = M01.Tables(0).Rows(0)("P01NO")
                End If

                _Code = _Code - 1
                _Loop1 = False
                _Loop2 = False
                _Loop3 = False
                X = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                   
                    I = 0
                    _LineItemQty = 0
                    _ColumCount = 4
                    _Used = 0
                    _Loop1 = False
                    _Loop2 = False
                    _Loop3 = False
                    '  X = 0
                    _LineItemQty = Trim(UltraGrid1.Rows(X).Cells(3).Text)
                    If cboStealing.Text <> "" Then
                        If Microsoft.VisualBasic.Day(Today) > 10 Then
                            _Date = Today.AddDays(+30)
                            _Date = Month(_Date) & "/1/" & Year(_Date)
                            vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                        Else
                            _Date = Today
                            _Date = Month(_Date) & "/1/" & Year(_Date)
                            vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "'"
                        End If
                    Else
                        If Microsoft.VisualBasic.Day(Today) > 10 Then
                            _Date = Today.AddDays(+30)
                            _Date = Month(_Date) & "/1/" & Year(_Date)
                            vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                        Else
                            _Date = Today
                            _Date = Month(_Date) & "/1/" & Year(_Date)
                            vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "' "
                        End If
                    End If
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow3 As DataRow In M01.Tables(0).Rows
                        ' X = 0
                        If _Loop3 = True Then
                            _Loop2 = True
                            Exit For
                        End If
                        _RowColunt = 0
                        If I = 0 Then
                            _ColumCount = 4
                        Else
                            _ColumCount = _ColumCount + 3
                        End If
                        For Each uRow1 As UltraGridRow In UltraGrid3.Rows
                            _Used = 0
                            If Trim(UltraGrid3.Rows(_RowColunt).Cells(_ColumCount).Text) <> "" Then
                                _Used = CDbl(UltraGrid3.Rows(_RowColunt).Cells(_ColumCount).Value)
                                'Check Available Projection
                                'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION TMPPROJECTION TABLE
                                'DEVELOPED BY SURANGA WIJESINGHE ON 01/04/2016
                                _Allocation = 0
                                vcWhere = "tmpCode='" & Trim(UltraGrid3.Rows(_RowColunt).Cells(0).Text) & "' and tmpShade='" & Trim(UltraGrid3.Rows(_RowColunt).Cells(1).Text) & "' and tmpQuality='" & txtQuality.Text & "' and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & " and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                                T01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP1"), New SqlParameter("@vcWhereClause1", vcWhere))
                                If isValidDataset(T01) Then
                                    _Allocation = T01.Tables(0).Rows(0)("tmpQty")
                                End If
                                '===========================================================END STATMENT
                                'CHECK AVAILABLE ALLOCATED QTY FOR PROJECTION T15Projection
                                'DEVELOPED BY SURANGA WIJESINGHE ON 01/04/2016

                                vcWhere = "T15Code='" & Trim(UltraGrid3.Rows(_RowColunt).Cells(0).Text) & "' and T15Shade='" & Trim(UltraGrid3.Rows(_RowColunt).Cells(1).Text) & "' and T15Quality='" & txtQuality.Text & "' and T15Year=" & M01.Tables(0).Rows(I)("M43Year") & " and T15Month=" & M01.Tables(0).Rows(I)("M43Product_Month") & ""
                                T01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMP2"), New SqlParameter("@vcWhereClause1", vcWhere))
                                If isValidDataset(T01) Then
                                    _Allocation = _Allocation + T01.Tables(0).Rows(0)("T15Qty")
                                End If

                                _Balance = Trim(UltraGrid3.Rows(_RowColunt).Cells(_ColumCount - 2).Text) - _Allocation
                                _Used = _Used - _Allocation
                                If _Balance >= _Used Then
                                    If _Used >= _LineItemQty Then
                                        If Trim(UltraGrid1.Rows(X).Cells(0).Text) = True Then

                                          

                                            vcWhere = "tmpSales_Order='" & cboSO.Text & "' and tmpLine_Item=" & Trim(UltraGrid1.Rows(X).Cells(1).Text) & " and tmpCode='" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "' and tmpShade='" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "' and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & ""
                                            M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                                            If isValidDataset(M02) Then
                                                nvcFieldList1 = "update tmpProjection_Allocation set tmpQty='" & _LineItemQty & "' where  tmpSales_Order='" & cboSO.Text & " and tmpLine_Item='" & txtLine_Item.Text & "' and tmpCode='" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "' and tmpShade='" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "' and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & ""
                                                ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                                            Else

                                                vcWhere = "tmpSales_Order='" & cboSO.Text & "'  and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & ""
                                                T01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                                                If isValidDataset(T01) Then

                                                Else
                                                    'Block This Code Using Suranga on 2016.3.30
                                                    'vcWhere = "tmpSales_Order='" & cboSO.Text & "' and tmpline_item=" & Trim(UltraGrid1.Rows(X).Cells(1).Text) & " "
                                                    'T01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                                                    'If isValidDataset(T01) Then
                                                    '    MsgBox("You can't allocate different sales month projection", MsgBoxStyle.Information, "Information ......")
                                                    '    DBEngin.CloseConnection(connection)
                                                    '    connection.ConnectionString = ""
                                                    '    connection.Close()
                                                    '    Exit Sub
                                                    'End If

                                                    End If

                                                    ncQryType = "ADD1"
                                                    vcWhere = ""
                                                    If cboStealing.Text <> "" Then
                                                        If _Used > 0 Then
                                                            nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpCode," & "tmpShade," & "tmpMonth," & "tmpYear," & "tmpQuality," & "tmpQty," & "tmpUser) " & "values('" & cboSO.Text & "'," & Trim(UltraGrid1.Rows(X).Cells(1).Text) & ",'" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "','" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "'," & M01.Tables(0).Rows(I)("M43Product_Month") & "," & M01.Tables(0).Rows(I)("M43Year") & ",'" & cboStealing.Text & "','" & _LineItemQty & "','" & strDisname & "')"
                                                            up_GetSettmp_ProjectAllocation(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                                            _Loop3 = True
                                                            Exit For
                                                        End If
                                                    Else
                                                        If _Used > 0 Then
                                                            nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpCode," & "tmpShade," & "tmpMonth," & "tmpYear," & "tmpQuality," & "tmpQty," & "tmpUser) " & "values('" & cboSO.Text & "'," & Trim(UltraGrid1.Rows(X).Cells(1).Text) & ",'" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "','" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "'," & M01.Tables(0).Rows(I)("M43Product_Month") & "," & M01.Tables(0).Rows(I)("M43Year") & ",'" & txtQuality.Text & "','" & _LineItemQty & "','" & strDisname & "')"
                                                            up_GetSettmp_ProjectAllocation(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                                            _Loop3 = True
                                                            Exit For
                                                        End If
                                                    End If
                                            End If
                                        End If
                                    Else


                                        vcWhere = "tmpSales_Order='" & cboSO.Text & "' and tmpLine_Item=" & Trim(UltraGrid1.Rows(X).Cells(1).Text) & " and tmpCode='" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "' and tmpShade='" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "' and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & ""
                                        M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                                        If isValidDataset(M02) Then
                                            nvcFieldList1 = "update tmpProjection_Allocation set tmpQty='" & _Balance & "' where  tmpSales_Order='" & cboSO.Text & " and tmpLine_Item='" & txtLine_Item.Text & "' and tmpCode='" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "' and tmpShade='" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "' and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & ""
                                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                                        Else
                                            vcWhere = "tmpSales_Order='" & cboSO.Text & "'  and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & ""
                                            T01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                                            If isValidDataset(T01) Then

                                            Else
                                                'vcWhere = "tmpSales_Order='" & cboSO.Text & "'  "
                                                'T01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                                                'If isValidDataset(T01) Then
                                                '    MsgBox("You can't allocate different sales month projection", MsgBoxStyle.Information, "Information ......")
                                                '    DBEngin.CloseConnection(connection)
                                                '    connection.ConnectionString = ""
                                                '    connection.Close()
                                                '    Exit Sub
                                                'End If
                                                End If

                                                ncQryType = "ADD1"
                                                vcWhere = ""
                                                If cboStealing.Text <> "" Then
                                                    If Trim(UltraGrid1.Rows(X).Cells(0).Text) = True Then
                                                        If _Used > 0 Then
                                                            nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpCode," & "tmpShade," & "tmpMonth," & "tmpYear," & "tmpQuality," & "tmpQty," & "tmpUser) " & "values('" & cboSO.Text & "'," & Trim(UltraGrid1.Rows(X).Cells(1).Text) & ",'" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "','" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "'," & M01.Tables(0).Rows(I)("M43Product_Month") & "," & M01.Tables(0).Rows(I)("M43Year") & ",'" & cboStealing.Text & "','" & _Used & "','" & strDisname & "')"
                                                            up_GetSettmp_ProjectAllocation(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                                            _LineItemQty = _LineItemQty - _Used
                                                        End If
                                                    End If
                                                Else
                                                    If Trim(UltraGrid1.Rows(X).Cells(0).Text) = True Then
                                                        If _Used > 0 Then
                                                            nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpCode," & "tmpShade," & "tmpMonth," & "tmpYear," & "tmpQuality," & "tmpQty," & "tmpUser) " & "values('" & cboSO.Text & "'," & Trim(UltraGrid1.Rows(X).Cells(1).Text) & ",'" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "','" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "'," & M01.Tables(0).Rows(I)("M43Product_Month") & "," & M01.Tables(0).Rows(I)("M43Year") & ",'" & txtQuality.Text & "','" & _Used & "','" & strDisname & "')"
                                                            up_GetSettmp_ProjectAllocation(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                                            _LineItemQty = _LineItemQty - _Used
                                                        End If
                                                    End If
                                                End If
                                        End If
                                    End If
                                Else
                                    MsgBox("Allocated Qty is grater than balance quantity", MsgBoxStyle.Information, "Information ...")
                                    DBEngin.CloseConnection(connection)
                                    connection.ConnectionString = ""
                                    connection.Close()
                                    Exit Sub
                                End If
                            End If
                            _RowColunt = _RowColunt + 1
                        Next
                        I = I + 1
                    Next
                    Y = 0
                    For Y = 0 To 19
                        With UltraGrid1

                            vcWhere = "tmpSales_Order='" & cboSO.Text & "' and tmpLine_Item=" & Trim(UltraGrid1.Rows(X).Cells(1).Text) & " "
                            M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(M02) Then
                                .Rows(X).Cells(Y).Appearance.BackColor = Color.Green
                            End If
                        End With
                    Next
                    UltraGrid1.Rows(X).Cells(0).Value = False

                    X = X + 1
                Next

            Else
                _RowColunt = 0
                vcWhere = "select * from P01PARAMETER where P01CODE='PRN'"
                M01 = DBEngin.ExecuteDataset(connection, transaction, vcWhere)
                If isValidDataset(M01) Then
                    _Code = M01.Tables(0).Rows(0)("P01NO")
                End If

                _Code = _Code - 1
                For Each uRow As UltraGridRow In UltraGrid3.Rows
                    'X = (UltraGrid3.DisplayLayout.Bands(0).Columns.Count)
                    'X = X - 2
                    _ColumCount = 4
                    _Used = 0
                    X = 0
                    If cboStealing.Text <> "" Then
                        If Microsoft.VisualBasic.Day(Today) > 10 Then
                            _Date = Today.AddDays(+30)
                            _Date = Month(_Date) & "/1/" & Year(_Date)
                            vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                        Else
                            _Date = Today
                            _Date = Month(_Date) & "/1/" & Year(_Date)
                            vcWhere = "M43Quality='" & Trim(cboStealing.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "'"
                        End If
                    Else
                        If Microsoft.VisualBasic.Day(Today) > 10 Then
                            _Date = Today.AddDays(+30)
                            _Date = Month(_Date) & "/1/" & Year(_Date)
                            vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "' and convert(datetime,Ddate,111)>='" & _Date & "'  and M43Count_No=" & _Code & ""
                        Else
                            _Date = Today
                            _Date = Month(_Date) & "/1/" & Year(_Date)
                            vcWhere = "M43Quality='" & Trim(txtQuality.Text) & "'  and M43Count_No=" & _Code & " and convert(datetime,Ddate,111)>='" & _Date & "' "
                        End If
                    End If
                    M01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "PRS2"), New SqlParameter("@vcWhereClause1", vcWhere))
                    For Each DTRow3 As DataRow In M01.Tables(0).Rows

                        _Used = 0
                        If Trim(UltraGrid3.Rows(_RowColunt).Cells(_ColumCount).Text) <> "" Then
                            _Used = CDbl(UltraGrid3.Rows(_RowColunt).Cells(_ColumCount).Value)
                        End If

                        vcWhere = "tmpSales_Order='" & cboSO.Text & "' and tmpLine_Item=" & txtLine_Item.Text & " and tmpCode='" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "' and tmpShade='" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "' and tmpMonth=" & M01.Tables(0).Rows(X)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(X)("M43Year") & ""
                        M02 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                        If isValidDataset(M02) Then
                            nvcFieldList1 = "update tmpProjection_Allocation set tmpQty='" & _Used & "' where tmpSales_Order='" & cboSO.Text & " and tmpLine_Item='" & txtLine_Item.Text & "' and tmpCode='" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "' and tmpShade='" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "' and tmpMonth=" & M01.Tables(0).Rows(X)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(X)("M43Year") & ""
                            ExecuteNonQueryText(connection, transaction, nvcFieldList1)
                        Else
                            vcWhere = "tmpSales_Order='" & cboSO.Text & "'  "
                            T01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                            If isValidDataset(T01) Then
                                vcWhere = "tmpSales_Order='" & cboSO.Text & "'  and tmpMonth=" & M01.Tables(0).Rows(I)("M43Product_Month") & " and tmpYear=" & M01.Tables(0).Rows(I)("M43Year") & ""
                                T01 = DBEngin.ExecuteDataset(connection, transaction, "up_GetSetDelivary_Quatation", New SqlParameter("@cQryType", "TMPA"), New SqlParameter("@vcWhereClause1", vcWhere))
                                If isValidDataset(T01) Then

                                Else
                                    'MsgBox("You can't allocate different sales month projection", MsgBoxStyle.Information, "Information ......")
                                    'DBEngin.CloseConnection(connection)
                                    'connection.ConnectionString = ""
                                    'connection.Close()
                                    'Exit Sub
                                End If
                            End If

                            ncQryType = "ADD1"
                            vcWhere = ""
                            If cboStealing.Text <> "" Then
                                If _Used > 0 Then
                                    nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpCode," & "tmpShade," & "tmpMonth," & "tmpYear," & "tmpQuality," & "tmpQty," & "tmpUser) " & "values('" & cboSO.Text & "'," & txtLine_Item.Text & ",'" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "','" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "'," & M01.Tables(0).Rows(X)("M43Product_Month") & "," & M01.Tables(0).Rows(X)("M43Year") & ",'" & cboStealing.Text & "','" & _Used & "','" & strDisname & "')"
                                    up_GetSettmp_ProjectAllocation(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                End If
                            Else
                                If _Used > 0 Then
                                    nvcFieldList1 = "(tmpSales_Order," & "tmpLine_Item," & "tmpCode," & "tmpShade," & "tmpMonth," & "tmpYear," & "tmpQuality," & "tmpQty," & "tmpUser) " & "values('" & cboSO.Text & "'," & txtLine_Item.Text & ",'" & UltraGrid3.Rows(_RowColunt).Cells(0).Value & "','" & UltraGrid3.Rows(_RowColunt).Cells(1).Value & "'," & M01.Tables(0).Rows(X)("M43Product_Month") & "," & M01.Tables(0).Rows(X)("M43Year") & ",'" & txtQuality.Text & "','" & _Used & "','" & strDisname & "')"
                                    up_GetSettmp_ProjectAllocation(ncQryType, nvcFieldList1, vcWhere, connection, transaction)
                                End If
                            End If
                            End If

                            X = X + 1
                            _ColumCount = _ColumCount + 3
                    Next

                    _RowColunt = _RowColunt + 1
                Next
                X = 0
                For Each uRow As UltraGridRow In UltraGrid1.Rows
                    If Trim(UltraGrid1.Rows(X).Cells(1).Value) = txtLine_Item.Text Then
                        I = 0
                        For I = 0 To 19
                            With UltraGrid1
                                .Rows(X).Cells(I).Appearance.BackColor = Color.Green
                            End With
                        Next
                        UltraGrid1.Rows(X).Cells(0).Value = False
                    End If
                    X = X + 1
                Next
            End If
            nvcFieldList1 = "delete from tmp_Block_Quality_Projection where tmpQuality='" & txtQuality.Text & "' and tmpUser='" & strDisname & "'"
            ExecuteNonQueryText(connection, transaction, nvcFieldList1)

            transaction.Commit()

            'UltraGrid3.DisplayLayout.Bands(0).Groups.Remove(0)
            UltraGrid3.DisplayLayout.Bands(0).Groups.Clear()
            UltraGrid3.DisplayLayout.Bands(0).Columns.Dispose()
            UltraGroupBox1.Visible = False

            Dim TipInfo As New UltraToolTipInfo()
            TipInfo.ToolTipText = ""
            Me.UltraToolTipManager1.SetUltraToolTip(Me.UltraGrid1, TipInfo)
            Me.UltraToolTipManager1.DisplayStyle = Infragistics.Win.ToolTipDisplayStyle.BalloonTip
            cboStealing.Text = ""
            chkAllocate.Checked = False
            DBEngin.CloseConnection(connection)
            connection.ConnectionString = ""
            connection.Close()
           
        Catch returnMessage As ExecutionEngineException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
                DBEngin.CloseConnection(connection)
                connection.ConnectionString = ""
                connection.Close()
            End If
        End Try
    End Sub

  
    Function Calculate_Balance()
        Dim Value As Double
        Dim Value1 As Double
        Dim Sql As String
        Dim _Rowcount As Integer
        Dim _ColumCount As Integer
        Dim x As Integer
        Dim vcWhere As String
        Dim _Code As Integer
        Dim i As Integer
        Dim _Balance As Double
        Dim _Projection As Double
        Dim _Alocated As Double
        Dim _Used As Double
        Dim _String As String

        _Balance = 0
        x = 0
        _Rowcount = 0
        Value = CDbl(lblQty.Text)
        _Projection = 0
        _Alocated = 0
        _Used = 0
        Try
            For Each uRow As UltraGridRow In UltraGrid3.Rows
                x = (UltraGrid3.DisplayLayout.Bands(0).Columns.Count)
                x = x - 2
                _ColumCount = 4
                _Projection = 0
                _Alocated = 0
                _Used = 0

                For i = 1 To x / 3

                    _Projection = 0
                    _Alocated = 0
                    _Used = 0

                    If Trim(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount - 1).Text) <> "" Then
                        _Used = CDbl(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount - 1).Value)
                    End If

                    If Trim(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount - 2).Text) <> "" Then
                        _Projection = CDbl(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount - 2).Value)
                    End If

                    If Trim(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount).Text) <> "" Then
                        If IsNumeric(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount).Value) Then
                            _Alocated = CDbl(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount).Value)
                            'Value1 = CDbl(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount).Value)
                            '_String = (Value1.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
                            '_String = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", Value1))
                            'UltraGrid3.Rows(_Rowcount).Cells(_ColumCount).Text = _String
                        Else
                            Dim windowInfo As New UltraDesktopAlertShowWindowInfo()
                            Dim strFileName As String
                            windowInfo.Caption = "Please enter the correct value"
                            windowInfo.FooterText = "Technova"
                            strFileName = ConfigurationManager.AppSettings("SoundPath") + "\REMINDER.wav"
                            windowInfo.Sound = strFileName
                            UltraDesktopAlert1.Show(windowInfo)
                            'windowInfo.DesktopAlert.Show()
                            UltraGrid3.Rows(_Rowcount).Cells(_ColumCount).Value = ""

                        End If
                    End If

                    If _Projection + _Used >= _Alocated Then
                        If IsNumeric(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount).Value) Then
                            _Balance = _Balance + CDbl(UltraGrid3.Rows(_Rowcount).Cells(_ColumCount).Value)
                        Else


                        End If
                    Else
                        Exit Function
                    End If
                    _ColumCount = _ColumCount + 3
                Next
                _Rowcount = _Rowcount + 1
            Next
            If CDbl(lblQty.Text) >= _Balance Then

            Else
                MsgBox("Please chech the Allocation Qty again", MsgBoxStyle.Information, "Information ....")
                Exit Function
            End If

            _Balance = CDbl(lblQty.Text) - _Balance
            lblBalance.Text = (_Balance.ToString("0,0.00", System.Globalization.CultureInfo.InvariantCulture))
            lblBalance.Text = (String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0,0.00}", _Balance))

          
        Catch returnMessage As ExecutionEngineException
            If returnMessage.Message <> Nothing Then
                MessageBox.Show(returnMessage.Message)
              
            End If
        End Try
    End Function

    Private Sub UltraGrid3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles UltraGrid3.AfterCellUpdate
        Call Calculate_Balance()
    End Sub

    Private Sub UltraGrid3_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid3.InitializeLayout
        'MsgBox("")
    End Sub

    Private Sub cboStealing_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboStealing.InitializeLayout

    End Sub

    Private Sub cboStealing_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboStealing.TextChanged
      
    End Sub

    Private Sub UltraButton8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton8.Click
        If cboStealing.Text <> "" Then
            UltraGrid3.DisplayLayout.Bands(0).Groups.Clear()
            UltraGrid3.DisplayLayout.Bands(0).Columns.Dispose()

            Call Projection_AllocationCapacity_stealing()
        End If
    End Sub

    Private Sub chkNPL1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNPL1.CheckedChanged
        If chkNPL1.Checked = True Then
            chkNPL2.Checked = False
        End If
    End Sub

    Private Sub chkNPL2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNPL2.CheckedChanged
        If chkNPL2.Checked = True Then
            chkNPL1.Checked = False
        End If
    End Sub

    Private Sub cboSO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cboSO.InitializeLayout

    End Sub
End Class